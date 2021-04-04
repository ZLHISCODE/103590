VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmChargeItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "收费项目设置"
   ClientHeight    =   7650
   ClientLeft      =   1155
   ClientTop       =   2520
   ClientWidth     =   7260
   Icon            =   "frmChargeItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   5850
      Index           =   2
      Left            =   60
      TabIndex        =   78
      Top             =   960
      Visible         =   0   'False
      Width           =   6840
      Begin VB.PictureBox pic价格等级 
         BorderStyle     =   0  'None
         Height          =   4545
         Index           =   0
         Left            =   240
         ScaleHeight     =   4545
         ScaleWidth      =   6345
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   660
         Width           =   6345
         Begin ZL9BillEdit.BillEdit msh价目 
            Height          =   3420
            Index           =   0
            Left            =   0
            TabIndex        =   115
            Top             =   100
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   6033
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
         Begin VB.TextBox txt调价说明 
            Height          =   300
            Index           =   0
            Left            =   1155
            MaxLength       =   100
            TabIndex        =   84
            Top             =   3960
            Width           =   5070
         End
         Begin VB.CheckBox chkNow 
            Caption         =   "立即执行(&N)"
            Height          =   225
            Index           =   0
            Left            =   4170
            TabIndex        =   82
            Top             =   3600
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   285
            Index           =   0
            Left            =   1170
            TabIndex        =   81
            Top             =   3570
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   111214595
            CurrentDate     =   36444
            MaxDate         =   401768
         End
         Begin VB.Label lbl调价执行时间 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行日期(&B)"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   80
            Top             =   3615
            Width           =   1050
         End
         Begin VB.Label lbl调价说明 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "调价说明(&X)"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   83
            Top             =   4020
            Width           =   990
         End
      End
      Begin XtremeSuiteControls.TabControl tbPriceGrade 
         Height          =   4800
         Left            =   240
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   870
         Width           =   6375
         _Version        =   589884
         _ExtentX        =   11245
         _ExtentY        =   8467
         _StockProps     =   64
      End
      Begin VB.Image img价目 
         Height          =   600
         Left            =   270
         Picture         =   "frmChargeItem.frx":000C
         Stretch         =   -1  'True
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblEdit 
         Caption         =   "    此处设置收费项目的价格，当它是变价时，只能选择一个收入项目。"
         Height          =   435
         Index           =   14
         Left            =   1170
         TabIndex        =   79
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6555
      Index           =   4
      Left            =   1455
      TabIndex        =   87
      Top             =   6165
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5040
         TabIndex        =   107
         Top             =   3405
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   106
         Top             =   3405
         Width           =   975
      End
      Begin VB.OptionButton opt使用科室 
         Caption         =   "全院"
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   99
         Top             =   3480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opt使用科室 
         Caption         =   "指定科室"
         Height          =   180
         Index           =   0
         Left            =   2040
         TabIndex        =   98
         Top             =   3480
         Width           =   1095
      End
      Begin ZL9BillEdit.BillEdit msh从属 
         Height          =   2175
         Left            =   240
         TabIndex        =   75
         Top             =   840
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   3836
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
      Begin MSComctlLib.ListView Lvw科室 
         Height          =   1980
         Left            =   240
         TabIndex        =   96
         Top             =   3840
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   3493
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "当前主项的使用范围"
         Height          =   180
         Left            =   240
         TabIndex        =   97
         Top             =   3480
         Width           =   1620
      End
      Begin VB.Label lbl从属合计 
         Alignment       =   1  'Right Justify
         Caption         =   "合计:##.##"
         Height          =   180
         Left            =   4680
         TabIndex        =   90
         Top             =   3060
         Width           =   1695
      End
      Begin VB.Image img从属 
         Height          =   600
         Left            =   270
         Picture         =   "frmChargeItem.frx":0246
         Stretch         =   -1  'True
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblEdit 
         Caption         =   "    从属项目是指用户在进行单据录入中，会随着主收费项目的增加而自动增加的收费项目。"
         Height          =   435
         Index           =   13
         Left            =   1140
         TabIndex        =   88
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6660
      Index           =   3
      Left            =   0
      TabIndex        =   85
      Top             =   450
      Visible         =   0   'False
      Width           =   6870
      Begin VB.Frame Frame1 
         Height          =   4785
         Left            =   150
         TabIndex        =   91
         Top             =   0
         Width           =   6585
         Begin VB.Frame Frame2 
            Height          =   120
            Left            =   195
            TabIndex        =   92
            Top             =   660
            Width           =   6135
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "开单人所在科室(&F)"
            Height          =   195
            Index           =   6
            Left            =   4380
            TabIndex        =   95
            Top             =   450
            Width           =   1860
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "院外执行(&E)"
            Height          =   195
            Index           =   5
            Left            =   4395
            TabIndex        =   94
            Top             =   825
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.TextBox txt门诊执行 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1065
            Width           =   1785
         End
         Begin VB.TextBox txt住院执行 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4065
            MaxLength       =   30
            TabIndex        =   9
            Top             =   1065
            Width           =   1905
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "无明确执行科室(&N)"
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   1
            Top             =   210
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "病人所在病区(&B)"
            Height          =   195
            Index           =   2
            Left            =   4380
            TabIndex        =   3
            Top             =   195
            Width           =   1755
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "操作员所在科室(&R)"
            Height          =   195
            Index           =   3
            Left            =   2265
            TabIndex        =   5
            Top             =   450
            Width           =   1920
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "病人所在科室(&K)"
            Height          =   180
            Index           =   1
            Left            =   2280
            TabIndex        =   2
            Top             =   210
            Width           =   1725
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "指定科室(&D)"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   4
            Top             =   450
            Width           =   2265
         End
         Begin ZL9BillEdit.BillEdit msf定向执行 
            Height          =   3000
            Left            =   405
            TabIndex        =   11
            Top             =   1680
            Width           =   5940
            _ExtentX        =   10478
            _ExtentY        =   5292
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
         Begin MSComctlLib.ImageList imgList 
            Left            =   -210
            Top             =   2640
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
                  Picture         =   "frmChargeItem.frx":0688
                  Key             =   "close"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmChargeItem.frx":0C22
                  Key             =   "expend"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmChargeItem.frx":11BC
                  Key             =   "Dept"
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl门诊执行 
            AutoSize        =   -1  'True
            Caption         =   "门诊(&O)"
            Height          =   180
            Left            =   645
            TabIndex        =   6
            Top             =   1125
            Width           =   630
         End
         Begin VB.Label lbl定向执行 
            AutoSize        =   -1  'True
            Caption         =   "2、指定病人科室："
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   225
            TabIndex        =   10
            Top             =   1455
            Width           =   1530
         End
         Begin VB.Label lbl一般情况 
            AutoSize        =   -1  'True
            Caption         =   "1、除指定病人科室外："
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   195
            TabIndex        =   93
            Top             =   855
            Width           =   1890
         End
         Begin VB.Label lbl住院执行 
            AutoSize        =   -1  'True
            Caption         =   "住院(&I)"
            Height          =   180
            Left            =   3405
            TabIndex        =   8
            Top             =   1125
            Width           =   630
         End
      End
      Begin VB.PictureBox picDept 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   3600
         ScaleHeight     =   2655
         ScaleWidth      =   3000
         TabIndex        =   101
         Top             =   1920
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CheckBox ChkSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "全选"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2115
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   88
            Width           =   675
         End
         Begin VB.ComboBox cboProperty 
            Height          =   300
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   50
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwItems 
            Height          =   2040
            Left            =   50
            TabIndex        =   104
            Top             =   380
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   3598
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "imgList"
            SmallIcons      =   "imgList"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label lbl工作性质 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "工作性质"
            Height          =   180
            Left            =   50
            TabIndex        =   105
            Top             =   110
            Width           =   720
         End
      End
      Begin VB.Frame fra批量 
         Caption         =   "应用范围"
         Height          =   1530
         Left            =   150
         TabIndex        =   86
         Top             =   4920
         Visible         =   0   'False
         Width           =   6585
         Begin VB.OptionButton optApply 
            Caption         =   "应用于该分类下所有项目(&L)"
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   14
            Top             =   885
            Width           =   6270
         End
         Begin VB.OptionButton optApply 
            Caption         =   "应用于该类别下所有项目(&U)"
            Height          =   225
            Index           =   3
            Left            =   210
            TabIndex        =   15
            Top             =   1215
            Width           =   6315
         End
         Begin VB.OptionButton optApply 
            Caption         =   "应用于同级的所有项目(&G)"
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   13
            Top             =   555
            Width           =   6285
         End
         Begin VB.OptionButton optApply 
            Caption         =   "仅对本项目起作用(&W)"
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   12
            Top             =   225
            Value           =   -1  'True
            Width           =   6240
         End
      End
   End
   Begin VB.CheckBox chk保留 
      Caption         =   "新建下一项时保存部分信息"
      Height          =   255
      Left            =   4665
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   "连续新增时是否清除所有内容"
      Top             =   75
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   345
      TabIndex        =   71
      Tag             =   "分类"
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   69
      Tag             =   "分类"
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   70
      Tag             =   "分类"
      Top             =   7200
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItem.frx":1756
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItem.frx":1A70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6405
      Index           =   1
      Left            =   210
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   6795
      Begin VB.PictureBox picTwo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   155
         ScaleHeight     =   1035
         ScaleWidth      =   6735
         TabIndex        =   59
         Top             =   5400
         Width           =   6735
         Begin VB.CommandButton cmd病案 
            Caption         =   "…"
            Height          =   240
            Left            =   6240
            TabIndex        =   108
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   750
            Width           =   255
         End
         Begin VB.ComboBox cbo录入限量范围 
            Height          =   300
            Left            =   4350
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   0
            Width           =   2205
         End
         Begin VB.TextBox txt录入限量 
            Height          =   300
            Left            =   1065
            MaxLength       =   13
            TabIndex        =   61
            Top             =   15
            Width           =   1170
         End
         Begin VB.ComboBox cmb费用确认 
            Height          =   300
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   360
            Width           =   1875
         End
         Begin VB.CheckBox chk费用确认范围 
            Caption         =   "费用确认应用于当前分类所有项目"
            Height          =   255
            Left            =   3120
            TabIndex        =   66
            Top             =   360
            Width           =   3495
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   690
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   15
            Left            =   4365
            MaxLength       =   40
            TabIndex        =   109
            ToolTipText     =   "按*打开选择器"
            Top             =   720
            Width           =   2205
         End
         Begin VB.Label lbl病案费目 
            Caption         =   "病案费目(&F)"
            Height          =   255
            Left            =   3360
            TabIndex        =   110
            Top             =   750
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "录入限量应用于"
            Height          =   180
            Left            =   3015
            TabIndex        =   62
            Top             =   75
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "录入限量(&P)"
            Height          =   180
            Left            =   0
            TabIndex        =   60
            Top             =   75
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "费用确认(&Q)"
            Height          =   180
            Left            =   0
            TabIndex        =   64
            Top             =   425
            Width           =   990
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "院区(&Z)"
            Height          =   180
            Left            =   345
            TabIndex        =   67
            Top             =   780
            Visible         =   0   'False
            Width           =   630
         End
      End
      Begin VB.PictureBox picOne 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   160
         ScaleHeight     =   1455
         ScaleWidth      =   6615
         TabIndex        =   44
         Top             =   3480
         Width           =   6615
         Begin VB.ComboBox cmb项目特性 
            Height          =   300
            Left            =   4335
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   0
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   1050
            MaxLength       =   100
            TabIndex        =   56
            Top             =   1125
            Width           =   5505
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1050
            MaxLength       =   72
            TabIndex        =   46
            Top             =   0
            Width           =   1770
         End
         Begin VB.ComboBox cmb服务对象 
            Height          =   300
            Left            =   4335
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   375
            Width           =   2205
         End
         Begin VB.ComboBox cmb计算单位 
            Height          =   300
            Left            =   1050
            TabIndex        =   48
            Top             =   375
            Width           =   1755
         End
         Begin VB.ComboBox cmb费用类型 
            Height          =   300
            Left            =   4335
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   750
            Width           =   2205
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   52
            Tag             =   "分类"
            Top             =   750
            Width           =   1755
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "项目特性(&B)"
            Height          =   180
            Index           =   10
            Left            =   3285
            TabIndex        =   112
            Top             =   60
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "建档日期(&B)"
            Height          =   180
            Index           =   12
            Left            =   0
            TabIndex        =   51
            Tag             =   "分类"
            Top             =   810
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "说明(&X)"
            ForeColor       =   &H80000007&
            Height          =   180
            Index           =   8
            Left            =   360
            TabIndex        =   55
            Top             =   1185
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "规格(&R)"
            ForeColor       =   &H80000007&
            Height          =   180
            Index           =   4
            Left            =   350
            TabIndex        =   45
            Top             =   60
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "计算单位(&L)"
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   47
            Top             =   435
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "服务对象(&J)"
            Height          =   180
            Index           =   6
            Left            =   3285
            TabIndex        =   49
            Top             =   435
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "费用类型(&F)"
            Height          =   180
            Index           =   7
            Left            =   3290
            TabIndex        =   53
            Top             =   810
            Width           =   990
         End
      End
      Begin VB.CheckBox chk自动计算 
         Caption         =   "不进行自动计算(&A)"
         Height          =   210
         Left            =   4965
         TabIndex        =   38
         ToolTipText     =   "在录入费用记录时，对该项目补充摘要"
         Top             =   2400
         Width           =   1890
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   13
         Left            =   1210
         TabIndex        =   41
         Top             =   3105
         Width           =   1785
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   14
         Left            =   4485
         TabIndex        =   43
         Top             =   3105
         Width           =   2205
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   12
         Left            =   4695
         MaxLength       =   40
         TabIndex        =   28
         Top             =   885
         Width           =   2025
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   11
         Left            =   1210
         MaxLength       =   100
         TabIndex        =   58
         Top             =   4995
         Width           =   5505
      End
      Begin VB.ComboBox cmbClass 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   150
         Width           =   1635
      End
      Begin MSComctlLib.ListView lvwSel 
         Height          =   1635
         Left            =   825
         TabIndex        =   30
         Top             =   -1500
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2884
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   3570
         MaxLength       =   40
         TabIndex        =   24
         Top             =   525
         Width           =   1605
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   10
         Left            =   6240
         MaxLength       =   40
         TabIndex        =   26
         Top             =   525
         Width           =   465
      End
      Begin VB.CheckBox chk急诊 
         Caption         =   "急诊(&Z)"
         Height          =   210
         Left            =   4965
         TabIndex        =   100
         Top             =   2745
         Width           =   1305
      End
      Begin VB.ComboBox cmb护理 
         Height          =   300
         ItemData        =   "frmChargeItem.frx":1D8A
         Left            =   4965
         List            =   "frmChargeItem.frx":1D8C
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   2700
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   33
         Tag             =   "分类"
         Top             =   1275
         Width           =   1605
      End
      Begin VB.CheckBox chk摘要 
         Caption         =   "补充摘要(&A)"
         Height          =   210
         Left            =   4965
         TabIndex        =   39
         ToolTipText     =   "在录入费用记录时，对该项目补充摘要"
         Top             =   2475
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   12
         TabIndex        =   31
         Tag             =   "分类"
         Top             =   1275
         Width           =   1620
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   73
         Tag             =   "分类"
         Text            =   "111111"
         Top             =   570
         Width           =   1485
      End
      Begin ZL9BillEdit.BillEdit mshAlias 
         Height          =   1335
         Left            =   180
         TabIndex        =   34
         Top             =   1665
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   2355
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
      Begin VB.CheckBox chk加班加价 
         Caption         =   "加班加价(&D)"
         Height          =   210
         Left            =   4965
         TabIndex        =   37
         Top             =   2190
         Width           =   1305
      End
      Begin VB.CheckBox chk屏蔽费别 
         Caption         =   "屏蔽费别(&M)"
         Height          =   240
         Left            =   4965
         TabIndex        =   36
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   40
         TabIndex        =   22
         Top             =   885
         Width           =   2790
      End
      Begin VB.CheckBox chk变价 
         Caption         =   "变价(&G)"
         Height          =   210
         Left            =   4965
         TabIndex        =   35
         Top             =   1665
         Width           =   945
      End
      Begin VB.CommandButton cmd上级 
         Caption         =   "…"
         Height          =   240
         Left            =   6435
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "分类"
         ToolTipText     =   "按*打开选择器"
         Top             =   180
         Width           =   255
      End
      Begin VB.TextBox txtTemp 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   840
         TabIndex        =   77
         TabStop         =   0   'False
         Tag             =   "分类"
         Text            =   "11"
         Top             =   525
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   3570
         MaxLength       =   40
         TabIndex        =   19
         ToolTipText     =   "按*打开选择器"
         Top             =   150
         Width           =   3150
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "最高限价(&M)"
         Height          =   180
         Index           =   20
         Left            =   150
         TabIndex        =   40
         Top             =   3180
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "最低限价(&N)"
         Height          =   180
         Index           =   21
         Left            =   3450
         TabIndex        =   42
         Top             =   3165
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "备选码(&B)"
         Height          =   180
         Index           =   19
         Left            =   3810
         TabIndex        =   27
         Top             =   945
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "产地(&T)"
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   18
         Left            =   520
         TabIndex        =   57
         Top             =   5055
         Width           =   630
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类别(&C)"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "标识主码(&P)"
         Height          =   180
         Index           =   17
         Left            =   2550
         TabIndex        =   23
         Top             =   585
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "标识子码(&I)"
         Height          =   180
         Index           =   16
         Left            =   5235
         TabIndex        =   25
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "五笔简码(&W)"
         Height          =   180
         Index           =   9
         Left            =   6840
         TabIndex        =   32
         Top             =   -90
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)                   (拼音)                   (五笔)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   29
         Top             =   1335
         Width           =   5130
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "上级(&V)"
         Height          =   180
         Index           =   3
         Left            =   2895
         TabIndex        =   18
         Tag             =   "分类"
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   72
         Tag             =   "分类"
         Top             =   585
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Tag             =   "分类"
         Top             =   945
         Width           =   630
      End
   End
   Begin MSComctlLib.TabStrip TabMain 
      Height          =   6990
      Left            =   120
      TabIndex        =   0
      Tag             =   "分类"
      Top             =   45
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   12330
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本信息"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "收费价目"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "从属项目"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "执行科室"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChargeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    text编码 = 0
    Text名称 = 1
    Text规格 = 2
    text简码 = 3
    Text说明 = 4
    Text分类 = 5
    Text建档时间 = 6
    text调价说明 = 7
    text五笔 = 8
    text标识主码 = 9
    Text标识子码 = 10
    text产地 = 11
    Text备选码 = 12
    text最高限价 = 13
    text最低限价 = 14
    text病案费目 = 15
End Enum

Private mlng编码长度 As Long
Private mlng单位长度 As Long
Private mlng别名长度 As Long
Private mlng简码长度 As Long

Private mstr类别 As String  '类别编码,只有一个字母
Private mstr分类编码 As String    '原始的上级编码的值
Private mstr编码 As String        '原始的本级编码的值
Private mdbl医价价格 As Double    '医价接口标准价格
Private mdbl最高限价 As Double
Private mdbl最低限价 As Double
Private mblnOk As Boolean
Private mlngFind As Long
Private mblnVerifyPris As Boolean   '审核调价单权限 true-有权限，false-无权限
Private mblnVerifyFlow As Boolean   '调价是否启用了审核流程，true-启用，false-未启用

'以下项目可能修改
Private mstr分类ID As String
Private mstrID As String
Private mint编码 As Integer       '修改前包括下级在内的编码最长的长度

Dim mcol价目() As Collection    '保存的是收费价目的ID，以收入项目的ID作Key。免得同一收入项目失去原有价目ID
Dim mblnNew() As Boolean  '新价格
Dim mblnChanged价目() As Boolean  '价目是否改变
Dim mlng末级 As Long    '末级
Dim medit方式 As EditMode   '0、新增；1、修改；2、调价；3、执行科室、4、从属项目、5、复制新增
Dim mblnChange As Boolean     '是否改变了
Dim mstr列表(1 To 4) As String '保存一些列表值3
Dim mblnCancel As Boolean
Dim mblnEditCancel As Boolean   '取消更新
Dim mstrSel  As String  '选择目标名称
Dim mblnShow收费价目 As Boolean '判断是否已经显示了收费价目页，用在医价系统中
Private mstrServerObj As String  '服务对象

'是否变价  通过控件chk变价判断
'加班加价  通过控件chk变价判断
Private strInputed As String

Private mblnIsSpecialItem As Boolean                '是否是特殊项目(特殊项目指的是：床位和护理类项目以及包含在"自动计价项目"中的其它自动计算项目(计算标志为6,7,8));或者是床位或护理项目的从属项目
Private mstrCurrentDateFormat As String             '当前使用的日期格式

Private mrs性质分类 As ADODB.Recordset
Private mrs部门 As ADODB.Recordset

Private mstr已选执行科室 As String
Private mblnRefresh As Boolean

'收费价目列表
Private Const mcstCol收费项目 As Integer = 0
Private Const mcstCol原价 As Integer = 1
Private Const mcstCol现价 As Integer = 2
Private Const mcstCol缺省价格 As Integer = 3
Private Const mcstCol附加手术收费率 As Integer = 4
Private Const mcstCol加班加价率 As Integer = 5
Private Const mcstCols As Integer = 6
Private mstrPrivs As String
Private mblnNotClick As Boolean
Private mblnCanUpdateAll As Boolean '是否允许操作所有项目：未启用价格等级或启用了价格等级有“所有院区”权限

Private Sub Ini性质分类()
    '取部门性质分类，如果已经提取了则退出
    On Error GoTo ErrHandle
    If Not mrs性质分类 Is Nothing Then
        mrs性质分类.Filter = ""
        If Not mrs性质分类.EOF Then
            Exit Sub
        End If
    End If
    
    gstrSQL = "Select 名称,服务病人 From 部门性质分类"
    Set mrs性质分类 = zlDatabase.OpenSQLRecord(gstrSQL, "取部门性质分类")
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load部门(ByVal intType As Integer, ByVal str工作性质 As String)
    'intType:0-执行科室（所有性质，服务于病人）；1-病人科室（临床性质）
    Dim rsData As ADODB.Recordset
    Dim ObjItem As ListItem
    
    On Error GoTo ErrHandle
    
    If intType = 1 Then
        gstrSQL = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and 工作性质=[1] " & _
                "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " order by 编码"
    Else
        gstrSQL = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and T.服务对象 in (1,2,3) " & _
                " and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
                
        If str工作性质 <> "所有性质" Then
            gstrSQL = gstrSQL & " and 工作性质=[1] "
        End If
                
        gstrSQL = gstrSQL & " order by 编码"
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str工作性质)
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsData.EOF
        
        If Me.lvwItems.Tag = "开单" Then
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & rsData!ID, rsData!名称)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsData!编码
            ObjItem.Checked = False
        
            If InStr(Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) & ",", rsData!ID & ",") > 0 Then
                ObjItem.Checked = True
            End If
        End If
        
        If Me.lvwItems.Tag = "执行" Then
            If InStr(mstr已选执行科室, rsData!ID & "," & "[" & rsData!编码 & "]" & rsData!名称) = 0 Then
                Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & rsData!ID, rsData!名称)
                ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsData!编码
                ObjItem.Checked = False
            End If
        End If
        
        rsData.MoveNext
    Loop
    rsData.Close
    
    '没有时退出
    If Me.lvwItems.ListItems.Count = 0 Then Exit Sub
    
    Me.lvwItems.ListItems(1).Selected = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Ini收费适用科室(ByVal str项目ID As String)
    Dim rsTmp As ADODB.Recordset
    Dim n As Integer
    
    '所有临床、医技科室和病区
'    gstrSQL = " Select Distinct 编码||'-'||名称 科室,ID From 部门表 " & _
'         " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('临床', '护理', '检查', '检验', '治疗', '手术') And 服务对象 IN(2,3))" & _
'         " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
'         " Order By 编码||'-'||名称 "
    'Oracle11g 会出现重复id，故改为如下SQL
    On Error GoTo ErrHandle
    gstrSQL = _
        "Select Distinct a.编码 || '-' || a.名称 科室, a.Id " & vbNewLine & _
        "From 部门表 A, 部门性质说明 B " & vbNewLine & _
        "Where a.Id = b.部门id And b.工作性质 In ('临床', '护理', '检查', '检验', '治疗', '手术') And b.服务对象 In (2, 3) And " & vbNewLine & _
        "      (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & vbNewLine & _
        "Order By 编码 || '-' || 名称 "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "所有临床、医技科室和病区")
    
    Lvw科室.ListItems.Clear
    With rsTmp
        Do While Not .EOF
            Lvw科室.ListItems.Add , "_" & !ID, !科室, 1, 1
            .MoveNext
        Loop
    End With
    
    If str项目ID = "" Then Exit Sub
    
    '收费适用科室
    gstrSQL = "Select 科室ID From 收费适用科室 Where 项目id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "收费适用科室", Val(str项目ID))
    
    With rsTmp
        If .RecordCount > 0 Then
            opt使用科室(0).value = True
            Lvw科室.Enabled = True
            Do While Not .EOF
                For n = 1 To Lvw科室.ListItems.Count
                    If Val(Mid(Lvw科室.ListItems(n).Key, 2)) = !科室ID Then
                        Lvw科室.ListItems(n).Checked = True
                    End If
                Next
                .MoveNext
            Loop
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsValid收入项目和从属关系() As Boolean
    '检查主项的价格如果存在多个收入项目时就不能再设置从属项目；如果有从属项目就不能设置多个收入项目
    Dim rs As New ADODB.Recordset
    Dim blnIs存在多个收入项目 As Boolean
    Dim blnIs存在从属项目 As Boolean
    Dim i As Integer
    
    '是否已存在多个收入项目,按价格等级分组判断
    On Error GoTo ErrHandle
    If mstrID <> "" Then
        gstrSQL = "Select 1 From 收费价目" & vbNewLine & _
                " Where 收费细目id=[1] And 执行日期 <= SYSDATE AND (终止日期 > SYSDATE OR 终止日期 IS NULL) " & vbNewLine & _
                " Group By 价格等级" & vbNewLine & _
                " Having Count(1) > 1"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        If Not rs.EOF Then
            blnIs存在多个收入项目 = True
        End If
        rs.Close
    End If
    
    '编辑后是否存在多个收入项目
    '只要一个价格等级存在多个收入项目，都认为是存在多个收入项目
    If medit方式 = EditNew Or medit方式 = EditCopy Or medit方式 = EditRaise Then
        For i = msh价目.LBound To msh价目.UBound
            If Me.msh价目(i).Rows > 2 Then
                If Me.msh价目(i).TextMatrix(2, mcstCol原价) <> "" Then
                    blnIs存在多个收入项目 = True
                    Exit For
                Else
                    blnIs存在多个收入项目 = False
                End If
            Else
                blnIs存在多个收入项目 = False
            End If
        Next
    End If
            
    '是否已存在从属
    If mstrID <> "" Then
        gstrSQL = "select 主项id from 收费从属项目 where 主项id=[1] "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        
        If rs.RecordCount > 0 Then
            blnIs存在从属项目 = True
        End If
        rs.Close
    End If
    
    '编辑后是否存在从属
    If medit方式 = EditNew Or medit方式 = EditCopy Or medit方式 = EditSlave Then
        If Me.msh从属.Rows > 1 Then
            If Me.msh从属.TextMatrix(1, 1) <> "" Then
                blnIs存在从属项目 = True
            Else
                blnIs存在从属项目 = False
            End If
        Else
            blnIs存在从属项目 = False
        End If
    End If
    
    '如果存在多个收入项目和从属关系的互斥，就提示
    If blnIs存在多个收入项目 And blnIs存在从属项目 Then
         '根据编辑状态显示提示窗口
        Select Case medit方式
        Case EditNew, EditCopy
            MsgBox "如果主项的价格设置了多个收入项目，就不能再设置从属项目；如果设置了从属，价格就不能有多个收入项目。", vbExclamation, gstrSysName
        Case EditRaise
            MsgBox "主项已经有从属项目，不能设置多个收入项目。", vbExclamation, gstrSysName
        Case EditSlave
            MsgBox "主项的价格有多个收入项目，不能再设置从属项目。", vbExclamation, gstrSysName
        End Select
        IsValid收入项目和从属关系 = False
        Exit Function
    End If
    
    IsValid收入项目和从属关系 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSpecialItem(ByVal strID As String) As Boolean
    '判断是否是调价需要特殊处理的项目
    '特殊项目指的是：1、床位和护理类项目以及包含在"自动计价项目"中的其它自动计算项目(计算标志为6,7)
    '                2、当前项目是否是其他床位或者护理类项目的从属项目
    '返回True－是特殊项目
    '返回False－不是特殊项目
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim blnTmp As Boolean
    
    On Error GoTo ErrHandle
    strSQL = "Select Id From 收费项目目录 " & _
        " Where Id=[1] And (类别='J' Or 类别='H')" & _
        " Or Id= (Select Distinct 收费细目id From 自动计价项目 Where 计算标志 In(6,7) And 收费细目id=[1])"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-判断是否是特殊项目。", Val(strID))
    
    blnTmp = (rs.RecordCount > 0)
    
    If Not blnTmp And Val(strID) <> 0 Then
        gstrSQL = "Select ID From 收费项目目录 Where ID In (Select 主项id From 收费从属项目 Where 从项id = [1]) And 类别 In ('J', 'H') And Rownum = 1"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "从属项目", Val(strID))
        
        blnTmp = (rs.RecordCount > 0)
    End If
    GetSpecialItem = blnTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub load性质分类(ByVal intType As Integer)
    'intType:0-执行科室（所有性质，服务于病人）；1-病人科室（临床性质）
    
    mblnRefresh = True
    
    With cboProperty
        .Clear
        
        If mrs性质分类 Is Nothing Then Exit Sub
        
        If intType = 0 Then
            mrs性质分类.Filter = "服务病人=1 Or 服务病人=2 Or 服务病人=3"
        Else
            mrs性质分类.Filter = "名称='临床'"
        End If
        
        If mrs性质分类.RecordCount = 0 Then Exit Sub
        
        If intType = 0 Then
            .AddItem "所有性质"
            
            Do While Not mrs性质分类.EOF
                .AddItem mrs性质分类!名称
                
                mrs性质分类.MoveNext
            Loop
        Else
            .AddItem "临床"
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    DoEvents
    
    mblnRefresh = False
End Sub

Private Function TabExist(ByVal strTabName As String) As Boolean
    Dim i As Integer
    
    For i = 1 To TabMain.Tabs.Count
        If TabMain.Tabs(i).Key = "_" & strTabName Then
            TabExist = True
            Exit Function
        End If
    Next
End Function

Public Function 编辑项目(ByVal strPrivs As String, ByVal blnCanUpdateAll As Boolean, _
    ByVal str分类ID As String, Optional strID As String = "", _
    Optional ByVal lng末级项目 As Long = 1, Optional ByVal edit方式 As EditMode = EditNew, _
    Optional ByVal PriceImp As Boolean = False) As Boolean
    '功能:用来与调用的收费细目管理窗口进行通讯的程序
    '参数:
    '     strPrivs 权限串
    '     str分类ID   收费项目的分类ID   '为数字表示ID，否则为类别名
    '     strID           本收费项目的的ID
    '     bln末级项目     本收费项目是否末级
    '     edit方式  取值为：0、新增；1、修改；2、调价；3、执行科室、4、从属项目、5、复制新增
    '     PriceImp  =True表示使用医价 =False表档不使用医价 默认为不使用医价
    '返回值:编辑成功返回True,否则为False
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    mstrPrivs = strPrivs
    mblnCanUpdateAll = blnCanUpdateAll
    mblnShow收费价目 = False
    
    mblnVerifyPris = IIF(InStr(1, ";" & strPrivs & ";", ";收费价目调价审核;") > 0, True, False)
    mblnVerifyFlow = IIF(Val(zlDatabase.GetPara("调价需要审核", glngSys, 1009, 0)) = 0, False, True)
        
    '不使用医价时屏蔽（标识主码和子码）
'    If PriceImp = False Then
'        Me.txtEdit(9).Enabled = False
'        Me.txtEdit(9).BackColor = &H80000004
'        Me.txtEdit(10).Enabled = False
'        Me.txtEdit(10).BackColor = &H80000004
'    End If
    
    medit方式 = edit方式
    mstrID = strID
    Call GetPriceGrade(gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    '根据收费项目的ID读出类别
    If medit方式 <> EditNew Then
        If IsNumeric(mstrID) Then
            strSQL = "select 类别 from 收费项目目录 where 类别<>'5' and 类别<>'6' and 类别<>'7' and id=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
            
            If rsTmp.RecordCount < 1 Then
                MsgBox "不存在的项目ID！", vbExclamation, gstrSysName
                Exit Function
            End If
            mstr类别 = rsTmp!类别
        Else
            MsgBox "无效的项目ID！", vbExclamation, gstrSysName
            Exit Function
        End If
        '判断是否是特殊项目
        mblnIsSpecialItem = GetSpecialItem(strID)
    Else
        '对于新增的只有从外面传入类别
        mstr类别 = Mid(str分类ID, 2, 1)
        strSQL = "select 1 from 收费项目类别 where 编码<>'4' And 编码<>'5' and 编码<>'6' and 编码<>'7' and Upper(编码)=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(mstr类别)))
        
        If rsTmp.RecordCount < 1 Then
            mstr类别 = ""
        End If
    End If
    If edit方式 <> EditNew Then
        '判断该收费项目是否存在,并根据项目ID求出分类ID
        strSQL = "select 分类ID from 收费项目目录 where 类别<>'5' and 类别<>'6' and 类别<>'7' and  id=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
        
        If rsTmp.RecordCount > 0 Then
            mstr分类ID = Nvl(rsTmp!分类id)
        Else
            MsgBox "选定收费项目不存在！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf edit方式 = EditNew Then
        If Len(str分类ID) > 2 Then
            If IsNumeric(Mid(str分类ID, 3)) Then
                '判断该收费项目是否存在,并根据项目ID求出分类ID
                strSQL = "select ID from 收费分类目录 where id=[1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(str分类ID, 3)))
                
                If rsTmp.RecordCount < 1 Then
                    mstr分类ID = ""
                Else
                    mstr分类ID = CStr(rsTmp!ID)
                End If
            Else
                mstr分类ID = ""
            End If
        Else
            mstr分类ID = ""
        End If
    End If
    
    mstr编码 = ""
    
    If Trim(mstr分类ID) = "0" Then
        mstr分类ID = ""
    End If
    frmChargeItem.Caption = "收费项目设置"
    Call GetDefineSize
    If edit方式 <> EditNew And edit方式 <> EditCopy Then chk保留.Visible = False
    msh从属.Cols = 3
    msh价目(0).Cols = mcstCols
    TabMain.Tabs.Clear
    Select Case edit方式
    Case EditNew, EditCopy
        TabMain.Tabs.Add , "_基本信息", "基本信息"
        If init基本 = False Then
            Exit Function
        End If
        TabMain.Tabs.Add , "_收费价目", "收费价目"
        TabMain.Tabs.Add , "_执行科室", "执行科室"
        If InStr(strPrivs, "项目组合设置") > 0 Then
            TabMain.Tabs.Add , "_从属项目", "从属项目"
            init从属
            Call Ini收费适用科室(mstrID)
        End If
        If init价目 = False Then Exit Function
        init执行
        chk保留.Visible = True
        '由于是复制新增，所以要清除一下内容
        If medit方式 = EditCopy Then
            ClearContext False
        End If
    Case EditModify
        TabMain.Tabs.Add , "_基本信息", "基本信息"
        init基本
    Case EditRaise
        TabMain.Tabs.Add , "_收费价目", "收费价目"
        If init价目 = False Then Exit Function
    Case EditDept
        TabMain.Tabs.Add , "_执行科室", "执行科室"
        init执行
    Case EditSlave
        TabMain.Tabs.Add , "_从属项目", "从属项目"
        init从属
        Call Ini收费适用科室(mstrID)
    End Select
    Call tabMain_Click
    mblnChange = False
    frmChargeItem.Show vbModal
    编辑项目 = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnChange = False
End Function

Private Function IsValid执行() As Boolean
    '检查是否需要设置执行科室
    On Error GoTo ErrHandle
    Dim i As Long
    Dim j As Long
    Dim blnEmpt As Boolean
    Dim strTemp As String

    If opt科室(4).value = True Then '指定科室
        Select Case mstrServerObj
            Case "1"
                If txt门诊执行.Tag = "" Then
                    blnEmpt = True
                    strTemp = ",收费项目服务对象为【门诊】时，应该至少设置一个执行科室！"
                End If
            Case "2"
                If txt住院执行.Tag = "" Then
                    blnEmpt = True
                    strTemp = ",收费项目服务对象为【住院】时，应该至少设置一个执行科室！"
                End If
            Case "3"
                If txt门诊执行.Tag = "" And txt住院执行.Tag = "" Then
                    blnEmpt = True
                    strTemp = ",收费项目服务对象为【门诊和住院】时，应该至少设置一个执行科室！"
                End If
        End Select
        
        If blnEmpt = True Then
            If msf定向执行.TextMatrix(1, 0) <> "" And msf定向执行.TextMatrix(1, 2) <> "" Then
                IsValid执行 = True
            Else
                MsgBox "指定科室" & strTemp, vbInformation, gstrSysName
                If medit方式 = EditNew Or medit方式 = EditCopy Then '新增、复制新增
                    TabMain.Tabs(3).Selected = True
                End If
                Select Case mstrServerObj
                    Case "1"
                        txt门诊执行.SetFocus
                    Case "2"
                        txt住院执行.SetFocus
                    Case "3"
                        txt门诊执行.SetFocus
                End Select
                IsValid执行 = False
            End If
        Else
            IsValid执行 = True
        End If
    Else
        IsValid执行 = True
    End If
'    If Trim(mstr类别) <> "1" And Trim(mstr类别) <> "H" And Trim(mstr类别) <> "J" Then
'        If sstAdmin.Enabled = True Then
'            txtOutIn.Visible = False
'            cmdSel开单科室(0).Visible = False
'            cmdSel执行科室(0).Visible = False
'            cmdSel开单科室(1).Visible = False
'            cmdSel执行科室(1).Visible = False
'ReOut:
'            For i = 2 To msfOut.Rows - 1
'                If Trim(msfOut.TextMatrix(i, 0)) = "" And Trim(msfOut.TextMatrix(i, 2)) = "" Then
'                    msfOut.RemoveItem i
'                    GoTo ReOut
'                End If
'            Next
'ReIn:
'            For i = 2 To msfIn.Rows - 1
'                If Trim(msfIn.TextMatrix(i, 0)) = "" And Trim(msfIn.TextMatrix(i, 2)) = "" Then
'                    msfIn.RemoveItem i
'                    GoTo ReIn
'                End If
'            Next
'            For i = 0 To msfOut.Rows - 1
'                If Trim(msfOut.TextMatrix(i, 0)) = "" And Trim(msfOut.TextMatrix(i, 2)) <> "" Then
'                    msfOut.Row = i: msfOut.Col = 0
'                    sstAdmin.Tab = 0
''                    msfOut_RowColChange
'                    MsgBox "开单科室不能为空！", vbExclamation, gstrSysName
'                    If msfOut.Enabled And msfOut.Visible Then
'                        msfOut.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                If Trim(msfOut.TextMatrix(i, 1)) = "" And Trim(msfOut.TextMatrix(i, 2)) <> "" Then
'                    msfOut.Row = i: msfOut.Col = 1
'                    sstAdmin.Tab = 0
''                    msfOut_RowColChange
'                    MsgBox "执行科室不能为空！", vbExclamation, gstrSysName
'                    If msfOut.Enabled And msfOut.Visible Then
'                        msfOut.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                For j = 1 To msfOut.Rows - 1
'                    If msfOut.TextMatrix(i, 0) = msfOut.TextMatrix(j, 0) And i <> j Then
'                        msfOut.Row = j: msfOut.Col = 0
'                        sstAdmin.Tab = 0
''                        msfOut_RowColChange
'                        MsgBox "开单科室 " & msfOut.Text & " 存在重复！", vbExclamation, gstrSysName
'                        If msfOut.Enabled And msfOut.Visible Then
'                            msfOut.SetFocus
'                            txtOutIn.Visible = True
'                        End If
'                        Exit Function
'                    End If
'                Next
'            Next
'            For i = 0 To msfIn.Rows - 1
'                If Trim(msfIn.TextMatrix(i, 0)) = "" And Trim(msfIn.TextMatrix(i, 2)) <> "" Then
'                    msfIn.Row = i: msfIn.Col = 0
'                    sstAdmin.Tab = 1
''                    msfIn_RowColChange
'                    MsgBox "开单科室不能为空！", vbExclamation, gstrSysName
'                    If msfIn.Enabled And msfIn.Visible Then
'                        msfIn.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                If Trim(msfIn.TextMatrix(i, 1)) = "" And Trim(msfIn.TextMatrix(i, 2)) <> "" Then
'                    msfIn.Row = i: msfIn.Col = 1
'                    sstAdmin.Tab = 1
''                    msfIn_RowColChange
'                    MsgBox "执行科室不能为空！", vbExclamation, gstrSysName
'                    If msfIn.Enabled And msfIn.Visible Then
'                        msfIn.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                For j = 1 To msfIn.Rows - 1
'                    If Trim(msfIn.TextMatrix(i, 0)) = Trim(msfIn.TextMatrix(j, 0)) And i <> j Then
'                        msfIn.Row = j: msfIn.Col = 0
'                        sstAdmin.Tab = 1
''                        msfIn_RowColChange
'                        MsgBox "开单科室 " & msfIn.Text & " 存在重复！", vbExclamation, gstrSysName
'                        If msfIn.Enabled And msfIn.Visible Then
'                            msfIn.SetFocus
'                            txtOutIn.Visible = True
'                        End If
'                        Exit Function
'                    End If
'                Next
'            Next
'        End If
'    End If
'    IsValid执行 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
    '检查非法性
    On Error GoTo ErrHandle
    Dim i As Long
    Select Case medit方式
    Case EditNew, EditCopy
        If IsValid基本 = False Then Exit Function
        If IsValid执行 = False Then Exit Function
        
        If IsValid收入项目和从属关系 = False Then Exit Function
        If IsValid价目 = False Then Exit Function
        If InStr(frmChargeManage.mstrPrivs, "项目组合设置") > 0 Then
            If IsValid从属 = False Then Exit Function
        End If
    Case EditModify
        If IsValid基本 = False Then Exit Function
        '如果显示了调价界面，则要检查价目
        If mblnShow收费价目 Then
            If IsValid价目 = False Then Exit Function
        End If
    Case EditRaise
        If IsValid收入项目和从属关系 = False Then Exit Function
        If IsValid价目 = False Then Exit Function
    Case EditDept
        If IsValid执行 = False Then Exit Function
        If optApply(0).value = False Then
            For i = 1 To 3
                If optApply(i).value = True Then
                    If MsgBox("你选择了“" & Mid(optApply(i).Caption, 1, InStr(optApply(i).Caption, "(") - 1) & "”应用模式。" & vbCrLf & _
                        "这会影响到其它项目，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Next
        End If
    Case EditSlave
        If IsValid收入项目和从属关系 = False Then Exit Function
        If IsValid从属 = False Then Exit Function
    End Select
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save细目() As Boolean
    '根据当前模式保存细目
    On Error GoTo errSave
    gcnOracle.BeginTrans
    Select Case medit方式
    Case EditNew, EditCopy
        If Save基本() = False Then Exit Function
        If Save价目 = False Then Exit Function
        If Save执行 = False Then Exit Function
        If InStr(frmChargeManage.mstrPrivs, "项目组合设置") > 0 Then
            If Save从属 = False Then Exit Function
        End If
    Case EditModify
        If Save基本 = False Then Exit Function
        '如果调出了调价界面，则必须要重新保存价目
        If mblnShow收费价目 Then
            If Save价目 = False Then Exit Function
        End If
    Case EditRaise
        If Save价目 = False Then Exit Function
    Case EditDept
        If Save执行 = False Then Exit Function
    Case EditSlave
        If Save从属 = False Then Exit Function
    End Select
    gcnOracle.CommitTrans
    Save细目 = True
    Exit Function
errSave:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save基本() As Boolean
    Dim str别名 As String
    Dim i  As Integer
    Dim intDept As Integer
    Dim int特殊项目 As Integer
    Dim str站点 As String
    
    On Error GoTo ErrHandle
    With mshAlias
        If Trim(txtEdit(text简码).Text) <> "" Then
            str别名 = "1''" & txtEdit(Text名称).Text & "''1''" & txtEdit(text简码).Text & "''"
        End If
        If Trim(txtEdit(text五笔).Text) <> "" Then
            str别名 = str别名 & "1''" & txtEdit(Text名称).Text & "''2''" & txtEdit(text五笔).Text & "''"
        End If
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) <> "" Then
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    str别名 = str别名 & "9''" & Trim(.TextMatrix(i, 0)) & "''1''" & Trim(.TextMatrix(i, 1)) & "''"
                End If
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    str别名 = str别名 & "9''" & Trim(.TextMatrix(i, 0)) & "''2''" & Trim(.TextMatrix(i, 2)) & "''"
                End If
            End If
        Next
    End With
    If mstr类别 = "1" Then
        If chk急诊.value = 0 Then
            int特殊项目 = 1
        Else
            int特殊项目 = 2
        End If
    ElseIf mstr类别 = "H" Then
        int特殊项目 = cmb护理.ListIndex + 3
    Else
        int特殊项目 = 0
    End If
    str站点 = zlStr.NeedCode(cmbStationNo.Text, "-")
    
    If medit方式 <> EditModify Then
        '新增
        For i = 0 To 6
            If opt科室(i).value = True Then
                intDept = i
                Exit For
            End If
        Next
        mstrID = sys.NextId("收费项目目录")
        gstrSQL = "zl_收费细目_insert(" & int特殊项目 & "," & mstrID & ",'" & mstr类别 & "','" & txtEdit(text编码).Text & "','" & txtEdit(text标识主码).Text & "','" & txtEdit(Text标识子码).Text & "','" & txtEdit(Text备选码).Text & "','" & txtEdit(Text名称).Text & _
            "'," & IIF(mstr分类ID = "", "Null", mstr分类ID) & ",'" & Replace(txtEdit(Text规格).Text, "'", "''") & "','" & Replace(txtEdit(Text说明).Text, "'", "''") & _
            "','" & cmb计算单位.Text & "'," & GetTextFromCombo(cmb费用类型, True) & "," & chk屏蔽费别.value & "," & chk变价.value & "," & chk加班加价.value & "," & intDept & "," & _
            Left(cmb服务对象.Text, 1) & "," & chk摘要.value & "," & txtEdit(text最高限价).Text & "," & txtEdit(text最低限价).Text & ",'" & str别名 & "'," & Val(Me.txt录入限量.Text) & "," & cbo录入限量范围.ListIndex & "," & cmb费用确认.ListIndex & "," & chk费用确认范围.value & "," & chk自动计算.value & _
            ",'" & str站点 & "','" & txtEdit(text病案费目).Text & "'," & cmb项目特性.ListIndex & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        '为材料更新产地
        If mstr类别 = "M" Then
            gstrSQL = "ZL_收费细目_材料产地(" & mstrID & ",'" & Replace(Me.txtEdit(text产地).Text, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Else
        '修改
        gstrSQL = "zl_收费细目_update(" & int特殊项目 & "," & mstrID & ",'" & mstr类别 & "','" & txtEdit(text编码).Text & "','" & txtEdit(text标识主码).Text & "','" & txtEdit(Text标识子码).Text & "','" & txtEdit(Text备选码).Text & "','" & txtEdit(Text名称).Text & _
            "'," & IIF(mstr分类ID = "", "Null", mstr分类ID) & _
            ",'" & Replace(txtEdit(Text规格).Text, "'", "''") & "','" & Replace(txtEdit(Text说明).Text, "'", "''") & "','" & cmb计算单位.Text & "'," & GetTextFromCombo(cmb费用类型, True) & _
            "," & chk屏蔽费别.value & "," & chk变价.value & "," & chk加班加价.value & "," & _
            Left(cmb服务对象.Text, 1) & "," & chk摘要.value & "," & txtEdit(text最高限价).Text & "," & txtEdit(text最低限价).Text & ",'" & str别名 & "'," & Val(Me.txt录入限量.Text) & "," & cbo录入限量范围.ListIndex & "," & cmb费用确认.ListIndex & "," & chk费用确认范围.value & "," & chk自动计算.value & _
            ",'" & str站点 & "','" & txtEdit(text病案费目).Text & "'," & cmb项目特性.ListIndex & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        '为材料更新产地
        If mstr类别 = "M" Then
            gstrSQL = "ZL_收费细目_材料产地(" & mstrID & ",'" & Replace(Me.txtEdit(text产地).Text, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
    Save基本 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save价目() As Boolean
    Dim intRow As Integer
    Dim i As Integer, k As Integer
    Dim lng调价ID As Long
    Dim lng价目ID As Long
    Dim dateExec As Date
    Dim str开始时间 As String
    Dim str终止时间 As String
    Dim strNo As String
    Dim str填制日期 As String
    Dim dtTemp As Date
    
    On Error GoTo ErrHandle
    dtTemp = sys.Currentdate
    For k = 0 To tbPriceGrade.ItemCount - 1
        If tbPriceGrade(k).Visible _
            And Not (mblnChanged价目(k) = False And (medit方式 = EditModify Or medit方式 = EditRaise)) Then    '隐藏的和未改变的不保存
            str开始时间 = Format(IIF(Me.chkNow(k).value = 1, dtTemp, dtpBegin(k).value), mstrCurrentDateFormat)
            str终止时间 = Format(DateAdd("s", -1, str开始时间), "yyyy-MM-dd hh:mm:ss")
            str填制日期 = Format(dtTemp, "yyyy-MM-dd hh:mm:ss")
            intRow = 0
        
            '1、启用流程则在表中插入数据，否则不插入数据
            '2、有权限则审核，没有权限则不审核
            With msh价目(k)
                If Trim(.TextMatrix(1, mcstCol收费项目)) <> "" Then
                    lng调价ID = sys.NextId("收费价目")
                    strNo = sys.GetNextNo(9)
            
                    If medit方式 = EditRaise Or (medit方式 = EditModify And mblnShow收费价目) Then
                        If mblnVerifyFlow = False And mblnVerifyPris = False Then
                            gcnOracle.RollbackTrans
                            MsgBox "在没有启用调价审核模式下，操作员必须要有审核权限才能调价！", vbInformation, gstrSysName
                            Exit Function
                        End If
                        '调价
                        If mblnVerifyFlow = True Then
                            For i = 1 To .Rows - 1
                                If .RowData(i) > 0 Then
                                    If intRow = 0 Then
                                        lng价目ID = lng调价ID
                                    Else
                                        lng价目ID = sys.NextId("收费价目")
                                    End If
                                    'Zl_收费调价记录_Insert(
                                    gstrSQL = "Zl_收费调价记录_Insert("
                                    '  Id_In         In 收费价目.Id%Type,
                                    gstrSQL = gstrSQL & "" & lng价目ID & ","
                                    '  原价id_In     In 收费价目.原价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(mcol价目(k)("C" & .RowData(i)) = 0, "null", mcol价目(k)("C" & .RowData(i))) & ","
                                    '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & mstrID & ","
                                    '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & .RowData(i) & ","
                                    '  原价_In       In 收费价目.原价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol原价)) & ","
                                    '  现价_In       In 收费价目.现价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol现价)) & ","
                                    '  缺省价格_In   In 收费价目.缺省价格%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstCol缺省价格)) = 0, "Null", Val(.TextMatrix(1, mcstCol缺省价格))) & ","
                                    '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol附加手术收费率) = "", 0, .TextMatrix(i, mcstCol附加手术收费率)) & ","
                                    '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol加班加价率) = "", 0, .TextMatrix(i, mcstCol加班加价率)) & ","
                                    '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                    gstrSQL = gstrSQL & "'" & txt调价说明(k).Text & "',"
                                    '  调价id_In     In 收费价目.调价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng调价ID & ","
                                    '  填制人_In     In 收费调价记录.填制人%Type := Null,
                                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                    '  填制日期_In   In 收费调价记录.填制日期%Type := Null,
                                    gstrSQL = gstrSQL & "" & "to_date('" & str填制日期 & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  执行日期_In   In 收费价目.执行日期%Type := Null,
                                    gstrSQL = gstrSQL & "" & "to_date('" & str开始时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  变动原因_In   In 收费价目.变动原因%Type := 1,
                                    gstrSQL = gstrSQL & "" & "1" & ","
                                    '  No_In         In 收费价目.No%Type := Null,
                                    gstrSQL = gstrSQL & "'" & strNo & "',"
                                    '  序号_In       In 收费价目.序号%Type := 1,
                                    gstrSQL = gstrSQL & "" & intRow + 1 & ","
                                    '  价格等级_In   In 收费价目.价格等级%Type := Null
                                    gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
                                    If mblnVerifyPris = True Then
                                        'Zl_收费调价记录_Verify(
                                        gstrSQL = "Zl_收费调价记录_Verify("
                                        '  Id_In       In 收费调价记录.Id%Type,
                                        gstrSQL = gstrSQL & "" & lng价目ID & ","
                                        '  审核标志_In In 收费调价记录.审核标志%Type := 1,
                                        gstrSQL = gstrSQL & "" & "1" & ","
                                        '  审核人_In   In 收费调价记录.审核人%Type := Null,
                                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                        '  审核日期_In In 收费调价记录.审核日期%Type := Null,
                                        gstrSQL = gstrSQL & "" & "to_date('" & str填制日期 & "','YYYY-MM-DD HH24:MI:SS')" & ")"
                                        '  说明_In     In 收费调价记录.说明%Type := Null
                                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                    End If
                                    intRow = intRow + 1
                                End If
                            Next
                        Else
                            '有审核权限，未启用审核模式则直接在收费调价记录表中产生已经生效的审核数据
                            If mblnVerifyPris = True Then
                                '当调价栏显示时如果是定价项目则要处理收费价目（设置以前价目的停用时间）
                                If (medit方式 = EditRaise Or (medit方式 = EditModify And mblnShow收费价目)) _
                                    And chk变价.value = 0 Then   '调价
                                    '填写以前价目的终止日期
                                    'ZL_收费价目_STOP(
                                    gstrSQL = "zl_收费价目_stop("
                                    '  收费细目id_In In 收费价目.收费细目id%Type,
                                    gstrSQL = gstrSQL & mstrID & ","
                                    '  终止日期_In   In 收费价目.终止日期%Type := Null,
                                    gstrSQL = gstrSQL & "To_Date('" & str终止时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  价格等级_In   In 收费价目.价格等级%Type := Null
                                    gstrSQL = gstrSQL & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                End If
            
                                If chk变价.value = 0 Or medit方式 <> EditRaise Or mblnNew(k) Then
                                    If chk变价.value = 1 And medit方式 = EditModify And mblnShow收费价目 Then
                                        'ZL_收费价目_UPDATE(
                                        gstrSQL = "ZL_收费价目_UPDATE("
                                        '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                        gstrSQL = gstrSQL & "" & mstrID & ","
                                        '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                        gstrSQL = gstrSQL & "" & .RowData(1) & ","
                                        '  原价_In       In 收费价目.原价%Type := Null,
                                        gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstCol原价)) & ","
                                        '  现价_In       In 收费价目.现价%Type := Null,
                                        gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstCol现价)) & ","
                                        '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                        gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol附加手术收费率) = "", 0, .TextMatrix(1, mcstCol附加手术收费率)) & ","
                                        '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                        gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol加班加价率) = "", 0, .TextMatrix(1, mcstCol加班加价率)) & ","
                                        '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                        gstrSQL = gstrSQL & "" & "'" & txt调价说明(k).Text & "',"
                                        '  调价id_In     In 收费价目.调价id%Type := Null,
                                        gstrSQL = gstrSQL & "" & lng调价ID & ","
                                        '  调价人_In     In 收费价目.调价人%Type := Null,
                                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                        '  缺省价格_In   In 收费价目.缺省价格%Type := Null,
                                        gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstCol缺省价格)) = 0, "Null", Val(.TextMatrix(1, mcstCol缺省价格))) & ","
                                        '  价格等级_In   In 收费价目.价格等级%Type := Null
                                        gstrSQL = gstrSQL & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                        Exit Function
                                    End If
                                    For i = 1 To .Rows - 1
                                        If .RowData(i) > 0 Then
                                            If intRow = 0 Then
                                                lng价目ID = lng调价ID
                                            Else
                                                lng价目ID = sys.NextId("收费价目")
                                            End If
                                            'Zl_收费价目_Insert(
                                            gstrSQL = "Zl_收费价目_Insert("
                                            '  Id_In         In 收费价目.Id%Type,
                                            gstrSQL = gstrSQL & "" & lng价目ID & ","
                                            '  原价id_In     In 收费价目.原价id%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(mcol价目(k)("C" & .RowData(i)) = 0, "null", mcol价目(k)("C" & .RowData(i))) & ","
                                            '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                            gstrSQL = gstrSQL & "" & mstrID & ","
                                            '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                            gstrSQL = gstrSQL & "" & .RowData(i) & ","
                                            '  原价_In       In 收费价目.原价%Type := Null,
                                            gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol原价)) & ","
                                            '  现价_In       In 收费价目.现价%Type := Null,
                                            gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol现价)) & ","
                                            '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol附加手术收费率) = "", 0, .TextMatrix(i, mcstCol附加手术收费率)) & ","
                                            '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol加班加价率) = "", 0, .TextMatrix(i, mcstCol加班加价率)) & ","
                                            '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                            gstrSQL = gstrSQL & "'" & txt调价说明(k).Text & "',"
                                            '  调价id_In     In 收费价目.调价id%Type := Null,
                                            gstrSQL = gstrSQL & "" & lng调价ID & ","
                                            '  调价人_In     In 收费价目.调价人%Type := Null,
                                            gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                            '  执行日期_In   In 收费价目.执行日期%Type := Null,
                                            gstrSQL = gstrSQL & "" & "to_date('" & str开始时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                            '  变动原因_In   In 收费价目.变动原因%Type := 1,
                                            gstrSQL = gstrSQL & "" & "1" & ","
                                            '  No_In         In 收费价目.No%Type := Null,
                                            gstrSQL = gstrSQL & "'" & strNo & "',"
                                            '  序号_In       In 收费价目.序号%Type := 1,
                                            gstrSQL = gstrSQL & "" & intRow + 1 & ","
                                            '  缺省价格_In   In 收费价目.缺省价格%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstCol缺省价格)) = 0, "Null", Val(.TextMatrix(1, mcstCol缺省价格))) & ","
                                            '  调价汇总号_In In 收费价目.调价汇总号%Type := Null,
                                            gstrSQL = gstrSQL & "" & "NULL" & ","
                                            '  价格等级_In   In 收费价目.价格等级%Type := Null
                                            gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                            intRow = intRow + 1
                                        End If
                                    Next
                                Else
                                    '变价直接修改
                                    'ZL_收费价目_UPDATE(
                                    gstrSQL = "ZL_收费价目_UPDATE("
                                    '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & mstrID & ","
                                    '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & .RowData(1) & ","
                                    '  原价_In       In 收费价目.原价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstCol原价)) & ","
                                    '  现价_In       In 收费价目.现价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstCol现价)) & ","
                                    '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol附加手术收费率) = "", 0, .TextMatrix(1, mcstCol附加手术收费率)) & ","
                                    '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol加班加价率) = "", 0, .TextMatrix(1, mcstCol加班加价率)) & ","
                                    '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                    gstrSQL = gstrSQL & "" & "'" & txt调价说明(k).Text & "',"
                                    '  调价id_In     In 收费价目.调价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng调价ID & ","
                                    '  调价人_In     In 收费价目.调价人%Type := Null,
                                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                    '  缺省价格_In   In 收费价目.缺省价格%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstCol缺省价格)) = 0, "Null", Val(.TextMatrix(1, mcstCol缺省价格))) & ","
                                    '  价格等级_In   In 收费价目.价格等级%Type := Null
                                    gstrSQL = gstrSQL & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                End If
                            End If
                        End If
                    Else
                        '正常新增收费细目直接插入到收费价目表中即可
                        If medit方式 = EditNew Or medit方式 = EditCopy Then
                            For i = 1 To .Rows - 1
                                If .RowData(i) > 0 Then
                                    If intRow = 0 Then
                                        lng价目ID = lng调价ID
                                    Else
                                        lng价目ID = sys.NextId("收费价目")
                                    End If
                                    'Zl_收费价目_Insert(
                                    gstrSQL = "Zl_收费价目_Insert("
                                    '  Id_In         In 收费价目.Id%Type,
                                    gstrSQL = gstrSQL & "" & lng价目ID & ","
                                    '  原价id_In     In 收费价目.原价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(mcol价目(k)("C" & .RowData(i)) = 0, "null", mcol价目(k)("C" & .RowData(i))) & ","
                                    '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & mstrID & ","
                                    '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & .RowData(i) & ","
                                    '  原价_In       In 收费价目.原价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol原价)) & ","
                                    '  现价_In       In 收费价目.现价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol现价)) & ","
                                    '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol附加手术收费率) = "", 0, .TextMatrix(i, mcstCol附加手术收费率)) & ","
                                    '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol加班加价率) = "", 0, .TextMatrix(i, mcstCol加班加价率)) & ","
                                    '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                    gstrSQL = gstrSQL & "'" & txt调价说明(k).Text & "',"
                                    '  调价id_In     In 收费价目.调价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng调价ID & ","
                                    '  调价人_In     In 收费价目.调价人%Type := Null,
                                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                    '  执行日期_In   In 收费价目.执行日期%Type := Null,
                                    gstrSQL = gstrSQL & "" & "to_date('" & str开始时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  变动原因_In   In 收费价目.变动原因%Type := 1,
                                    gstrSQL = gstrSQL & "" & "1" & ","
                                    '  No_In         In 收费价目.No%Type := Null,
                                    gstrSQL = gstrSQL & "'" & strNo & "',"
                                    '  序号_In       In 收费价目.序号%Type := 1,
                                    gstrSQL = gstrSQL & "" & intRow + 1 & ","
                                    '  缺省价格_In   In 收费价目.缺省价格%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstCol缺省价格)) = 0, "Null", Val(.TextMatrix(1, mcstCol缺省价格))) & ","
                                    '  调价汇总号_In In 收费价目.调价汇总号%Type := Null,
                                    gstrSQL = gstrSQL & "" & "NULL" & ","
                                    '  价格等级_In   In 收费价目.价格等级%Type := Null
                                    gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                    intRow = intRow + 1
                                End If
                            Next
                        End If
                    End If
                ElseIf medit方式 = EditModify Or medit方式 = EditRaise Then
                    'Zl_收费价目_Delete(
                    gstrSQL = "Zl_收费价目_Delete("
                    '  细目id_In   In 收费价目.收费细目id%Type,
                    gstrSQL = gstrSQL & "" & mstrID & ","
                    '  站点_In     In 收费项目目录.站点%Type := Null,
                    gstrSQL = gstrSQL & "" & "NULL" & ","
                    '  价格等级_In In 收费价目.价格等级%Type := Null
                    gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            End With
        ElseIf tbPriceGrade(k).Visible = False And medit方式 = EditModify Then
            '删除多余的价格等级数据
            'Zl_收费价目_Delete(
            gstrSQL = "Zl_收费价目_Delete("
            '  细目id_In   In 收费价目.收费细目id%Type,
            gstrSQL = gstrSQL & "" & mstrID & ","
            '  站点_In     In 收费项目目录.站点%Type := Null,
            gstrSQL = gstrSQL & "" & "NULL" & ","
            '  价格等级_In In 收费价目.价格等级%Type := Null
            gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "缺省", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    Save价目 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save执行() As Boolean
    Dim i As Long
    Dim str科室 As String
    Dim lng科室 As Long
    Dim lng应用 As Long
    Dim strTemp As String
    Dim strMid As Variant
    Dim intCount As Integer
    Dim strIn As String
    Dim strOut As String
    
    If medit方式 <> EditDept And opt科室(4).value = False And Not (opt科室(0).value And mstrServerObj <> "1") Then
        Save执行 = True: Exit Function
    End If
    
    '定向执行检查
    On Error GoTo ErrHandle
    With Me.msf定向执行
        strTemp = ""
        For intCount = 1 To .Rows - 1
            If Val(.TextMatrix(intCount, 0)) <> 0 Then
                '不再检查是否重复 By 赵彤宇
                'If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 0)) & "-" & .TextMatrix(intCount, 1) & ";") > 0 Then
                If InStr(1, strTemp & ";", ";" & .TextMatrix(intCount, 0) & ";") > 0 Then
                    MsgBox "重复指定了执行科室“" & .TextMatrix(intCount, 1) & "”！", vbInformation, gstrSysName
                    .SetFocus: Exit Function
                Else
                    strTemp = strTemp & ";" & .TextMatrix(intCount, 0)
                End If
'                    If Val(.TextMatrix(intCount, 2)) = 0 Then
'                        MsgBox "“" & .TextMatrix(intCount, 1) & "”未指定执行科室！", vbInformation, gstrSysName
'                        Me.stbInfo.Tab = 1: .SetFocus: Exit Sub
'                    End If
            End If
        Next
        
        strTemp = ""
        
        For intCount = 1 To .Rows - 1
            If Val(.TextMatrix(intCount, 0)) <> 0 Then
                strMid = Split(.TextMatrix(intCount, 2), ",")
                For i = LBound(strMid) To UBound(strMid)
                    strTemp = strTemp & "|" & Trim(IIF(strMid(i) = "（所有部门）", 0, strMid(i))) & "^" & Trim(.TextMatrix(intCount, 0))
                Next
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        str科室 = strTemp
        
    End With
    
    If Len(Me.txt门诊执行.Tag) > 0 And opt科室(4).value Then
        strOut = Me.txt门诊执行.Tag
    End If
    
    If Len(Me.txt住院执行.Tag) > 0 And (opt科室(4).value Or opt科室(0).value And mstrServerObj <> "1") Then
        strIn = Me.txt住院执行.Tag
    End If
    
    For i = 0 To 6
        If opt科室(i).value = True Then lng科室 = i: Exit For
    Next
    For i = 0 To 3
        If optApply(i).value = True Then lng应用 = i: Exit For
    Next
    
    'Zl_收费细目_Dept(
    gstrSQL = "Zl_收费细目_Dept("
    '  收费细目id_In In 收费项目目录.Id%Type,
    gstrSQL = gstrSQL & "" & mstrID & ","
    '  执行科室_In   In Number,
    gstrSQL = gstrSQL & "" & lng科室 & ","
    '  应用范围_In   In Number,
    gstrSQL = gstrSQL & "" & lng应用 & ","
    '  分类id_In     In 收费项目目录.分类id%Type,
    gstrSQL = gstrSQL & "" & IIF(mstr分类ID = "", "Null", mstr分类ID) & ","
    '  类别_In       In 收费项目目录.类别%Type,
    gstrSQL = gstrSQL & "'" & mstr类别 & "',"
    '  科室列表_In   In Varchar2, --开单科室定向执行的说明串，以|分割，每个定向按开单科室id^执行科室id形式组织
    gstrSQL = gstrSQL & "'" & str科室 & "',"
    '  门诊执行_In   In 诊疗执行科室.执行科室id%Type := Null,
    gstrSQL = gstrSQL & "'" & strOut & "',"
    '  住院执行_In   In 诊疗执行科室.执行科室id%Type := Null,
    gstrSQL = gstrSQL & "'" & strIn & "',"
    '  站点_In       In 收费项目目录.站点%Type := Null
    gstrSQL = gstrSQL & "" & IIF(mblnCanUpdateAll, "NULL", "'" & gstrNodeNo & "'") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save执行 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save从属() As Boolean
    Dim i As Integer
    Dim str科室id As String
    
    On Error GoTo ErrHandle
    If medit方式 = EditSlave Then
        gstrSQL = "zl_收费从属项目_delete(" & mstrID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    With msh从属
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 Then
                gstrSQL = "zl_收费从属项目_insert(" & _
                mstrID & "," & .RowData(i) & "," & .TextMatrix(i, 1) & "," & Left(.TextMatrix(i, 2), 1) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    
    If opt使用科室(0).value = True Then
        With Lvw科室
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked = True Then
                    str科室id = IIF(str科室id = "", "", str科室id & ",") & Mid(.ListItems(i).Key, 2)
                End If
            Next
        End With
    Else
        str科室id = ""
    End If
    gstrSQL = "Zl_收费适用科室_Update(" & mstrID & "," & IIF(str科室id = "", "Null", "'" & str科室id & "'") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save从属 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboProperty_Click()
    If Me.msf定向执行.Col = 3 Then
        Load部门 1, cboProperty.Text
    Else
        Load部门 0, cboProperty.Text
    End If
    
    ChkSelect.value = 0
End Sub


Private Sub cboProperty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
    End Select
End Sub


Private Sub cboProperty_LostFocus()
    Call picDept_LostFocus
End Sub



Private Sub chkNow_Click(Index As Integer)
  '当前是否是立即生效
    If Me.chkNow(Index).value = 1 Then
        Me.dtpBegin(Index).Enabled = False
        '超过当前时间不能立即生效
        If Me.dtpBegin(Index).MinDate > sys.Currentdate Then
            MsgBox "上次执行时间已超过当前时间不能使用立即生效，请手动调整时间！", vbInformation
            Me.chkNow(Index).value = 0
        End If
    ElseIf medit方式 = EditModify And txtEdit(text标识主码).Text <> txtEdit(text标识主码).Tag Then
        MsgBox "你已经改变了医价项目，对应的价格只能选择立即生效！", vbInformation
        Me.chkNow(Index).value = 1
    Else
        Me.dtpBegin(Index).Enabled = True
    End If
End Sub

Private Sub chkNow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub ChkSelect_Click()
    Dim i As Integer
    Dim str名称 As String
    
    If mblnRefresh = True Then Exit Sub
    
    If ChkSelect.value = 2 Then Exit Sub
    Call SetSelect(lvwItems, ChkSelect.value)
    
    If cboProperty.Text = "所有性质" Then
        mstr已选执行科室 = ""
    End If
    
    If ChkSelect.value = 1 Then
        '当前性质全选
        For i = 1 To lvwItems.ListItems.Count
            str名称 = Mid(lvwItems.ListItems(i).Key, 2) & "," & "[" & lvwItems.ListItems(i).SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) & "]" & lvwItems.ListItems(i).Text
            
            If InStr(mstr已选执行科室, str名称) = 0 Or cboProperty.Text = "所有性质" Then
                mstr已选执行科室 = IIF(mstr已选执行科室 = "", "", mstr已选执行科室 & ";") & str名称
            End If
        Next
    ElseIf cboProperty.Text <> "所有性质" Then
        '当前性质全清

        For i = 1 To lvwItems.ListItems.Count
            str名称 = Mid(lvwItems.ListItems(i).Key, 2) & "," & "[" & lvwItems.ListItems(i).SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) & "]" & lvwItems.ListItems(i).Text
               
            If InStr(mstr已选执行科室, str名称) > 0 Then
                mstr已选执行科室 = Replace(mstr已选执行科室, str名称, "")
            End If
        Next
    End If
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ChkSelect_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
    End Select
End Sub

Private Sub chk急诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbStationNo_Change()
    Call init价目(True)
End Sub

Private Sub cmbStationNo_Click()
    Dim strStationNo As String
    
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If Val(mstrID) <> 0 Then
        strStationNo = zlStr.NeedCode(cmbStationNo.Text, "-")
        If strStationNo <> "" And cmbStationNo.Tag <> strStationNo Then
            If CanChangeStation(mstrID) = False Then
                mblnNotClick = True
                cbo.SeekIndex cmbStationNo, cmbStationNo.Tag
                mblnNotClick = False
            End If
        End If
    End If
    
    Call init价目(True)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CanChangeStation(ByVal lngId As Long)
    '判断是否能够调整站点
    '问题号：110164
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select 1" & vbNewLine & _
        " From 收费价目 A" & vbNewLine & _
        " Where a.收费细目id = [1] And a.价格等级 Is Not Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
    If Not rsTemp.EOF Then
        MsgBox "当前项目启用了价格等级，不允许调整为其它院区！", vbInformation, gstrSysName
        Exit Function
    End If
    CanChangeStation = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmb护理_Click()
    Me.chk自动计算.Visible = (Me.cmb护理.ListIndex <> 0)
End Sub
Private Sub cmb护理_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To Lvw科室.ListItems.Count
        If zlStr.GetCodeByVB(Mid(Lvw科室.ListItems(i).Text, InStr(Lvw科室.ListItems(i).Text, "-") + 1)) Like UCase(IIF(gstrLike <> "", "*", "") & strFind & "*") Or _
                UCase(Lvw科室.ListItems(i).Text) Like UCase(IIF(gstrLike <> "", "*", "") & strFind & "*") Then
            Lvw科室.ListItems(i).Selected = True
            Lvw科室.ListItems(i).EnsureVisible
            Lvw科室.SetFocus
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "没有找到您查找的科室。", vbInformation, Me.Caption
        Else
            MsgBox "已经是最后一个科室了。", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdHelp_Click()
    If Me.Caption = "收费项目设置" Then
        ShowHelp App.ProductName, Me.hwnd, "frmChargeItem", Int((glngSys) / 100)
'    ElseIf Me.Caption = "收费分类设置" Then
'        ShowHelp App.ProductName, Me.hwnd, "frm收费项目设置1", Int((glngSys) / 100)
    End If
End Sub

Private Sub cmdOK_Click()
    If IsValid() = False Then Exit Sub
    If Save细目() = False Then Exit Sub
    '刷新主窗口的显示
    Call frmChargeManage.FillTree
    If medit方式 <> EditNew And medit方式 <> EditCopy Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    '连续增加
    ClearContext (chk保留.value = 0)
    ShowTab "基本信息"
    txtEdit(text编码).SetFocus
    mblnChange = False
    mblnOk = True
End Sub

Private Sub ChangeCode(nod As Node, ByVal strOldCode As String, ByVal strNewCode As String)
    '功能:改变下级的编码内容
    Dim nodChild As Node
    
    Set nodChild = nod.Child
    Do Until nodChild Is Nothing
        nodChild.Text = strNewCode & Mid(nodChild.Text, Len(strOldCode))
        ChangeCode nodChild, strOldCode, strNewCode
        Set nodChild = nodChild.Next
    Loop
End Sub

Private Sub chk变价_Click()
    Dim i As Integer
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        If msh价目(i).Rows > 2 Then
            chk变价.value = 0
            Exit Sub
        End If
    Next
    
    '当修改项目改变了“变价/定价”属性时调出调价栏
    If medit方式 = EditModify Then
        If chk变价.value <> chk变价.Tag Then
            If Not mblnShow收费价目 Then
                TabMain.Tabs.Add , "_收费价目", "收费价目"
                mblnShow收费价目 = True
                Call init价目: Call init价目(True)
                MsgBox "请重新确认收费价目。", vbInformation, gstrSysName
            End If
        Else
            If mblnShow收费价目 Then
                TabMain.Tabs.Remove "_收费价目"
                mblnShow收费价目 = False
            End If
        End If
    End If
    For i = 0 To tbPriceGrade.ItemCount - 1
        With msh价目(i)
            If chk变价.value = 1 Then
                .Rows = 2
                .TextMatrix(0, mcstCol原价) = "最低限价"
                .TextMatrix(0, mcstCol现价) = "最高限价"
                .ColData(mcstCol原价) = IIF(gstr医价接口编号 <> "" And gbln允许医价收费项目 = True, 5, 4)
                .ColData(mcstCol现价) = IIF(gstr医价接口编号 <> "" And gbln允许医价收费项目 = True, 5, 4)
                .TextMatrix(1, mcstCol原价) = txtEdit(text最低限价).Text
                .TextMatrix(1, mcstCol现价) = txtEdit(text最高限价).Text
                .ColWidth(mcstCol缺省价格) = 1000
            Else
                .TextMatrix(0, mcstCol原价) = "原价"
                .TextMatrix(0, mcstCol现价) = "现价"
                .TextMatrix(1, mcstCol原价) = "0.000"
                .ColData(mcstCol原价) = 5
                .ColData(mcstCol现价) = 4
                .ColWidth(mcstCol缺省价格) = 0
            End If
        End With
    Next
    mblnChange = True
End Sub

Private Sub chk变价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk加班加价_Click()
    Dim i As Integer
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        With msh价目(i)
            If chk加班加价.value = 1 Then
                .ColWidth(mcstCol加班加价率) = 1500
                .TextMatrix(0, mcstCol加班加价率) = "加班加价率"
            Else
                .ColWidth(mcstCol加班加价率) = 0
            End If
        End With
    Next
    mblnChange = True
End Sub

Private Sub chk加班加价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk屏蔽费别_Click()
    mblnChange = True
End Sub

Private Sub chk屏蔽费别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb费用类型_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandle
    Dim lngIdx As Long
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    Else
        lngIdx = cbo.MatchIndex(cmb费用类型.hwnd, KeyAscii)
        If lngIdx <> -2 Then cmb费用类型.ListIndex = lngIdx
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk摘要_Click()
    mblnChange = True
End Sub

Private Sub chk摘要_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb服务对象_Click()
    mblnChange = True
End Sub

Private Sub cmb费用类型_Click()
    mblnChange = True
End Sub

Private Sub cmb服务对象_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb计算单位_Change()
    mblnChange = True
End Sub

Private Sub cmb计算单位_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub








Private Sub cmdOK_GotFocus()
    ''
End Sub

Private Sub cmdOkDept_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim strArr
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
    
    With Me.lvwItems
        Select Case .Tag
            Case "执行"
                '删除不在已选择列表中的执行科室
                For i = msf定向执行.Rows - 1 To 1 Step -1
                    If InStr(mstr已选执行科室, msf定向执行.TextMatrix(i, 0) & "," & msf定向执行.TextMatrix(i, 1)) = 0 Then
                        If i > 1 Then
                            msf定向执行.MsfObj.RemoveItem i
                        Else
                            msf定向执行.TextMatrix(1, 0) = ""
                            msf定向执行.TextMatrix(1, 1) = ""
                            msf定向执行.TextMatrix(1, 2) = ""
                            msf定向执行.TextMatrix(1, 3) = ""
                        End If
                    End If
                Next
                
                '增加新执行科室
                mstr已选执行科室 = mstr已选执行科室 & ";"
                strArr = Split(mstr已选执行科室, ";")
                
                For i = 0 To UBound(strArr) - 1
                    blnNew = True
                    If strArr(i) <> "" Then
                        For n = 1 To msf定向执行.Rows - 1
                            If strArr(i) = msf定向执行.TextMatrix(n, 0) & "," & msf定向执行.TextMatrix(n, 1) Then
                                blnNew = False
                            End If
                        Next
                        If blnNew = True Then
                            strNew = IIF(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            If msf定向执行.TextMatrix(msf定向执行.Rows - 1, 1) <> "" Then
                                msf定向执行.Rows = msf定向执行.Rows + 1
                            End If
                            msf定向执行.TextMatrix(msf定向执行.Rows - 1, 0) = Split(strArr(i), ",")(0)
                            msf定向执行.TextMatrix(msf定向执行.Rows - 1, 1) = Split(strArr(i), ",")(1)
                        End If
                    Next
                End If
        End Select
    End With
    
    picDept.Visible = False
End Sub

Private Sub cmd病案_Click()
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    
    strSQL = "Select 编码 id,上级 as 上级id, 名称, 简码, 末级 From 病案费目 Start With 上级 Is Null Connect By Prior 编码 = 上级"
    blnRe = frmTreeLeafSel.ShowTree(strSQL, strID, str名称, "病案费目")
    '成功返回
    If blnRe Then
        '新的本级的宽度
        lbl病案费目.Tag = strID
        txtEdit(text病案费目).Text = str名称
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd上级_Click()
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim str编码 As String
    
    strSQL = "select ID,上级ID,名称,编码,简码 from 收费分类目录 " & _
        " start with 上级ID is null   connect by prior ID =上级ID"
    strID = mstr分类ID
    str名称 = txtEdit(Text分类).Text
    str编码 = txtTemp.Text
    blnRe = frmTreeSel.ShowTree(strSQL, strID, str名称, str编码, mstrID, "收费项目选择", "所有收费项目分类", , mstr编码, 3, 4, 4, False)
    '成功返回
    If blnRe Then
        '新的本级的宽度
        mstr分类ID = strID
        txtEdit(Text分类).Text = str名称
        mstr分类编码 = str编码
        Call SetCodeNO
        txtEdit(text编码).MaxLength = mlng编码长度
        mblnChange = True
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsRaiseByDate(ByVal strID As String) As Boolean
    '判断该收费项目是否是按日调价
    '返回True-是按天条件
    '返回False-不是按天调价
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select a.Id" & _
            " From 收费项目目录 A, 收费价目 D" & _
            " Where a.ID = d.收费细目ID And Nvl(d.终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate " & _
            "       And d.执行日期<>trunc(d.执行日期,'dd') And d.收费细目id=[1] "
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
    IsRaiseByDate = Not (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub dtpBegin_Change(Index As Integer)
    mblnChange = True
    If Mid(cmbClass.Text, 1, 1) <> "J" And Mid(cmbClass.Text, 1, 1) <> "H" Then
        If DateDiff("s", Me.dtpBegin(Index).value, Format(sys.Currentdate, "yyyy-mm-dd h:m:s")) > 0 Then
            MsgBox "调价执行时间不能小于当前时间！", vbInformation, gstrSysName
            Me.dtpBegin(Index).value = DateAdd("n", 1, sys.Currentdate)
        End If
    End If
End Sub

Private Sub dtpBegin_GotFocus(Index As Integer)
    msh价目(Index).CmdVisible = False
End Sub

Private Sub dtpBegin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Integer
    
    picOne.Top = mshAlias.Top + mshAlias.Height + IIF(txtEdit(text最低限价).Visible, txtEdit(text最低限价).Height + 100, 0) + 100
    picTwo.Top = picOne.Top + picOne.Height + IIF(txtEdit(text产地).Visible, txtEdit(text产地).Height + 50, 0) + 50
    lblEdit(18).Top = picOne.Top + picOne.Height + 85
    txtEdit(text产地).Top = picOne.Top + picOne.Height + 50
    fra(1).Height = 6660 - (IIF(Not txtEdit(text最低限价).Visible, txtEdit(text最低限价).Height + 50, 0) + 50) - (IIF(Not txtEdit(text产地).Visible, txtEdit(text产地).Height + 50, 0) + 50) - 100
    For i = 2 To 4
        fra(i).Left = fra(1).Left
        fra(i).Top = fra(1).Top
        fra(i).Height = fra(1).Height
    Next
    TabMain.Height = fra(1).Height + 350
    Me.Height = TabMain.Height + 1080
    cmdOK.Top = TabMain.Top + TabMain.Height + 100
    cmdCancel.Top = cmdOK.Top
    cmdHelp.Top = cmdOK.Top
    tbPriceGrade.Height = fra(2).Height - tbPriceGrade.Top - 100
    Frame1.Height = fra(3).Height - IIF(fra批量.Visible, fra批量.Height, 0) - 200
    msf定向执行.Height = Frame1.Height - (lbl定向执行.Top + lbl定向执行.Height) - 150
    fra批量.Top = Frame1.Top + Frame1.Height + 100
    Lvw科室.Height = fra(4).Height - Label2.Top - Label2.Height - 250
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str病人科室ID As String
    Dim str病人科室名称 As String
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Select Case .Tag
        Case "门诊"
            Me.txt门诊执行.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt门诊执行.Text = .SelectedItem.Text
            Me.txt门诊执行.SetFocus: Call OS.PressKey(vbKeyTab)
        Case "住院"
            Me.txt住院执行.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt住院执行.Text = .SelectedItem.Text
            Me.txt住院执行.SetFocus: Call OS.PressKey(vbKeyTab)
        Case "开单"
            With Me.lvwItems
                If Me.msf定向执行.Col = 3 And Me.lvwItems.Checkboxes = True Then
                    For i = 1 To .ListItems.Count
                        If .ListItems(i).Checked = True Then
                            If Me.msf定向执行.Text = "" Then
                                Me.msf定向执行.Text = "[" & .ListItems(i).SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .ListItems(i).Text
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) = Mid(.ListItems(i).Key, 2)
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 3) = Me.msf定向执行.Text
                            Else
                                Me.msf定向执行.Text = Me.msf定向执行.Text & ",[" & .ListItems(i).SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .ListItems(i).Text
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) = Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) & "," & Mid(.ListItems(i).Key, 2)
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 3) = Me.msf定向执行.Text
                            End If
                            m = m + 1
                        End If
                    Next
                    If m = 0 Then
                        Me.msf定向执行.Text = ""
                        Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) = "（所有部门）"
                        Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 3) = "（所有部门）"
                    End If
                Else
                    Me.msf定向执行.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                    Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) = Mid(.SelectedItem.Key, 2)
                    Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 3) = Me.msf定向执行.Text
                End If
            End With
            
            '如果有其他未填的行，询问是否按同一方案增加
            For i = 1 To Me.msf定向执行.Rows - 1
                If Me.msf定向执行.TextMatrix(i, 0) <> "" And Me.msf定向执行.TextMatrix(i, 3) = "" Then
                    blnBatch = True
                    Exit For
                End If
            Next
            
            If blnBatch = True Then
                If MsgBox("是否应用与其他未设置的列？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    str病人科室ID = Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2)
                    str病人科室名称 = Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 3)
                    For i = 1 To Me.msf定向执行.Rows - 1
                        If Me.msf定向执行.TextMatrix(i, 3) = "" Then
                            Me.msf定向执行.TextMatrix(i, 2) = str病人科室ID
                            Me.msf定向执行.TextMatrix(i, 3) = str病人科室名称
                        End If
                    Next
                End If
            End If
            
            Me.msf定向执行.SetFocus
            Call OS.PressKey(vbKeyReturn)
        Case "执行"
            Dim strTmp As String
            Dim strArr
            Dim n As Integer
            Dim strNew As String
            Dim blnNew As Boolean
            
            If Val(Me.picDept.Tag) = 1 And lbl工作性质.Visible = True Then
                '删除不在已选择列表中的执行科室
                For i = msf定向执行.Rows - 1 To 1 Step -1
                    If InStr(mstr已选执行科室, msf定向执行.TextMatrix(i, 0) & "," & msf定向执行.TextMatrix(i, 1)) = 0 Then
                        If i > 1 Then
                            msf定向执行.MsfObj.RemoveItem i
                        Else
                            msf定向执行.TextMatrix(1, 0) = ""
                            msf定向执行.TextMatrix(1, 1) = ""
                            msf定向执行.TextMatrix(1, 2) = ""
                            msf定向执行.TextMatrix(1, 3) = ""
                            
                            If msf定向执行.Rows > 2 Then
                                msf定向执行.MsfObj.RemoveItem 1
                            End If
                        End If
                    End If
                Next
                
                '增加新执行科室
                mstr已选执行科室 = mstr已选执行科室 & ";"
                strArr = Split(mstr已选执行科室, ";")
                
                For i = 0 To UBound(strArr) - 1
                    blnNew = True
                    If strArr(i) <> "" Then
                        For n = 1 To msf定向执行.Rows - 1
                            If strArr(i) = msf定向执行.TextMatrix(n, 0) & "," & msf定向执行.TextMatrix(n, 1) Then
                                blnNew = False
                            End If
                        Next
                        If blnNew = True Then
                            strNew = IIF(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            If msf定向执行.TextMatrix(msf定向执行.Rows - 1, 1) <> "" Then
                                msf定向执行.Rows = msf定向执行.Rows + 1
                            End If
                            msf定向执行.TextMatrix(msf定向执行.Rows - 1, 0) = Split(strArr(i), ",")(0)
                            msf定向执行.TextMatrix(msf定向执行.Rows - 1, 1) = Split(strArr(i), ",")(1)
                        End If
                    Next
                End If
                
                msf定向执行.Row = msf定向执行.Rows - 1
                Me.msf定向执行.SetFocus
                Call OS.PressKey(vbKeyRight)
            Else
                Me.msf定向执行.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 0) = Mid(.SelectedItem.Key, 2)
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 1) = Me.msf定向执行.Text
                Me.msf定向执行.SetFocus
                Call OS.PressKey(vbKeyRight)
            End If

            picDept.Visible = False
        End Select
    End With
End Sub

Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim str名称 As String
    
    If Me.lvwItems.Tag = "执行" Then
        str名称 = Mid(Item.Key, 2) & "," & "[" & Item.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) & "]" & Item.Text
        
        If Item.Checked = True Then
            If InStr(mstr已选执行科室, str名称) = 0 Then
                mstr已选执行科室 = IIF(mstr已选执行科室 = "", "", mstr已选执行科室 & ";") & str名称
            End If
        Else
            If InStr(mstr已选执行科室, str名称) > 0 Then
                mstr已选执行科室 = Replace(mstr已选执行科室, str名称, "")
            End If
        End If
    End If
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.lvwItems.Tag = "开单" Or Me.lvwItems.Tag = "执行" Then
        If KeyCode = vbKeyA And Shift = vbCtrlMask Then '全选 Ctrl+A
            If Me.lvwItems.Tag = "执行" Then
                If Me.ChkSelect.value = 0 Then
                    Me.ChkSelect.value = 1
                    Call SetSelect(lvwItems, True)
                End If
            Else
                Call SetSelect(lvwItems, True)
            End If
        End If
        
        If KeyCode = vbKeyR And Shift = vbCtrlMask Then     '全消 Ctrl+R
            If Me.lvwItems.Tag = "执行" Then
                If Me.ChkSelect.value = 1 Then
                    Me.ChkSelect.value = 0
                    Call SetSelect(lvwItems, False)
                End If
            Else
                Call SetSelect(lvwItems, False)
            End If
        End If
    End If
End Sub

Private Sub lvwItems_GotFocus()
    If Me.lvwItems.Tag = "开单" Or Me.lvwItems.Tag = "执行" Then
        Me.lvwItems.ToolTipText = "全选Ctrl+A；全清Ctrl+R"
    Else
        Me.lvwItems.ToolTipText = ""
    End If
End Sub
Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
         picDept.Visible = False
    End Select

End Sub

Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub Lvw科室_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub Lvw科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmdFind_Click
End Sub

Private Sub Lvw科室_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    For i = 1 To Lvw科室.ListItems.Count
        If zlStr.GetCodeByVB(Mid(Lvw科室.ListItems(i).Text, InStr(Lvw科室.ListItems(i).Text, "-") + 1)) Like UCase(Chr(KeyAscii)) & "*" Then
            Lvw科室.ListItems(i).Selected = True: Exit For
        End If
    Next
End Sub

Private Sub msf定向执行_CommandClick()
    Dim i As Integer
    
    mstr已选执行科室 = ""
    If Me.msf定向执行.Col = 1 Then
        With Me.msf定向执行
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    mstr已选执行科室 = IIF(mstr已选执行科室 = "", "", mstr已选执行科室 & ";") & .TextMatrix(i, 0) & "," & .TextMatrix(i, 1)
                End If
            Next
        End With
    End If
    
    With Me.picDept
        If Me.msf定向执行.Col = 3 Then
            .Tag = ""
            Me.lvwItems.Tag = "开单"
            .Left = Me.fra(3).Left + Me.msf定向执行.Left + Me.msf定向执行.ColWidth(0) + Me.msf定向执行.ColWidth(1) + Me.msf定向执行.ColWidth(2)
            .Width = IIF(Me.msf定向执行.ColWidth(3) < 3000, 3000, Me.msf定向执行.ColWidth(3))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "执行"
            .Left = Me.fra(3).Left + Me.msf定向执行.Left + Me.msf定向执行.ColWidth(0)
            .Width = IIF(Me.msf定向执行.ColWidth(2) < 3000, 3000, Me.msf定向执行.ColWidth(2))
        End If
        
        .Top = Me.fra(3).Top + Me.Frame1.Top + Me.msf定向执行.Top + (Me.msf定向执行.Row - Me.msf定向执行.MsfObj.TopRow + 1) * Me.msf定向执行.RowHeight(0) - 50
        
        If fra批量.Top + fra批量.Height - .Top - 50 > 0 Then
            .Height = fra批量.Top + fra批量.Height - .Top - 50
        Else
            .Height = (fra批量.Top - Frame1.Top - Frame1.Height) + fra批量.Height
        End If
        
        lbl工作性质.Visible = (Me.msf定向执行.Col = 1)
        cboProperty.Visible = lbl工作性质.Visible
        ChkSelect.Visible = lbl工作性质.Visible
        
        If Me.lvwItems.Tag = "执行" Then
            lbl工作性质.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        .ZOrder 0
        .Visible = True
    End With
    
    With Me.lvwItems
        If .Tag = "执行" Then
            .Left = lbl工作性质.Left
            .Top = cboProperty.Top + cboProperty.Height + 50
            .Width = Me.picDept.Width - .Left - 50
            .Height = Me.picDept.Height - .Top - 50
        Else
            .Left = 0
            .Top = 0
            .Width = Me.picDept.Width
            .Height = Me.picDept.Height
        End If
        
        .SetFocus
        .Refresh
    End With
    
    If Me.msf定向执行.Col = 3 Then
        load性质分类 1
    Else
        load性质分类 0
    End If
     
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub msf定向执行_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msf定向执行.TextMatrix(Row, Col)
End Sub

Private Sub msf定向执行_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msf定向执行_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    Dim ObjItem As ListItem
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf定向执行
        If .Active = False Then Call OS.PressKey(vbKeyTab): Exit Sub
        If .TxtVisible = False Then
            If .Col = 1 And .TextMatrix(.Row, 1) = "" Then
                ShowTab "从属项目"
                '----------------------------------
                '没有找到焦点问题,暂用下面的方面处理
                OS.PressKey (vbKeyTab)
                OS.PressKey (vbKeyTab)
                OS.PressKey (vbKeyTab)
                If .Row = 1 Then
                    OS.PressKey (vbKeyTab)
                End If
                '-----------------------------------
                Exit Sub
            End If
            If .Col = 3 And (.TextMatrix(.Row, 3) = "") Then
                .TextMatrix(.Row, 3) = "（所有部门）"
                .TextMatrix(.Row, 2) = "（所有部门）"
                Exit Sub
            End If
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If .Col = 1 And Trim(.Text) = "" Then
                If .Row = 1 Then .SetFocus: Call OS.PressKey(vbKeyTab)
                Exit Sub
            End If
            
            If .Col = 3 And Trim(.Text) = "" Then
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 2) = ""
                Exit Sub
            End If
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    If Me.msf定向执行.Col = 3 Then
        gstrSQL = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and 工作性质='临床'" & _
                "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
                " order by 编码"
    Else
        gstrSQL = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and T.服务对象 in (1,2,3)" & _
                "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
                " order by 编码"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp & "%")
    
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msf定向执行.Text = "[" & !编码 & "]" & !名称
            If Me.msf定向执行.Col = 1 Then
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 0) = !ID
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 1) = Me.msf定向执行.Text
            Else
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 2) = !ID
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 3) = Me.msf定向执行.Text
            End If
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        If Me.msf定向执行.Col = 3 Then
            .Tag = ""
            Me.lvwItems.Tag = "开单"
            .Left = Me.fra(3).Left + Me.msf定向执行.Left + Me.msf定向执行.ColWidth(0) + Me.msf定向执行.ColWidth(1) + Me.msf定向执行.ColWidth(2)
            .Width = IIF(Me.msf定向执行.ColWidth(3) < 3000, 3000, Me.msf定向执行.ColWidth(3))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "执行"
            .Left = Me.fra(3).Left + Me.msf定向执行.Left
            .Width = IIF(Me.msf定向执行.ColWidth(2) < 3000, 3000, Me.msf定向执行.ColWidth(2))
        End If
        
        .Top = Me.fra(3).Top + Me.Frame1.Top + Me.msf定向执行.Top + (Me.msf定向执行.Row - Me.msf定向执行.MsfObj.TopRow + 1) * Me.msf定向执行.RowHeight(0) - 50
        
        If fra批量.Top + fra批量.Height - .Top - 50 > 0 Then
            .Height = fra批量.Top + fra批量.Height - .Top - 50
        Else
            .Height = (fra批量.Top - Frame1.Top - Frame1.Height) + fra批量.Height
        End If
        
        lbl工作性质.Visible = False
        cboProperty.Visible = lbl工作性质.Visible
        ChkSelect.Visible = lbl工作性质.Visible
        
        If Me.msf定向执行.Col = 1 Then
            lbl工作性质.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        .ZOrder 0
        .Visible = True
    End With
    
    With Me.lvwItems
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        
        .SetFocus
        .Refresh
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshAlias_EditKeyPress(KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub mshAlias_KeyPress(KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msh价目_AfterAddRow(Index As Integer, Row As Long)
    If chk变价.value = 1 Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub opt使用科室_Click(Index As Integer)
    If Index = 1 Then
        Lvw科室.Enabled = False
        txtFind.Enabled = False
        cmdFind.Enabled = False
        txtFind.BackColor = &H8000000F
    Else
        Lvw科室.Enabled = True
        txtFind.Enabled = True
        cmdFind.Enabled = True
        txtFind.BackColor = &H80000005
    End If
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDOKDEPT,CMDCANCELDEPT,LVWITEMS,CBOPROPERTY,PICDEPT,CHKSELECT", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
End Sub

Private Sub pic价格等级_Resize(Index As Integer)
    On Error Resume Next
    txt调价说明(Index).Top = pic价格等级(Index).ScaleHeight - txt调价说明(Index).Height - 100
    lbl调价说明(Index).Top = txt调价说明(Index).Top + (txt调价说明(Index).Height - lbl调价说明(Index).Height) \ 2
    dtpBegin(Index).Top = txt调价说明(Index).Top - dtpBegin(Index).Height - 100
    lbl调价执行时间(Index).Top = dtpBegin(Index).Top + (dtpBegin(Index).Height - lbl调价执行时间(Index).Height) \ 2
    chkNow(Index).Top = dtpBegin(Index).Top + (dtpBegin(Index).Height - chkNow(Index).Height) \ 2
    msh价目(Index).Height = dtpBegin(Index).Top - msh价目(Index).Top - 100
End Sub

Private Sub tbPriceGrade_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    msh价目(Val(tbPriceGrade.Tag)).
    tbPriceGrade.Tag = Item.Index
End Sub

Private Sub txt录入限量_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt录入限量_LostFocus()
    Me.txt录入限量.Text = FormatEx(Val(Me.txt录入限量.Text), 5)
End Sub


Private Sub txt门诊执行_Change()
    If Trim(txt门诊执行.Text) = "" Then
        txt门诊执行.Tag = ""
    End If
End Sub

Private Sub txt门诊执行_GotFocus()
     Me.txt门诊执行.SelStart = 0: Me.txt门诊执行.SelLength = 100
End Sub
Private Sub mshAlias_AfterDeleteRow()
    mblnChange = True
End Sub
Private Sub mshAlias_EnterCell(Row As Long, Col As Long)
    If Col = 0 Then
        OS.OpenIme True
        mshAlias.MaxLength = mlng别名长度
    Else
        OS.OpenIme False
        mshAlias.MaxLength = mlng简码长度
    End If
End Sub
Private Sub mshAlias_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strTemp As String
    
    If KeyCode = vbKeyReturn Then
        If mshAlias.TxtVisible = False Then
            If mshAlias.Col = 0 And mshAlias.Row = 1 Then OS.PressKey vbKeyTab
            Exit Sub
        End If
        strTemp = mshAlias.Text
        If mshAlias.Col = 0 Then
            If zlCommFun.StrIsValid(strTemp, mlng别名长度) = False Then
                Cancel = True
                If mshAlias.Active And mshAlias.Visible Then
                    mshAlias.TxtSetFocus
                End If
            Else
                mshAlias.TextMatrix(mshAlias.Row, 1) = zlStr.GetCodeByORCL(strTemp, False, mlng简码长度)
                mshAlias.TextMatrix(mshAlias.Row, 2) = zlStr.GetCodeByORCL(strTemp, True, mlng简码长度)
                
                If mshAlias.TextMatrix(mshAlias.Row, 1) = "" Then mshAlias.TextMatrix(mshAlias.Row, 1) = " "
                If mshAlias.TextMatrix(mshAlias.Row, 2) = "" Then mshAlias.TextMatrix(mshAlias.Row, 2) = " "
            End If
        Else
            Cancel = Not zlCommFun.StrIsValid(strTemp, mlng简码长度)
            If Cancel = True Then
                If mshAlias.Active And mshAlias.Visible Then
                    mshAlias.TxtSetFocus
                End If
            Else
                If strTemp = "" Then mshAlias.Text = " "
            End If
        End If
    End If
    If Cancel = False Then mblnChange = True
End Sub

Private Sub msh价目_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    If msh价目(Index).RowData(Row) <> 0 Then
        mcol价目(Index).Remove "C" & msh价目(Index).RowData(Row)
        mblnChange = True: mblnChanged价目(Index) = True
    End If
End Sub

Private Sub msh价目_LostFocus(Index As Integer)
    If chk变价.value = 1 Then
        msh价目(Index).Rows = 2
    End If
    msh价目(Index).CmdVisible = False
End Sub

Private Sub msh价目_CommandClick(Index As Integer)
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim strTemp As String
    Dim strID As String
    Dim lngRow As Long
    
    With msh价目(Index)
        lngRow = .Row
        strTemp = .TextMatrix(lngRow, mcstCol收费项目)
        strID = .RowData(lngRow)
        strSQL = "select ID,上级ID,名称,末级  from 收入项目 where " & Where撤档时间() & _
            "  start with 上级ID is null  connect by prior ID =上级ID"
        blnRe = frmTreeLeafSel.ShowTree(strSQL, strID, strTemp, "收入项目")
        If blnRe Then
            On Error Resume Next
            If .RowData(lngRow) <> strID Then
                mcol价目(Index).Add 0, "C" & strID
                If Err <> 0 Then
                    MsgBox "该收入项目已设置了价目。", vbExclamation, gstrSysName
                    Exit Sub
                End If
                If .RowData(lngRow) > 0 Then mcol价目(Index).Remove "C" & .RowData(lngRow)
                .RowData(lngRow) = strID
            End If
            .TextMatrix(lngRow, mcstCol收费项目) = strTemp
            If .TextMatrix(lngRow, mcstCol附加手术收费率) = "" Then .TextMatrix(lngRow, mcstCol附加手术收费率) = "100.0"
            If .TextMatrix(lngRow, mcstCol原价) = "" Then .TextMatrix(lngRow, mcstCol原价) = "0.000"
            mblnChange = True
            mblnChanged价目(Index) = True
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh从属_CommandClick()
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim strTemp As String
    Dim strID As String
    Dim i As Integer
    Dim lngRow As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strReturn As String
    Dim strHyID As Long
    Dim strWherePriceGrade As String
    
    With msh从属
        '没设置收费类别就不能用
        lngRow = .Row '用变量保存
        strTemp = .TextMatrix(lngRow, 0)
        strID = .RowData(lngRow)
        
        If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
           strWherePriceGrade = " And d.价格等级 Is Null"
        Else
           strWherePriceGrade = "" & _
               " And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And d.价格等级 = [1])" & vbNewLine & _
               "      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And d.价格等级 = [2])" & vbNewLine & _
               "      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And d.价格等级 = [3])" & vbNewLine & _
               "      Or (d.价格等级 Is Null" & vbNewLine & _
               "          And Not Exists (Select 1" & vbNewLine & _
               "                          From 收费价目" & vbNewLine & _
               "                          Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
               "                                And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And 价格等级 = [1])" & vbNewLine & _
               "                                      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
               "                                      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And 价格等级 = [3])))))"
        End If
        strSQL = _
            "SELECT A.编码,A.名称,A.规格,A.计算单位," & _
            "       ltrim(rtrim(to_char(Sum(nvl(D.现价,0)),'9999999990.00'))) 价格,A.ID" & _
            " FROM 收费项目目录 A,收费价目 D" & _
            " WHERE A.ID=D.收费细目ID(+) And a.ID>0" & _
            "       And (A.撤档时间=to_date('3000-01-01','yyyy-mm-dd') or A.撤档时间 is null)" & _
            "       And D.执行日期 <= SYSDATE AND (D.终止日期 > SYSDATE OR D.终止日期 IS NULL)" & _
                    strWherePriceGrade & vbNewLine & _
            " Group By A.编码,A.名称,A.规格,A.计算单位,A.ID"

'            "SELECT DISTINCT C.名称 类别,B.名称 分类,A.编码,A.名称," & vbCrLf & _
'            "      A.规格,A.产地,A.计算单位,ltrim(rtrim(to_char(nvl(D.价格,0),'9999999990.00'))) 价格,A.ID" & vbCrLf & _
'            "  FROM 收费项目目录 A, 收费分类目录 B, 收费项目类别 C,(SELECT 收费细目ID,SUM(价格) AS 价格  FROM (" & vbCrLf & _
'            "        SELECT 收费细目ID,SUM(现价) AS 价格 FROM 收费价目 " & vbCrLf & _
'            "          WHERE 执行日期 <= SYSDATE AND (终止日期 > SYSDATE OR 终止日期 IS NULL) " & vbCrLf & _
'            "          GROUP BY  收费细目ID " & vbCrLf & _
'            "          UNION All " & vbCrLf & _
'            "        SELECT m.主项ID 收费细目ID,SUM(n.现价) AS 价格 FROM 收费价目 n ,收费从属项目 m " & vbCrLf & _
'            "         WHERE m.从项id = n.收费细目Id " & vbCrLf & _
'            "          AND  n.执行日期<=SYSDATE AND (n.终止日期> SYSDATE OR n.终止日期 IS null) " & vbCrLf & _
'            "          GROUP BY m.主项ID) GROUP BY 收费细目ID  ) D" & vbCrLf & _
'            " WHERE A.分类ID = B.ID(+) AND A.类别 = C.编码 AND  (A.撤档时间=to_date('3000-01-01','yyyy-mm-dd') or A.撤档时间 is null)  " & vbCrLf & _
'            "   AND A.ID = D.收费细目ID(+) AND " & Where撤档时间("A")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
        If rsTmp.RecordCount < 1 Then Exit Sub
        
        strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "编码,1000,0,2;名称,2500,0,2;规格,1500,0,2;单位,1500,0,2;价格,1000,1,2;ID,0,0,2", _
            "项目选择器", True, strTemp, 3, 1000 + 2500 + 1500 + 1500 + 1000 + 1800)
        If Trim(strReturn) = "" Then Exit Sub
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 And .RowData(i) = Split(strReturn, ",")(UBound(Split(strReturn, ","))) Then
                MsgBox "该收费项目已被作为从属项了。", vbExclamation, gstrSysName
                Exit Sub
            End If
            
        Next
        If Val(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) = Val(mstrID) And Val(mstrID) > 0 Then
            MsgBox "收费项目本身不能作为自己的从属项目。", vbExclamation, gstrSysName
            Exit Sub
        End If
        '递归检查
        strHyID = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
        If CheckHypotaxis(strHyID) = True Then
            MsgBox "该收费项目已存在从主关联不能再作为主从关联。", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        '如果是特殊项目，则从属项目的价格执行日期只能按日调价
        If mblnIsSpecialItem Then
            If Not IsRaiseByDate(Val(strHyID)) Then
                 MsgBox "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1) & "的价格调整不是按天来执行的，不能做为从属项目。", vbOKOnly + vbInformation, gstrSysName
                 Exit Sub
            End If
        End If
        
        .RowData(lngRow) = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
        .TextMatrix(lngRow, 0) = "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1)
        If .TextMatrix(lngRow, 1) = "" Then .TextMatrix(lngRow, 1) = "0"
        
        If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
           strWherePriceGrade = " And b.价格等级 Is Null"
        Else
           strWherePriceGrade = "" & _
               " And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And b.价格等级 = [2])" & vbNewLine & _
               "      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And b.价格等级 = [3])" & vbNewLine & _
               "      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And b.价格等级 = [4])" & vbNewLine & _
               "      Or (b.价格等级 Is Null" & vbNewLine & _
               "          And Not Exists (Select 1" & vbNewLine & _
               "                          From 收费价目" & vbNewLine & _
               "                          Where b.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
               "                                And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
               "                                      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
               "                                      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And 价格等级 = [4])))))"
        End If
        strSQL = "SELECT a.id,a.是否变价,sum(b.原价) 原价,sum(b.现价) 现价," & vbCrLf & _
                "       decode(nvl(a.是否变价,0),1,ltrim(rtrim(to_char(sum(b.原价),'9999999990.00')))||'～'||ltrim(rtrim(to_char(sum(b.现价),'9999999990.00'))),ltrim(rtrim(to_char(sum(b.现价),'9999999990.00'))))  AS  价格 " & vbCrLf & _
                " FROM 收费项目目录 a,收费价目 b " & vbCrLf & _
                " WHERE a.id=b.收费细目id AND  a.id=[1] " & vbCrLf & _
                "       And b.执行日期 <= SYSDATE AND (b.终止日期 > SYSDATE OR b.终止日期 IS NULL)" & _
                        strWherePriceGrade & vbNewLine & _
                "GROUP BY a.id,a.是否变价"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(lngRow)), gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
        
        If rsTmp.RecordCount > 0 Then
            .TextMatrix(lngRow, 3) = Trim(rsTmp!价格)
        End If
        mblnChange = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh价目_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With msh价目(Index)
        If .TxtVisible = False Then Exit Sub
        Select Case .Col
        Case mcstCol收费项目
            If IsRecord("收入项目", .Text, Index) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .Text = .TextMatrix(.Row, mcstCol收费项目)
            If .TextMatrix(.Row, mcstCol附加手术收费率) = "" Then .TextMatrix(.Row, mcstCol附加手术收费率) = "100.0"
        Case mcstCol原价, mcstCol现价, mcstCol缺省价格
            If chk变价.value = 1 And gstr医价接口编号 <> "" And gbln允许医价收费项目 = True Then
                Cancel = True
                Exit Sub
            End If
            If NumIsValid(.Text) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .Text = Format(Val(.Text), "###########0.000;-##########0.000;0.000;0.000")
            If .Col = mcstCol原价 Then
                If .Text = .TextMatrix(.Row, mcstCol现价) Then
                    Cancel = True
                    .TxtSetFocus
                End If
                If chk变价.value = 1 And Val(.TextMatrix(.Row, mcstCol现价)) <> 0 Then
                    If Val(.Text) > Val(.TextMatrix(.Row, mcstCol现价)) Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
                If chk变价.value = 1 And Val(.TextMatrix(.Row, mcstCol缺省价格)) <> 0 Then
                    If Val(.Text) > Val(.TextMatrix(.Row, mcstCol缺省价格)) Then
                        .TextMatrix(.Row, mcstCol缺省价格) = .Text
                    End If
                End If
            ElseIf .Col = mcstCol现价 Then
                If .Text = .TextMatrix(.Row, mcstCol原价) Then
                    If MsgBox("两个价格相同了，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                If chk变价.value = 1 And .TextMatrix(.Row, mcstCol原价) <> "" Then
                    If Val(.Text) < Val(.TextMatrix(.Row, mcstCol原价)) Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
                If chk变价.value = 1 And Val(.TextMatrix(.Row, mcstCol缺省价格)) <> 0 Then
                    If Val(.Text) < Val(.TextMatrix(.Row, mcstCol缺省价格)) Then
                        .TextMatrix(.Row, mcstCol缺省价格) = .Text
                    End If
                End If
            ElseIf .Col = mcstCol缺省价格 Then
                If Val(.Text) <> 0 Then
                    If Val(.Text) < Val(.TextMatrix(.Row, mcstCol原价)) Or Val(.Text) > Val(.TextMatrix(.Row, mcstCol现价)) Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
            End If
'            If chk变价.value = 1 And Not Cancel Then
'                If .Col = mcstCol原价 Then
'                    Me.txtEdit(text最低限价) = .Text
'                ElseIf .Col = mcstCol现价 Then
'                    Me.txtEdit(text最高限价) = .Text
'                End If
'            End If
        Case mcstCol附加手术收费率, mcstCol加班加价率
            If NumIsValid(.Text) = False Then
                Cancel = True
                Exit Sub
            End If
            .Text = Format(Val(.Text), "###########0.0;-##########0.0;0.0;0.0")
        End Select
    End With
    If Cancel = False Then mblnChange = True: mblnChanged价目(Index) = True
End Sub

Private Sub msh从属_EnterCell(Row As Long, Col As Long)
    Dim var列表 As Variant
    Dim lngCount As Long
    Dim i As Long
    
    On Error Resume Next
    '显示合计
    Me.lbl从属合计.Tag = 0
    For i = 0 To msh从属.Rows - 1
        Me.lbl从属合计.Tag = Me.lbl从属合计.Tag + Val(msh从属.TextMatrix(i, 1)) * Val(msh从属.TextMatrix(i, 3))
    Next
    Me.lbl从属合计.Caption = "合计:" & Format(Me.lbl从属合计.Tag, "0.00")
    On Error GoTo 0
    '设置固定关系
    var列表 = Split(mstr列表(Col + 1), ";")
    msh从属.Clear
    For lngCount = LBound(var列表) To UBound(var列表)
        msh从属.AddItem var列表(lngCount)
    Next
    If msh从属.ListCount = 0 Or Row = 0 Then Exit Sub
    If Row > 1 And msh从属.TextMatrix(Row - 1, Col) <> "" Then
        If msh从属.TextMatrix(Row, Col) = "" Then msh从属.TextMatrix(Row, Col) = msh从属.TextMatrix(Row - 1, Col)
    Else
        msh从属.ListIndex = 0
    End If
End Sub

Private Sub msh从属_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Long
    Dim strTmp As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With msh从属
        If msh从属.TxtVisible = False And .CboVisible = False Then
            If msh从属.Col = 0 And msh从属.TextMatrix(msh从属.Row, 0) = "" Then cmdOK.SetFocus
            Exit Sub
        End If
        Select Case msh从属.Col
        Case 0
            If IsRecord("收费项目目录", .Text) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .Text = .TextMatrix(.Row, 0)
        Case 1
            If NumIsValid(.Text) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            strTmp = .TextMatrix(.Row, 1)
            If .TextMatrix(.Row, 2) <> "0-不固定" And Val(.Text) = 0 Then
                .Text = strTmp
                Exit Sub
            End If
            
            .Text = Val(.Text)
            Me.lbl从属合计.Tag = 0
            For i = 0 To msh从属.Rows - 1
                If IsNumeric(msh从属.TextMatrix(i, 1)) And IsNumeric(msh从属.TextMatrix(i, 3)) Then
                    Me.lbl从属合计.Tag = Me.lbl从属合计.Tag + Val(msh从属.TextMatrix(i, 1)) * Val(msh从属.TextMatrix(i, 3))
                End If
            Next
            Me.lbl从属合计.Caption = "合计:" & Format(Me.lbl从属合计.Tag, "0.00")
        Case 2
            If .TextMatrix(.Row, 2) <> "0-不固定" And Val(.TextMatrix(.Row, 1)) = 0 Then
                .TextMatrix(.Row, 1) = "1"
            End If
        End Select
    End With
    If Cancel = False Then mblnChange = True
End Sub

Private Sub optApply_Click(Index As Integer)
    Dim i As Integer
    
    mblnChange = True
    For i = 1 To optApply.UBound
        If i = Index Then
            optApply(i).FontBold = True
        Else
            optApply(i).FontBold = False
        End If
    Next
End Sub

Private Sub optApply_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub opt科室_Click(Index As Integer)
    Dim sngLeft As Single
    
    '101736,服务对象包含住院的允许设置手工记帐缺省执行科室
    lbl一般情况.Caption = "1、除指定病人科室外："
    txt门诊执行.Enabled = False: txt门诊执行.BackColor = &H8000000F: txt门诊执行.Text = "": txt门诊执行.Tag = ""
    txt住院执行.Enabled = False: txt住院执行.BackColor = &H8000000F: txt住院执行.Text = "": txt住院执行.Tag = ""
    msf定向执行.Active = False: msf定向执行.Enabled = False
    msf定向执行.BackColorBkg = &H8000000F: msf定向执行.ClearBill
    
    If Index = 4 Then
        txt门诊执行.Enabled = True: txt门诊执行.BackColor = &H80000005
        txt住院执行.Enabled = True: txt住院执行.BackColor = &H80000005
        Select Case mstrServerObj
            Case "1"
                txt住院执行.Enabled = False: txt门诊执行.BackColor = &H8000000F
            Case "2"
                txt门诊执行.Enabled = False: txt住院执行.BackColor = &H8000000F
        End Select
        msf定向执行.Active = True: msf定向执行.Enabled = True
        msf定向执行.BackColorBkg = &H80000005
        
        '2010-05-10 解决无执行科室显示
        Ini性质分类
        load性质分类 0
    ElseIf Index = 0 Then
        '无明确执行科室
        If mstrServerObj <> "1" Then
            lbl一般情况.Caption = "1、手工记帐缺省执行科室设置："
            txt住院执行.Enabled = True: txt住院执行.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub opt科室_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub SetCodeNO()
    '设置编码
    On Error GoTo ErrHandle
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngMaxLen As Long
    Dim strTmp As String
    Dim strTmp1 As String
    
    '对于新增的要重新设置项目编码
    If medit方式 = EditNew Or medit方式 = EditCopy Then
        '先得到本级编码最大长度
        lngMaxLen = 2
        
        If mstr分类编码 = "" Then
            strSQL = "select max(length(编码)) from 收费项目目录 where " & IIF(Trim(mstr分类ID) = "" Or Trim(mstr分类ID) = "0", " 分类id is null ", "  分类id=[1] ")
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr分类ID))
        
            If rsTmp.RecordCount > 0 Then
                lngMaxLen = Nvl(rsTmp(0), 2)
            End If
            
            '再通过GetMax得到本级最大的编码+1
            strTmp = sys.MaxCode("收费项目目录", "编码", lngMaxLen, " where 分类id=" & mstr分类ID)
        
            strTmp1 = String(lngMaxLen, "0")
            RSet strTmp1 = strTmp
            strTmp = Replace(strTmp1, " ", "0")
            '判断该分类下没有没项目，如果没有就应初始加上分类编码
            strSQL = "select count(*) 项目数 from 收费项目目录 where 分类id=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr分类ID))
            
            If Nvl(rsTmp!项目数, 0) > 0 Then
                txtEdit(text编码).Text = strTmp
            Else
                txtEdit(text编码).Text = mstr分类编码 & strTmp
            End If
        Else
            '本分类下最大编码（按分类编码规则）
            strSQL = "select max(编码) as 最大编码 from 收费项目目录 where 分类id=[1] And 编码 Like [2] And Length(编码) > Length([3])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr分类ID), mstr分类编码 & "%", mstr分类编码)
            
            If Nvl(rsTmp!最大编码, "") = "" Then
                strTmp = mstr分类编码 & "01"
            Else
                strTmp = zlStr.Increase(rsTmp!最大编码)
            End If
            
            '检查其他分类下是否存在本分类的类似编码（主要是由于项目改变分类造成的）
            strSQL = "select max(编码) as 最大编码 from 收费项目目录 where 分类id<>[1] And 编码 Like [2] And Length(编码) > Length([3])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr分类ID), mstr分类编码 & "%", mstr分类编码)
            
            If Nvl(rsTmp!最大编码, "") <> "" Then
                '其他分类存在比本分类更大的编码
                If strTmp <= rsTmp!最大编码 Then
                    strTmp = zlStr.Increase(rsTmp!最大编码)
                End If
            End If
            
            txtEdit(text编码).Text = strTmp
        End If
        
        mstr编码 = txtEdit(text编码).Text
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmbClass_Click()
    On Error GoTo ErrHandle
    Dim strClass As String
    Dim i As Integer
    
    If Trim(cmbClass.Text) = "" Or InStr(cmbClass.Text, "-") < 1 Then Exit Sub
    Me.chk变价.Visible = True
    
    txtEdit(Text规格).BackColor = RGB(255, 255, 255)
    Me.chk屏蔽费别.Visible = True
    Me.chk加班加价.Visible = True
    Me.chk摘要.Visible = True
    Me.chk急诊.Visible = False
    Me.cmb护理.Visible = False
    Me.chk自动计算.Visible = False
    lblEdit(10).Visible = False
    cmb项目特性.Visible = False
    
    '规格
    Me.lblEdit(4).Enabled = True
    Me.txtEdit(2).Enabled = True
    '设置服务对象
    '服务对象
    Me.lblEdit(6).Enabled = True
    Me.cmb服务对象.Enabled = True
    If Me.cmb服务对象.ListCount > 3 Then
        Me.cmb服务对象.ListIndex = 3
    End If
    '类用类型
    Me.lblEdit(7).Enabled = True
    If InStr(1, mstrPrivs, ";医保类型;") = 0 Then
        cmb费用类型.Enabled = False
    Else
        cmb费用类型.Enabled = True
    End If
    
    '得到类型编码
    If mblnEditCancel = False Then
        mstr类别 = Left(Me.cmbClass.Text, 1)
    End If
    
    If TabExist("收费价目") = True Then
        For i = 0 To tbPriceGrade.ItemCount - 1
            If mstr类别 = "F" Then
                msh价目(i).TextMatrix(0, mcstCol附加手术收费率) = "附加手术收费率"
                If msh价目(i).ColWidth(mcstCol附加手术收费率) = 0 Then
                   msh价目(i).ColWidth(mcstCol附加手术收费率) = 1500
                End If
            Else
                msh价目(i).ColWidth(mcstCol附加手术收费率) = 0
            End If
        Next
    End If
    
    '设置编码
    Call SetCodeNO
    '严格检查类别是不是正确
    strClass = cmbClass.Text
    strClass = Trim(zlStr.NeedName(strClass))
    
    '显示当前应用哪个类别
    Me.optApply(3).Caption = "应用于" & IIF(mblnCanUpdateAll, "", "本院区") & " " & strClass & " 类别下所有项目(&U)"
    '产地
    If strClass = "材料" Then
        Me.lblEdit(18).Visible = True
        Me.lblEdit(18).Enabled = True
        Me.txtEdit(text产地).Visible = True
        Me.txtEdit(text产地).Enabled = True
        Call Form_Resize
    Else
        Me.lblEdit(18).Visible = False
        Me.lblEdit(18).Enabled = False
        Me.txtEdit(text产地).Visible = False
        Me.txtEdit(text产地).Enabled = False
        Call Form_Resize
    End If
    If strClass = "输血" Then
        lblEdit(10).Visible = True
        cmb项目特性.Visible = True
        If cmb服务对象.ListCount > 0 Then
            cmb服务对象.ListIndex = 0
        End If
    End If
    If Not (strClass = "挂号" Or strClass = "护理" Or strClass = "床位") Then
        Exit Sub
    End If
    '设置禁止输入的项目
    Me.chk变价.value = 0
    Me.chk变价.Visible = False
    Me.chk屏蔽费别.Visible = False
    Me.chk加班加价.value = 0
    Me.chk加班加价.Visible = False
    Me.chk摘要.Visible = False
    '规格
    Me.lblEdit(4).Enabled = False
    Me.txtEdit(2).Enabled = False
    Me.txtEdit(Text规格).BackColor = Me.BackColor
    '服务对象
    Me.lblEdit(6).Enabled = False
    Me.cmb服务对象.Enabled = False
    If Me.cmb服务对象.ListCount > 3 Then
        Me.cmb服务对象.ListIndex = 3
    End If
    Select Case strClass
    Case "挂号"
        '类用类型
        Me.lblEdit(7).Enabled = False
        Me.cmb费用类型.Enabled = False
        If Me.cmb费用类型.ListCount > 0 Then
            Me.cmb费用类型.ListIndex = 0
        End If
        Me.chk急诊.Visible = True
        If Me.cmb服务对象.ListCount > 1 Then
            Me.cmb服务对象.ListIndex = 1
        End If
        Exit Sub
    Case "护理"
        Me.cmb护理.Visible = True
        Me.chk自动计算.Visible = (Me.cmb护理.ListIndex <> 0)
        If InStr(1, mstrPrivs, ";医保类型;") = 0 Then
            cmb费用类型.Enabled = False
        End If
        Exit Sub
    Case "床位"
        '规格
        Me.lblEdit(4).Enabled = True
        Me.txtEdit(Text规格).Enabled = True
        txtEdit(Text规格).BackColor = RGB(255, 255, 255)
        If InStr(1, mstrPrivs, ";医保类型;") = 0 Then
            cmb费用类型.Enabled = False
        End If
        Exit Sub
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtEdit_Change(Index As Integer)
    Dim i As Integer
    
On Error GoTo ErrHandle
    mblnChange = True
    
    Select Case Index
    Case Text名称
        Dim strTmp As String
        '重新检查名称，并去 掉特殊字符
        strTmp = MoveSpecialChar(txtEdit(Text名称).Text)
        If txtEdit(Text名称).Text <> strTmp Then
            txtEdit(Text名称).Text = strTmp
            Me.txtEdit(text简码).Text = zlStr.GetCodeByORCL(strTmp, False, mlng简码长度)
            Me.txtEdit(text五笔).Text = zlStr.GetCodeByORCL(strTmp, True, mlng简码长度)
        End If
        txtEdit(text简码).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, False, mlng简码长度)
        txtEdit(text五笔).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, True, mlng简码长度)
    Case text标识主码, Text标识子码
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
        txtEdit(Index).SelStart = Len(txtEdit(Index).Text)
    Case Text备选码
        txtEdit(Index).SelStart = Len(txtEdit(Index).Text)
    Case Text分类
        '对于新增的要重新设置项目编码
        Call SetCodeNO
    Case text最高限价, text最低限价
        If IsNumeric(txtEdit(Index).Text) Then
            If text最高限价 = Index Then
                mdbl最高限价 = Val(txtEdit(Index).Text)
                If chk变价.value = 1 Then
                    For i = 0 To tbPriceGrade.ItemCount - 1
                        msh价目(i).TextMatrix(1, mcstCol现价) = Format(mdbl最高限价, "0.000")
                    Next
                End If
            Else
                mdbl最低限价 = Val(txtEdit(Index).Text)
                If chk变价.value = 1 Then
                    For i = 0 To tbPriceGrade.ItemCount - 1
                        msh价目(i).TextMatrix(1, mcstCol原价) = Format(mdbl最低限价, "0.000")
                    Next
                End If
            End If
        End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    mstrSel = 0
    Select Case Index
    Case Text名称, Text说明
        OS.OpenIme True
    Case text编码, text简码, Text建档时间, text最高限价, text最低限价
        OS.OpenIme False
    Case Text分类
        mstrSel = 1
    Case text标识主码
        OS.OpenIme False
        mstrSel = 2
    End Select
End Sub

Private Sub InitLvwSel()
    '初始化lvwSel控件
    lvwSel.View = lvwReport
    lvwSel.Visible = False
    lvwSel.GridLines = True
    lvwSel.FullRowSelect = True
    lvwSel.Width = 5000
    zlControl.LvwSelectColumns lvwSel, "编码,1000,0,2;名称,1500,0,2", True
    Select Case True
        Case mstrSel = 1
            lvwSel.Top = txtEdit(Text分类).Top + txtEdit(Text分类).Height + Screen.TwipsPerPixelY * 1
            lvwSel.Left = txtEdit(Text分类).Left
            lvwSel.Height = 1635
            lvwSel.Width = txtEdit(Text分类).Width
        Case mstrSel = 2
            lvwSel.Top = txtEdit(text标识主码).Top + txtEdit(text标识主码).Height + Screen.TwipsPerPixelY * 1
            lvwSel.Left = txtEdit(text标识主码).Left
            lvwSel.Width = 3200
            lvwSel.Height = 2500
    End Select
    lvwSel.Tag = False
    zlControl.LvwFlatColumnHeader lvwSel
End Sub

Private Sub lvwSel_LostFocus()
    lvwSel.Visible = False
    If mstrSel = 1 Then txtEdit(Text分类).SetFocus
    If mstrSel = 2 Then txtEdit(text标识主码).SetFocus
End Sub

Private Sub lvwSel_DblClick()
    lvwSel_KeyPress 13
End Sub

Private Sub lvwSel_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    '控件选择器
    On Error GoTo ErrHandle
    Select Case KeyAscii
    Case 13, Asc(" ")
        Select Case True
            Case mstrSel = 1
                If Not lvwSel.SelectedItem Is Nothing Then
                    mstr分类ID = lvwSel.SelectedItem.Tag
                    mstr分类编码 = lvwSel.SelectedItem.Text
                    txtEdit(Text分类).Text = lvwSel.SelectedItem.SubItems(1)
                    lvwSel.Visible = False
                    txtEdit(Text分类).SetFocus
                End If
                OS.PressKey vbKeyTab
            Case mstrSel = 2
                If Not lvwSel.SelectedItem Is Nothing Then
                    '先检查是不是有重复
                    If medit方式 <> EditNew And IsNumeric(mstrID) Then
                        strSQL = " SELECT 编码,名称 FROM  收费项目目录 WHERE UPPER(标识主码) = [1] AND ID<>[2] "
                    Else
                        strSQL = " SELECT 编码,名称 FROM  收费项目目录 WHERE UPPER(标识主码) = [1] "
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lvwSel.SelectedItem.Text, Val(mstrID))
                    
                    If rsTmp.RecordCount > 0 Then
                        strSQL = ""
                        rsTmp.MoveFirst
                        For i = 1 To rsTmp.RecordCount
                            If i = rsTmp.RecordCount Then
                                strSQL = strSQL & "[" & Nvl(rsTmp!编码) & "]" & Nvl(rsTmp!名称)
                            Else
                                strSQL = strSQL & "[" & Nvl(rsTmp!编码) & "]" & Nvl(rsTmp!名称) & vbCrLf
                            End If
                            rsTmp.MoveNext
                        Next
                        MsgBox "项目：“" & strSQL & "”已经使用该标准价格！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '开始取那个价格项目
                    txtEdit(text标识主码).Text = lvwSel.SelectedItem.Text
                    strSQL = "select 项目编码, 项目名称, 拼音码, 项目别名, 计价单位, 项目内涵, 除外内容, 项目说明, 项目价格, 重复标志, 医院等级, 注销标志, 财务编码, 最高限价, 最低限价, 调价日期 from 标准医价规范 where nvl(注销标志,0) =0 and  项目编码 = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(text标识主码).Text)
                    
                    If rsTmp.RecordCount = 1 Then
                        txtEdit(text标识主码).Text = Nvl(rsTmp!项目编码)
                        If medit方式 = EditNew Then
                            '名称
                            txtEdit(Text名称).Text = Nvl(rsTmp!项目名称)
                            txtEdit(text简码).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, False, mlng简码长度)
                            txtEdit(text五笔).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, True, mlng简码长度)
                            '单位
                            If Nvl(rsTmp!计价单位) <> "" Then
                                cmb计算单位.Text = Nvl(rsTmp!计价单位)
                            End If
                            '别名
                            If mshAlias.Rows > 2 And Trim(mshAlias.TextMatrix(mshAlias.Rows - 1, 0)) <> "" Then
                                mshAlias.Rows = mshAlias.Rows + 1
                            End If
                            mshAlias.TextMatrix(mshAlias.Rows - 1, 0) = Nvl(rsTmp!项目别名)
                            mshAlias.TextMatrix(mshAlias.Rows - 1, 1) = zlStr.GetCodeByORCL(Nvl(rsTmp!项目别名), False, mlng简码长度)
                            mshAlias.TextMatrix(mshAlias.Rows - 1, 2) = zlStr.GetCodeByORCL(Nvl(rsTmp!项目别名), True, mlng简码长度)
                            '最高与最低限价
                            txtEdit(text最高限价).Text = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                            txtEdit(text最低限价).Text = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                            If chk变价.value = 1 Then
                                For i = 0 To tbPriceGrade.ItemCount - 1
                                    msh价目(i).Rows = 2
                                    msh价目(i).TextMatrix(1, mcstCol现价) = txtEdit(text最高限价).Text
                                    msh价目(i).TextMatrix(1, mcstCol原价) = txtEdit(text最低限价).Text
                                Next
                            End If
                            mdbl最高限价 = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                            mdbl最低限价 = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                            '项目价格
                            mdbl医价价格 = Nvl(rsTmp!项目价格, 0)
                        ElseIf medit方式 = EditModify Then
                            '最高与最低限价
                            txtEdit(text最高限价).Text = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                            txtEdit(text最低限价).Text = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                            mdbl最高限价 = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                            mdbl最低限价 = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                            
                            If Not mblnShow收费价目 Then
                                TabMain.Tabs.Add , "_收费价目", "收费价目"
                                mblnShow收费价目 = True
                            End If
                            Call init价目
                            MsgBox "请重新确认收费价目。", vbInformation, gstrSysName
                        End If
                        OS.PressKey vbKeyTab
                        txtEdit(Text标识子码).SetFocus
                    End If
                End If
        End Select
    Case vbKeyEscape
        lvwSel.Visible = False
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnMatching As Boolean
    Dim i As Long
    Dim ObjItem As ListItem
    
    
    blnMatching = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0") = "0", True, False)
    Select Case Index
    Case Text分类   '分类
        If KeyCode = 13 Then
            KeyCode = 0
            strSQL = "Select ID,编码,名称 From 收费分类目录 Where Upper(名称) Like [1] or  Upper(编码) Like [2] Or Upper(Zlspellcode(名称)) Like [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(blnMatching = True, "%", "") & UCase(txtEdit(Text分类).Text) & "%", UCase(txtEdit(Text分类).Text) & "%")
            
            If rsTmp.RecordCount = 1 Then
                txtEdit(Text分类).Text = Nvl(rsTmp!名称)
                txtEdit(text编码).Text = Nvl(rsTmp!编码)
                mstr分类编码 = Nvl(rsTmp!编码)
                mstr分类ID = rsTmp!ID
                Call SetCodeNO
                txtEdit(text编码).MaxLength = mlng编码长度
                OS.PressKey vbKeyTab
            ElseIf rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                lvwSel.ListItems.Clear
                '初始化选择器
                Call InitLvwSel
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = lvwSel.ListItems.Add(, , Nvl(rsTmp!编码))
                    ObjItem.SubItems(1) = Nvl(rsTmp!名称)
                    ObjItem.Tag = rsTmp!ID
                    rsTmp.MoveNext
                Next
                lvwSel.ListItems(1).Selected = True
                lvwSel.SelectedItem.EnsureVisible
                lvwSel.Visible = True
                lvwSel.Enabled = True
                lvwSel.ZOrder
                lvwSel.SetFocus
            ElseIf Trim(txtEdit(Text分类).Text) = "无" Then
                txtEdit(Text分类).Text = "无"
                mstr分类ID = "0"
            Else
                strSQL = "Select ID,编码,名称 From 收费分类目录 Where Nvl(名称,'') = [1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(Text分类).Text)
                
                If rsTmp.RecordCount > 0 Then
                    txtEdit(Text分类).Text = Nvl(rsTmp!名称)
                    txtEdit(text编码).Text = Nvl(rsTmp!编码)
                    mstr分类ID = rsTmp!ID
                    Call SetCodeNO
                    txtEdit(text编码).MaxLength = mlng编码长度
                    OS.PressKey vbKeyTab
                Else
                    mstr分类ID = 0
                    txtEdit(Text分类).Text = ""
                End If
            End If
        End If
    Case text标识主码    '标识主码
        If KeyCode = 13 And txtEdit(text标识主码).Text = txtEdit(text标识主码).Tag Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        If KeyCode = 13 And gstr医价接口编号 <> "" And gbln允许医价收费项目 = True Then
            txtEdit(text最高限价).Enabled = False
            txtEdit(text最低限价).Enabled = False
            '先检查是不是有重复
            If medit方式 <> EditNew And IsNumeric(mstrID) Then
                strSQL = " SELECT 编码,名称 FROM  收费项目目录 WHERE UPPER(标识主码) = [1] AND ID<>[2] "
            Else
                strSQL = " SELECT 编码,名称 FROM  收费项目目录 WHERE UPPER(标识主码) = [1] "
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(Index).Text, Val(mstrID))
            
            If rsTmp.RecordCount > 0 Then
                strSQL = ""
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    If i = rsTmp.RecordCount Then
                        strSQL = strSQL & "[" & Nvl(rsTmp!编码) & "]" & Nvl(rsTmp!名称)
                    Else
                        strSQL = strSQL & "[" & Nvl(rsTmp!编码) & "]" & Nvl(rsTmp!名称) & vbCrLf
                    End If
                    rsTmp.MoveNext
                Next
                MsgBox "项目：“" & strSQL & "”已经使用该标准价格！", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(Index)
                Exit Sub
            End If
            '取那个价格项目
            strSQL = "select 项目编码, 项目名称, 拼音码, 项目别名, 计价单位, 项目内涵, 除外内容," & vbNewLine & _
                    "        项目说明, 项目价格, 重复标志, 医院等级, 注销标志, 财务编码, 最高限价, 最低限价, 调价日期" & vbNewLine & _
                    " from 标准医价规范" & vbNewLine & _
                    " where nvl(注销标志,0) =0 and  upper(项目编码) like [1] or upper(项目名称) LIKE [2] or upper(拼音码) LIKE [2] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(UCase(txtEdit(Index))) & "%", IIF(blnMatching = True, "%", "") & Trim(UCase(txtEdit(Index))) & "%")
            
            If rsTmp.RecordCount = 1 Then
                txtEdit(Index).Text = Nvl(rsTmp!项目编码)
                If medit方式 = EditNew Then
                    '名称
                    txtEdit(Text名称).Text = Nvl(rsTmp!项目名称)
                    txtEdit(text简码).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, False, mlng简码长度)
                    txtEdit(text五笔).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, True, mlng简码长度)
                    '单位
                    If Nvl(rsTmp!计价单位) <> "" Then
                        cmb计算单位.Text = Nvl(rsTmp!计价单位)
                    End If
                    '别名
                    If mshAlias.Rows > 2 And Trim(mshAlias.TextMatrix(mshAlias.Rows - 1, 0)) <> "" Then
                        mshAlias.Rows = mshAlias.Rows + 1
                    End If
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 0) = Nvl(rsTmp!项目别名)
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 1) = zlStr.GetCodeByORCL(Nvl(rsTmp!项目别名), False, mlng简码长度)
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 2) = zlStr.GetCodeByORCL(Nvl(rsTmp!项目别名), True, mlng简码长度)
                    '最高与最低限价
                    txtEdit(text最高限价).Text = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                    txtEdit(text最低限价).Text = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                    If chk变价.value = 1 Then
                        For i = 0 To tbPriceGrade.ItemCount - 1
                            msh价目(i).Rows = 2
                            msh价目(i).TextMatrix(1, mcstCol现价) = txtEdit(text最高限价).Text
                            msh价目(i).TextMatrix(1, mcstCol原价) = txtEdit(text最低限价).Text
                        Next
                    End If
                    mdbl最高限价 = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                    mdbl最低限价 = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                    '项目价格
                    mdbl医价价格 = Nvl(rsTmp!项目价格, 0)
                ElseIf medit方式 = EditModify Then
                    '最高与最低限价
                    txtEdit(text最高限价).Text = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                    txtEdit(text最低限价).Text = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                    If chk变价.value = 1 Then
                        For i = 0 To tbPriceGrade.ItemCount - 1
                            msh价目(i).Rows = 2
                            msh价目(i).TextMatrix(1, mcstCol现价) = txtEdit(text最高限价).Text
                            msh价目(i).TextMatrix(1, mcstCol原价) = txtEdit(text最低限价).Text
                        Next
                    End If
                    mdbl最高限价 = Format(Nvl(rsTmp!最高限价, 0), "0.00")
                    mdbl最低限价 = Format(Nvl(rsTmp!最低限价, 0), "0.00")
                End If
                If medit方式 = EditModify Then
                    If Not mblnShow收费价目 Then
                        TabMain.Tabs.Add , "_收费价目", "收费价目"
                        mblnShow收费价目 = True
                    End If
                    Call init价目
                    MsgBox "请重新确认收费价目。", vbInformation, gstrSysName
                End If
                
                OS.PressKey vbKeyTab
            ElseIf rsTmp.RecordCount > 1 Then
                KeyCode = 0
                lvwSel.ListItems.Clear
                '初始化选择器
                Call InitLvwSel
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = lvwSel.ListItems.Add(, , Nvl(rsTmp!项目编码))
                    ObjItem.SubItems(1) = Nvl(rsTmp!项目名称)
                    ObjItem.Tag = Nvl(rsTmp!项目编码)
                    rsTmp.MoveNext
                Next
                lvwSel.ListItems(1).Selected = True
                lvwSel.SelectedItem.EnsureVisible
                lvwSel.Visible = True
                lvwSel.Enabled = True
                lvwSel.ZOrder
                lvwSel.SetFocus
            Else
                KeyCode = 0
                MsgBox "不存在的“标识主码”！", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(Index)
            End If
        End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 4, 8
        OS.OpenIme False
    End Select
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandle
    mblnEditCancel = True
    Select Case True
    Case Index = text编码 Or Index = text简码 Or Index = text五笔 Or Index = Text分类 Or _
        Index = Text名称 Or Index = Text建档时间 Or Index = Text规格 Or _
        Index = Text备选码 Or Index = Text标识子码 Or Index = text最高限价 Or Index = text最低限价
'        ShowTab "基本信息"
        If Index = text最高限价 Or Index = text最低限价 Then
            If Trim(txtEdit(Index).Text) = "" Then txtEdit(Index).Text = 0
            If IsNumeric(txtEdit(Index).Text) = False Then
                Cancel = True
                mblnEditCancel = False
                MsgBox "请输入一个合法的价格！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(txtEdit(text最高限价).Text) <> 0 And Val(txtEdit(text最高限价).Text) < Val(txtEdit(text最低限价).Text) Then
                MsgBox "最高限价必须大于或等于最低限价！", vbInformation, gstrSysName
                Cancel = True
                mblnEditCancel = False
                Exit Sub
            End If
            '检查现价是否与限价冲突
            If Len(Trim(mstrID)) > 0 And (Val(txtEdit(text最高限价).Text) <> 0 Or Val(txtEdit(text最低限价).Text) <> 0) Then
                strSQL = "Select Max(现价) As 最高现价,Min(现价) As 最低现价 From 收费价目 Where" & _
                    " Decode(终止日期,to_date('3000-01-01','YYYY-MM-DD'),Null,终止日期) is Null And 收费细目ID =" & mstrID
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                If Not rsTmp.EOF Then
                    If Val(txtEdit(text最高限价).Text) <> 0 And Val(Nvl(rsTmp!最高现价)) > Val(txtEdit(text最高限价).Text) Then Cancel = True
                    If Val(txtEdit(text最低限价).Text) <> 0 And Val(Nvl(rsTmp!最低现价)) < Val(txtEdit(text最低限价).Text) Then Cancel = True
                    If Cancel Then
                        MsgBox "现行价格不在当前设置的限价内，请调价或重新设置限价！", vbInformation, gstrSysName
                        mblnEditCancel = False
                        Exit Sub
                    End If
                End If
            ElseIf Val(txtEdit(text最高限价).Text) <> 0 Or Val(txtEdit(text最低限价).Text) <> 0 Then
                For i = 0 To tbPriceGrade.ItemCount - 1
                    If Len(Trim(Me.msh价目(i).TextMatrix(1, mcstCol现价))) > 0 Then
                        If Val(txtEdit(text最高限价).Text) <> 0 And Val(Me.msh价目(i).TextMatrix(1, mcstCol现价)) > Val(txtEdit(text最高限价).Text) Then Cancel = True
                        If Val(txtEdit(text最低限价).Text) <> 0 And Val(Me.msh价目(i).TextMatrix(1, mcstCol现价)) < Val(txtEdit(text最低限价).Text) Then Cancel = True
                        If Cancel Then
                            MsgBox tbPriceGrade(i).Caption & "现行价格（" & Format(Me.msh价目(i).TextMatrix(1, mcstCol现价), "#0.000") & "）不在当前设置的限价内，请调价或重新设置限价！", vbInformation, gstrSysName
                            mblnEditCancel = False
                            Exit Sub
                        End If
                    End If
                Next
            End If
        End If
    Case Index = Text说明
'        ShowTab "收费价目"
    Case Index = text标识主码
        ShowTab "基本信息"
        Cancel = Not zlCommFun.StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength)
        If Not Cancel And (gstr医价接口编号 <> "" And gbln允许医价收费项目) Then
            '检查不是包含有非法字符
            If Trim(txtEdit(Index)) = "" Then
                Cancel = True
                MsgBox "“标识主码”不能为空！", vbInformation, gstrSysName
            Else
                strSQL = "select 1 from 标准医价规范 where 项目编码= [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(Index).Text)
                
                If rsTmp.RecordCount < 1 Then
                    Cancel = True
                    MsgBox "不存在的“标识主码”！", vbInformation, gstrSysName
                    txtEdit(text标识主码).Text = txtEdit(text标识主码).Tag
'                    zlControl.TxtSelAll txtEdit(Index)
                End If
            End If
        End If
        mblnEditCancel = False
        Exit Sub
    End Select
    mblnEditCancel = False
    If Index <> Text规格 And Index <> Text说明 And Index <> text产地 Then Cancel = Not zlCommFun.StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = text病案费目 Then
         If KeyAscii = vbKeyDelete Then
            txtEdit(text病案费目).Text = ""
            Exit Sub
        Else
            KeyAscii = 0
            Exit Sub
         End If
    End If
'    If InStr("~@%^&_|`'""/?", Chr(KeyAscii)) > 0 And _
'        index <> Text规格 And index <> Text说明 And index <> Text产地 Then KeyAscii = 0: Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 And _
        Index <> Text规格 And Index <> Text说明 And Index <> text产地 Then KeyAscii = 0: Exit Sub
    If (Index = 5) And KeyAscii = Asc("*") Then
        KeyAscii = 0
        cmd上级_Click
        Exit Sub
    ElseIf KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        KeyAscii = 0
        If (mstr类别 = "M" And Index = text产地) Or (mstr类别 <> "M" And Index = Text说明) Then
            If TabMain.Tabs.Count > 1 Then
                ShowTab "执行科室"
            Else
                cmdOK.SetFocus
            End If
        ElseIf Not (Index = Text分类 Or Index = text标识主码) _
            Or (Index = text标识主码 And gstr医价接口编号 = "") Then
            OS.PressKey vbKeyTab
        End If
    ElseIf Index = Text名称 Then
        Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("？")
        Case Asc("%")
            KeyAscii = Asc("％")
        Case Asc("_")
            KeyAscii = Asc("＿")
        End Select
        txtEdit(text简码).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, False, mlng简码长度)
        txtEdit(text五笔).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, True, mlng简码长度)
    ElseIf Index = text简码 Or Index = text五笔 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf Index = Text标识子码 Then
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    ElseIf Index = Text备选码 Then
        If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    ElseIf Index = text最高限价 Or Index = text最低限价 Then
        If InStr("0123456789.", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(text编码).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(text编码).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

Private Sub Form_Activate()
    
    On Error Resume Next
    
    If lblStationNo.Visible = False Then
        lbl病案费目.Left = Label1.Left
        txtEdit(text病案费目).Left = lbl病案费目.Left + lbl病案费目.Width + 50
        cmd病案.Left = txtEdit(text病案费目).Left + txtEdit(text病案费目).Width - cmd病案.Width
    Else
        lbl病案费目.Left = lblEdit(7).Left
        txtEdit(text病案费目).Left = lbl病案费目.Left + lbl病案费目.Width + 50
        cmd病案.Left = txtEdit(text病案费目).Left + txtEdit(text病案费目).Width - cmd病案.Width
    End If
    
    Select Case TabMain.SelectedItem.Caption
    Case "基本信息"
        If txtEdit(Text名称).Enabled And txtEdit(Text名称).Visible Then
            txtEdit(Text名称).SetFocus
        End If
    Case "收费价目"
        If msh价目(0).Visible And msh价目(0).Active Then
            msh价目(0).SetFocus
        End If
    Case "从属项目"
        If msh从属.Visible And msh从属.Active Then
            msh从属.SetFocus
        End If
    Case "执行科室"
        Dim i As Integer
        For i = 0 To 3
            If opt科室(i).value = True Then
                If opt科室(i).Visible And opt科室(i).Enabled Then
                    opt科室(i).SetFocus
                End If
                Exit Sub
            End If
        Next
    End Select
End Sub

Private Sub Form_Load()
    '个性化设置
    chk保留.value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "保留状态", "0"))
    With Me.msf定向执行
        .Active = True
        .MsfObj.ScrollBars = flexScrollBarVertical
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "执行科室ID": .TextMatrix(0, 1) = "执行科室"
        .TextMatrix(0, 2) = "病人科室ID": .TextMatrix(0, 3) = "病人科室"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 5: .ColData(3) = 1
        .ColWidth(0) = 0: .ColWidth(1) = 1550: .ColWidth(2) = 0: .ColWidth(3) = 3600
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1500
        .Add , "编码", "编码", 900
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    mlngFind = 1
    mblnOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mblnVerifyFlow = False
    mblnVerifyPris = False
    If mblnChange = False Then
        Exit Sub
    End If
    i = MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If i = vbNo Then
        Cancel = 1
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "保留状态", chk保留.value
End Sub

Private Sub tabMain_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If mblnEditCancel = True Then Exit Sub
    If Val(TabMain.Tag) = TabMain.SelectedItem.Index Then Exit Sub
    TabMain.Tag = TabMain.SelectedItem.Index
    
    fra(1).Visible = False
    fra(2).Visible = False
    fra(3).Visible = False
    fra(4).Visible = False
    On Error Resume Next
    Select Case TabMain.SelectedItem.Caption
    Case "收费价目"
        If mstrID = "" Then
            If Mid(cmbClass.Text, 1, 1) = "J" Or Mid(cmbClass.Text, 1, 1) = "H" Then
                For i = 0 To tbPriceGrade.ItemCount - 1
                    dtpBegin(i).CustomFormat = "yyyy年MM月dd日"
                    dtpBegin(i).Width = 1600
                    dtpBegin(i).value = DateAdd("d", 1, sys.Currentdate)
                    dtpBegin(i).MinDate = sys.Currentdate
                Next
                mstrCurrentDateFormat = "yyyy-mm-dd"
            Else
                For i = 0 To tbPriceGrade.ItemCount - 1
                    dtpBegin(i).CustomFormat = "yyyy年MM月dd日 HH:mm:ss"
                    dtpBegin(i).Width = 2535
                    dtpBegin(i).value = DateAdd("d", 1, sys.Currentdate)
                    dtpBegin(i).MinDate = sys.Currentdate
                Next
                mstrCurrentDateFormat = "yyyy-mm-dd hh:mm:ss"
            End If
        End If
        fra(2).Visible = True
        fra(2).ZOrder
        If msh价目(0).Active And msh价目(0).Visible Then
            msh价目(0).SetFocus
        End If
        
        '在编辑状态，如果启用医价系统并且选用了新的医价项目，则收费价目页调整后价格只能选择立即执行。
        If medit方式 = EditModify And (gstr医价接口编号 <> "" And gbln允许医价收费项目) Then
            If txtEdit(text标识主码).Text <> txtEdit(text标识主码).Tag Then
                For i = 0 To tbPriceGrade.ItemCount - 1
                    dtpBegin(i).Enabled = False
                    chkNow(i).value = 1
                Next
            End If
        End If
    Case "执行科室"
        fra(3).Visible = True
        fra(3).ZOrder
        
        mstrServerObj = ""
        If medit方式 = EditDept Then '单独执行科室时不能直接通过控件获取当前的服务对象
            gstrSQL = "select nvl(服务对象,0) as 服务对象 from 收费项目目录 where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "服务对象", mstrID)
            If rsTmp.RecordCount > 0 Then
                mstrServerObj = rsTmp!服务对象
            End If
        Else
            mstrServerObj = Mid(cmb服务对象.Text, 1, 1)
        End If
        
        For i = 0 To 3
            If opt科室(i).value = True And opt科室(i).Enabled And opt科室(i).Visible Then
                opt科室(i).SetFocus
                Exit Sub
            End If
        Next
    Case "从属项目"
        fra(4).Visible = True
        fra(4).ZOrder
        '处理没有找到焦点问(暂时这么处理)
'        If msh从属.Active And msh从属.Visible Then
'            msh从属.SetFocus
'        End If
    Case Else
        fra(1).Visible = True
        fra(1).ZOrder
        If txtEdit(text编码).Enabled And txtEdit(text编码).Visible Then
            txtEdit(text编码).SetFocus
        End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearContext(Optional ByVal bln完全 As Boolean = True)
    On Error GoTo ErrHandle
    Dim lngCol As Long, i As Integer
    
    mstrID = ""
    If Trim(mstr分类ID) = "0" Then mstr分类ID = ""
    '设置编码
    Call SetCodeNO
    
    If txtEdit(text编码).Text = "" Then txtEdit(text编码).Text = 1
    mstr编码 = txtEdit(text编码).Text
    txtEdit(Text名称).Text = ""
    txtEdit(text标识主码).Text = ""
    txtEdit(Text标识子码).Text = ""
    txtEdit(text最高限价).Text = ""
    txtEdit(text最低限价).Text = ""
    txtEdit(text简码).Text = ""
    txtEdit(text五笔).Text = ""
    
    txtEdit(Text规格).Text = ""
    mshAlias.Rows = 2
    mshAlias.TextMatrix(1, 0) = "": mshAlias.TextMatrix(1, 1) = "": mshAlias.TextMatrix(1, 2) = ""
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        For lngCol = 1 To mcol价目(i).Count
            mcol价目(i).Remove 1
        Next
        For lngCol = 1 To msh价目(i).Rows - 1
            If msh价目(i).RowData(lngCol) > 0 Then
                mcol价目(i).Add 0, "C" & msh价目(i).RowData(lngCol)
                msh价目(i).TextMatrix(lngCol, 1) = "0.000"
            End If
        Next
    Next
    
    mshAlias.Col = 0
    msh从属.Col = 0
    For i = 0 To tbPriceGrade.ItemCount - 1
        msh价目(i).Col = 0
    Next
    
    If bln完全 = False Then Exit Sub
    txtEdit(Text说明).Text = ""
    txtEdit(text产地).Text = ""
    cmb计算单位.Text = ""
    chk变价.value = 0
    chk加班加价.value = 0
    chk屏蔽费别.value = 0
    chk摘要.value = 0
    chk急诊.value = 0
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        For lngCol = 1 To mcol价目(i).Count
            mcol价目(i).Remove 1
        Next
        
        msh价目(i).ClearBill
        msh价目(i).Rows = 2
        msh价目(i).RowData(1) = 0
        For lngCol = 0 To msh价目(i).Cols - 1
            msh价目(i).TextMatrix(1, lngCol) = ""
        Next
        txt调价说明(i).Text = ""
    Next
    
    msh从属.ClearBill
    msh从属.Rows = 2
    msh从属.RowData(1) = 0
    For lngCol = 0 To msh从属.Cols - 1
        msh从属.TextMatrix(1, lngCol) = ""
    Next
    
    opt科室(0).value = True
    optApply(0).value = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function init基本() As Boolean
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockReadOnly
    rsTmp.CursorLocation = adUseClient
    
    mdbl最高限价 = 0
    mdbl最低限价 = 0
    
    mblnNotClick = True
    Call IniStationNo
    mblnNotClick = False
    
    mshAlias.Cols = 3
    mshAlias.ColAlignment(0) = 1
    mshAlias.ColAlignment(1) = 1
    mshAlias.ColAlignment(2) = 1
    mshAlias.ColWidth(0) = 1800
    mshAlias.ColWidth(1) = 1200
    mshAlias.ColWidth(2) = 1200
    mshAlias.TextMatrix(0, 0) = "别名"
    mshAlias.TextMatrix(0, 1) = "拼音简码"
    mshAlias.TextMatrix(0, 2) = "五笔简码"
    mshAlias.PrimaryCol = 0
    mshAlias.ColData(0) = 4 '文本框
    mshAlias.ColData(1) = 4 '文本框
    mshAlias.ColData(2) = 4 '文本框
    mshAlias.Rows = 2
    mshAlias.Active = True
    
    '初始化类别
    strSQL = "select 编码,名称 from 收费项目类别 where 编码<>'4' And 编码<>'5' and 编码<>'6' and 编码<>'7'"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    mblnEditCancel = True
    Me.cmbClass.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            cmbClass.AddItem Nvl(rsTmp!编码) & "-" & Nvl(rsTmp!名称)
            If i = 1 Then
                cmbClass.ListIndex = 0
            ElseIf Nvl(rsTmp!编码) = "E" Then
                cmbClass.ListIndex = cmbClass.NewIndex
            End If
            
            rsTmp.MoveNext
        Next
    End If
    mblnEditCancel = False
    
    txtTemp.Text = ""
    If Trim(mstr分类ID) = "0" Or Trim(mstr分类ID) = "" Then
        '顶级节点，包含类别信息
        mstr分类编码 = ""
        txtEdit(Text分类).Text = "无"
    Else
        '一般节点，直接从数库中读取
        strSQL = "select 编码,名称 from 收费分类目录 where ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr分类ID))
                
        mstr分类编码 = rsTmp("编码")
        txtEdit(Text分类).Text = rsTmp("名称")
    End If
    
    '取得上级编码，本级编码长度等值
    txtTemp.MaxLength = 0
    
    '设置项目特性
    cmb项目特性.Clear
    cmb项目特性.AddItem "0-血液"
    cmb项目特性.AddItem "1-附加"
    cmb项目特性.ListIndex = 0
    
    cmb服务对象.Clear
    cmb服务对象.AddItem "0-无"
    cmb服务对象.AddItem "1-门诊"
    cmb服务对象.AddItem "2-住院"
    cmb服务对象.AddItem "3-门诊与住院"
    cmb服务对象.ListIndex = 3
    
    '设置护理的设置项目
    cmb护理.Clear
    cmb护理.AddItem "0-一般项目"
    cmb护理.AddItem "1-护理等级"
    cmb护理.AddItem "2-基本护理等级"
    cmb护理.ListIndex = 0
    
    '设置费用确认项目
    cmb费用确认.Clear
    cmb费用确认.AddItem "0-不需要确认环节"
    cmb费用确认.AddItem "1-需要确认环节"
    cmb费用确认.ListIndex = 0
    
    '设置录入限量的应用范围
    With Me.cbo录入限量范围
        .Clear
        .AddItem "本项目"
        .AddItem "本级"
        .AddItem "本分类"
        .AddItem "本类别"
        .AddItem "所有"
        .ListIndex = 0
    End With
    
    '设置费用类型
    strSQL = "select 编码,名称,缺省标志 from 费用类型 where 性质<>'1' order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    cmb费用类型.Clear
    cmb费用类型.AddItem ""
    Do Until rsTmp.EOF
        cmb费用类型.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        If cmb费用类型.ListIndex = -1 And rsTmp("缺省标志") = 1 Then
            cmb费用类型.ListIndex = cmb费用类型.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If InStr(1, mstrPrivs, "医保类型") = 0 Then
        cmb费用类型.Enabled = False
    End If
    '取出用过的计算单位
    strSQL = "select distinct 计算单位 from 收费项目目录 where 类别=[1] and rownum<500"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别)
        
    cmb计算单位.Clear
    Do Until rsTmp.EOF
        If Not IsNull(rsTmp("计算单位")) Then
            cmb计算单位.AddItem rsTmp("计算单位")
        End If
        rsTmp.MoveNext
    Loop
    
    chk急诊.Visible = False
    cmb护理.Visible = False
    'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
    msh从属.RowData(1) = 0
    If medit方式 <> EditNew And IsNumeric(mstrID) Then
        strSQL = "select A.类别,b.名称 类别名称,A.编码,A.标识主码,A.标识子码,a.备选码, A.名称,A.规格,A.计算单位,A.费用类型,A.项目特性,A.服务对象" & _
        "    ,A.补充摘要,A.说明,A.产地,A.屏蔽费别,A.是否变价,A.加班加价,A.建档时间,A.最高限价,A.最低限价,A.录入限量,A.费用确认,A.站点,A.计算方式,a.病案费目 " & _
            " From 收费项目目录 A,收费项目类别 B  " & _
            " Where A.类别=B.编码 and  A.ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
        
        mblnEditCancel = True
        If Me.cmbClass.ListCount > 0 Then
            For i = 0 To Me.cmbClass.ListCount
                If Me.cmbClass.List(i) = Nvl(rsTmp!类别) & "-" & Nvl(rsTmp!类别名称) Then
                    Me.cmbClass.ListIndex = i
                    Exit For
                End If
            Next
        End If
        mblnEditCancel = False
        If medit方式 <> EditCopy Then
            txtEdit(text编码).Text = Mid(rsTmp("编码"), Len(txtTemp.Text) + 1)
            mstr编码 = rsTmp("编码")
        Else
            Call SetCodeNO
        End If
        '
        If gstr医价接口编号 <> "" And gbln允许医价收费项目 = True Then
            txtEdit(text最高限价).Enabled = False
            txtEdit(text最低限价).Enabled = False
        Else
            txtEdit(text最高限价).Visible = False
            txtEdit(text最低限价).Visible = False
            lblEdit(20).Visible = False
            lblEdit(21).Visible = False
        End If
        txtEdit(text最高限价).Text = Format(Nvl(rsTmp("最高限价"), 0), "0.00")
        txtEdit(text最低限价).Text = Format(Nvl(rsTmp("最低限价"), 0), "0.00")
        mdbl最高限价 = Format(Nvl(rsTmp("最高限价"), 0), "0.00")
        mdbl最低限价 = Format(Nvl(rsTmp("最低限价"), 0), "0.00")
        
        '求出包括子节点在内的最长编码
        txtEdit(text标识主码).Text = Nvl(rsTmp("标识主码"))
        txtEdit(text标识主码).Tag = Nvl(rsTmp("标识主码"))
        txtEdit(Text标识子码).Text = Nvl(rsTmp("标识子码"))
        txtEdit(Text备选码).Text = Nvl(rsTmp("备选码"))
        
        txtEdit(Text名称).Text = rsTmp("名称")
        txtEdit(Text规格).Text = Nvl(rsTmp("规格"))
        txtEdit(Text说明).Text = Nvl(rsTmp("说明"))
        txtEdit(text产地).Text = Nvl(rsTmp("产地"))
        txtEdit(Text建档时间).Text = Format(rsTmp("建档时间"), "yyyy-MM-dd")
        
        chk屏蔽费别.value = IIF(rsTmp("屏蔽费别") = 1, 1, 0)
        chk加班加价.value = IIF(rsTmp("加班加价") = 1, 1, 0)
        chk变价.Tag = IIF(rsTmp("是否变价") = 1, 1, 0)
        chk变价.value = IIF(rsTmp("是否变价") = 1, 1, 0)
        chk摘要.value = IIF(rsTmp("补充摘要") = 1, 1, 0)
        txt录入限量.Text = IIF(IsNull(rsTmp("录入限量")), "", rsTmp("录入限量"))
        
        mblnNotClick = True
        cbo.SeekIndex cmbStationNo, Nvl(rsTmp("站点"))
        cmbStationNo.Tag = Nvl(rsTmp("站点"))
        mblnNotClick = False
        
        chk自动计算.value = IIF(rsTmp("计算方式") = 1, 1, 0)
        txtEdit(text病案费目).Text = IIF(IsNull(rsTmp!病案费目), "", rsTmp!病案费目)
        
        Select Case rsTmp!类别
        Case "1"   '挂号
            chk急诊.value = IIF(rsTmp("项目特性") = 1, 1, 0)
            chk急诊.Visible = True
            cmb护理.Visible = False
            
            chk变价.Visible = False
            chk屏蔽费别.Visible = chk变价.Visible
            chk加班加价.Visible = chk变价.Visible
            chk摘要.Visible = chk变价.Visible
            txtEdit(Text规格).Enabled = chk变价.Visible
            txtEdit(Text规格).BackColor = Me.BackColor
            cmb服务对象.ListIndex = 1
            cmb服务对象.Enabled = chk变价.Visible
        Case "H"    '护理
            If IsNull(rsTmp!项目特性) = False Then
                cmb护理.ListIndex = rsTmp!项目特性
            End If
            cmb护理.Visible = True
            chk急诊.Visible = False
            
            chk变价.Visible = False
            chk屏蔽费别.Visible = chk变价.Visible
            chk加班加价.Visible = chk变价.Visible
            chk摘要.Visible = chk变价.Visible
            txtEdit(Text规格).Enabled = chk变价.Visible
            txtEdit(Text规格).BackColor = Me.BackColor
            cmb服务对象.Enabled = chk变价.Visible
        Case "J"    '床位
            chk变价.Visible = False
            chk屏蔽费别.Visible = chk变价.Visible
            chk加班加价.Visible = chk变价.Visible
            chk摘要.Visible = chk变价.Visible
            txtEdit(Text规格).Enabled = True
            cmb服务对象.Enabled = chk变价.Visible
        Case "K" '输血
            If IsNull(rsTmp!项目特性) = True Then
                cmb项目特性.ListIndex = 0
            Else
                cmb项目特性.ListIndex = rsTmp!项目特性
            End If
        End Select
        
        cmb计算单位.Text = IIF(IsNull(rsTmp("计算单位")), "", rsTmp("计算单位"))
        cmb服务对象.ListIndex = IIF(IsNull(rsTmp("服务对象")), 0, rsTmp("服务对象"))
        cmb费用确认.ListIndex = IIF(IsNull(rsTmp("费用确认")), 0, rsTmp("费用确认"))
        
        Call SetComboByText(cmb费用类型, IIF(IsNull(rsTmp("费用类型")), "", rsTmp("费用类型")), True)
        '得到别名表
        strSQL = "select 名称,nvl(码类,1) 码类,nvl(简码,'') 简码 From 收费项目别名 where 性质 in (1,9) And 收费细目ID=[1] order by 名称"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
                
        Dim blnYes As Boolean
        Do Until rsTmp.EOF
            If rsTmp("名称") = txtEdit(Text名称).Text Then
                If rsTmp!码类 = 1 Then
                    txtEdit(text简码).Text = IIF(IsNull(rsTmp("简码")), "", rsTmp("简码"))
                Else
                    txtEdit(text五笔).Text = IIF(IsNull(rsTmp("简码")), "", rsTmp("简码"))
                End If
            Else
                blnYes = False
                For i = 1 To mshAlias.Rows - 1
                    If mshAlias.TextMatrix(i, 0) = rsTmp!名称 Then
                        If rsTmp!码类 = 1 Or rsTmp!码类 = 2 Then
                            mshAlias.TextMatrix(i, rsTmp!码类) = rsTmp!简码
                        End If
                        blnYes = True
                    End If
                Next
                If blnYes = False Then
                    If Not (mshAlias.Rows = 2 And mshAlias.TextMatrix(1, 0) = "") Then
                        mshAlias.Rows = mshAlias.Rows + 1
                    End If
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 0) = rsTmp!名称
                    If rsTmp!码类 = 1 Or rsTmp!码类 = 2 Then
                        mshAlias.TextMatrix(mshAlias.Rows - 1, rsTmp!码类) = rsTmp!简码
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
        '得到当前收费价目条数
        
'        strSQL = "select a.ID " & _
'            " from 收费价目 A  Where decode(a.终止日期,to_date('3000-01-01','YYYY-MM-DD'),null,a.终止日期) is null And a.收费细目ID = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
'
'        msh价目.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
        '修改
    ElseIf Trim(mstrID) = "" Then    '新增
        If gstr医价接口编号 <> "" And gbln允许医价收费项目 = True Then
            txtEdit(text最高限价).Enabled = False
            txtEdit(text最低限价).Enabled = False
        Else
            txtEdit(text最高限价).Visible = False
            txtEdit(text最低限价).Visible = False
            lblEdit(20).Visible = False
            lblEdit(21).Visible = False
        End If
        
        strSQL = "select 病案费目 from 收费项目目录 a,(select max(建档时间) as 建档时间 from 收费项目目录 where 分类id=[1]) b where a.建档时间=b.建档时间 and a.分类id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病案费目", mstr分类ID)
        If rsTmp.RecordCount > 0 Then
            txtEdit(15).Text = IIF(IsNull(rsTmp!病案费目), "", rsTmp!病案费目)
        End If
        
        strSQL = "select 编码,名称 from 收费项目类别 where Upper(编码)=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(mstr类别)))
        
        mblnEditCancel = True
        If rsTmp.RecordCount > 0 Then
            If Me.cmbClass.ListCount > 0 Then
                For i = 0 To Me.cmbClass.ListCount
                    If Me.cmbClass.List(i) = Nvl(rsTmp!编码) & "-" & Nvl(rsTmp!名称) Then
                        Me.cmbClass.ListIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
        mblnEditCancel = False
        
        '设置编码
        Call SetCodeNO
        mstr编码 = txtEdit(text编码).Text
        txtEdit(Text建档时间).Text = Format(sys.Currentdate, "yyyy-MM-dd")
        txtEdit(text简码).Text = ""
        txtEdit(text五笔).Text = ""
'        dtpBegin.value = sys.Currentdate
        Select Case mstr类别
        Case "1"    '挂号
            chk急诊.Visible = True
            cmb护理.Visible = False
            
            chk变价.Visible = False
            chk屏蔽费别.Visible = chk变价.Visible
            chk加班加价.Visible = chk变价.Visible
            chk摘要.Visible = chk变价.Visible
            txtEdit(Text规格).Enabled = chk变价.Visible
            txtEdit(Text规格).BackColor = Me.BackColor
            cmb服务对象.ListIndex = 1
            cmb服务对象.Enabled = chk变价.Visible
        Case "H"    '护理
            cmb护理.Visible = True
            chk急诊.Visible = False
            
            chk变价.Visible = False
            chk屏蔽费别.Visible = chk变价.Visible
            chk加班加价.Visible = chk变价.Visible
            chk摘要.Visible = chk变价.Visible
            txtEdit(Text规格).Enabled = chk变价.Enabled
            txtEdit(Text规格).BackColor = Me.BackColor
            cmb服务对象.Enabled = chk变价.Visible
        Case "J"    '床位
            chk变价.Visible = False
            chk屏蔽费别.Visible = chk变价.Visible
            chk加班加价.Visible = chk变价.Visible
            chk摘要.Visible = chk变价.Visible
            txtEdit(Text规格).Enabled = True
            cmb服务对象.Enabled = chk变价.Visible
        End Select
    End If
    init基本 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub IniStationNo()
    Dim strSQL As String
    Dim rsRecord As ADODB.Recordset
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        On Error GoTo ErrHandle
        cmbStationNo.Clear
        cmbStationNo.Tag = ""
        strSQL = "select 编号,名称 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "站点查询")
        If mblnCanUpdateAll Then
            With cmbStationNo
                .AddItem ""
                Do While Not rsRecord.EOF
                    .AddItem rsRecord!编号 & "-" & rsRecord!名称
                    rsRecord.MoveNext
                Loop
            End With
        Else
            rsRecord.Filter = "编号='" & gstrNodeNo & "'"
            With cmbStationNo
                Do While Not rsRecord.EOF
                    .AddItem rsRecord!编号 & "-" & rsRecord!名称
                    rsRecord.MoveNext
                Loop
                If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = .NewIndex
            End With
        End If
        
'        With cmbStationNo
'            .Clear
'            .AddItem ""
'            .AddItem "0"
'            .AddItem "1"
'            .AddItem "2"
'            .AddItem "3"
'            .AddItem "4"
'            .AddItem "5"
'            .AddItem "6"
'            .AddItem "7"
'            .AddItem "8"
'            .AddItem "9"
'
'            .ListIndex = 0
'        End With
'    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function init价目(Optional ByVal blnStationChanged As Boolean) As Boolean
    On Error GoTo ErrHandle
    '功能:初始化收费价目表和从属项目表
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim str名称 As String
    Dim lngCol As Long
    Dim lng原有ID As Long
    Dim lngRow As Long, i As Integer
    Dim strWhere As String, objTabItem As TabControlItem
    Dim blnFind As Boolean
    
    If blnStationChanged = False Then
        With rsTmp
            strSQL = "select ID,名称,编码 from 收入项目 where 末级=1 and rownum<2"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
            
            If .RecordCount = 0 Then
                MsgBox "请在“收入项目管理”中添加收入项目后，再使用本功能。", vbExclamation, gstrSysName
                .Close
                Exit Function
            End If
            .Close
        End With
    End If
    
    '初始化页签
    If mblnCanUpdateAll Then
        '如果该收费项目是设置了站点的，则肯定只需要设置这个站点的价格等级即可
        If Not (medit方式 = EditNew Or medit方式 = EditCopy Or medit方式 = EditModify) Then
            strWhere = "       And Exists(Select 1 From 收费项目目录 Where ID = [1] And (站点 Is Null Or 站点 = b.站点))"
        ElseIf cmbStationNo.Text <> "" And blnStationChanged Then
            strWhere = "       And b.站点 =[3]"
        End If
        strSQL = "Select '00' As 编码, '缺省' As 价格等级 From Dual" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select Distinct a.编码, a.名称 As 价格等级" & vbNewLine & _
                " From 收费价格等级 A, 收费价格等级应用 B" & vbNewLine & _
                " Where a.名称 = b.价格等级 And Nvl(a.是否适用普通项目, 0) = 1" & vbNewLine & _
                        strWhere & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                " Order By 编码"
    Else
        strSQL = "Select '00' As 编码, '缺省' As 价格等级 From Dual" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select Distinct a.编码, a.名称 As 价格等级" & vbNewLine & _
                " From 收费价格等级 A, 收费价格等级应用 B" & vbNewLine & _
                " Where a.名称 = b.价格等级 And b.站点 = [2] And Nvl(a.是否适用普通项目, 0) = 1" & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                " Order By 编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取价格等级", Val(mstrID), gstrNodeNo, zlStr.NeedCode(cmbStationNo.Text, "-"))
    If rsTmp.RecordCount = 0 Then
        MsgBox "你没有调整该收费项目价格的权限！", vbInformation + vbOKOnly, gstrSysName
        Unload Me: Exit Function
    End If
        
    tbPriceGrade.RemoveAll
    lngRow = 0
    Do While Not rsTmp.EOF
        If lngRow > pic价格等级.UBound Then
            Load pic价格等级(lngRow): pic价格等级(lngRow).Visible = True
            Load msh价目(lngRow): Set msh价目(lngRow).Container = pic价格等级(lngRow): msh价目(lngRow).Visible = True
            Load lbl调价执行时间(lngRow): Set lbl调价执行时间(lngRow).Container = pic价格等级(lngRow): lbl调价执行时间(lngRow).Visible = True
            Load dtpBegin(lngRow): Set dtpBegin(lngRow).Container = pic价格等级(lngRow): dtpBegin(lngRow).Visible = True
            Load chkNow(lngRow): Set chkNow(lngRow).Container = pic价格等级(lngRow): chkNow(lngRow).Visible = True
            Load lbl调价说明(lngRow): Set lbl调价说明(lngRow).Container = pic价格等级(lngRow): lbl调价说明(lngRow).Visible = True
            Load txt调价说明(lngRow): Set txt调价说明(lngRow).Container = pic价格等级(lngRow): txt调价说明(lngRow).Visible = True
        End If
        Set objTabItem = tbPriceGrade.InsertItem(lngRow, Nvl(rsTmp!价格等级), pic价格等级(lngRow).hwnd, 0)
        lngRow = lngRow + 1
        rsTmp.MoveNext
    Loop
    If blnStationChanged Then Exit Function
    
    With tbPriceGrade.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003 '显示风格
        .BoldSelected = True '显示页标题字体加粗
        .ClientFrame = xtpTabFrameSingleLine '页面边框
        .Layout = xtpTabLayoutAutoSize
        .Position = xtpTabPositionBottom
        .StaticFrame = True
    End With
    If tbPriceGrade.ItemCount > 0 Then tbPriceGrade(0).Selected = True
    
    ReDim mcol价目(tbPriceGrade.ItemCount - 1)
    ReDim mblnNew(tbPriceGrade.ItemCount - 1)
    ReDim mblnChanged价目(tbPriceGrade.ItemCount - 1)
    For i = 0 To tbPriceGrade.ItemCount - 1
        Set mcol价目(i) = New Collection
        With msh价目(i)
            .Cols = mcstCols
            .ColWidth(mcstCol收费项目) = 1500
            .ColWidth(mcstCol原价) = 1000
            .ColWidth(mcstCol现价) = 1000
            .ColWidth(mcstCol缺省价格) = IIF(chk变价.value = 1, 1000, 0)
            .TextMatrix(0, mcstCol收费项目) = "收入项目"
            .TextMatrix(0, mcstCol原价) = "原价"
            .TextMatrix(0, mcstCol现价) = "现价"
            .TextMatrix(0, mcstCol缺省价格) = "缺省价格"
            If mstr类别 = "F" Then
                .TextMatrix(0, mcstCol附加手术收费率) = "附加手术收费率"
                .ColWidth(mcstCol附加手术收费率) = 1500
            Else
                .ColWidth(mcstCol附加手术收费率) = 0
            End If
            .TextMatrix(0, mcstCol加班加价率) = "加班加价率"
            .ColWidth(mcstCol加班加价率) = 0
            '对齐方式
            .ColAlignment(mcstCol收费项目) = 1
            .ColAlignment(mcstCol原价) = 7
            .ColAlignment(mcstCol现价) = 7
            .ColAlignment(mcstCol缺省价格) = 7
            .ColAlignment(mcstCol附加手术收费率) = 7
            .ColAlignment(mcstCol加班加价率) = 7
            '实现方式
            .ColData(mcstCol收费项目) = 1 '可以输入，且有一个按钮
            .ColData(mcstCol原价) = 5 '不允许选择
            .ColData(mcstCol现价) = 4 '直接输入
            .ColData(mcstCol缺省价格) = 4 '直接输入
            .ColData(mcstCol附加手术收费率) = 4
            .ColData(mcstCol加班加价率) = 4
            
            .PrimaryCol = 0
            .Active = True
        End With
        Me.dtpBegin(i).value = DateAdd("d", 1, Now)
    Next
    If mstrID = "" Then
        init价目 = True
        Exit Function
    End If
        
    '装入数据
    strSQL = "select 名称,是否变价,加班加价,最高限价,最低限价  from 收费项目目录 where ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
        
    If Not mblnShow收费价目 Then
        chk变价.Tag = IIF(rsTmp("是否变价") = 1, 1, 0)
        chk变价.value = IIF(rsTmp("是否变价") = 1, 1, 0)
        For i = 0 To tbPriceGrade.ItemCount - 1
            msh价目(i).ColWidth(mcstCol缺省价格) = IIF(rsTmp("是否变价") = 1, 1000, 0)
        Next
    End If
    chk加班加价.value = IIF(rsTmp("加班加价") = 1, 1, 0)
    If Not mblnShow收费价目 Then
        mdbl最高限价 = Nvl(rsTmp!最高限价, 0)
        mdbl最低限价 = Nvl(rsTmp!最低限价, 0)
    Else
        mdbl最高限价 = Val(txtEdit(text最高限价).Text)
        mdbl最低限价 = Val(txtEdit(text最低限价).Text)
    End If
    '根据具体数据改变列头
    Call chk变价_Click
    Call chk加班加价_Click
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        '显示收费价目
        strSQL = "Select a.ID,a.原价ID,a.收费细目ID,Nvl(a.原价,0) As 原价,Nvl(a.现价,0) As 现价," & vbNewLine & _
                "        Nvl(a.缺省价格,0) As 缺省价格,a.收入项目ID,b.名称,a.加班加价率,a.附术收费率," & vbNewLine & _
                "        a.变动原因,a.调价说明,a.执行日期,a.终止日期 " & vbNewLine & _
                " From 收费价目 A,收入项目 B" & vbNewLine & _
                " Where a.收入项目ID=b.ID And (a.终止日期 Is Null Or a.终止日期 = To_Date('3000-01-01','YYYY-MM-DD'))" & vbNewLine & _
                "       And a.收费细目ID = [1] " & IIF(tbPriceGrade(i).Caption = "缺省", " And a.价格等级 Is Null", " And a.价格等级=[2]")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID), tbPriceGrade(i).Caption)
            
        msh价目(i).Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
        msh价目(i).Tag = msh价目(i).Rows
        
        mblnNew(i) = rsTmp.RecordCount = 0 '新价格
        If rsTmp.RecordCount = 0 Then
            For lngCol = 0 To mcstCols - 1
                msh价目(i).TextMatrix(1, lngCol) = ""
            Next
            dtpBegin(i).value = sys.Currentdate
            If medit方式 = EditCopy Or medit方式 = EditNew Then
                txt调价说明(i).Text = "初始价格"
            Else
                txt调价说明(i).Text = ""
            End If
        Else
            lngCol = 1
            If medit方式 = EditCopy Then
                dtpBegin(i).value = Format(sys.Currentdate, "yyyy-mm-dd h:m:s")
            Else
                If mblnIsSpecialItem Then
                    dtpBegin(i).CustomFormat = "yyyy年MM月dd日"
                    dtpBegin(i).Width = 1600
                    mstrCurrentDateFormat = "yyyy-mm-dd"
                Else
                    dtpBegin(i).CustomFormat = "yyyy年MM月dd日 HH:mm:ss"
                    dtpBegin(i).Width = 2535
                    mstrCurrentDateFormat = "yyyy-mm-dd hh:mm:ss"
                End If
                
                If chk变价.value = 0 Then '1 不是变价项目
                    If mblnIsSpecialItem Then       '1.1 是特殊项目
                        If DateDiff("s", rsTmp("执行日期"), sys.Currentdate) > 0 Then        '1.1.1 上次开始时间小于当前时间
                            If DateDiff("s", rsTmp("执行日期"), Format(sys.Currentdate, "yyyy-mm-dd 00:00:00")) > 0 Then     '1.1.1.1 上次开始时间大于当天零点时间
                                chkNow(i).Visible = True
                                dtpBegin(i).MinDate = sys.Currentdate
                            Else        '1.1.1.2 上次开始时间小于当天零点时间
                                chkNow(i).Visible = False
                                dtpBegin(i).MinDate = DateAdd("d", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s"))
                            End If
                            dtpBegin(i).value = DateAdd("d", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s"))
                        Else        '1.1.2 上次开始时间大于当前时间
                            dtpBegin(i).value = DateAdd("d", 1, Format(rsTmp("执行日期"), "yyyy-mm-dd h:m:s"))
                            dtpBegin(i).MinDate = DateAdd("d", 1, Format(rsTmp("执行日期"), "yyyy-mm-dd h:m:s"))
                            chkNow(i).Visible = False
                        End If
                    Else        '1.2 不是特殊项目
                        If DateDiff("s", rsTmp("执行日期"), sys.Currentdate) > 0 Then        '1.2.1 上次开始时间小于当前时间
                            dtpBegin(i).value = Format(DateAdd("d", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s")), "yyyy-mm-dd 00:00:00")
                            dtpBegin(i).MinDate = DateAdd("s", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s"))
                        Else    '1.2.2 上次开始时间大于当前时间
                            dtpBegin(i).value = Format(DateAdd("d", 1, Format(rsTmp("执行日期"), "yyyy-mm-dd h:m:s")), "yyyy-mm-dd 00:00:00")
                            dtpBegin(i).MinDate = DateAdd("s", 1, Format(rsTmp("执行日期"), "yyyy-mm-dd h:m:s"))
                        End If
                        chkNow(i).Visible = True
                    End If
                    txt调价说明(i).Text = ""
                Else    '2 是变价项目
                    dtpBegin(i).value = Format(rsTmp("执行日期"), "yyyy-mm-dd h:m:s")
                    dtpBegin(i).Enabled = False
                End If
            End If
            Do Until rsTmp.EOF
                msh价目(i).TextMatrix(lngCol, mcstCol收费项目) = rsTmp("名称")
                If chk变价.value = 1 Then '变价
                    msh价目(i).TextMatrix(lngCol, mcstCol原价) = Format(rsTmp("原价"), "###########0.000;-##########0.000;0.000;0.000")
                    msh价目(i).TextMatrix(lngCol, mcstCol现价) = Format(rsTmp("现价"), "###########0.000;-##########0.000;0.000;0.000")
                    msh价目(i).TextMatrix(lngCol, mcstCol缺省价格) = Format(rsTmp("缺省价格"), "###########0.000;-##########0.000;0.000;0.000")
                Else
                    msh价目(i).TextMatrix(lngCol, mcstCol原价) = Format(rsTmp("现价"), "###########0.000;-##########0.000;0.000;0.000")
                    If medit方式 = EditCopy Then msh价目(i).TextMatrix(lngCol, mcstCol现价) = Format(rsTmp("现价"), "###########0.000;-##########0.000;0.000;0.000")
                End If
                msh价目(i).TextMatrix(lngCol, mcstCol附加手术收费率) = IIF(IsNull(rsTmp("附术收费率")), 0, rsTmp("附术收费率"))
                msh价目(i).TextMatrix(lngCol, mcstCol加班加价率) = IIF(IsNull(rsTmp("加班加价率")), 0, rsTmp("加班加价率"))
                msh价目(i).RowData(lngCol) = rsTmp("收入项目ID")
                lng原有ID = rsTmp("ID")
                mcol价目(i).Add lng原有ID, "C" & rsTmp("收入项目ID")
                lngCol = lngCol + 1
                rsTmp.MoveNext
            Loop
        End If
        If medit方式 = EditRaise Then
            msh价目(i).Col = 2
        End If
    Next
    
    If mblnCanUpdateAll = False Then
        strSQL = "Select 1 From 收费项目目录 Where ID = [1] And 站点 = [2] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断该项目是否属于当前站点", Val(mstrID), gstrNodeNo)
        If rsTmp.EOF Then
            '不是全院项目，不能调整缺省价格
            msh价目(0).MsfObj.Enabled = False
            msh价目(0).BackColor = vbButtonFace
            msh价目(0).BackColorBkg = vbButtonFace
            dtpBegin(0).Enabled = False
            chkNow(0).Enabled = False
            txt调价说明(0).Enabled = False
            txt调价说明(0).BackColor = vbButtonFace
            If tbPriceGrade.ItemCount > 1 Then tbPriceGrade.Item(1).Selected = True
        Else
            msh价目(0).MsfObj.Enabled = True
            msh价目(0).BackColor = vbWindowBackground
            msh价目(0).BackColorBkg = vbWindowBackground
            'dtpBegin(0).Enabled = True
            chkNow(0).Enabled = True
            txt调价说明(0).Enabled = True
            txt调价说明(0).BackColor = vbWindowBackground
        End If
    End If
    
    init价目 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub init从属()
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strWherePriceGrade As String
    
    With msh从属
        .Cols = 4
        .ColWidth(0) = 2000
        .ColWidth(1) = 800
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColAlignment(0) = 1
        .TextMatrix(0, 0) = "收费项目"
        .TextMatrix(0, 1) = "次数"
        .TextMatrix(0, 2) = "固定关系"
        .TextMatrix(0, 3) = "单价"
        .ColAlignment(2) = 1
        '实现方式
        .ColData(0) = 1 '表示该列可以输入，外部显示为按钮选择
        .ColData(1) = 4 '直接输入
        .ColData(2) = 3
        
        .PrimaryCol = 1
        .Active = True
    End With
    Me.lbl从属合计.Caption = ""
    Me.lbl从属合计.Tag = 0
    
    mstr列表(3) = "0-不固定;1-固定;2-按比例计算"
    '显示从属项目
    If mstrID = "" Then Exit Sub
    
    If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
       strWherePriceGrade = " And d.价格等级 Is Null"
    Else
       strWherePriceGrade = "" & _
           " And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And d.价格等级 = [2])" & vbNewLine & _
           "      Or (Instr(';4;', ';' || b.类别 || ';') > 0 And d.价格等级 = [3])" & vbNewLine & _
           "      Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And d.价格等级 = [4])" & vbNewLine & _
           "      Or (d.价格等级 Is Null" & vbNewLine & _
           "          And Not Exists (Select 1" & vbNewLine & _
           "                          From 收费价目" & vbNewLine & _
           "                          Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
           "                                And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
           "                                      Or (Instr(';4;', ';' || b.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
           "                                      Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And 价格等级 = [4])))))"
    End If
    gstrSQL = "Select a.主项ID,a.从项ID,a.固有从属,a.从项数次,b.名称,b.编码 项目编码,c.编码 ,c.名称 类别, " & vbCrLf & _
            "           decode(nvl(b.是否变价,0),1,ltrim(rtrim(to_char(sum(d.原价),'9999999990.00')))||'～'||ltrim(rtrim(to_char(sum(d.现价),'9999999990.00'))),ltrim(rtrim(to_char(sum(d.现价),'9999999990.00'))))  AS  价格 " & vbCrLf & _
            " From 收费从属项目 a,收费项目目录 b,收费项目类别 c ,收费价目 d " & vbCrLf & _
            " Where c.编码=b.类别 and  a.从项ID=b.id  and b.id=d.收费细目id  and 主项ID=[1] " & vbCrLf & _
            "       AND NVL (D.终止日期, TO_DATE ('3000-01-01', 'YYYY-MM-DD')) = TO_DATE ('3000-01-01', 'YYYY-MM-DD') " & _
                    strWherePriceGrade & vbNewLine & _
            " GROUP BY a.ROWID,a.主项ID,b.是否变价,a.从项ID,a.固有从属,a.从项数次,b.名称,b.编码 ,c.编码 ,c.名称 " & _
            " ORDER BY a.ROWID "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID), gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    
    msh从属.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
    If rsTmp.RecordCount = 0 Then
        For i = 0 To 3
            msh从属.TextMatrix(1, i) = ""
        Next
    Else
        i = 1
        Do Until rsTmp.EOF
            msh从属.TextMatrix(i, 0) = "[" & rsTmp("项目编码") & "]" & rsTmp("名称")
            msh从属.TextMatrix(i, 1) = rsTmp("从项数次")
            
            If rsTmp("固有从属") = 0 Then
                msh从属.TextMatrix(i, 2) = "0-不固定"
            ElseIf rsTmp("固有从属") = 2 Then
                msh从属.TextMatrix(i, 2) = "2-按比例计算"
            Else
                msh从属.TextMatrix(i, 2) = "1-固定"
            End If
            msh从属.TextMatrix(i, 3) = rsTmp("价格")
            msh从属.RowData(i) = rsTmp("从项ID")
            i = i + 1
            rsTmp.MoveNext
        Loop
        msh从属_EnterCell 1, 0
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub init执行()
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim lngSel As Long
    Dim strTmp As String
    
    opt科室_Click 0
    If mstrID = "" Then Exit Sub
    '显示科室
    gstrSQL = "select A.类别,A.ID,A.分类ID,A.执行科室,B.名称,C.名称  类别 from 收费项目目录 A,收费分类目录 B,收费项目类别 C where A.分类ID=B.ID(+) and A.类别=C.编码 and A.ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        
    '显示当前应用哪个类别
    Me.optApply(1).Caption = "应用于" & IIF(mblnCanUpdateAll, "", "本院区") & " " & Nvl(rsTmp("名称")) & " 分类同级的所有项目(&G)"
    Me.optApply(2).Caption = "应用于" & IIF(mblnCanUpdateAll, "", "本院区") & " " & Nvl(rsTmp("名称")) & " 分类下所有项目(&L)"
    Me.optApply(3).Caption = "应用于" & IIF(mblnCanUpdateAll, "", "本院区") & " " & Nvl(rsTmp("类别")) & " 类别下所有项目(&U)"
    lngSel = IIF(rsTmp("执行科室") < 7, rsTmp("执行科室"), 0)
    opt科室(lngSel).value = True
    opt科室_Click IIF(rsTmp("执行科室") < 7, rsTmp("执行科室"), 0)
    
    If opt科室(4).value = True Or opt科室(0).value = True Then
        '门诊住院执行科室
        gstrSQL = "select R.病人来源,E.ID,E.名称" & _
                " from 收费执行科室 R,部门表 E" & _
                " where R.执行科室ID=E.ID and R.病人来源 in (1,2) and R.开单科室id is null and R.收费细目ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
                
        With rsTmp
            Do While Not .EOF
                If !病人来源 = 1 Then Me.txt门诊执行.Text = !名称: Me.txt门诊执行.Tag = !ID
                If !病人来源 = 2 Then Me.txt住院执行.Text = !名称: Me.txt住院执行.Tag = !ID
                .MoveNext
            Loop
        End With
        
        If opt科室(4).value = True Then
            gstrSQL = _
            "select a.收费细目id 收费细目Id,a.病人来源," & vbCrLf & _
                "       b.id 开单id,b.编码 开单编码,b.名称 开单名称," & vbCrLf & _
                "       c.id 执行id,c.编码 执行编码,c.名称 执行名称  " & vbCrLf & _
                "  from 收费执行科室 a,部门表 b,部门表 c" & vbCrLf & _
                " where a.执行科室ID=c.id(+) And a.开单科室ID=b.id(+)  and a.病人来源 is null and a.收费细目ID=[1] and " & Where撤档时间("B") & vbCrLf & _
                " Order By c.名称"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
            
            Me.msf定向执行.ClearBill
            
            With rsTmp
                Do While Not .EOF
                    If strTmp <> !执行名称 Then
                        i = i + 1
                        Me.msf定向执行.Rows = i + 1
                        Me.msf定向执行.TextMatrix(i, 2) = IIF(IsNull(!开单ID), "（所有部门）", !开单ID)
                        Me.msf定向执行.TextMatrix(i, 3) = IIF(IsNull(!开单ID), "（所有部门）", "[" & !开单编码 & "]" & !开单名称)
                        Me.msf定向执行.TextMatrix(i, 0) = !执行ID
                        Me.msf定向执行.TextMatrix(i, 1) = "[" & !执行编码 & "]" & !执行名称
                    Else
                        Me.msf定向执行.TextMatrix(i, 2) = Me.msf定向执行.TextMatrix(i, 2) & "," & !开单ID
                        Me.msf定向执行.TextMatrix(i, 3) = Me.msf定向执行.TextMatrix(i, 3) & ",[" & !开单编码 & "]" & !开单名称
                    End If
                    strTmp = !执行名称
                    .MoveNext
                Loop
            End With
        End If
    End If
    
    '批量修改
    If medit方式 = 3 Then
        fra批量.Visible = True
    End If
    
    '取部门性质分类
    Ini性质分类
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowTab(ByVal strTab As String)
    '功能:显示指定页
    '参数:strTab 页名
    On Error Resume Next
    TabMain.Tabs("_" & strTab).Selected = True
    tabMain_Click
End Sub

Private Sub ShowItem(lst As ListItem)
    On Error GoTo ErrHandle
    '重新显示某一行,用于刷新
    Dim rsTmp As New ADODB.Recordset
    Dim lngCol  As Long
    Dim varValue As Variant
    
    rsTmp.CursorLocation = adUseClient
    gstrSQL = "Select A.ID,A.类别,A.编码,A.名称,A.规格,A.计算单位,A.费用类型," & vbCrLf & _
        " decode(A.服务对象,1,'门诊',2,'住院',3,'门诊与住院','无') as 服务对象,decode(A.补充摘要,1,'√','') as 补充摘要," & vbCrLf & _
        " decode(A.类别,'1',decode(A.项目特性,1,'急诊',''),'H',decode(A.项目特性,1,'护理等级',2,'基本护理', '')) 项目特性," & vbCrLf & _
        " A.说明,decode(A.屏蔽费别,1,'√','') as 屏蔽费别,decode(A.是否变价,1,'√','') as 是否变价,decode(A.加班加价,1,'√','') as 加班加价,A.执行科室," & vbCrLf & _
        " to_char(A.建档时间,'yyyy-mm-dd') as 建档时间,to_char(A.撤档时间,'yyyy-mm-dd') as 撤档时间," & vbCrLf & _
        " '" & txtEdit(Text分类).Text & "' as 所属分类 From 收费项目目录 A " & vbCrLf & _
        " Where A.ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        
    '根据ListView的列名从数据库取数
    lst.Text = rsTmp("名称")
    For lngCol = 2 To frmChargeManage.lvwMain_S.ColumnHeaders.Count
        varValue = rsTmp(frmChargeManage.lvwMain_S.ColumnHeaders(lngCol).Text).value
        lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
        lst.Tag = rsTmp!ID
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValid基本() As Boolean
    '功能:分析基本信息页的输入内容是否有效
    '参数:intTab 页号
    '返回值:有效返回True,否则为False
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim j As Long
    Dim str别名 As String
    
    IsValid基本 = False
    For i = 0 To 5
        strTemp = Trim(txtEdit(i).Text)
        If i <> Text规格 And i <> text产地 Then
        If zlCommFun.StrIsValid(txtEdit(i).Text, txtEdit(i).MaxLength) = False Then
            ShowTab "基本信息"
            If txtEdit(i).Enabled And txtEdit(i).Visible Then
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
            End If
            Exit Function
        End If
        End If
    Next
    '类别检查
    If InStr(cmbClass.Text, "-") > 0 Then
        strTemp = Left(cmbClass.Text, 1)
        strSQL = "select 编码 from 收费项目类别 where 编码<>'4' And 编码<>'5' and 编码<>'6' and 编码<>'7' and upper(编码) =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(strTemp)))
        
        If rsTmp.RecordCount < 1 Then
            ShowTab "基本信息"
            MsgBox "输入的类别不正确，请重新输入！", vbExclamation, gstrSysName
            cmbClass.SetFocus
            Exit Function
        Else
            mstr类别 = Nvl(rsTmp!编码)
        End If
    Else
        ShowTab "基本信息"
        If Trim(cmbClass.Text) = "" Then
            MsgBox "类别不能为空，请重新输入！", vbExclamation, gstrSysName
        Else
            MsgBox "类别不正确，请重新输入！", vbExclamation, gstrSysName
        End If
        If cmbClass.Visible And cmbClass.Enabled Then
            cmbClass.SetFocus
        End If
        Exit Function
    End If
    '分类检查
    If Trim(txtEdit(Text分类).Text) = "无" Or Trim(txtEdit(Text分类).Text) = "" Then
        txtEdit(Text分类).Text = "无"
        mstr分类ID = "0"
        MsgBox "分类不能为空，请重新输入！", vbExclamation, gstrSysName
        If txtEdit(Text分类).Visible And txtEdit(Text分类).Enabled Then
            txtEdit(Text分类).SetFocus
        End If
        Exit Function
    Else
        strSQL = "Select 1 From 收费分类目录 Where ID " & IIF(Trim(mstr分类ID) = "" Or Trim(mstr分类ID) = "0", " is null ", " = [2]") & " And 名称=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtEdit(Text分类).Text), Val(mstr分类ID))

        If rsTmp.RecordCount < 1 Then
            ShowTab "基本信息"
            MsgBox "分类输入错误，请重新输入！", vbExclamation, gstrSysName
            If txtEdit(Text分类).Visible And txtEdit(Text分类).Enabled Then
                txtEdit(Text分类).SetFocus
            End If
            Exit Function
        End If
    End If
    
    txtEdit(text编码).Text = Trim(txtEdit(text编码).Text)
    '计算单位
    If zlCommFun.StrIsValid(cmb计算单位.Text, mlng单位长度, , "计算单位") = False Then
        ShowTab "基本信息"
        If cmb计算单位.Enabled And cmb计算单位.Visible Then
            cmb计算单位.SetFocus
        End If
        Exit Function
    End If
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(text编码).Text) = 0 Then
            ShowTab "基本信息"
            txtEdit(text编码).SetFocus
            MsgBox "编码不能为空。", vbExclamation, gstrSysName
            Exit Function
        End If
    Else
        If Len(txtEdit(text编码).Text) < txtEdit(text编码).MaxLength Then
            ShowTab "基本信息"
            txtEdit(text编码).SetFocus
            MsgBox "编码的长度不够。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    If medit方式 = EditCopy Or medit方式 = EditNew Or medit方式 = EditModify Then
        gstrSQL = "select 类别,编码,名称 from 收费项目目录 where 编码=[1] " & IIF(medit方式 = EditCopy Or medit方式 = EditNew, "", " And ID <> [2] ")
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtTemp.Text) & txtEdit(text编码).Text, Val(mstrID))
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            strTemp = ""
            For j = 0 To rsTmp.RecordCount - 1
                strTemp = strTemp & "   [" & rsTmp!类别 & rsTmp!编码 & "]" & rsTmp!名称 & IIF(j = rsTmp.RecordCount - 1, "", vbCrLf)
                rsTmp.MoveNext
            Next
            ShowTab "基本信息"
            txtEdit(text编码).SetFocus
            MsgBox "编码与以下项目编码重复： " & vbCrLf & strTemp & vbCrLf & " 请重新输入其他编码！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '将标识码置为大写
    txtEdit(text标识主码).Text = UCase(txtEdit(text标识主码).Text)
    txtEdit(Text标识子码).Text = UCase(txtEdit(Text标识子码).Text)
'    txtEdit(Text备选码).Text = UCase(txtEdit(Text备选码).Text)
    If Len(Trim(txtEdit(text标识主码).Text)) < 1 And (gstr医价接口编号 <> "" And gbln允许医价收费项目) Then
        MsgBox "“标识主码”不允许为空！", vbExclamation, gstrSysName
        ShowTab "基本信息"
        If txtEdit(text标识主码).Enabled And txtEdit(text标识主码).Visible Then
            txtEdit(text标识主码).SetFocus
        End If
        Exit Function
    End If
    If Len(Trim(txtEdit(Text备选码).Text)) > 0 Then
        For i = 1 To Len(Trim(txtEdit(Text备选码).Text))
            If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(Trim(txtEdit(Text备选码).Text), i, 1)) < 1 Then
                MsgBox "备选码必须是由字母与数字组成。", vbExclamation, gstrSysName
                ShowTab "基本信息"
                If txtEdit(Text备选码).Enabled And txtEdit(Text备选码).Visible Then
                    txtEdit(Text备选码).SetFocus
                End If
                Exit Function
            End If
        Next
    End If
    If Len(Trim(txtEdit(Text名称).Text)) = 0 Then
        ShowTab "基本信息"
        txtEdit(Text名称).SetFocus
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(Text名称).Text = ""
        Exit Function
    End If
    If Len(Trim(txtEdit(text最低限价).Text)) = 0 Then txtEdit(text最低限价).Text = 0
    If Len(Trim(txtEdit(text最高限价).Text)) = 0 Then txtEdit(text最高限价).Text = 0
    For i = 1 To mshAlias.Rows - 1
        If Trim(mshAlias.TextMatrix(i, 0)) = Trim(txtEdit(Text名称).Text) Then
            ShowTab "基本信息"
            mshAlias.Row = i
            mshAlias.Col = 0
            If mshAlias.Active And mshAlias.Visible Then
                mshAlias.SetFocus
            End If
            MsgBox "别名与名称相同了。", vbExclamation, gstrSysName
            Exit Function
        End If
        For j = 1 To mshAlias.Rows - 1
            If Trim(mshAlias.TextMatrix(i, 0)) = Trim(mshAlias.TextMatrix(j, 0)) And i <> j Then
                ShowTab "基本信息"
                mshAlias.Row = j
                mshAlias.Col = 0
                If mshAlias.Active And mshAlias.Visible Then
                    mshAlias.SetFocus
                End If
                MsgBox "别名重复。", vbExclamation, gstrSysName
                Exit Function
            End If
        Next
    Next
    
    '检查别名字符串长度
    If Trim(txtEdit(text简码).Text) = "" Then
        txtEdit(text简码).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, False, mlng简码长度)
    End If

    If Trim(txtEdit(text五笔).Text) = "" Then
        txtEdit(text五笔).Text = zlStr.GetCodeByORCL(txtEdit(Text名称).Text, True, mlng简码长度)
    End If
    
    With mshAlias
        If Trim(txtEdit(text简码).Text) <> "" Then
            str别名 = "1''" & txtEdit(Text名称).Text & "''1''" & txtEdit(text简码).Text & "''"
        End If
        If Trim(txtEdit(text五笔).Text) <> "" Then
            str别名 = str别名 & "1''" & txtEdit(Text名称).Text & "''2''" & txtEdit(text五笔).Text & "''"
        End If
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) <> "" Then
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    str别名 = str别名 & "9''" & Trim(.TextMatrix(i, 0)) & "''1''" & Trim(.TextMatrix(i, 1)) & "''"
                End If
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    str别名 = str别名 & "9''" & Trim(.TextMatrix(i, 0)) & "''2''" & Trim(.TextMatrix(i, 2)) & "''"
                End If
            End If
        Next
    End With
    If LenB(str别名) > 4000 Then
        ShowTab "基本信息"
        If mshAlias.Active And mshAlias.Visible Then
            mshAlias.SetFocus
        End If
        MsgBox "别名字符串太长，请减少别名个数或者别名长度。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    IsValid基本 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid价目() As Boolean
    '功能:分析收费价目页的输入内容是否有效
    '参数:intTab 页号
    '返回值:有效返回True,否则为False
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim j As Integer, k As Integer
    Dim dbl合计价格 As Double
    Dim blnNothing As Boolean
    Dim blnHaveData As Boolean
    
    IsValid价目 = False
    For k = 0 To tbPriceGrade.ItemCount - 1
        If tbPriceGrade(k).Visible _
            And Not (mblnChanged价目(k) = False And (medit方式 = EditModify Or medit方式 = EditRaise)) Then   '隐藏的和未改变的不检查
            blnNothing = False
            With msh价目(k)
                If Trim(.TextMatrix(1, mcstCol收费项目)) = "" Then
                    If tbPriceGrade(k).Caption = "缺省" Or tbPriceGrade.ItemCount = 1 Then
                        ShowTab "收费价目"
                        tbPriceGrade.Item(k).Selected = True
                        If .Active And .Visible Then
                            .SetFocus
                        End If
                        .Row = 1
                        MsgBox "请为本收费项目设置价格。", vbExclamation, gstrSysName
                        Exit Function
                    Else
                        '除缺省外其它价格等级都可不设置收费项目
                        blnNothing = True
                    End If
                End If
                If blnNothing = False Then
                    For i = 1 To .Rows - 1
                        If .RowData(i) > 0 Then
                            For j = 1 To .Cols - 1
                                If Not IsNumeric(.TextMatrix(i, j)) And .ColWidth(j) > 0 Then
                                    ShowTab "收费价目"
                                    tbPriceGrade.Item(k).Selected = True
                                    If .Active And .Visible Then
                                        .SetFocus
                                    End If
                                    .Row = i: .Col = j
                                    MsgBox "收费价目第" & i & "行" & j + 1 & "列应输入数值。", vbExclamation, gstrSysName
                                    Exit Function
                                End If
                            Next
                            If Val(.TextMatrix(i, mcstCol现价)) < 0 Then
                                ShowTab "收费价目"
                                tbPriceGrade.Item(k).Selected = True
                                MsgBox "价格不允许为负数，请在第 " & i & " 行输入正确的价格。", vbExclamation, gstrSysName
                                Exit Function
                            End If
                            
                            '变价项目检查缺省价格
                            If Me.chk变价.value = 1 Then
                                If Val(.TextMatrix(i, mcstCol缺省价格)) > 0 Then
                                    If Val(.TextMatrix(i, mcstCol缺省价格)) < Val(.TextMatrix(i, mcstCol原价)) Or Val(.TextMatrix(i, mcstCol缺省价格)) > Val(.TextMatrix(i, mcstCol现价)) Then
                                        ShowTab "收费价目"
                                        tbPriceGrade.Item(k).Selected = True
                                        MsgBox "缺省价格应介于最低价和最高价之间，请在第 " & i & " 行输入正确的缺省价格。", vbExclamation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    If chk变价.value = 0 And gstr医价接口编号 <> "" And gbln允许医价收费项目 Then
                        For i = 1 To .Rows - 1
                            If .RowData(i) > 0 Then
                                dbl合计价格 = dbl合计价格 + Val(.TextMatrix(i, mcstCol现价))
                            End If
                        Next
                        
                        If dbl合计价格 > mdbl最高限价 Or dbl合计价格 < mdbl最低限价 Then
                            ShowTab "收费价目"
                            tbPriceGrade.Item(k).Selected = True
                            MsgBox "价格必须设定在最高限价(" & Format(mdbl最高限价, "0.00") & ")和最低限价(" & Format(mdbl最低限价, "0.00") & ")之间。", vbExclamation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End With
            
            If blnNothing = False Then
                If Me.chkNow(k).value = 0 Then
                    If DateDiff("s", sys.Currentdate, Me.dtpBegin(k)) < 0 Then
                        MsgBox "调价执行时间不能小于当前时间！", vbInformation, gstrSysName
                        Me.dtpBegin(k).value = DateAdd("n", 1, sys.Currentdate)
                        If TabMain.Tabs.Count > 1 Then
                            TabMain.Tabs(2).Selected = True
                        End If
                        tbPriceGrade.Item(k).Selected = True
                        If Me.dtpBegin(k).Enabled = True Then
                            Me.dtpBegin(k).SetFocus
                        End If
                        tabMain_Click
                        Exit Function
                    End If
                End If
                If zlCommFun.StrIsValid(txt调价说明(k).Text, txt调价说明(k).MaxLength) = False Then
                    ShowTab "收费价目"
                    tbPriceGrade.Item(k).Selected = True
                    If txt调价说明(k).Enabled And txt调价说明(k).Visible Then
                        txt调价说明(k).SetFocus
                        zlControl.TxtSelAll txt调价说明(k)
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
    IsValid价目 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid从属() As Boolean
    '功能:分析从属项目页的输入内容是否有效
    '参数:intTab 页号
    '返回值:有效返回True,否则为False
    On Error GoTo ErrHandle
    Dim i As Integer
    
    IsValid从属 = False
    With msh从属
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 Then
                If .TextMatrix(i, 1) = "" Then
                    ShowTab "从属项目"
                    If .Enabled And .Visible Then
                        .SetFocus
                    End If
                    .Row = i: .Col = 1
                    MsgBox "请输入次数。", vbExclamation, gstrSysName
                    Exit Function
                End If
                If .TextMatrix(i, 2) = "" Then
                    ShowTab "从属项目"
                    If .Enabled And .Visible Then
                        .SetFocus
                    End If
                    .Row = i: .Col = 2
                    MsgBox "请选择从属关系。", vbExclamation, gstrSysName
                    Exit Function
                End If
                If .TextMatrix(i, 2) <> "0-不固定" And Val(.TextMatrix(i, 1)) = 0 Then
                    ShowTab "从属项目"
                    If .Enabled And .Visible Then
                        .SetFocus
                    End If
                    .Row = i: .Col = 1
                    MsgBox "对于固定关系，其次数不能为0。", vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    IsValid从属 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NumIsValid(ByVal strNumber As String) As Boolean
    '功能:分析输入内容是否是有效的数字
    '参数:strNumber  输入内容
    '返回值:有效返回True,否则为False
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "请输入一个数值。", vbExclamation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) > 9999999999.999 Then
        MsgBox "这个数太大了。", vbExclamation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) < 0 Then
        MsgBox "不能为负数。", vbExclamation, gstrSysName
        Exit Function
    End If
    NumIsValid = True
End Function

Private Function IsRecord(ByVal strTable As String, ByVal strWhere As String, _
    Optional ByVal Index As Long) As Boolean
    '功能:分析输入内容是否是有效的数据库中表的记录
    '参数:strTable 表名;
    '     strWhere SQL语句的条件
    '     index  选择收入项目时传入
    '返回值:有效返回True,否则为False
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim strReturn As String '选择器返回字符串
    Dim strHyID As Long
    Dim strWherePriceGrade As String
    
    rsTmp.CursorLocation = adUseClient
    IsRecord = False
    If InStr(strWhere, "'") > 0 Then
        MsgBox "输入了非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If strTable = "收入项目" Then
        gstrSQL = "select 编码,名称,收据费目,病案费目,id from 收入项目 where 末级=1 and ( 编码 like [1] or 名称 like [1] or 简码 like [2] ) and " & Where撤档时间
    Else
        If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
            strWherePriceGrade = " And d.价格等级 Is Null"
        Else
            strWherePriceGrade = "" & _
                " And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And d.价格等级 = [4])" & vbNewLine & _
                "      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And d.价格等级 = [5])" & vbNewLine & _
                "      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And d.价格等级 = [6])" & vbNewLine & _
                "      Or (d.价格等级 Is Null" & vbNewLine & _
                "          And Not Exists (Select 1" & vbNewLine & _
                "                          From 收费价目" & vbNewLine & _
                "                          Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And 价格等级 = [4])" & vbNewLine & _
                "                                      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And 价格等级 = [5])" & vbNewLine & _
                "                                      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And 价格等级 = [6])))))"
        End If
        
        gstrSQL = _
            "SELECT A.编码,A.名称," & _
            "       A.规格,A.计算单位,ltrim(rtrim(to_char(Sum(nvl(D.现价,0)),'9999999990.00'))) 价格,A.ID" & _
            " FROM(Select Distinct a.类别,A.ID,A.编码,A.名称,A.规格,A.计算单位" & _
            "       From 收费项目目录 A,收费项目别名 B" & _
            "       WHERE A.ID = B.收费细目ID" & _
            "             And (A.撤档时间=to_date('3000-01-01','yyyy-mm-dd') or A.撤档时间 is null)" & _
            "             And (A.编码 like [1] or A.名称 like [1] or  ('['||A.编码||']'||A.名称  =[3])  or  B.简码 like [2])" & _
            "   ) A,收费价目 D" & _
            " Where A.ID=D.收费细目ID(+)" & _
            "       And D.执行日期 <= SYSDATE AND (D.终止日期 > SYSDATE OR D.终止日期 IS NULL)" & _
                    strWherePriceGrade & vbNewLine & _
            " Group By A.编码,A.名称,A.规格,A.计算单位,A.ID"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strWhere & "%", "%" & UCase(strWhere) & "%", strWhere, _
        gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    
    If rsTmp.RecordCount < 1 Then MsgBox "没有找到您查找的收费项目。", vbInformation, Me.Caption: Exit Function
    If rsTmp.RecordCount > 1 Then
        If strTable = "收入项目" Then
            strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "编码,800,0,2;名称,1500,0,2;收据费目,1200,0,2;病案费目,1200,0,2;ID,0,1,2", "收入项目选择器", True, , , 800 + 1500 + 1200 + 1200 + 800)
        Else
            strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "编码,1000,0,2;名称,1500,0,2;规格,1500,0,2;计算单位,800,0,2;价格,1000,1,2;ID,0,1,2", "收费项目选择器", True, , , 1000 + 1500 + 1500 + 800 + 800 + 2000)
        End If
        If Trim(strReturn) = "" Then
            Exit Function
        End If
    Else
        If strTable = "收入项目" Then
            strReturn = Nvl(rsTmp!编码) & "," & Nvl(rsTmp!名称) & "," & Nvl(rsTmp!收据费目) & "," & Nvl(rsTmp!病案费目) & "," & Nvl(rsTmp!ID, 0)
        Else
            strReturn = Nvl(rsTmp!编码) & "," & Nvl(rsTmp!名称) & "," & Nvl(rsTmp!规格) & "," & Nvl(rsTmp!计算单位) & "," & Nvl(rsTmp!价格) & "," & Nvl(rsTmp!ID, 0)
        End If
    End If
    If strTable = "收入项目" Then
        On Error Resume Next
        With msh价目(Index)
            strTemp = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
            If .RowData(.Row) <> strTemp Then
                mcol价目(Index).Add 0, "C" & strTemp
                If Err <> 0 Then
                    MsgBox "收入项目“" & Split(strReturn, ",")(1) & "”已设置了价目。", vbExclamation, gstrSysName
                    Exit Function
                End If
                If .RowData(.Row) > 0 Then mcol价目(Index).Remove "C" & .RowData(.Row)
                .RowData(.Row) = CLng(strTemp)
            End If
            .TextMatrix(.Row, mcstCol收费项目) = Split(strReturn, ",")(1)
            If .TextMatrix(.Row, mcstCol原价) = "" Then .TextMatrix(.Row, mcstCol原价) = "0.000"
        End With
    Else
        For i = 0 To msh从属.Rows - 1
            If msh从属.RowData(i) > 0 And msh从属.RowData(i) = CLng(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) And i <> msh从属.Row Then
                MsgBox "收费项目“" & Split(strReturn, ",")(1) & "”已被作为从属项了。", vbExclamation, gstrSysName
                Exit Function
            End If
        Next
        If Val(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) = Val(mstrID) And Val(mstrID) > 0 Then
            MsgBox "收费项目本身不能作为自己的从属项目。", vbExclamation, gstrSysName
            Exit Function
        End If
        '递归检查
        strHyID = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
        If CheckHypotaxis(strHyID) = True Then
            MsgBox "该收费项目已存在从主关联不能再作为主从关联。", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '如果是特殊项目，则从属项目的价格执行日期只能按日调价
        If mblnIsSpecialItem Then
            If Not IsRaiseByDate(Val(strHyID)) Then
                 MsgBox "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1) & "的价格调整不是按天来执行的，不能做为从属项目。", vbOKOnly + vbInformation, gstrSysName
                 Exit Function
            End If
        End If
        
        msh从属.RowData(msh从属.Row) = CLng(Split(strReturn, ",")(UBound(Split(strReturn, ","))))
        msh从属.TextMatrix(msh从属.Row, 0) = "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1)
        If msh从属.TextMatrix(msh从属.Row, 1) = "" Then
            msh从属.TextMatrix(msh从属.Row, 1) = "0"
            msh从属.TextMatrix(msh从属.Row, 2) = "0-不固定"
        End If
        If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
            strWherePriceGrade = " And b.价格等级 Is Null"
        Else
            strWherePriceGrade = "" & _
                " And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And b.价格等级 = [2])" & vbNewLine & _
                "      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And b.价格等级 = [3])" & vbNewLine & _
                "      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And b.价格等级 = [4])" & vbNewLine & _
                "      Or (b.价格等级 Is Null" & vbNewLine & _
                "          And Not Exists (Select 1" & vbNewLine & _
                "                          From 收费价目" & vbNewLine & _
                "                          Where b.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
                "                                      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
                "                                      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And 价格等级 = [4])))))"
        End If
        gstrSQL = "SELECT a.id,a.是否变价,sum(b.原价) 原价,sum(b.现价) 现价," & vbCrLf & _
                "         Decode(nvl(a.是否变价,0),1,ltrim(rtrim(to_char(sum(b.原价),'9999999990.00')))||'～'||ltrim(rtrim(to_char(sum(b.现价),'9999999990.00'))),ltrim(rtrim(to_char(sum(b.现价),'9999999990.00'))))  AS  价格 " & vbCrLf & _
                " FROM 收费项目目录 a,收费价目 b " & vbCrLf & _
                " WHERE a.id=b.收费细目id AND  a.id=[1] " & vbCrLf & _
                "       And b.执行日期 <= SYSDATE AND (b.终止日期 > SYSDATE OR b.终止日期 IS NULL)" & _
                        strWherePriceGrade & vbNewLine & _
                " GROUP BY a.id,a.是否变价"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(msh从属.RowData(msh从属.Row)), _
            gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
        
        If rsTmp.RecordCount > 0 Then
             msh从属.TextMatrix(msh从属.Row, 3) = Trim(rsTmp!价格)
        End If
    End If
    IsRecord = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "Select A.编码, A.计算单位, A.名称, A.标识主码, A.标识子码, A.最高限价, A.最低限价, A.规格, A.说明, A.产地, A.备选码, B.名称 别名, B.简码 " & _
            " From 收费项目目录 A, 收费项目别名 B " & _
            " Where A.ID = B.收费细目id And A.ID = 0 And B.码类 = 1 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng编码长度 = rsTmp.Fields("编码").DefinedSize
    mlng单位长度 = rsTmp.Fields("计算单位").DefinedSize
    mlng简码长度 = rsTmp.Fields("简码").DefinedSize
    mlng别名长度 = rsTmp.Fields("别名").DefinedSize
    
    txtEdit(text编码).MaxLength = mlng编码长度
    txtEdit(Text名称).MaxLength = rsTmp.Fields("名称").DefinedSize
    txtEdit(text标识主码).MaxLength = rsTmp.Fields("标识主码").DefinedSize
    txtEdit(Text标识子码).MaxLength = rsTmp.Fields("标识子码").DefinedSize
    txtEdit(text最高限价).MaxLength = rsTmp.Fields("最高限价").DefinedSize - 2
    txtEdit(text最低限价).MaxLength = rsTmp.Fields("最低限价").DefinedSize - 2
    txtEdit(Text规格).MaxLength = rsTmp.Fields("规格").DefinedSize
    txtEdit(Text说明).MaxLength = rsTmp.Fields("说明").DefinedSize
    txtEdit(text产地).MaxLength = rsTmp.Fields("产地").DefinedSize
    txtEdit(Text备选码).MaxLength = rsTmp.Fields("备选码").DefinedSize
    txtEdit(text简码).MaxLength = mlng简码长度
    txtEdit(text五笔).MaxLength = mlng简码长度
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt门诊执行_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    Dim ObjItem As ListItem
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txt门诊执行.Text) = "" Then Me.txt门诊执行.Tag = "": Me.txt门诊执行.Text = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txt门诊执行.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSQL = "select distinct ID,编码,名称" & _
            " from 部门表 D,部门性质说明 T" & _
            " where D.ID=T.部门ID and T.服务对象 in (1,2,3)" & _
            "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
            " order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp & "%")
    
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName:  Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt门诊执行.Tag = !ID: Me.txt门诊执行.Text = !名称: Call OS.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.TabMain.Left + Me.fra(3).Left + Me.txt门诊执行.Left - 130
        .Top = Me.TabMain.Top + Me.fra(3).Top + Me.txt门诊执行.Top + Me.txt门诊执行.Height - Me.Frame2.Top + 160
        
        lbl工作性质.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        
        .ZOrder 0: .Visible = True
    End With
    
    With Me.lvwItems
        .Tag = "门诊"
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt调价说明_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txt调价说明_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt住院执行_Change()
    If Trim(txt住院执行.Text) = "" Then
        txt住院执行.Tag = ""
    End If
End Sub

Private Sub txt住院执行_GotFocus()
    Me.txt住院执行.SelStart = 0: Me.txt住院执行.SelLength = 100
End Sub

Private Sub txt住院执行_KeyPress(KeyAscii As Integer)
    Dim ObjItem As ListItem
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txt住院执行.Text) = "" Then Me.txt住院执行.Tag = "": Me.txt住院执行.Text = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txt住院执行.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
   
    gstrSQL = "select distinct ID,编码,名称" & _
            " from 部门表 D,部门性质说明 T" & _
            " where D.ID=T.部门ID and T.服务对象 in (1,2,3)" & _
            "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
            " order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp & "%")
        
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt住院执行.Tag = !ID: Me.txt住院执行.Text = !名称: Call OS.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.TabMain.Left + Me.fra(3).Left + Me.txt住院执行.Left - 1300
        .Top = Me.TabMain.Top + Me.fra(3).Top + Me.txt住院执行.Top + Me.txt住院执行.Height - Me.Frame2.Top + 130
        
        lbl工作性质.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        
        .ZOrder 0: .Visible = True
    End With
    
    With Me.lvwItems
        .Tag = "住院"
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        .SetFocus
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckHypotaxis(HypotaxisID As Long) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''
    '功能           检查从属项目是否递归
    '参数
    '               hypotaxisID从属项目ID
    '返回           Flase=没有重复 True=有重复
    '''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "select 1 from 收费从属项目 where 主项ID= [1] "
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, HypotaxisID)
    
    If rsTmp.EOF = True Then
        CheckHypotaxis = False
    Else
        CheckHypotaxis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
