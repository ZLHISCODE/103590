VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmParFee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "费用参数设置"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   Icon            =   "frmParFee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8602.637
   ScaleMode       =   0  'User
   ScaleWidth      =   11910
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   587
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11910
      TabIndex        =   119
      Top             =   7980
      Width           =   11910
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   130
         Top             =   145
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   124
         Top             =   145
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   122
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11400
         TabIndex        =   121
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   10245
         TabIndex        =   120
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   131
         Top             =   165
         Width           =   4215
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病区查找(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   129
         Top             =   168
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参数查找(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   123
         Top             =   168
         Width           =   1095
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7980
      Left            =   0
      ScaleHeight     =   7980
      ScaleWidth      =   2415
      TabIndex        =   117
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   6060
         Left            =   2260
         MousePointer    =   9  'Size W E
         ScaleHeight     =   6060
         ScaleWidth      =   45
         TabIndex        =   118
         Top             =   120
         Width           =   45
      End
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   30
         ScaleHeight     =   6375
         ScaleWidth      =   2370
         TabIndex        =   125
         Top             =   0
         Width           =   2370
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5490
            Left            =   0
            TabIndex        =   126
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9684
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParFee.frx":6852
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   7005
         Left            =   0
         TabIndex        =   128
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   12356
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParFee.frx":10A52
      End
   End
   Begin TabDlg.SSTab stabDesign 
      Height          =   8205
      Left            =   2340
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   -15
      Visible         =   0   'False
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   14473
      _Version        =   393216
      Style           =   1
      Tabs            =   18
      Tab             =   11
      TabsPerRow      =   12
      TabHeight       =   520
      TabCaption(0)   =   "挂号"
      TabPicture(0)   =   "frmParFee.frx":1A13E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "预交"
      TabPicture(1)   =   "frmParFee.frx":1A15A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPar(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "一卡通"
      TabPicture(2)   =   "frmParFee.frx":1A176
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "医疗卡"
      TabPicture(3)   =   "frmParFee.frx":1A192
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(8)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "分诊"
      TabPicture(4)   =   "frmParFee.frx":1A1AE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(9)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "门诊划价"
      TabPicture(5)   =   "frmParFee.frx":1A1CA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(10)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "门诊收费"
      TabPicture(6)   =   "frmParFee.frx":1A1E6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "picPar(11)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "门诊记帐"
      TabPicture(7)   =   "frmParFee.frx":1A202
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "picPar(12)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "补结算"
      TabPicture(8)   =   "frmParFee.frx":1A21E
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "picPar(13)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "住院记帐"
      TabPicture(9)   =   "frmParFee.frx":1A23A
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "picPar(14)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "病人结帐"
      TabPicture(10)  =   "frmParFee.frx":1A256
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "picPar(15)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "公共"
      TabPicture(11)  =   "frmParFee.frx":1A272
      Tab(11).ControlEnabled=   -1  'True
      Tab(11).Control(0)=   "picPar(0)"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "财务监控"
      TabPicture(12)  =   "frmParFee.frx":1A28E
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "picPar(16)"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "医嘱附费"
      TabPicture(13)  =   "frmParFee.frx":1A2AA
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "picPar(17)"
      Tab(13).Control(0).Enabled=   0   'False
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "费用基础"
      TabPicture(14)  =   "frmParFee.frx":1A2C6
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "picPar(5)"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "操作员控制"
      TabPicture(15)  =   "frmParFee.frx":1A2E2
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "picPar(2)"
      Tab(15).ControlCount=   1
      TabCaption(16)  =   "记帐报警"
      TabPicture(16)  =   "frmParFee.frx":1A2FE
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "picPar(3)"
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "自动计算"
      TabPicture(17)  =   "frmParFee.frx":1A31A
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "picPar(4)"
      Tab(17).Control(0).Enabled=   0   'False
      Tab(17).ControlCount=   1
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7185
         Index           =   4
         Left            =   -75240
         ScaleHeight     =   7155
         ScaleWidth      =   9585
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   630
         Width           =   9615
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   15
            ItemData        =   "frmParFee.frx":1A336
            Left            =   1305
            List            =   "frmParFee.frx":1A338
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   6315
            Width           =   2880
         End
         Begin VB.OptionButton opt护理 
            Caption         =   "以价格最高的护理等级为标准"
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   10
            Top             =   5940
            Width           =   2670
         End
         Begin VB.OptionButton opt护理 
            Caption         =   "以最后一次护理等级为标准"
            Height          =   255
            Index           =   0
            Left            =   2775
            TabIndex        =   9
            Top             =   5970
            Value           =   -1  'True
            Width           =   2625
         End
         Begin VB.CheckBox chk 
            Caption         =   "修正上期自动计费(表示是否自动修改上一核算期间的自动费用计算数据。)"
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   7
            Top             =   5715
            Width           =   6510
         End
         Begin VB.TextBox txtDateInput 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   6
            Top             =   1560
            Visible         =   0   'False
            Width           =   1380
         End
         Begin ZL9BillEdit.BillEdit Bill 
            Height          =   5235
            Index           =   0
            Left            =   4965
            TabIndex        =   5
            Top             =   420
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   9234
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAutoCalc 
            Height          =   5235
            Left            =   90
            TabIndex        =   3
            Top             =   420
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   9234
            _Version        =   393216
            Cols            =   3
            RowHeightMin    =   315
            BackColorBkg    =   -2147483643
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.Frame fraAutoCharge 
            Caption         =   "自动记帐模式"
            Height          =   975
            Left            =   90
            TabIndex        =   756
            Top             =   6375
            Width           =   9360
            Begin VB.CheckBox chk 
               Caption         =   "下午算半天模式 (指以半天为计算单位,上午入院算1天,下午算半天,上午出院当天不算费用,下午算半天)"
               Height          =   225
               Index           =   43
               Left            =   135
               TabIndex        =   12
               Top             =   405
               Width           =   8775
            End
            Begin VB.Label lblAutoChargeNM 
               AutoSize        =   -1  'True
               Caption         =   "自动记帐规则说明："
               Height          =   180
               Left            =   120
               TabIndex        =   757
               Top             =   330
               Width           =   1620
            End
         End
         Begin VB.Label lbl护理 
            AutoSize        =   -1  'True
            Caption         =   "同天不同护理等级的护理费计算"
            Height          =   180
            Left            =   90
            TabIndex        =   8
            Top             =   6000
            Width           =   2520
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "按病区对指定费用进行自动计算"
            Height          =   180
            Index           =   13
            Left            =   5070
            TabIndex        =   4
            Top             =   120
            Width           =   2520
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "对床位费或护理费进行自动计算"
            Height          =   180
            Index           =   12
            Left            =   165
            TabIndex        =   2
            Top             =   135
            Width           =   2520
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7095
         Index           =   2
         Left            =   -74805
         ScaleHeight     =   7065
         ScaleWidth      =   8970
         TabIndex        =   13
         Top             =   705
         Visible         =   0   'False
         Width           =   9000
         Begin VB.CommandButton cmdOperate 
            Caption         =   "增加(&A)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   0
            Left            =   8310
            TabIndex        =   16
            Top             =   405
            Width           =   1100
         End
         Begin VB.CommandButton cmdOperate 
            Caption         =   "修改(&M)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   1
            Left            =   8310
            TabIndex        =   17
            Top             =   885
            Width           =   1100
         End
         Begin VB.CommandButton cmdOperate 
            Caption         =   "删除(&D)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   2
            Left            =   8310
            TabIndex        =   18
            Top             =   1365
            Width           =   1100
         End
         Begin VB.CommandButton cmdOperate 
            Caption         =   "清除(&L)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   3
            Left            =   8310
            TabIndex        =   19
            Top             =   1845
            Width           =   1100
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2070
            MaxLength       =   12
            TabIndex        =   24
            Top             =   6705
            Width           =   1350
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   6255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   11033
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
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
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "历史天数"
               Object.Width           =   2187
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "允许操作他人单据"
               Object.Width           =   2893
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "金额上限"
               Object.Width           =   2187
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "按操作员对不同单据的操作权限，针对单据的历史天数和最初操作人进行限制"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   6120
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单笔费用最大提醒金额："
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   6765
            Width           =   1980
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7335
         Index           =   5
         Left            =   -74850
         ScaleHeight     =   7305
         ScaleWidth      =   9060
         TabIndex        =   32
         Top             =   690
         Width           =   9090
         Begin VB.CommandButton cmd公费费用类型 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   5880
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.CommandButton cmd公费费用类型 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   4950
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.CommandButton cmd医保费用类型 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   8280
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.CommandButton cmd医保费用类型 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   7350
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Height          =   2790
            Index           =   1
            Left            =   4950
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   510
            Width           =   2220
         End
         Begin VB.Frame fra特定收费项目 
            Caption         =   " 特定收费项目 "
            Height          =   2265
            Left            =   360
            TabIndex        =   33
            Top             =   240
            Width           =   3495
            Begin VB.CommandButton cmdSelect 
               Caption         =   "…"
               Height          =   240
               Index           =   1
               Left            =   2775
               TabIndex        =   643
               TabStop         =   0   'False
               Top             =   780
               Width           =   255
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "…"
               Height          =   240
               Index           =   0
               Left            =   2775
               TabIndex        =   642
               TabStop         =   0   'False
               Top             =   315
               Width           =   255
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "…"
               Height          =   240
               Index           =   3
               Left            =   2775
               TabIndex        =   641
               TabStop         =   0   'False
               Top             =   1245
               Width           =   255
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "…"
               Height          =   240
               Index           =   4
               Left            =   2775
               TabIndex        =   640
               TabStop         =   0   'False
               Top             =   1710
               Width           =   255
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   0
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   285
               Width           =   1710
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   1
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   750
               Width           =   1710
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   3
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   1215
               Width           =   1710
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   4
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   1680
               Width           =   1710
            End
            Begin VB.Label lbl 
               Caption         =   "工本费"
               Height          =   225
               Index           =   7
               Left            =   630
               TabIndex        =   36
               Top             =   795
               Width           =   585
            End
            Begin VB.Label lbl 
               Caption         =   "病历费"
               Height          =   225
               Index           =   6
               Left            =   630
               TabIndex        =   34
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lbl 
               Caption         =   "普通配置费"
               Height          =   225
               Index           =   18
               Left            =   285
               TabIndex        =   38
               Top             =   1245
               Width           =   930
            End
            Begin VB.Label lbl 
               Caption         =   "肿瘤配置费"
               Height          =   225
               Index           =   56
               Left            =   285
               TabIndex        =   40
               Top             =   1725
               Width           =   930
            End
         End
         Begin VB.Frame fra票据 
            Caption         =   "票据号码控制"
            Height          =   3855
            Left            =   360
            TabIndex        =   42
            Top             =   2880
            Width           =   3495
            Begin VB.TextBox txtUD 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Index           =   4
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   44
               Text            =   "7"
               Top             =   480
               Width           =   300
            End
            Begin VB.CheckBox chk 
               Caption         =   "严格控制"
               Height          =   285
               Index           =   13
               Left            =   1860
               TabIndex        =   46
               Top             =   495
               Width           =   1020
            End
            Begin MSComCtl2.UpDown ud 
               Height          =   300
               Index           =   4
               Left            =   1440
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   7
               BuddyControl    =   "txtUD(4)"
               BuddyDispid     =   196647
               BuddyIndex      =   4
               OrigLeft        =   1350
               OrigTop         =   240
               OrigRight       =   1590
               OrigBottom      =   540
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   2805
               Index           =   0
               Left            =   240
               TabIndex        =   47
               Top             =   855
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   4948
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
                  Object.Width           =   1765
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "号码长度"
                  Object.Width           =   1588
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "严格控制"
                  Object.Width           =   1588
               EndProperty
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "号码长度"
               Height          =   180
               Index           =   19
               Left            =   240
               TabIndex        =   43
               Top             =   540
               Width           =   720
            End
         End
         Begin VB.ListBox lst 
            Height          =   2160
            Index           =   3
            Left            =   4920
            Style           =   1  'Checkbox
            TabIndex        =   57
            Top             =   4440
            Width           =   1665
         End
         Begin VB.ListBox lst 
            Height          =   2790
            Index           =   0
            Left            =   7245
            Style           =   1  'Checkbox
            TabIndex        =   53
            Top             =   495
            Width           =   2220
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "医保病人适用费用类型"
            Height          =   180
            Index           =   20
            Left            =   7350
            TabIndex        =   52
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "公费病人适用费用类型"
            Height          =   180
            Index           =   21
            Left            =   4950
            TabIndex        =   48
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "刷卡要求输密码"
            Height          =   180
            Index           =   41
            Left            =   4920
            TabIndex        =   56
            Top             =   4080
            Width           =   1260
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7035
         Index           =   17
         Left            =   -74895
         ScaleHeight     =   7005
         ScaleWidth      =   9015
         TabIndex        =   626
         TabStop         =   0   'False
         Top             =   705
         Visible         =   0   'False
         Width           =   9045
         Begin VB.Frame fra 
            Caption         =   "记帐后发药方式设置"
            Height          =   870
            Index           =   10
            Left            =   135
            TabIndex        =   636
            Top             =   3315
            Width           =   4530
            Begin VB.OptionButton optSendDrugFF 
               Caption         =   "选择是否发药"
               Height          =   180
               Index           =   2
               Left            =   2895
               TabIndex        =   639
               Top             =   450
               Width           =   1470
            End
            Begin VB.OptionButton optSendDrugFF 
               Caption         =   "不发药"
               Height          =   180
               Index           =   0
               Left            =   210
               TabIndex        =   637
               Top             =   450
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optSendDrugFF 
               Caption         =   "自动发药"
               Height          =   180
               Index           =   1
               Left            =   1515
               TabIndex        =   638
               Top             =   450
               Width           =   1470
            End
         End
         Begin VB.Frame fra 
            Caption         =   "药品显示单位"
            Height          =   810
            Index           =   9
            Left            =   135
            TabIndex        =   633
            Top             =   2325
            Width           =   4530
            Begin VB.OptionButton optDrugUnitFF 
               Caption         =   "售价单位"
               Height          =   180
               Index           =   0
               Left            =   195
               TabIndex        =   634
               Top             =   405
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optDrugUnitFF 
               Caption         =   "门诊/住院单位"
               Height          =   180
               Index           =   1
               Left            =   1515
               TabIndex        =   635
               Top             =   405
               Width           =   1470
            End
         End
         Begin VB.Frame fra 
            Height          =   2010
            Index           =   8
            Left            =   135
            TabIndex        =   627
            Top             =   165
            Width           =   4530
            Begin VB.CheckBox chk 
               Caption         =   "备货卫材只能输入内部条码"
               Height          =   195
               Index           =   187
               Left            =   195
               TabIndex        =   632
               Top             =   1545
               Width           =   2865
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药房库存"
               Height          =   195
               Index           =   166
               Left            =   195
               TabIndex        =   630
               Top             =   930
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药库库存"
               Height          =   195
               Index           =   167
               Left            =   195
               TabIndex        =   631
               Top             =   1230
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "变价允许输入数次"
               Height          =   195
               Index           =   165
               Left            =   195
               TabIndex        =   628
               Top             =   345
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "中药可以输入付数"
               Height          =   195
               Index           =   164
               Left            =   210
               TabIndex        =   629
               Top             =   630
               Value           =   1  'Checked
               Width           =   1740
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6705
         Index           =   16
         Left            =   -74865
         ScaleHeight     =   6675
         ScaleWidth      =   9165
         TabIndex        =   608
         TabStop         =   0   'False
         Top             =   765
         Visible         =   0   'False
         Width           =   9195
         Begin VB.CheckBox chk 
            Caption         =   "预交轧帐按门诊或住院分别轧帐"
            Height          =   375
            Index           =   198
            Left            =   435
            TabIndex        =   786
            Top             =   5355
            Width           =   3060
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   2
            Left            =   1410
            TabIndex        =   784
            Top             =   5070
            Width           =   3810
         End
         Begin VB.CheckBox chk 
            Caption         =   "借出款项后打印借款单"
            Height          =   375
            Index           =   162
            Left            =   435
            TabIndex        =   625
            Top             =   4500
            Width           =   3060
         End
         Begin VB.CheckBox chk 
            Caption         =   "借款申请后打印借款单"
            Height          =   375
            Index           =   161
            Left            =   435
            TabIndex        =   624
            Top             =   4200
            Width           =   2265
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   1
            Left            =   1170
            TabIndex        =   623
            Top             =   3945
            Width           =   4065
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   0
            Left            =   1185
            TabIndex        =   620
            Top             =   2970
            Width           =   4065
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   7
            Left            =   1185
            TabIndex        =   610
            Top             =   435
            Width           =   4065
         End
         Begin VB.Frame fraPrintModeDraw 
            Caption         =   "备用金领用单打印方式"
            Height          =   810
            Left            =   435
            TabIndex        =   615
            Top             =   1845
            Width           =   4620
            Begin VB.OptionButton optPrintModeDraw 
               Caption         =   "不打印(&1)"
               Height          =   300
               Index           =   0
               Left            =   180
               TabIndex        =   616
               Top             =   315
               Width           =   1230
            End
            Begin VB.OptionButton optPrintModeDraw 
               Caption         =   "自动打印"
               Height          =   300
               Index           =   1
               Left            =   1455
               TabIndex        =   617
               Top             =   315
               Value           =   -1  'True
               Width           =   1305
            End
            Begin VB.OptionButton optPrintModeDraw 
               Caption         =   "选择是否打印"
               Height          =   300
               Index           =   2
               Left            =   2880
               TabIndex        =   618
               Top             =   315
               Width           =   1650
            End
         End
         Begin VB.Frame fraPrintModeSJ 
            Caption         =   "收款收据打印方式"
            Height          =   840
            Left            =   435
            TabIndex        =   611
            Top             =   765
            Width           =   4620
            Begin VB.OptionButton optPrintModeSJ 
               Caption         =   "选择是否打印"
               Height          =   180
               Index           =   2
               Left            =   2880
               TabIndex        =   614
               Top             =   420
               Width           =   1710
            End
            Begin VB.OptionButton optPrintModeSJ 
               Caption         =   "不打印"
               Height          =   180
               Index           =   0
               Left            =   180
               TabIndex        =   612
               Top             =   420
               Width           =   1185
            End
            Begin VB.OptionButton optPrintModeSJ 
               Caption         =   "自动打印"
               Height          =   180
               Index           =   1
               Left            =   1455
               TabIndex        =   613
               Top             =   420
               Value           =   -1  'True
               Width           =   1365
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "领用票据时,必须进行签字确认."
            Height          =   180
            Index           =   160
            Left            =   435
            TabIndex        =   621
            Top             =   3315
            Width           =   3720
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "收费员轧账管理"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   785
            Top             =   5010
            Width           =   1260
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "人员借款管理"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   622
            Top             =   3885
            Width           =   1080
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "票据使用监控"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   619
            Top             =   2910
            Width           =   1080
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "收费财务监控"
            Height          =   180
            Index           =   3
            Left            =   135
            TabIndex        =   609
            Top             =   375
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7785
         Index           =   0
         Left            =   -180
         ScaleHeight     =   7755
         ScaleWidth      =   9570
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   690
         Visible         =   0   'False
         Width           =   9600
         Begin VB.CheckBox chk 
            Caption         =   "指定发料部门时不显示无库存卫材"
            Height          =   195
            Index           =   204
            Left            =   390
            TabIndex        =   80
            Top             =   3735
            Width           =   3060
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用免挂号模式"
            Height          =   195
            Index           =   290
            Left            =   390
            TabIndex        =   75
            Top             =   2445
            Width           =   3120
         End
         Begin VB.CheckBox chk 
            Caption         =   "同一身份证只能对应一个建档病人"
            Height          =   195
            Index           =   28
            Left            =   390
            TabIndex        =   76
            Top             =   2715
            Width           =   3120
         End
         Begin MSComCtl2.DTPicker dtpRegistTime 
            Height          =   300
            Left            =   1935
            TabIndex        =   68
            Top             =   1320
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   167903235
            UpDown          =   -1  'True
            CurrentDate     =   42804
         End
         Begin VB.Frame fra 
            Caption         =   "挂号公共"
            Height          =   1200
            Index           =   11
            Left            =   4470
            TabIndex        =   677
            Top             =   6510
            Width           =   4995
            Begin VB.CheckBox chk 
               Caption         =   "预约排队按时点显示"
               Height          =   195
               Index           =   186
               Left            =   180
               TabIndex        =   116
               Top             =   885
               Width           =   2280
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   180
               Index           =   24
               Left            =   2085
               TabIndex        =   115
               Text            =   "0"
               Top             =   600
               Width           =   360
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   180
               Index           =   4
               Left            =   2085
               TabIndex        =   113
               Text            =   "0"
               Top             =   315
               Width           =   360
            End
            Begin VB.CheckBox chk 
               Caption         =   "专家号同一病人限挂    个号"
               Height          =   180
               Index           =   178
               Left            =   180
               TabIndex        =   112
               Top             =   315
               Width           =   2775
            End
            Begin VB.CheckBox chk 
               Caption         =   "专家号同一病人限约    个号"
               Height          =   180
               Index           =   179
               Left            =   180
               TabIndex        =   114
               Top             =   600
               Width           =   3135
            End
            Begin VB.Line lnEpr 
               Index           =   2
               X1              =   2070
               X2              =   2460
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line lnEpr 
               Index           =   0
               X1              =   2070
               X2              =   2460
               Y1              =   510
               Y2              =   510
            End
         End
         Begin VB.Frame fra 
            Caption         =   "脱机医保结算方式"
            Height          =   2160
            Index           =   5
            Left            =   360
            TabIndex        =   676
            Top             =   5550
            Width           =   3795
            Begin VB.ListBox lst 
               Height          =   1740
               Index           =   5
               Left            =   165
               Style           =   1  'Checkbox
               TabIndex        =   88
               Top             =   300
               Width           =   3420
            End
         End
         Begin VB.Frame fra 
            Caption         =   "消费卡管理"
            Height          =   1260
            Index           =   7
            Left            =   4470
            TabIndex        =   98
            Top             =   1980
            Width           =   4995
            Begin VB.CheckBox chk 
               Caption         =   "门诊退费时消费卡需要刷卡验证"
               Height          =   180
               Index           =   59
               Left            =   180
               TabIndex        =   101
               Top             =   930
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "消费卡刷卡消费须定位到密码框"
               Height          =   180
               Index           =   193
               Left            =   180
               TabIndex        =   100
               Top             =   600
               Width           =   2970
            End
            Begin VB.CheckBox chk 
               Caption         =   "缴款后立即打印缴款单"
               Height          =   180
               Index           =   163
               Left            =   180
               TabIndex        =   99
               Top             =   270
               Width           =   2415
            End
         End
         Begin VB.Frame fra 
            Caption         =   "门诊费用转住院相关"
            Height          =   1470
            Index           =   4
            Left            =   4470
            TabIndex        =   107
            Top             =   4950
            Width           =   4995
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   16
               ItemData        =   "frmParFee.frx":1A33A
               Left            =   1620
               List            =   "frmParFee.frx":1A33C
               Style           =   2  'Dropdown List
               TabIndex        =   111
               Top             =   1110
               Width           =   2115
            End
            Begin VB.CheckBox chk 
               Caption         =   "产生的所有预交单据一次性打印"
               Height          =   180
               Index           =   197
               Left            =   180
               TabIndex        =   110
               Top             =   845
               Width           =   2865
            End
            Begin VB.CheckBox chk 
               Caption         =   "出院病人允许门诊转住院"
               Height          =   180
               Index           =   183
               Left            =   180
               TabIndex        =   109
               Top             =   580
               Width           =   2685
            End
            Begin VB.CheckBox chk 
               Caption         =   "门诊转住院必须先审核"
               Height          =   180
               Index           =   159
               Left            =   180
               TabIndex        =   108
               Top             =   315
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "预交票据打印格式"
               Height          =   180
               Index           =   3
               Left            =   150
               TabIndex        =   787
               Top             =   1170
               Width           =   1440
            End
         End
         Begin VB.Frame fra 
            Caption         =   "执行登记管理"
            Height          =   1530
            Index           =   3
            Left            =   4470
            TabIndex        =   102
            Top             =   3345
            Width           =   4995
            Begin VB.Frame fraRegPrint 
               Caption         =   "执行登记单打印方式"
               Height          =   765
               Left            =   135
               TabIndex        =   678
               Top             =   630
               Width           =   3585
               Begin VB.OptionButton optRegPrint 
                  Caption         =   "不打印"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   104
                  Top             =   330
                  Width           =   915
               End
               Begin VB.OptionButton optRegPrint 
                  Caption         =   "自动打印"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   105
                  Top             =   330
                  Width           =   1050
               End
               Begin VB.OptionButton optRegPrint 
                  Caption         =   "选择是否打印"
                  Height          =   255
                  Index           =   2
                  Left            =   2160
                  TabIndex        =   106
                  Top             =   330
                  Width           =   1380
               End
               Begin VB.CommandButton cmdRegPrint 
                  Caption         =   "执行登记单打印设置"
                  Height          =   350
                  Left            =   4230
                  TabIndex        =   679
                  Top             =   255
                  Width           =   1860
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示医嘱发送的单据"
               Height          =   195
               Index           =   158
               Left            =   165
               TabIndex        =   103
               Top             =   315
               Width           =   2100
            End
         End
         Begin VB.Frame fra 
            Caption         =   "门诊收费、划价、记帐时输入"
            Height          =   870
            Index           =   6
            Left            =   360
            TabIndex        =   83
            Top             =   4575
            Width           =   3795
            Begin VB.CheckBox chk 
               Caption         =   "病人姓名"
               Height          =   210
               Index           =   7
               Left            =   510
               TabIndex        =   84
               Top             =   270
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号单号"
               Height          =   210
               Index           =   10
               Left            =   2040
               TabIndex        =   87
               Top             =   540
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "病人标识"
               Height          =   180
               Index           =   8
               Left            =   2040
               TabIndex        =   85
               Top             =   270
               Value           =   1  'Checked
               Width           =   1035
            End
            Begin VB.CheckBox chk 
               Caption         =   "刷就诊卡"
               Height          =   210
               Index           =   9
               Left            =   510
               TabIndex        =   86
               Top             =   540
               Value           =   1  'Checked
               Width           =   1020
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   10
            Left            =   2340
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   62
            Text            =   "1"
            Top             =   585
            Width           =   520
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   73
            Text            =   "5"
            Top             =   2060
            Width           =   930
         End
         Begin VB.CheckBox chk 
            Caption         =   "输入费用项目首位当类别简码"
            Height          =   195
            Index           =   56
            Left            =   390
            TabIndex        =   78
            Top             =   3195
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "从属项目汇总计算折扣额"
            Height          =   195
            Index           =   39
            Left            =   390
            TabIndex        =   79
            Top             =   3435
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "输入费用项目时先输类别"
            Height          =   195
            Index           =   25
            Left            =   390
            TabIndex        =   77
            Top             =   2955
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   65
            Text            =   "15"
            Top             =   960
            Width           =   930
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   70
            Text            =   "0"
            Top             =   1700
            Width           =   930
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2340
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   59
            Text            =   "1"
            Top             =   240
            Width           =   520
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            ItemData        =   "frmParFee.frx":1A33E
            Left            =   1530
            List            =   "frmParFee.frx":1A340
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   4100
            Width           =   2625
         End
         Begin VB.Frame fra 
            Caption         =   "零钞处理规则"
            Height          =   1680
            Index           =   15
            Left            =   4470
            TabIndex        =   89
            Top             =   180
            Width           =   4995
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   17
               Left            =   3075
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   360
               Width           =   1695
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   14
               Left            =   765
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   1230
               Width           =   1695
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   13
               Left            =   765
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   780
               Width           =   1695
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   12
               Left            =   765
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "消费卡"
               Height          =   180
               Index           =   4
               Left            =   2520
               TabIndex        =   96
               Top             =   405
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "结帐"
               Height          =   180
               Index           =   38
               Left            =   360
               TabIndex        =   94
               Top             =   1275
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "收费"
               Height          =   180
               Index           =   37
               Left            =   360
               TabIndex        =   92
               Top             =   855
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "挂号"
               Height          =   180
               Index           =   15
               Left            =   360
               TabIndex        =   90
               Top             =   420
               Width           =   360
            End
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   5
            Left            =   2865
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1700
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(5)"
            BuddyDispid     =   196647
            BuddyIndex      =   5
            OrigLeft        =   2625
            OrigTop         =   1560
            OrigRight       =   2880
            OrigBottom      =   1860
            Max             =   4
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   1
            Left            =   2865
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(1)"
            BuddyDispid     =   196647
            BuddyIndex      =   1
            OrigLeft        =   2625
            OrigTop         =   120
            OrigRight       =   2880
            OrigBottom      =   420
            Max             =   7
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   0
            Left            =   2865
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   15
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196647
            BuddyIndex      =   0
            OrigLeft        =   2625
            OrigTop         =   1200
            OrigRight       =   2880
            OrigBottom      =   1500
            Max             =   365
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   6
            Left            =   2865
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   2060
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(6)"
            BuddyDispid     =   196647
            BuddyIndex      =   6
            OrigLeft        =   2625
            OrigTop         =   1920
            OrigRight       =   2880
            OrigBottom      =   2220
            Max             =   5
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   10
            Left            =   2865
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   585
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(10)"
            BuddyDispid     =   196647
            BuddyIndex      =   10
            OrigLeft        =   2625
            OrigTop         =   465
            OrigRight       =   2880
            OrigBottom      =   765
            Max             =   7
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "号源开放时间"
            Height          =   180
            Index           =   2
            Left            =   750
            TabIndex        =   67
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "急诊挂号单有效的天数"
            Height          =   180
            Index           =   49
            Left            =   390
            TabIndex        =   61
            Top             =   630
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用单价保留位数"
            Height          =   180
            Index           =   35
            Left            =   390
            TabIndex        =   72
            Top             =   2110
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "挂号允许预约天数"
            Height          =   180
            Index           =   30
            Left            =   390
            TabIndex        =   64
            Top             =   1020
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用金额保留位数"
            Height          =   180
            Index           =   28
            Left            =   390
            TabIndex        =   69
            Top             =   1750
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "普通挂号单有效的天数"
            Height          =   180
            Index           =   16
            Left            =   390
            TabIndex        =   58
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "病人审核方式"
            Height          =   180
            Index           =   52
            Left            =   390
            TabIndex        =   81
            Top             =   4160
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7215
         Index           =   13
         Left            =   -74895
         ScaleHeight     =   7185
         ScaleWidth      =   9180
         TabIndex        =   555
         TabStop         =   0   'False
         Top             =   795
         Visible         =   0   'False
         Width           =   9210
         Begin VB.Frame fra票据格式 
            Caption         =   "退费票据格式"
            Height          =   1545
            Index           =   6
            Left            =   240
            TabIndex        =   579
            Top             =   5970
            Width           =   5910
            Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
               Height          =   1215
               Index           =   6
               Left            =   120
               TabIndex        =   580
               Top             =   240
               Width           =   5715
               _cx             =   10081
               _cy             =   2143
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A342
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   2
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin VB.Frame fra票据格式 
            Caption         =   "收费票据格式"
            Height          =   1545
            Index           =   2
            Left            =   240
            TabIndex        =   577
            Top             =   4335
            Width           =   5910
            Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
               Height          =   1215
               Index           =   2
               Left            =   120
               TabIndex        =   578
               Top             =   240
               Width           =   5715
               _cx             =   10081
               _cy             =   2143
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A3D0
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   2
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin VB.Frame fraSupplementaryPrint 
            Caption         =   "结算清单打印方式"
            Height          =   795
            Left            =   255
            TabIndex        =   573
            Top             =   3435
            Width           =   5910
            Begin VB.OptionButton optSupplementaryPrint 
               Caption         =   "不打印"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   574
               Top             =   390
               Value           =   -1  'True
               Width           =   900
            End
            Begin VB.OptionButton optSupplementaryPrint 
               Caption         =   "选择是否打印"
               Height          =   180
               Index           =   2
               Left            =   2505
               TabIndex        =   576
               Top             =   390
               Width           =   1455
            End
            Begin VB.OptionButton optSupplementaryPrint 
               Caption         =   "自动打印"
               Height          =   180
               Index           =   1
               Left            =   1275
               TabIndex        =   575
               Top             =   390
               Width           =   1065
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   12
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   563
            Text            =   "3"
            Top             =   1005
            Width           =   375
         End
         Begin VB.Frame fra结算方式 
            Caption         =   "支持收费结算方式"
            Height          =   6705
            Left            =   6750
            TabIndex        =   581
            Top             =   270
            Width           =   2295
            Begin VB.ListBox lst 
               Height          =   6360
               Index           =   2
               ItemData        =   "frmParFee.frx":1A45E
               Left            =   165
               List            =   "frmParFee.frx":1A460
               Style           =   1  'Checkbox
               TabIndex        =   582
               Top             =   270
               Width           =   2010
            End
         End
         Begin VB.Frame fraSupplementaryMode 
            Caption         =   "药品摆药后退费方式"
            Height          =   795
            Left            =   255
            TabIndex        =   565
            Top             =   1440
            Width           =   5910
            Begin VB.OptionButton optDrugSupplementary 
               Caption         =   "提醒"
               Height          =   180
               Index           =   2
               Left            =   2505
               TabIndex        =   568
               Top             =   420
               Width           =   690
            End
            Begin VB.OptionButton optDrugSupplementary 
               Caption         =   "禁止"
               Height          =   180
               Index           =   1
               Left            =   1470
               TabIndex        =   567
               Top             =   420
               Width           =   690
            End
            Begin VB.OptionButton optDrugSupplementary 
               Caption         =   "不检查"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   566
               Top             =   405
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   11
            Left            =   1275
            Locked          =   -1  'True
            TabIndex        =   560
            Text            =   "10"
            Top             =   585
            Width           =   480
         End
         Begin VB.Frame fra 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   18
            Left            =   2865
            TabIndex        =   558
            Top             =   465
            Width           =   285
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   19
            Left            =   2865
            MaxLength       =   3
            TabIndex        =   557
            Text            =   "0"
            Top             =   270
            Width           =   285
         End
         Begin VB.Frame fra单位 
            Caption         =   " 药品单位 "
            Height          =   795
            Index           =   3
            Left            =   255
            TabIndex        =   569
            Top             =   2415
            Width           =   5910
            Begin VB.OptionButton optSupplementaryUnit 
               Caption         =   "门诊(或住院)单位"
               Height          =   180
               Index           =   1
               Left            =   2505
               TabIndex        =   572
               Top             =   405
               Width           =   1770
            End
            Begin VB.OptionButton optSupplementaryUnit 
               Caption         =   "售价单位"
               Height          =   180
               Index           =   0
               Left            =   1275
               TabIndex        =   571
               Top             =   405
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl单位 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "补结算时按"
               Height          =   180
               Index           =   3
               Left            =   300
               TabIndex        =   570
               Top             =   405
               Width           =   900
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
            Height          =   195
            Index           =   136
            Left            =   240
            TabIndex        =   556
            Top             =   270
            Width           =   4260
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   11
            Left            =   1740
            TabIndex        =   561
            TabStop         =   0   'False
            Top             =   585
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "txtUD(11)"
            BuddyDispid     =   196647
            BuddyIndex      =   11
            OrigLeft        =   1740
            OrigTop         =   585
            OrigRight       =   1995
            OrigBottom      =   885
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "票据剩余         张时开始提醒收费员"
            Height          =   285
            Index           =   137
            Left            =   240
            TabIndex        =   559
            Top             =   585
            Width           =   3450
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   12
            Left            =   3510
            TabIndex        =   564
            TabStop         =   0   'False
            Top             =   1005
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   3
            BuddyControl    =   "txtUD(12)"
            BuddyDispid     =   196647
            BuddyIndex      =   12
            OrigLeft        =   1605
            OrigTop         =   4860
            OrigRight       =   1860
            OrigBottom      =   5160
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblVaildDays 
            Caption         =   "可进行保险补充结算的费用有效天数"
            Height          =   225
            Left            =   240
            TabIndex        =   562
            Top             =   1035
            Width           =   2895
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7530
         Index           =   12
         Left            =   -74880
         ScaleHeight     =   7500
         ScaleWidth      =   9660
         TabIndex        =   360
         TabStop         =   0   'False
         Top             =   810
         Visible         =   0   'False
         Width           =   9690
         Begin VB.Frame fraNormal 
            Height          =   1725
            Index           =   0
            Left            =   285
            TabIndex        =   361
            Top             =   120
            Width           =   4575
            Begin VB.CheckBox chk 
               Caption         =   "允许录入特殊使用的抗生素"
               Height          =   195
               Index           =   195
               Left            =   180
               TabIndex        =   369
               Top             =   1440
               Width           =   2490
            End
            Begin VB.Frame fraLineDays 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   2
               Left            =   2820
               TabIndex        =   367
               Top             =   1050
               Width           =   285
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   16
               Left            =   2835
               MaxLength       =   3
               TabIndex        =   366
               Text            =   "0"
               Top             =   840
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "只查找合约单位病人"
               Height          =   195
               Index           =   132
               Left            =   180
               TabIndex        =   368
               Top             =   1155
               Width           =   2100
            End
            Begin VB.CheckBox chk 
               Caption         =   "中药输入付数"
               Height          =   195
               Index           =   71
               Left            =   180
               TabIndex        =   362
               Top             =   285
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "变价输入数次"
               Height          =   195
               Index           =   72
               Left            =   2640
               TabIndex        =   363
               Top             =   285
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "开单人含护士"
               Height          =   195
               Index           =   77
               Left            =   180
               TabIndex        =   364
               Top             =   570
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
               Height          =   195
               Index           =   87
               Left            =   180
               TabIndex        =   365
               Top             =   855
               Width           =   4260
            End
         End
         Begin VB.Frame fraPrintBill 
            Caption         =   "单据打印"
            Height          =   1305
            Left            =   270
            TabIndex        =   374
            Top             =   2940
            Width           =   4575
            Begin VB.CheckBox chk 
               Caption         =   "审核时打印记帐单据"
               Height          =   195
               Index           =   135
               Left            =   285
               TabIndex        =   378
               Top             =   930
               Width           =   2505
            End
            Begin VB.CheckBox chk 
               Caption         =   "划价时打印划价记帐单据"
               Height          =   195
               Index           =   134
               Left            =   285
               TabIndex        =   376
               Top             =   630
               Width           =   2670
            End
            Begin VB.CheckBox chk 
               Caption         =   "记帐时打印记帐单据"
               Height          =   195
               Index           =   133
               Left            =   285
               TabIndex        =   375
               Top             =   330
               Width           =   2220
            End
         End
         Begin VB.Frame fra库存显示 
            Caption         =   "库存显示"
            Height          =   1305
            Index           =   2
            Left            =   270
            TabIndex        =   379
            Top             =   4455
            Width           =   4575
            Begin VB.OptionButton opt记帐库存显示方式 
               Caption         =   "仅显示有无"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   384
               Top             =   945
               Width           =   1215
            End
            Begin VB.OptionButton opt记帐库存显示方式 
               Caption         =   "显示库存数"
               Height          =   180
               Index           =   0
               Left            =   1365
               TabIndex        =   383
               Top             =   945
               Width           =   1290
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药房库存"
               Height          =   195
               Index           =   83
               Left            =   150
               TabIndex        =   380
               Top             =   375
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药库库存"
               Height          =   195
               Index           =   82
               Left            =   2250
               TabIndex        =   381
               Top             =   390
               Width           =   1770
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H00FFFFFF&
               Index           =   5
               X1              =   15
               X2              =   4545
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H80000000&
               Index           =   4
               X1              =   15
               X2              =   4545
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label lbl库存显示方式 
               AutoSize        =   -1  'True
               Caption         =   "库存显示方式"
               Height          =   180
               Index           =   2
               Left            =   150
               TabIndex        =   382
               Top             =   945
               Width           =   1080
            End
         End
         Begin VB.Frame fra单位 
            Caption         =   " 药品单位 "
            Height          =   735
            Index           =   2
            Left            =   285
            TabIndex        =   370
            Top             =   2025
            Width           =   4575
            Begin VB.OptionButton opt记帐单位 
               Caption         =   "门诊(或住院)单位"
               Height          =   180
               Index           =   1
               Left            =   2085
               TabIndex        =   373
               Top             =   375
               Width           =   1755
            End
            Begin VB.OptionButton opt记帐单位 
               Caption         =   "售价单位"
               Height          =   180
               Index           =   0
               Left            =   1020
               TabIndex        =   372
               Top             =   375
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl单位 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "记帐时按"
               Height          =   180
               Index           =   2
               Left            =   15
               TabIndex        =   371
               Top             =   360
               Width           =   975
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7530
         Index           =   10
         Left            =   -74925
         ScaleHeight     =   7500
         ScaleWidth      =   10065
         TabIndex        =   358
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   10095
         Begin VB.Frame fraNormal 
            Caption         =   "单据控制"
            Height          =   2910
            Index           =   1
            Left            =   285
            TabIndex        =   385
            Top             =   240
            Width           =   4680
            Begin VB.CheckBox chk 
               Caption         =   "允许录入特殊使用的抗生素"
               Height          =   195
               Index           =   194
               Left            =   240
               TabIndex        =   400
               Top             =   2640
               Width           =   2490
            End
            Begin VB.CheckBox chk 
               Caption         =   "住院病人按门诊收费"
               Height          =   270
               Index           =   168
               Left            =   240
               TabIndex        =   392
               Top             =   1294
               Width           =   1920
            End
            Begin VB.TextBox txtUD 
               ForeColor       =   &H80000012&
               Height          =   285
               Index           =   7
               Left            =   1365
               MaxLength       =   3
               TabIndex        =   398
               Text            =   "0"
               Top             =   2310
               Width           =   600
            End
            Begin VB.TextBox txt 
               ForeColor       =   &H80000012&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   1365
               MaxLength       =   12
               TabIndex        =   396
               Text            =   "0.00"
               Top             =   1935
               Width           =   1335
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   14
               Left            =   2895
               MaxLength       =   3
               TabIndex        =   394
               Text            =   "0"
               Top             =   1635
               Width           =   285
            End
            Begin VB.Frame fraLineDays 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   0
               Left            =   2895
               TabIndex        =   401
               Top             =   1815
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "不使用缺省开单人"
               Height          =   195
               Index           =   110
               Left            =   240
               TabIndex        =   388
               Top             =   672
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "必须要输入开单人"
               Height          =   195
               Index           =   111
               Left            =   2550
               TabIndex        =   389
               Top             =   672
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "缺省科室优先"
               Height          =   195
               Index           =   112
               Left            =   240
               TabIndex        =   390
               Top             =   1014
               Width           =   1620
            End
            Begin VB.CheckBox chk 
               Caption         =   "中药输入付数"
               Height          =   195
               Index           =   69
               Left            =   240
               TabIndex        =   386
               Top             =   330
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "变价输入数次"
               Height          =   195
               Index           =   74
               Left            =   2550
               TabIndex        =   387
               Top             =   330
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "开单人含护士"
               Height          =   195
               Index           =   75
               Left            =   2550
               TabIndex        =   391
               Top             =   1014
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
               Height          =   195
               Index           =   85
               Left            =   240
               TabIndex        =   393
               Top             =   1650
               Width           =   4260
            End
            Begin MSComCtl2.UpDown ud 
               Height          =   270
               Index           =   7
               Left            =   1950
               TabIndex        =   399
               TabStop         =   0   'False
               Top             =   2310
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   476
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtUD(7)"
               BuddyDispid     =   196647
               BuddyIndex      =   7
               OrigLeft        =   5265
               OrigTop         =   555
               OrigRight       =   5505
               OrigBottom      =   825
               Max             =   32767
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lblDay 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "取消划价超过           天未处理的划价单"
               Height          =   180
               Left            =   240
               TabIndex        =   397
               Top             =   2355
               Width           =   3510
            End
            Begin VB.Label lblMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单据最大金额"
               Height          =   180
               Left            =   240
               TabIndex        =   395
               Top             =   1995
               Width           =   1080
            End
         End
         Begin VB.Frame fra划价通知单打印 
            Caption         =   "划价通知单打印"
            Height          =   1290
            Left            =   5565
            TabIndex        =   420
            Top             =   240
            Width           =   3720
            Begin VB.OptionButton optPrintRequisition 
               Caption         =   "自动打印"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   422
               Top             =   585
               Value           =   -1  'True
               Width           =   1260
            End
            Begin VB.OptionButton optPrintRequisition 
               Caption         =   "不打印"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   421
               Top             =   330
               Width           =   1020
            End
            Begin VB.OptionButton optPrintRequisition 
               Caption         =   "选择是否打印"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   423
               Top             =   855
               Width           =   1500
            End
         End
         Begin VB.Frame fra 
            Caption         =   "汇总栏显示方式"
            Height          =   1290
            Index           =   1
            Left            =   5565
            TabIndex        =   424
            Top             =   1665
            Width           =   3720
            Begin VB.OptionButton optBillTotalShow 
               Caption         =   "以收据费目显示分类合计"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   425
               Top             =   345
               Value           =   -1  'True
               Width           =   2280
            End
            Begin VB.OptionButton optBillTotalShow 
               Caption         =   "以收入项目显示分类合计"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   426
               Top             =   615
               Width           =   2280
            End
            Begin VB.OptionButton optBillTotalShow 
               Caption         =   "按单据分类汇总显示"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   427
               Top             =   870
               Width           =   2280
            End
         End
         Begin VB.Frame fraBillInputItem 
            Caption         =   "划价时要输入的项目"
            Height          =   1050
            Index           =   0
            Left            =   285
            TabIndex        =   412
            Top             =   5730
            Width           =   4680
            Begin VB.CheckBox chk 
               Caption         =   "医疗付款方式"
               Height          =   210
               Index           =   95
               Left            =   2940
               TabIndex        =   419
               Top             =   675
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "性别"
               Height          =   210
               Index           =   88
               Left            =   240
               TabIndex        =   413
               Top             =   360
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "年龄"
               Height          =   210
               Index           =   91
               Left            =   2940
               TabIndex        =   415
               Top             =   360
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "费别"
               Height          =   210
               Index           =   92
               Left            =   3810
               TabIndex        =   416
               Top             =   360
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "是否加班"
               Height          =   210
               Index           =   89
               Left            =   1500
               TabIndex        =   414
               Top             =   360
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "开单日期"
               Height          =   210
               Index           =   93
               Left            =   240
               TabIndex        =   417
               Top             =   675
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "开单人"
               Height          =   210
               Index           =   94
               Left            =   1500
               TabIndex        =   418
               Top             =   675
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.Frame fra库存显示 
            Caption         =   "库存显示"
            Height          =   1275
            Index           =   0
            Left            =   285
            TabIndex        =   406
            Top             =   4260
            Width           =   4680
            Begin VB.OptionButton opt划价库存显示方式 
               Caption         =   "仅显示有无"
               Height          =   180
               Index           =   1
               Left            =   2835
               TabIndex        =   411
               Top             =   900
               Width           =   1215
            End
            Begin VB.OptionButton opt划价库存显示方式 
               Caption         =   "显示库存数"
               Height          =   180
               Index           =   0
               Left            =   1440
               TabIndex        =   410
               Top             =   900
               Width           =   1290
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药房库存"
               Height          =   195
               Index           =   78
               Left            =   240
               TabIndex        =   407
               Top             =   375
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药库库存"
               Height          =   195
               Index           =   79
               Left            =   2355
               TabIndex        =   408
               Top             =   375
               Width           =   1770
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   15
               X2              =   4625
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H80000000&
               Index           =   1
               X1              =   15
               X2              =   4640
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label lbl库存显示方式 
               AutoSize        =   -1  'True
               Caption         =   "库存显示方式"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   409
               Top             =   900
               Width           =   1080
            End
         End
         Begin VB.Frame fra单位 
            Caption         =   " 药品单位 "
            Height          =   780
            Index           =   0
            Left            =   285
            TabIndex        =   402
            Top             =   3315
            Width           =   4680
            Begin VB.OptionButton opt划价单位 
               Caption         =   "门诊(或住院)单位"
               Height          =   180
               Index           =   1
               Left            =   2250
               TabIndex        =   405
               Top             =   405
               Width           =   1740
            End
            Begin VB.OptionButton opt划价单位 
               Caption         =   "售价单位"
               Height          =   180
               Index           =   0
               Left            =   1050
               TabIndex        =   404
               Top             =   405
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl单位 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "划价时按"
               Height          =   180
               Index           =   0
               Left            =   15
               TabIndex        =   403
               Top             =   405
               Width           =   975
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6810
         Index           =   9
         Left            =   -74925
         ScaleHeight     =   6780
         ScaleWidth      =   9000
         TabIndex        =   327
         TabStop         =   0   'False
         Top             =   780
         Width           =   9030
         Begin VB.CheckBox chk 
            Caption         =   "分诊台签到开始排队"
            Height          =   330
            Index           =   65
            Left            =   4890
            TabIndex        =   352
            Top             =   210
            Width           =   1935
         End
         Begin VB.Frame fra分诊台签到排队 
            Height          =   7245
            Left            =   4740
            TabIndex        =   351
            Top             =   240
            Width           =   4635
            Begin VB.CommandButton cmdDepClearAll 
               Caption         =   "全清"
               Height          =   350
               Left            =   3300
               TabIndex        =   355
               Top             =   6420
               Width           =   1100
            End
            Begin VB.CommandButton cmdDepSelectAll 
               Caption         =   "全选"
               Height          =   350
               Left            =   2100
               TabIndex        =   354
               Top             =   6420
               Width           =   1100
            End
            Begin VB.CheckBox chk 
               Caption         =   "再次签到需重新排队"
               Height          =   330
               Index           =   177
               Left            =   2520
               TabIndex        =   357
               Top             =   6870
               Width           =   1935
            End
            Begin VB.CheckBox chk 
               Caption         =   "回诊病人需重新排队"
               Height          =   330
               Index           =   171
               Left            =   210
               TabIndex        =   356
               Top             =   6870
               Width           =   1935
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfTriageQueuingDep 
               Height          =   5985
               Left            =   150
               TabIndex        =   353
               Top             =   300
               Width           =   4305
               _cx             =   7594
               _cy             =   10557
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16772055
               ForeColorSel    =   0
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   12698049
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   255
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A462
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   1
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   -1  'True
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin VB.Frame fra排序 
            Caption         =   "候诊病人排序方式"
            Height          =   1260
            Left            =   210
            TabIndex        =   335
            Top             =   2040
            Width           =   4350
            Begin VB.OptionButton optTriageSort 
               Caption         =   "科室编码,号码,单据号"
               Height          =   210
               Index           =   0
               Left            =   435
               TabIndex        =   336
               Top             =   330
               Value           =   -1  'True
               Width           =   2280
            End
            Begin VB.OptionButton optTriageSort 
               Caption         =   "科室编码,号码,挂号时间"
               Height          =   210
               Index           =   1
               Left            =   435
               TabIndex        =   337
               Top             =   600
               Width           =   2280
            End
            Begin VB.OptionButton optTriageSort 
               Caption         =   "科室编码,号码,发生时间,登记时间"
               Height          =   210
               Index           =   2
               Left            =   435
               TabIndex        =   338
               Top             =   885
               Width           =   3555
            End
         End
         Begin VB.Frame fra排队单 
            Caption         =   "排队单打印"
            Height          =   825
            Left            =   210
            TabIndex        =   343
            Top             =   4515
            Width           =   4350
            Begin VB.OptionButton optTriagePrintMode 
               Caption         =   "不打印"
               Height          =   195
               Index           =   0
               Left            =   435
               TabIndex        =   344
               Top             =   375
               Value           =   -1  'True
               Width           =   990
            End
            Begin VB.OptionButton optTriagePrintMode 
               Caption         =   "自动打印"
               Height          =   195
               Index           =   1
               Left            =   1455
               TabIndex        =   345
               Top             =   375
               Width           =   1125
            End
            Begin VB.OptionButton optTriagePrintMode 
               Caption         =   "提示选择打印"
               Height          =   195
               Index           =   2
               Left            =   2700
               TabIndex        =   346
               Top             =   375
               Width           =   1455
            End
         End
         Begin VB.Frame fra条码 
            Caption         =   "条码打印"
            Height          =   855
            Left            =   210
            TabIndex        =   339
            Top             =   3495
            Width           =   4350
            Begin VB.OptionButton optTriageBarcodePrintMode 
               Caption         =   "不打印"
               Height          =   195
               Index           =   0
               Left            =   435
               TabIndex        =   340
               Top             =   360
               Value           =   -1  'True
               Width           =   990
            End
            Begin VB.OptionButton optTriageBarcodePrintMode 
               Caption         =   "自动打印"
               Height          =   195
               Index           =   1
               Left            =   1455
               TabIndex        =   341
               Top             =   360
               Width           =   1170
            End
            Begin VB.OptionButton optTriageBarcodePrintMode 
               Caption         =   "提示选择打印"
               Height          =   195
               Index           =   2
               Left            =   2700
               TabIndex        =   342
               Top             =   360
               Width           =   1395
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "医生诊室忙时允许分诊"
            Height          =   300
            Index           =   68
            Left            =   210
            TabIndex        =   333
            Top             =   1140
            Width           =   2460
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   11
            Left            =   960
            MaxLength       =   2
            TabIndex        =   332
            Text            =   "0"
            Top             =   585
            Width           =   435
         End
         Begin VB.CheckBox chk 
            Caption         =   "预约挂号进入队列"
            Height          =   270
            Index           =   66
            Left            =   210
            TabIndex        =   334
            Top             =   1485
            Width           =   1905
         End
         Begin VB.Frame fra排队叫号 
            Caption         =   "排队叫号模式"
            Height          =   1875
            Left            =   210
            TabIndex        =   347
            Top             =   5610
            Width           =   4350
            Begin VB.OptionButton optTriageQueuingMode 
               Caption         =   "禁止全院排队叫号"
               Height          =   240
               Index           =   0
               Left            =   435
               TabIndex        =   348
               Top             =   390
               Width           =   1770
            End
            Begin VB.OptionButton optTriageQueuingMode 
               Caption         =   "分诊台分诊呼叫或医生主动呼叫"
               Height          =   240
               Index           =   1
               Left            =   435
               TabIndex        =   350
               Top             =   1005
               Width           =   3045
            End
            Begin VB.OptionButton optTriageQueuingMode 
               Caption         =   "先分诊呼叫,再医生呼叫就诊"
               Height          =   240
               Index           =   2
               Left            =   435
               TabIndex        =   349
               Top             =   705
               Width           =   2625
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   329
            Text            =   "0"
            Top             =   210
            Width           =   420
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   3
            Left            =   1380
            TabIndex        =   330
            TabStop         =   0   'False
            Top             =   210
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtUD(3)"
            BuddyDispid     =   196647
            BuddyIndex      =   3
            OrigLeft        =   1635
            OrigTop         =   195
            OrigRight       =   1890
            OrigBottom      =   495
            Max             =   7
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl提前分诊 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提前      小时分诊"
            Height          =   180
            Left            =   525
            TabIndex        =   331
            Top             =   615
            Width           =   1620
         End
         Begin VB.Label lbl有效天数 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自动刷新         天内的挂号病人"
            Height          =   180
            Left            =   180
            TabIndex        =   328
            Top             =   270
            Width           =   3045
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7470
         Index           =   8
         Left            =   -74850
         ScaleHeight     =   7440
         ScaleWidth      =   9060
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   690
         Visible         =   0   'False
         Width           =   9090
         Begin VB.CheckBox chk 
            Caption         =   "发卡或绑定卡时自动生成病人门诊号"
            Height          =   180
            Index           =   201
            Left            =   150
            TabIndex        =   167
            Top             =   1620
            Width           =   3225
         End
         Begin VB.CheckBox chk 
            Caption         =   "收取病历费"
            Height          =   180
            Index           =   199
            Left            =   150
            TabIndex        =   166
            Top             =   1290
            Width           =   2535
         End
         Begin VB.Frame fra票据格式 
            Caption         =   "收费票据格式"
            Height          =   1455
            Index           =   9
            Left            =   150
            TabIndex        =   766
            Top             =   5160
            Width           =   4155
            Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
               Height          =   1125
               Index           =   9
               Left            =   90
               TabIndex        =   185
               Top             =   255
               Width           =   3975
               _cx             =   7011
               _cy             =   1984
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   2
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A4EC
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   2
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "卡费使用门诊收费医疗收据"
            Height          =   180
            Index           =   192
            Left            =   150
            TabIndex        =   165
            Top             =   960
            Width           =   2535
         End
         Begin VB.Frame fra票据格式 
            Height          =   120
            Index           =   8
            Left            =   1575
            TabIndex        =   680
            Top             =   6645
            Width           =   7935
         End
         Begin VB.Frame fraPrintMode_SendCard 
            Caption         =   "发卡和绑定卡凭据打印方式"
            Height          =   1440
            Left            =   150
            TabIndex        =   168
            Top             =   1935
            Width           =   4155
            Begin VB.OptionButton optPrintMode_SendCard 
               Caption         =   "自动打印"
               Height          =   180
               Index           =   1
               Left            =   300
               TabIndex        =   170
               Top             =   660
               Width           =   1020
            End
            Begin VB.OptionButton optPrintMode_SendCard 
               Caption         =   "选择是否打印"
               Height          =   180
               Index           =   2
               Left            =   300
               TabIndex        =   173
               Top             =   975
               Width           =   1380
            End
            Begin VB.OptionButton optPrintMode_SendCard 
               Caption         =   "不打印"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   169
               Top             =   345
               Value           =   -1  'True
               Width           =   900
            End
         End
         Begin VB.Frame fraShortLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   2790
            TabIndex        =   176
            Top             =   810
            Width           =   285
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   2790
            MaxLength       =   3
            TabIndex        =   164
            Text            =   "0"
            Top             =   630
            Width           =   285
         End
         Begin VB.CheckBox chk 
            Caption         =   "卡费以记账方式收取"
            Height          =   180
            Index           =   14
            Left            =   150
            TabIndex        =   160
            Top             =   285
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.Frame fra退卡方式 
            Caption         =   "退卡方式设置"
            Height          =   1695
            Left            =   150
            TabIndex        =   174
            Top             =   3420
            Width           =   4155
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "输入单据号退卡或刷卡退卡"
               Height          =   180
               Index           =   3
               Left            =   300
               TabIndex        =   183
               Top             =   1305
               Width           =   2460
            End
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "输入单据号后才刷卡退卡"
               Height          =   180
               Index           =   2
               Left            =   315
               TabIndex        =   182
               Top             =   1005
               Width           =   2460
            End
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "必须刷卡退卡"
               Height          =   180
               Index           =   1
               Left            =   315
               TabIndex        =   177
               Top             =   690
               Width           =   1740
            End
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "不进行刷卡验证"
               Height          =   180
               Index           =   0
               Left            =   315
               TabIndex        =   175
               Top             =   375
               Value           =   -1  'True
               Width           =   1740
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsInputItemSet 
            Height          =   6060
            Index           =   0
            Left            =   4830
            TabIndex        =   186
            Top             =   540
            Width           =   4380
            _cx             =   7726
            _cy             =   10689
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483634
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483634
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483643
            FloodColor      =   192
            SheetBorder     =   -2147483637
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParFee.frx":1A54D
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许通过输入姓名来模糊查找    天内的病人"
            Height          =   195
            Index           =   15
            Left            =   150
            TabIndex        =   162
            Top             =   630
            Width           =   4260
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   975
            Index           =   8
            Left            =   210
            TabIndex        =   187
            Top             =   6975
            Width           =   8925
            _cx             =   15743
            _cy             =   1720
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParFee.frx":1A5F1
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   2
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblCardDeposit 
            AutoSize        =   -1  'True
            Caption         =   "预交票据打印设置"
            Height          =   180
            Left            =   105
            TabIndex        =   681
            Top             =   6675
            Width           =   1440
         End
         Begin VB.Label lblInputSendCardSet 
            AutoSize        =   -1  'True
            Caption         =   "输入项控制"
            Height          =   180
            Left            =   4800
            TabIndex        =   184
            Top             =   285
            Width           =   900
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7860
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   7830
         ScaleWidth      =   9210
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   9240
         Begin VB.Frame fra 
            Caption         =   "门诊退费刷卡控制(退回预存款)"
            Height          =   1035
            Index           =   13
            Left            =   4770
            TabIndex        =   783
            Top             =   6600
            Width           =   4320
            Begin VB.OptionButton optBrushCard 
               Caption         =   "禁止刷卡"
               Height          =   180
               Index           =   10
               Left            =   240
               TabIndex        =   666
               Top             =   330
               Width           =   1035
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "必须刷卡验证"
               Height          =   180
               Index           =   11
               Left            =   1560
               TabIndex        =   667
               Top             =   360
               Value           =   -1  'True
               Width           =   1425
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "医疗卡设置了密码，必须刷卡验证"
               Height          =   180
               Index           =   12
               Left            =   240
               TabIndex        =   668
               Top             =   720
               Width           =   3045
            End
         End
         Begin VB.Frame fra 
            Caption         =   "门诊消费刷卡控制(使用预存款)"
            Height          =   1395
            Index           =   12
            Left            =   270
            TabIndex        =   782
            Top             =   6360
            Width           =   4215
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   26
               Left            =   480
               MaxLength       =   8
               TabIndex        =   792
               Text            =   "0"
               Top             =   1020
               Width           =   795
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "_________元内免密支付"
               Height          =   180
               Index           =   3
               Left            =   210
               TabIndex        =   791
               Top             =   1080
               Width           =   3045
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "医疗卡设置了密码，必须刷卡验证"
               Height          =   180
               Index           =   2
               Left            =   210
               TabIndex        =   661
               Top             =   720
               Width           =   3045
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "必须刷卡验证"
               Height          =   180
               Index           =   1
               Left            =   1530
               TabIndex        =   660
               Top             =   375
               Value           =   -1  'True
               Width           =   1425
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "禁止刷卡"
               Height          =   180
               Index           =   0
               Left            =   210
               TabIndex        =   659
               Top             =   375
               Width           =   1035
            End
         End
         Begin VB.Frame fra 
            Caption         =   "门诊收费控制"
            Height          =   1095
            Index           =   14
            Left            =   270
            TabIndex        =   651
            Top             =   1860
            Width           =   4215
            Begin VB.CheckBox chk 
               Caption         =   "项目执行前必须先收费或先记帐审核"
               Height          =   210
               Index           =   67
               Left            =   180
               TabIndex        =   652
               Top             =   360
               Width           =   3540
            End
            Begin VB.CheckBox chk 
               Caption         =   "项目开单后立即收费或记帐审核"
               Height          =   210
               Index           =   90
               Left            =   180
               TabIndex        =   653
               Top             =   675
               Width           =   3120
            End
         End
         Begin VB.Frame fraCharge 
            Caption         =   "消费确定后收费票据设置"
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   270
            TabIndex        =   649
            Top             =   3015
            Width           =   4215
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   8
               ItemData        =   "frmParFee.frx":1A67F
               Left            =   1335
               List            =   "frmParFee.frx":1A681
               Style           =   2  'Dropdown List
               TabIndex        =   654
               Top             =   390
               Width           =   2460
            End
            Begin VB.Label lblPrintFormat 
               AutoSize        =   -1  'True
               Caption         =   "票据打印格式"
               Height          =   180
               Left            =   150
               TabIndex        =   650
               Top             =   450
               Width           =   1080
            End
         End
         Begin VB.Frame fraSetMoneyMode 
            Caption         =   "门诊收费、结帐刷卡缺省金额操作（三方卡）"
            Height          =   1260
            Left            =   270
            TabIndex        =   648
            Top             =   5010
            Width           =   4215
            Begin VB.OptionButton optSetMoneyMode 
               Caption         =   "不缺省刷卡金额"
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   656
               Top             =   345
               Value           =   -1  'True
               Width           =   2640
            End
            Begin VB.OptionButton optSetMoneyMode 
               Caption         =   "缺省刷卡金额且金额不允许更改"
               Height          =   210
               Index           =   2
               Left            =   210
               TabIndex        =   658
               Top             =   975
               Width           =   3120
            End
            Begin VB.OptionButton optSetMoneyMode 
               Caption         =   "缺省刷卡金额且金额允许更改"
               Height          =   210
               Index           =   1
               Left            =   210
               TabIndex        =   657
               Top             =   660
               Width           =   2760
            End
         End
         Begin VB.Frame fra 
            Caption         =   "三方接口药房设置"
            Height          =   4530
            Index           =   16
            Left            =   4770
            TabIndex        =   647
            Top             =   1860
            Width           =   4320
            Begin TabDlg.SSTab stabDrug 
               Height          =   4170
               Left            =   120
               TabIndex        =   662
               Top             =   270
               Width           =   4110
               _ExtentX        =   7250
               _ExtentY        =   7355
               _Version        =   393216
               Style           =   1
               TabHeight       =   520
               TabCaption(0)   =   "西药房"
               TabPicture(0)   =   "frmParFee.frx":1A683
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "vsfDrugStore(0)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "中药房"
               TabPicture(1)   =   "frmParFee.frx":1A69F
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "vsfDrugStore(1)"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "成药房"
               TabPicture(2)   =   "frmParFee.frx":1A6BB
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "vsfDrugStore(2)"
               Tab(2).ControlCount=   1
               Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
                  Height          =   3780
                  Index           =   0
                  Left            =   60
                  TabIndex        =   663
                  Top             =   330
                  Width           =   4020
                  _cx             =   7091
                  _cy             =   6667
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1A6D7
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   0
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
                  Height          =   3780
                  Index           =   1
                  Left            =   -74940
                  TabIndex        =   664
                  Top             =   330
                  Width           =   3990
                  _cx             =   7038
                  _cy             =   6667
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1A74B
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   0
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
                  Height          =   3780
                  Index           =   2
                  Left            =   -74940
                  TabIndex        =   665
                  Top             =   330
                  Width           =   3990
                  _cx             =   7038
                  _cy             =   6667
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1A7BF
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   0
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
         End
         Begin VB.Frame fraRecored 
            Caption         =   "消费记帐审核后记帐票据设置"
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   270
            TabIndex        =   171
            Top             =   4005
            Width           =   4215
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   9
               Left            =   1335
               Style           =   2  'Dropdown List
               TabIndex        =   655
               Top             =   345
               Width           =   2460
            End
            Begin VB.Label lblRecordPrint 
               AutoSize        =   -1  'True
               Caption         =   "票据打印格式"
               Height          =   180
               Left            =   210
               TabIndex        =   172
               Top             =   405
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdOneCard 
            Height          =   345
            Index           =   0
            Left            =   8925
            Picture         =   "frmParFee.frx":1A833
            Style           =   1  'Graphical
            TabIndex        =   159
            Top             =   405
            Width           =   345
         End
         Begin VB.CommandButton cmdOneCard 
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   8925
            Picture         =   "frmParFee.frx":1ADBD
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   795
            Width           =   345
         End
         Begin VB.CommandButton cmdOneCard 
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   8925
            Picture         =   "frmParFee.frx":1B347
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   1200
            Width           =   345
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   1305
            Index           =   3
            Left            =   240
            TabIndex        =   157
            Top             =   390
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   2302
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "NO"
               Text            =   "编号"
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Name"
               Text            =   "名称"
               Object.Width           =   7410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   "PayType"
               Text            =   "结算方式"
               Object.Width           =   2998
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "OrgCode"
               Text            =   "医院编码"
               Object.Width           =   1677
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "Enable"
               Text            =   "启用"
               Object.Width           =   970
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "一卡通接口(老版)"
            Height          =   180
            Index           =   45
            Left            =   285
            TabIndex        =   156
            Top             =   120
            Width           =   1440
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   8115
         Index           =   6
         Left            =   -74865
         ScaleHeight     =   8085
         ScaleWidth      =   8985
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   615
         Width           =   9015
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   975
            Index           =   0
            Left            =   255
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   60
            Width           =   5025
            _Version        =   589884
            _ExtentX        =   8864
            _ExtentY        =   1720
            _StockProps     =   64
         End
         Begin VB.PictureBox picOtherRegister 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7050
            Left            =   345
            ScaleHeight     =   7020
            ScaleWidth      =   8970
            TabIndex        =   294
            TabStop         =   0   'False
            Top             =   0
            Width           =   9000
            Begin VB.Frame fraOrder 
               Caption         =   "医生站挂号排序控制"
               Height          =   2520
               Index           =   1
               Left            =   4050
               TabIndex        =   320
               Top             =   4080
               Visible         =   0   'False
               Width           =   4410
               Begin VB.CommandButton cmdStationRegOrder 
                  Caption         =   "↑"
                  Height          =   510
                  Index           =   0
                  Left            =   3720
                  TabIndex        =   322
                  Top             =   825
                  Width           =   375
               End
               Begin VB.CommandButton cmdStationRegOrder 
                  Caption         =   "↓"
                  Height          =   510
                  Index           =   1
                  Left            =   3720
                  TabIndex        =   323
                  Top             =   1425
                  Width           =   375
               End
               Begin VSFlex8Ctl.VSFlexGrid vsStationRegSort 
                  Height          =   1995
                  Left            =   240
                  TabIndex        =   321
                  Top             =   360
                  Width           =   3330
                  _cx             =   5874
                  _cy             =   3519
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483634
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   2
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   6
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   300
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1B8D1
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   4
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   8
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   2
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "医生站预约包含科室安排"
               Height          =   330
               Index           =   184
               Left            =   90
               TabIndex        =   752
               Top             =   1980
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "医生站挂未设置医生的号别时必须输入医生"
               Height          =   195
               Index           =   169
               Left            =   90
               TabIndex        =   303
               Top             =   2655
               Width           =   3870
            End
            Begin VB.Frame fraSlip 
               Caption         =   "挂号凭条打印"
               Height          =   765
               Left            =   4050
               TabIndex        =   316
               Top             =   3165
               Width           =   4410
               Begin VB.OptionButton optPrintSlip 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   318
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.OptionButton optPrintSlip 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   317
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optPrintSlip 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2670
                  TabIndex        =   319
                  Top             =   360
                  Width           =   1380
               End
            End
            Begin VB.Frame fraAppoint 
               Caption         =   "预约挂号单打印"
               Height          =   765
               Left            =   4050
               TabIndex        =   312
               Top             =   2220
               Width           =   4410
               Begin VB.OptionButton optPrintAppoint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2670
                  TabIndex        =   315
                  Top             =   375
                  Width           =   1380
               End
               Begin VB.OptionButton optPrintAppoint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   313
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optPrintAppoint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   314
                  Top             =   375
                  Width           =   1020
               End
            End
            Begin VB.Frame fraInvoice 
               Caption         =   "挂号票据打印"
               Height          =   735
               Left            =   4050
               TabIndex        =   308
               Top             =   1380
               Width           =   4410
               Begin VB.OptionButton optPrintFact 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2670
                  TabIndex        =   311
                  Top             =   390
                  Width           =   1380
               End
               Begin VB.OptionButton optPrintFact 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   309
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optPrintFact 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   310
                  Top             =   390
                  Width           =   1020
               End
            End
            Begin VB.Frame fraRegistMode 
               Caption         =   "挂号费支付模式"
               Height          =   1155
               Left            =   4050
               TabIndex        =   304
               Top             =   150
               Width           =   4410
               Begin VB.OptionButton optRegist 
                  Caption         =   "立即支付或窗口支付模式"
                  Height          =   225
                  Index           =   2
                  Left            =   300
                  TabIndex        =   307
                  Top             =   675
                  Width           =   3195
               End
               Begin VB.OptionButton optRegist 
                  Caption         =   "窗口支付模式"
                  Height          =   225
                  Index           =   1
                  Left            =   1920
                  TabIndex        =   306
                  Top             =   360
                  Width           =   1425
               End
               Begin VB.OptionButton optRegist 
                  Caption         =   "立即支付模式"
                  Height          =   225
                  Index           =   0
                  Left            =   300
                  TabIndex        =   305
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1425
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号必须刷卡"
               Height          =   330
               Index           =   102
               Left            =   90
               TabIndex        =   295
               Top             =   135
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号时缴费优先使用预交款"
               Height          =   330
               Index           =   103
               Left            =   90
               TabIndex        =   296
               Top             =   450
               Width           =   2850
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   180
               Index           =   106
               Left            =   1815
               TabIndex        =   300
               Text            =   "7"
               Top             =   1425
               Width           =   360
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号预约时收款"
               Height          =   330
               Index           =   108
               Left            =   90
               TabIndex        =   302
               Top             =   2310
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "医生站挂号包含科室安排"
               Height          =   330
               Index           =   107
               Left            =   90
               TabIndex        =   301
               Top             =   1680
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "随机序号选择"
               Height          =   330
               Index           =   105
               Left            =   90
               TabIndex        =   298
               Top             =   1080
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许住院病人进行挂号"
               Height          =   330
               Index           =   104
               Left            =   90
               TabIndex        =   297
               Top             =   765
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "输入姓名模糊查找     天内的病人"
               Height          =   180
               Index           =   106
               Left            =   90
               TabIndex        =   299
               Top             =   1425
               Width           =   3135
            End
            Begin VB.Line lnEpr 
               Index           =   1
               X1              =   1800
               X2              =   2190
               Y1              =   1620
               Y2              =   1620
            End
         End
         Begin VB.PictureBox picRegist 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6705
            Left            =   15
            ScaleHeight     =   6675
            ScaleWidth      =   9165
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   30
            Width           =   9195
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   2280
               MaxLength       =   2
               TabIndex        =   234
               Text            =   "0"
               Top             =   6660
               Width           =   660
            End
            Begin VB.Frame fra缴款输入控制 
               Caption         =   "挂号缴款输入控制"
               Height          =   1050
               Left            =   4335
               TabIndex        =   778
               Top             =   3705
               Width           =   4500
               Begin VB.OptionButton optMoneyControl 
                  Caption         =   "不进行控制"
                  Height          =   180
                  Index           =   0
                  Left            =   225
                  TabIndex        =   781
                  Top             =   255
                  Width           =   1260
               End
               Begin VB.OptionButton optMoneyControl 
                  Caption         =   "输入缴款金额之后才结束本次挂号收费"
                  Height          =   180
                  Index           =   1
                  Left            =   225
                  TabIndex        =   780
                  Top             =   510
                  Width           =   3825
               End
               Begin VB.OptionButton optMoneyControl 
                  Caption         =   "必须输入缴款金额"
                  Height          =   180
                  Index           =   2
                  Left            =   225
                  TabIndex        =   779
                  Top             =   780
                  Width           =   3825
               End
            End
            Begin VB.PictureBox pic提前颜色 
               BackColor       =   &H00000000&
               Height          =   270
               Left            =   2010
               ScaleHeight     =   210
               ScaleWidth      =   210
               TabIndex        =   758
               Top             =   5565
               Width           =   270
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号和预约时禁止输入年龄"
               Height          =   300
               Index           =   190
               Left            =   135
               TabIndex        =   755
               Top             =   5280
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "计划排班模式精简界面"
               Height          =   300
               Index           =   185
               Left            =   135
               TabIndex        =   753
               Top             =   4995
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "病人同科挂号限制用于急诊"
               Height          =   210
               Index           =   175
               Left            =   405
               TabIndex        =   226
               Top             =   4455
               Width           =   3735
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1890
               MaxLength       =   2
               TabIndex        =   225
               Text            =   "0"
               Top             =   4185
               Width           =   660
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   2100
               MaxLength       =   2
               TabIndex        =   228
               Text            =   "0"
               Top             =   4755
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂未设置医生的号别时必须输入医生"
               Height          =   195
               Index           =   27
               Left            =   135
               TabIndex        =   212
               Top             =   1230
               Width           =   4080
            End
            Begin VB.Frame fraLine 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   840
               TabIndex        =   268
               Top             =   345
               Width           =   285
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   180
               Index           =   2
               Left            =   840
               MaxLength       =   2
               TabIndex        =   208
               Text            =   "5"
               Top             =   150
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号必须建病案时自动产生病人门诊号"
               Height          =   195
               Index           =   23
               Left            =   135
               TabIndex        =   209
               Top             =   450
               Width           =   3360
            End
            Begin VB.CheckBox chk 
               Caption         =   "建档病人挂号存为划价单"
               Height          =   240
               Index           =   24
               Left            =   135
               TabIndex        =   211
               Top             =   960
               Width           =   4005
            End
            Begin VB.CheckBox chk 
               Caption         =   "优先使用预交款缴费"
               Height          =   195
               Index           =   30
               Left            =   135
               TabIndex        =   213
               Top             =   1485
               Width           =   2340
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   180
               Index           =   3
               Left            =   2025
               MaxLength       =   3
               TabIndex        =   215
               Text            =   "0"
               Top             =   1740
               Width           =   285
            End
            Begin VB.Frame fraLine2 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   2055
               TabIndex        =   267
               Top             =   1935
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号费用为零时也打印票据"
               Height          =   255
               Index           =   32
               Left            =   135
               TabIndex        =   216
               Top             =   1980
               Width           =   2460
            End
            Begin VB.Frame fraInput 
               Caption         =   "要求输入项"
               Height          =   1050
               Left            =   4335
               TabIndex        =   235
               Top             =   135
               Width           =   4485
               Begin VB.CheckBox chk 
                  Caption         =   "联系电话"
                  Height          =   195
                  Index           =   20
                  Left            =   1440
                  TabIndex        =   243
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "病人"
                  Height          =   195
                  Index           =   33
                  Left            =   240
                  TabIndex        =   236
                  Top             =   285
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "性别"
                  Height          =   195
                  Index           =   34
                  Left            =   1440
                  TabIndex        =   237
                  Top             =   285
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "年龄"
                  Height          =   195
                  Index           =   35
                  Left            =   2745
                  TabIndex        =   238
                  Top             =   285
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "费别"
                  Height          =   195
                  Index           =   38
                  Left            =   240
                  TabIndex        =   239
                  Top             =   525
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "结算方式"
                  Height          =   195
                  Index           =   40
                  Left            =   2745
                  TabIndex        =   241
                  Top             =   525
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "付款方式"
                  Height          =   195
                  Index           =   37
                  Left            =   1440
                  TabIndex        =   240
                  Top             =   525
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "家庭地址"
                  Height          =   195
                  Index           =   36
                  Left            =   240
                  TabIndex        =   242
                  Top             =   765
                  Width           =   1020
               End
            End
            Begin VB.Frame fraRegistBillMode 
               Caption         =   "挂号票据"
               Height          =   510
               Left            =   4335
               TabIndex        =   244
               Top             =   1260
               Width           =   4500
               Begin VB.OptionButton optRegistPrintMode 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   247
                  Top             =   240
                  Width           =   1380
               End
               Begin VB.OptionButton optRegistPrintMode 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   245
                  Top             =   240
                  Width           =   900
               End
               Begin VB.OptionButton optRegistPrintMode 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   246
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号后打印病历标签"
               Height          =   255
               Index           =   45
               Left            =   135
               TabIndex        =   217
               Top             =   2250
               Width           =   1980
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许住院病人挂号"
               Height          =   255
               Index           =   49
               Left            =   135
               TabIndex        =   218
               Top             =   2520
               Width           =   1845
            End
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   2
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   230
               Top             =   5910
               Width           =   1800
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许操作员随机选择挂号安排序号"
               Height          =   255
               Index           =   51
               Left            =   135
               TabIndex        =   219
               Top             =   2790
               Width           =   3090
            End
            Begin VB.Frame fraBarCodePrint 
               Caption         =   "病人条码打印"
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   4335
               TabIndex        =   248
               Top             =   1845
               Width           =   4500
               Begin VB.OptionButton optBarCodePrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   249
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optBarCodePrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   250
                  Top             =   225
                  Width           =   1020
               End
               Begin VB.OptionButton optBarCodePrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   251
                  Top             =   225
                  Width           =   1380
               End
            End
            Begin VB.Frame fraClearMZInfor 
               Caption         =   "退号清除门诊号信息(挂号有效天数内的病人)"
               Height          =   615
               Left            =   4335
               TabIndex        =   256
               Top             =   3030
               Width           =   4500
               Begin VB.OptionButton optRegistClearMzInfor 
                  Caption         =   "提示清除"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   259
                  Top             =   285
                  Width           =   1110
               End
               Begin VB.OptionButton optRegistClearMzInfor 
                  Caption         =   "自动清除"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   258
                  Top             =   285
                  Width           =   1110
               End
               Begin VB.OptionButton optRegistClearMzInfor 
                  Caption         =   "不清除"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   257
                  Top             =   285
                  Width           =   1110
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "家庭地址联想输入"
               Height          =   255
               Index           =   54
               Left            =   135
               TabIndex        =   220
               Top             =   3075
               Value           =   1  'Checked
               Width           =   1980
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   10
               Left            =   240
               MaxLength       =   2
               TabIndex        =   231
               Text            =   "0"
               Top             =   6285
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "已退序号允许挂号"
               Height          =   300
               Index           =   60
               Left            =   135
               TabIndex        =   222
               Top             =   3600
               Width           =   2550
            End
            Begin VB.CheckBox chk 
               Caption         =   "分时段号别严格按时段挂号"
               Height          =   300
               Index           =   61
               Left            =   135
               TabIndex        =   221
               Top             =   3330
               Width           =   2550
            End
            Begin VB.Frame fraSlipPrint 
               Caption         =   "挂号凭条"
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   4335
               TabIndex        =   252
               Top             =   2445
               Width           =   4500
               Begin VB.OptionButton optSlipPrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   255
                  Top             =   240
                  Width           =   1380
               End
               Begin VB.OptionButton optSlipPrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   254
                  Top             =   240
                  Width           =   1020
               End
               Begin VB.OptionButton optSlipPrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   253
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   900
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号时默认勾选购买病历选项"
               Height          =   300
               Index           =   62
               Left            =   135
               TabIndex        =   223
               Top             =   3870
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "门诊号有效性检查"
               Height          =   255
               Index           =   63
               Left            =   135
               TabIndex        =   210
               Top             =   690
               Width           =   2325
            End
            Begin VB.Frame fraRegistCards 
               Caption         =   "挂号发卡相关"
               Height          =   2145
               Left            =   4335
               TabIndex        =   260
               Top             =   4845
               Width           =   4500
               Begin VB.CheckBox chk 
                  Caption         =   "卡费与挂号费一起收(否则卡费存为划价单)"
                  Height          =   195
                  Index           =   26
                  Left            =   165
                  TabIndex        =   261
                  Top             =   315
                  Width           =   3960
               End
               Begin VB.CheckBox chk 
                  Caption         =   "发卡新病人自动产生临时姓名"
                  Height          =   195
                  Index           =   29
                  Left            =   165
                  TabIndex        =   262
                  Top             =   555
                  Width           =   3345
               End
               Begin VB.CheckBox chk 
                  Caption         =   "退号不退卡时重打票据"
                  Height          =   195
                  Index           =   44
                  Left            =   165
                  TabIndex        =   263
                  Top             =   825
                  Width           =   2400
               End
               Begin VB.CheckBox chk 
                  Caption         =   "发卡不弹出病人信息登记窗口"
                  Height          =   195
                  Index           =   42
                  Left            =   165
                  TabIndex        =   264
                  Top             =   1110
                  Width           =   3345
               End
               Begin VB.CheckBox chk 
                  Caption         =   "扫描身份证签约"
                  Height          =   195
                  Index           =   58
                  Left            =   165
                  TabIndex        =   265
                  Top             =   1395
                  Width           =   3345
               End
               Begin VB.CheckBox chk 
                  Caption         =   "非严格控制卡时始终为发卡"
                  Height          =   195
                  Index           =   64
                  Left            =   165
                  TabIndex        =   266
                  Top             =   1695
                  Width           =   3345
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "每隔     分钟自动刷新挂号安排表"
               Height          =   195
               Index           =   22
               Left            =   135
               TabIndex        =   207
               Top             =   150
               Width           =   3480
            End
            Begin VB.CheckBox chk 
               Caption         =   "输入姓名后模糊查找    天内的病人"
               Height          =   195
               Index           =   31
               Left            =   135
               TabIndex        =   214
               Top             =   1740
               Width           =   3300
            End
            Begin VB.CheckBox chk 
               Caption         =   "病人同一科室限挂        个号"
               Height          =   210
               Index           =   174
               Left            =   135
               TabIndex        =   224
               Top             =   4170
               Width           =   3735
            End
            Begin VB.CheckBox chk 
               Caption         =   "同一病人最多能挂号        个科室"
               Height          =   210
               Index           =   176
               Left            =   135
               TabIndex        =   227
               Top             =   4740
               Width           =   3705
            End
            Begin VB.CheckBox chk 
               Caption         =   "同一病人同一号源限挂         个号"
               Height          =   210
               Index           =   203
               Left            =   135
               TabIndex        =   233
               Top             =   6660
               Width           =   3465
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   2280
               X2              =   2880
               Y1              =   6900
               Y2              =   6900
            End
            Begin VB.Label lblColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "新版提前挂号安排颜色"
               Height          =   180
               Left            =   135
               TabIndex        =   759
               Top             =   5610
               Width           =   1800
            End
            Begin VB.Line Line1 
               Index           =   10
               X1              =   1845
               X2              =   2580
               Y1              =   4380
               Y2              =   4380
            End
            Begin VB.Line Line1 
               Index           =   9
               X1              =   2055
               X2              =   2775
               Y1              =   4950
               Y2              =   4950
            End
            Begin VB.Label lblSortMode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "排序方式"
               Height          =   180
               Left            =   135
               TabIndex        =   229
               Top             =   5970
               Width           =   720
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   240
               X2              =   960
               Y1              =   6525
               Y2              =   6525
            End
            Begin VB.Label lblGuardian 
               AutoSize        =   -1  'True
               Caption         =   "岁以下必须录入监护人"
               Height          =   180
               Left            =   960
               TabIndex        =   232
               Top             =   6285
               Width           =   1800
            End
         End
         Begin VB.PictureBox picRegistPlan 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6810
            Left            =   840
            ScaleHeight     =   6780
            ScaleWidth      =   8595
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   30
            Width           =   8625
            Begin VB.Frame fraNewPaln 
               BorderStyle     =   0  'None
               Height          =   3855
               Left            =   150
               TabIndex        =   682
               Top             =   1560
               Visible         =   0   'False
               Width           =   5145
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   18
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   195
                  Top             =   930
                  Width           =   1695
               End
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   19
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   196
                  Top             =   1290
                  Width           =   1695
               End
               Begin VB.CheckBox chk 
                  Caption         =   "替诊医生职务级别检查"
                  Height          =   195
                  Index           =   182
                  Left            =   2940
                  TabIndex        =   193
                  Top             =   30
                  Width           =   2115
               End
               Begin VB.CheckBox chk 
                  Caption         =   "按替诊医生同步更新预约挂号单"
                  Height          =   195
                  Index           =   180
                  Left            =   0
                  TabIndex        =   194
                  Top             =   420
                  Width           =   2835
               End
               Begin VB.Frame fraVisitTablePrintMode 
                  Caption         =   "出诊表打印方式"
                  Height          =   735
                  Left            =   0
                  TabIndex        =   685
                  Top             =   3150
                  Width           =   5055
                  Begin VB.OptionButton optVisitTablePrintMode 
                     Caption         =   "不打印"
                     Height          =   180
                     Index           =   0
                     Left            =   300
                     TabIndex        =   203
                     Top             =   360
                     Value           =   -1  'True
                     Width           =   855
                  End
                  Begin VB.OptionButton optVisitTablePrintMode 
                     Caption         =   "自动打印"
                     Height          =   180
                     Index           =   1
                     Left            =   1680
                     TabIndex        =   204
                     Top             =   360
                     Width           =   1035
                  End
                  Begin VB.OptionButton optVisitTablePrintMode 
                     Caption         =   "选择是否打印"
                     Height          =   180
                     Index           =   2
                     Left            =   3090
                     TabIndex        =   205
                     Top             =   360
                     Width           =   1395
                  End
               End
               Begin VB.Frame fraPrintMode 
                  Caption         =   "预约清单打印方式"
                  Height          =   1305
                  Left            =   2940
                  TabIndex        =   684
                  Top             =   1710
                  Width           =   2115
                  Begin VB.OptionButton optPrintMode 
                     Caption         =   "选择是否打印"
                     Height          =   180
                     Index           =   2
                     Left            =   300
                     TabIndex        =   202
                     Top             =   930
                     Width           =   1395
                  End
                  Begin VB.OptionButton optPrintMode 
                     Caption         =   "自动打印"
                     Height          =   180
                     Index           =   1
                     Left            =   300
                     TabIndex        =   201
                     Top             =   615
                     Width           =   1035
                  End
                  Begin VB.OptionButton optPrintMode 
                     Caption         =   "不打印"
                     Height          =   180
                     Index           =   0
                     Left            =   300
                     TabIndex        =   200
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   855
                  End
               End
               Begin VB.Frame fraToExcelMode 
                  Caption         =   "预约清单控制方式"
                  Height          =   1305
                  Left            =   0
                  TabIndex        =   683
                  Top             =   1710
                  Width           =   2715
                  Begin VB.OptionButton optToExcelMode 
                     Caption         =   "选择是否输出到Excel"
                     Height          =   225
                     Index           =   2
                     Left            =   300
                     TabIndex        =   199
                     Top             =   930
                     Width           =   2025
                  End
                  Begin VB.OptionButton optToExcelMode 
                     Caption         =   "自动输出到Excel"
                     Height          =   225
                     Index           =   1
                     Left            =   300
                     TabIndex        =   198
                     Top             =   615
                     Width           =   1665
                  End
                  Begin VB.OptionButton optToExcelMode 
                     Caption         =   "不输出到Excel"
                     Height          =   225
                     Index           =   0
                     Left            =   300
                     TabIndex        =   197
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   1485
                  End
               End
               Begin VB.CheckBox chk 
                  Caption         =   "仅限于对院内医生进行挂号安排"
                  Height          =   180
                  Index           =   181
                  Left            =   0
                  TabIndex        =   192
                  Top             =   30
                  Width           =   2820
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "将未区分站点的号源分配给                   进行出诊安排"
                  Height          =   180
                  Index           =   5
                  Left            =   0
                  TabIndex        =   790
                  Top             =   990
                  Width           =   4950
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "排序时号源号码的比较方式"
                  Height          =   180
                  Index           =   8
                  Left            =   0
                  TabIndex        =   789
                  Top             =   1350
                  Width           =   2160
               End
            End
            Begin VB.Frame fraRegistPlanMode 
               Caption         =   "排班模式"
               Height          =   690
               Index           =   5
               Left            =   75
               TabIndex        =   189
               Top             =   135
               Width           =   8400
               Begin VB.OptionButton optRegistPlanMode 
                  Caption         =   "计划排班模式"
                  Height          =   285
                  Index           =   0
                  Left            =   255
                  TabIndex        =   178
                  Top             =   240
                  Width           =   1590
               End
               Begin VB.OptionButton optRegistPlanMode 
                  Caption         =   "出诊表排班模式"
                  Height          =   285
                  Index           =   1
                  Left            =   1935
                  TabIndex        =   179
                  Top             =   240
                  Width           =   1980
               End
               Begin MSComCtl2.DTPicker dtpRegistPlanMode 
                  Height          =   300
                  Left            =   4710
                  TabIndex        =   180
                  Top             =   225
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   529
                  _Version        =   393216
                  CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
                  Format          =   166789123
                  CurrentDate     =   42481
               End
               Begin VB.CommandButton cmdInstantActive 
                  Caption         =   "立即启用出诊表排班模式"
                  Height          =   480
                  Left            =   6900
                  TabIndex        =   181
                  Top             =   135
                  Width           =   1380
               End
               Begin VB.Label lblRegistPlanMode 
                  AutoSize        =   -1  'True
                  Caption         =   "启用日期"
                  Height          =   180
                  Left            =   3945
                  TabIndex        =   675
                  Top             =   285
                  Width           =   720
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "存在预约挂号单禁止删除安排"
               Height          =   240
               Index           =   21
               Left            =   150
               TabIndex        =   191
               Top             =   1245
               Width           =   2655
            End
            Begin VB.CheckBox chk 
               Caption         =   "仅限于对院内医生进行挂号安排"
               Height          =   270
               Index           =   17
               Left            =   150
               TabIndex        =   190
               Top             =   945
               Width           =   2820
            End
         End
         Begin VB.PictureBox pic预约 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7245
            Left            =   645
            ScaleHeight     =   7215
            ScaleWidth      =   9105
            TabIndex        =   269
            TabStop         =   0   'False
            Top             =   -60
            Width           =   9135
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   10
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   750
               Top             =   2865
               Width           =   1350
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   195
               Index           =   22
               Left            =   1815
               MaxLength       =   2
               TabIndex        =   274
               Text            =   "0"
               Top             =   705
               Width           =   660
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   195
               Index           =   21
               Left            =   2025
               MaxLength       =   2
               TabIndex        =   272
               Text            =   "0"
               Top             =   450
               Width           =   660
            End
            Begin VB.Frame fraReceiveMode 
               Caption         =   "预约接收模式"
               Height          =   615
               Left            =   120
               TabIndex        =   644
               Top             =   3990
               Width           =   6210
               Begin VB.OptionButton optReceiveMode 
                  Caption         =   "仅预约接收"
                  Height          =   255
                  Index           =   1
                  Left            =   2715
                  TabIndex        =   646
                  Top             =   255
                  Width           =   3165
               End
               Begin VB.OptionButton optReceiveMode 
                  Caption         =   "预约接收就诊"
                  Height          =   255
                  Index           =   0
                  Left            =   285
                  TabIndex        =   645
                  Top             =   255
                  Width           =   2130
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "预约显示所有号码"
               Height          =   195
               Index           =   46
               Left            =   75
               TabIndex        =   270
               Top             =   195
               Width           =   1935
            End
            Begin VB.CheckBox chk 
               Caption         =   "挂号费用以预约接收时间为准!"
               Height          =   210
               Index           =   48
               Left            =   75
               TabIndex        =   275
               Top             =   975
               Width           =   3375
            End
            Begin VB.CheckBox chk 
               Caption         =   "预约时不生成门诊号"
               Height          =   285
               Index           =   50
               Left            =   75
               TabIndex        =   276
               Top             =   1185
               Width           =   2655
            End
            Begin VB.Frame fra 
               BorderStyle     =   0  'None
               Caption         =   "默名单相关控制"
               Height          =   1350
               Index           =   0
               Left            =   135
               TabIndex        =   289
               Top             =   4545
               Width           =   8850
               Begin VB.ComboBox cbo 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  Index           =   11
                  Left            =   2955
                  Style           =   2  'Dropdown List
                  TabIndex        =   754
                  Top             =   555
                  Width           =   780
               End
               Begin VB.Frame fraSplitregister 
                  Height          =   45
                  Left            =   1335
                  TabIndex        =   325
                  Top             =   270
                  Width           =   8475
               End
               Begin VB.TextBox txt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   6
                  Left            =   1335
                  MaxLength       =   4
                  TabIndex        =   293
                  Text            =   "0"
                  Top             =   945
                  Width           =   660
               End
               Begin VB.TextBox txt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   5
                  Left            =   3750
                  MaxLength       =   4
                  TabIndex        =   291
                  Text            =   "0"
                  Top             =   615
                  Width           =   660
               End
               Begin VB.Label lblAvailabilityTimes 
                  AutoSize        =   -1  'True
                  Caption         =   "预约有效时间：预约单在预约时间                  分钟未接收的为失约！"
                  Height          =   180
                  Left            =   225
                  TabIndex        =   290
                  Top             =   615
                  Width           =   6120
               End
               Begin VB.Label lblRegisterCtl 
                  AutoSize        =   -1  'True
                  Caption         =   "黑名单相关控制"
                  Height          =   180
                  Left            =   15
                  TabIndex        =   326
                  Top             =   210
                  Width           =   1260
               End
               Begin VB.Label lblBreakAnAppointmentNums 
                  AutoSize        =   -1  'True
                  Caption         =   "病人预约失约         次自动进入黑名单"
                  Height          =   180
                  Left            =   210
                  TabIndex        =   292
                  Top             =   945
                  Width           =   3330
               End
               Begin VB.Line Line1 
                  Index           =   0
                  X1              =   1350
                  X2              =   2070
                  Y1              =   1185
                  Y2              =   1185
               End
               Begin VB.Line Line1 
                  Index           =   2
                  X1              =   3780
                  X2              =   4500
                  Y1              =   855
                  Y2              =   855
               End
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   195
               Index           =   7
               Left            =   660
               MaxLength       =   2
               TabIndex        =   279
               Text            =   "0"
               Top             =   1800
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "退号审核:N天内取消预约需要通过审核"
               Height          =   210
               Index           =   52
               Left            =   3150
               TabIndex        =   280
               Top             =   1785
               Width           =   3735
            End
            Begin VB.CheckBox chk 
               Caption         =   "预约失约用于挂号"
               Height          =   210
               Index           =   53
               Left            =   75
               TabIndex        =   277
               Top             =   1500
               Width           =   2055
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   8
               Left            =   1275
               MaxLength       =   4
               TabIndex        =   288
               Text            =   "0"
               Top             =   2505
               Width           =   540
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   9
               Left            =   1185
               MaxLength       =   4
               TabIndex        =   286
               Text            =   "0"
               Top             =   2175
               Width           =   540
            End
            Begin VB.Frame fraBespeak 
               Caption         =   "预约挂号单打印"
               Height          =   615
               Left            =   120
               TabIndex        =   281
               Top             =   3270
               Width           =   6210
               Begin VB.OptionButton optPrintBespeak 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   2715
                  TabIndex        =   284
                  Top             =   255
                  Width           =   1380
               End
               Begin VB.OptionButton optPrintBespeak 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   285
                  TabIndex        =   282
                  Top             =   255
                  Width           =   900
               End
               Begin VB.OptionButton optPrintBespeak 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   1395
                  TabIndex        =   283
                  Top             =   255
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "同一病人最多能预约        个科室"
               Height          =   210
               Index           =   47
               Left            =   75
               TabIndex        =   271
               Top             =   450
               Width           =   3705
            End
            Begin VB.CheckBox chk 
               Caption         =   "病人同一科室限约        个号"
               Height          =   210
               Index           =   172
               Left            =   75
               TabIndex        =   273
               Top             =   705
               Width           =   3735
            End
            Begin VB.Label lblAppStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "缺省预约方式"
               Height          =   180
               Left            =   75
               TabIndex        =   751
               Top             =   2925
               Width           =   1080
            End
            Begin VB.Line Line1 
               Index           =   8
               X1              =   1785
               X2              =   2505
               Y1              =   915
               Y2              =   915
            End
            Begin VB.Line Line1 
               Index           =   7
               X1              =   1980
               X2              =   2700
               Y1              =   660
               Y2              =   660
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   660
               X2              =   1380
               Y1              =   2010
               Y2              =   2010
            End
            Begin VB.Label lblCancelBespeak 
               AutoSize        =   -1  'True
               Caption         =   "预约号         天内不能取消预约"
               Height          =   180
               Left            =   75
               TabIndex        =   278
               Top             =   1800
               Width           =   2790
            End
            Begin VB.Label lblBespeakMinTime 
               AutoSize        =   -1  'True
               Caption         =   "预约限制时间        分钟：指预约时间距离现在时刻的最小间隔"
               Height          =   180
               Left            =   75
               TabIndex        =   287
               Top             =   2520
               Width           =   5220
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   1185
               X2              =   1905
               Y1              =   2745
               Y2              =   2745
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   1185
               X2              =   1905
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Label lblBespeakDefaultDays 
               AutoSize        =   -1  'True
               Caption         =   "预约缺省天数        天"
               Height          =   180
               Left            =   75
               TabIndex        =   285
               Top             =   2160
               Width           =   1980
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7395
         Index           =   7
         Left            =   -74850
         ScaleHeight     =   7365
         ScaleWidth      =   9060
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   1020
         Width           =   9090
         Begin VB.CheckBox chk 
            Caption         =   "预交款分站点显示"
            Height          =   300
            Index           =   202
            Left            =   285
            TabIndex        =   148
            Top             =   2790
            Width           =   1785
         End
         Begin VB.CheckBox chk 
            Caption         =   "禁止在院病人缴门诊预交"
            Height          =   300
            Index           =   57
            Left            =   285
            TabIndex        =   141
            Top             =   1425
            Width           =   2880
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许在院病人进行住院余额退款"
            Height          =   300
            Index           =   188
            Left            =   285
            TabIndex        =   147
            Top             =   2505
            Width           =   2925
         End
         Begin VB.Frame fra票据格式 
            Height          =   120
            Index           =   4
            Left            =   1740
            TabIndex        =   672
            Top             =   5685
            Width           =   7770
         End
         Begin VB.CheckBox chk 
            Caption         =   "退住院预交刷卡验证"
            Height          =   300
            Index           =   6
            Left            =   285
            TabIndex        =   143
            Top             =   1935
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许通过姓名模糊查找病人"
            Height          =   300
            Index           =   5
            Left            =   285
            TabIndex        =   142
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3840
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许出院病人缴住院预交"
            Height          =   300
            Index           =   4
            Left            =   285
            TabIndex        =   140
            Top             =   1155
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   "在院病人未入科不准收预交"
            Height          =   300
            Index           =   2
            Left            =   285
            TabIndex        =   137
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1290
            TabIndex        =   145
            Text            =   "10"
            Top             =   2220
            Width           =   510
         End
         Begin VB.Frame fra票据格式 
            Height          =   120
            Index           =   0
            Left            =   1755
            TabIndex        =   138
            Top             =   4200
            Width           =   7770
         End
         Begin VB.CheckBox chk 
            Caption         =   "缴预交后不清除界面信息"
            Height          =   300
            Index           =   0
            Left            =   285
            TabIndex        =   135
            Top             =   105
            Width           =   2340
         End
         Begin VB.Frame fra退款设置 
            Caption         =   "退款设置"
            Height          =   930
            Left            =   285
            TabIndex        =   150
            Top             =   3120
            Width           =   4050
            Begin VB.OptionButton optDepsoitDelSet 
               Caption         =   "余额不足时提醒退款"
               Height          =   285
               Index           =   0
               Left            =   495
               TabIndex        =   151
               Top             =   270
               Width           =   2625
            End
            Begin VB.OptionButton optDepsoitDelSet 
               Caption         =   "余额不足时禁止退款"
               Height          =   285
               Index           =   1
               Left            =   495
               TabIndex        =   152
               Top             =   555
               Value           =   -1  'True
               Width           =   2220
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许更改病人的缴款科室"
            Height          =   300
            Index           =   3
            Left            =   285
            TabIndex        =   139
            Top             =   885
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "只显示有剩余的历史缴款"
            Height          =   300
            Index           =   1
            Left            =   285
            TabIndex        =   136
            Top             =   345
            Width           =   3120
         End
         Begin VSFlex8Ctl.VSFlexGrid vs代收 
            Height          =   3840
            Left            =   4620
            TabIndex        =   153
            Top             =   195
            Width           =   4590
            _cx             =   8096
            _cy             =   6773
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483628
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParFee.frx":1B989
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   0
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   2
            Left            =   1770
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   2220
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "txtUD(2)"
            BuddyDispid     =   196647
            BuddyIndex      =   2
            OrigLeft        =   1755
            OrigTop         =   2505
            OrigRight       =   2010
            OrigBottom      =   2805
            Max             =   1000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "票据剩余         张时开始提醒收费员"
            Height          =   300
            Index           =   11
            Left            =   285
            TabIndex        =   144
            Top             =   2235
            Width           =   3450
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1050
            Index           =   0
            Left            =   315
            TabIndex        =   154
            Top             =   4530
            Width           =   8940
            _cx             =   15769
            _cy             =   1852
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParFee.frx":1B9EA
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   2
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1125
            Index           =   4
            Left            =   315
            TabIndex        =   673
            Top             =   6015
            Width           =   8940
            _cx             =   15769
            _cy             =   1984
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParFee.frx":1BA78
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   2
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblDepositPrintRedSet 
            AutoSize        =   -1  'True
            Caption         =   "预交红票打印设置"
            Height          =   180
            Left            =   285
            TabIndex        =   674
            Top             =   5715
            Width           =   1440
         End
         Begin VB.Label lblDepositPrintSet 
            AutoSize        =   -1  'True
            Caption         =   "预交票据打印设置"
            Height          =   180
            Left            =   285
            TabIndex        =   149
            Top             =   4230
            Width           =   1440
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   9810
         Index           =   11
         Left            =   -74640
         ScaleHeight     =   9780
         ScaleWidth      =   11655
         TabIndex        =   359
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   11685
         Begin VB.PictureBox picChargePg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   7620
            Index           =   0
            Left            =   60
            ScaleHeight     =   7590
            ScaleWidth      =   8445
            TabIndex        =   429
            TabStop         =   0   'False
            Top             =   -960
            Width           =   8475
            Begin VB.Frame fra退费缺省方式 
               Caption         =   "退费缺省方式"
               Height          =   1275
               Left            =   5190
               TabIndex        =   497
               Top             =   6210
               Width           =   4215
               Begin VSFlex8Ctl.VSFlexGrid vsfDelFeeDefaultType 
                  Height          =   1035
                  Left            =   60
                  TabIndex        =   498
                  Top             =   210
                  Width           =   4095
                  _cx             =   7223
                  _cy             =   1826
                  Appearance      =   2
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   0
                  HighLight       =   1
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1BB06
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   0
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   2
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.Frame fraNormal 
               Height          =   5085
               Index           =   2
               Left            =   165
               TabIndex        =   430
               Top             =   90
               Width           =   4800
               Begin VB.CheckBox chk 
                  Caption         =   "允许录入特殊使用的抗生素"
                  Height          =   195
                  Index           =   196
                  Left            =   2280
                  TabIndex        =   460
                  Top             =   4560
                  Width           =   2490
               End
               Begin VB.CheckBox chk 
                  Caption         =   "多单据分单据结算时，只对医保结算成功的单据收费"
                  Height          =   195
                  Index           =   170
                  Left            =   240
                  TabIndex        =   461
                  Top             =   4830
                  Width           =   4500
               End
               Begin VB.CheckBox chk 
                  Caption         =   "未挂号时自动加收收费项目"
                  Height          =   195
                  Index           =   119
                  Left            =   240
                  TabIndex        =   449
                  Top             =   3195
                  Width           =   2460
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H00C00000&
                  Height          =   270
                  Index           =   17
                  Left            =   2700
                  Locked          =   -1  'True
                  TabIndex        =   450
                  Top             =   3157
                  Width           =   1575
               End
               Begin VB.CommandButton cmdAddedItem 
                  Caption         =   "…"
                  Height          =   280
                  Left            =   4290
                  TabIndex        =   451
                  TabStop         =   0   'False
                  Top             =   3152
                  Width           =   280
               End
               Begin VB.CheckBox chk 
                  Caption         =   "门诊退费须先申请"
                  Height          =   195
                  Index           =   16
                  Left            =   240
                  TabIndex        =   459
                  Top             =   4560
                  Width           =   1920
               End
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   3
                  Left            =   1845
                  Style           =   2  'Dropdown List
                  TabIndex        =   456
                  Top             =   3840
                  Width           =   1170
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000012&
                  Height          =   180
                  Index           =   18
                  Left            =   1680
                  MaxLength       =   2
                  TabIndex        =   448
                  Text            =   "1"
                  Top             =   2940
                  Width           =   285
               End
               Begin VB.Frame fraLineDays 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   15
                  Index           =   3
                  Left            =   1650
                  TabIndex        =   500
                  Top             =   3120
                  Width           =   285
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   15
                  Left            =   2895
                  MaxLength       =   3
                  TabIndex        =   446
                  Text            =   "0"
                  Top             =   2655
                  Width           =   285
               End
               Begin VB.Frame fraLineDays 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   15
                  Index           =   1
                  Left            =   2865
                  TabIndex        =   499
                  Top             =   2835
                  Width           =   285
               End
               Begin VB.CheckBox chk 
                  Caption         =   "中药输入付数"
                  Height          =   195
                  Index           =   70
                  Left            =   240
                  TabIndex        =   431
                  Top             =   195
                  Value           =   1  'Checked
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "变价输入数次"
                  Height          =   195
                  Index           =   73
                  Left            =   2820
                  TabIndex        =   432
                  Top             =   195
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "开单人含护士"
                  Height          =   195
                  Index           =   76
                  Left            =   240
                  TabIndex        =   433
                  Top             =   480
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "显示收款累计"
                  Height          =   195
                  Index           =   116
                  Left            =   2820
                  TabIndex        =   434
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "提取划价单收费时检查皮试结果"
                  Height          =   195
                  Index           =   117
                  Left            =   240
                  TabIndex        =   444
                  Top             =   2325
                  Width           =   2820
               End
               Begin VB.CheckBox chk 
                  Caption         =   "优先使用预交款缴费"
                  Height          =   195
                  Index           =   118
                  Left            =   240
                  TabIndex        =   435
                  Top             =   780
                  Width           =   2040
               End
               Begin VB.CheckBox chk 
                  Caption         =   "提取划价单后立即缴款"
                  Height          =   300
                  Index           =   127
                  Left            =   240
                  TabIndex        =   437
                  Top             =   1065
                  Width           =   2160
               End
               Begin VB.CheckBox chk 
                  Caption         =   "收费时允许同时输入多张单据"
                  Height          =   195
                  Index           =   120
                  Left            =   240
                  TabIndex        =   443
                  Top             =   2040
                  Width           =   3000
               End
               Begin VB.CheckBox chk 
                  Caption         =   "收费时检查病人挂号科室"
                  Height          =   195
                  Index           =   123
                  Left            =   240
                  TabIndex        =   441
                  Top             =   1755
                  Width           =   2295
               End
               Begin VB.CheckBox chk 
                  Caption         =   "不弹出划价单选择窗口"
                  Height          =   195
                  Index           =   122
                  Left            =   240
                  TabIndex        =   439
                  Top             =   1455
                  Width           =   2160
               End
               Begin VB.TextBox txtUD 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   8
                  Left            =   1260
                  TabIndex        =   453
                  Text            =   "10"
                  Top             =   3480
                  Width           =   495
               End
               Begin VB.CheckBox chk 
                  Caption         =   "住院病人按门诊收费"
                  Height          =   270
                  Index           =   126
                  Left            =   2820
                  TabIndex        =   442
                  Top             =   1710
                  Width           =   1920
               End
               Begin VB.CheckBox chk 
                  Caption         =   "缺省科室优先"
                  Height          =   300
                  Index           =   115
                  Left            =   2820
                  TabIndex        =   438
                  Top             =   1035
                  Width           =   1620
               End
               Begin VB.CheckBox chk 
                  Caption         =   "必须要输入开单人"
                  Height          =   195
                  Index           =   114
                  Left            =   2820
                  TabIndex        =   440
                  Top             =   1425
                  Width           =   1740
               End
               Begin VB.CheckBox chk 
                  Caption         =   "不使用缺省开单人"
                  Height          =   195
                  Index           =   113
                  Left            =   2820
                  TabIndex        =   436
                  Top             =   765
                  Width           =   1740
               End
               Begin VB.TextBox txt 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Index           =   13
                  Left            =   1380
                  MaxLength       =   12
                  TabIndex        =   458
                  Text            =   "0.00"
                  Top             =   4200
                  Width           =   1335
               End
               Begin MSComCtl2.UpDown ud 
                  Height          =   300
                  Index           =   8
                  Left            =   1755
                  TabIndex        =   454
                  Top             =   3480
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   10
                  BuddyControl    =   "txtUD(8)"
                  BuddyDispid     =   196647
                  BuddyIndex      =   8
                  OrigLeft        =   6345
                  OrigTop         =   3420
                  OrigRight       =   6600
                  OrigBottom      =   3705
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.CheckBox chk 
                  Caption         =   "自动搜寻病人    天内的划价单据"
                  Height          =   195
                  Index           =   121
                  Left            =   240
                  TabIndex        =   447
                  Top             =   2910
                  Width           =   3000
               End
               Begin VB.CheckBox chk 
                  Caption         =   "票据剩余         张时开始提醒收费员"
                  Height          =   285
                  Index           =   125
                  Left            =   240
                  TabIndex        =   452
                  Top             =   3495
                  Width           =   3450
               End
               Begin VB.CheckBox chk 
                  Caption         =   "收费明细自动按              组合单据"
                  Height          =   195
                  Index           =   124
                  Left            =   240
                  TabIndex        =   455
                  Top             =   3900
                  Width           =   3570
               End
               Begin VB.CheckBox chk 
                  Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
                  Height          =   195
                  Index           =   86
                  Left            =   240
                  TabIndex        =   445
                  Top             =   2625
                  Width           =   4260
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单据最大金额"
                  Height          =   210
                  Left            =   240
                  TabIndex        =   457
                  Top             =   4245
                  Width           =   1080
               End
            End
            Begin VB.Frame fra单位 
               Caption         =   " 药品单位 "
               Height          =   810
               Index           =   1
               Left            =   165
               TabIndex        =   462
               Top             =   5325
               Width           =   4785
               Begin VB.OptionButton opt收费单位 
                  Caption         =   "售价单位"
                  Height          =   180
                  Index           =   0
                  Left            =   1365
                  TabIndex        =   464
                  Top             =   405
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton opt收费单位 
                  Caption         =   "门诊(或住院)单位"
                  Height          =   180
                  Index           =   1
                  Left            =   2505
                  TabIndex        =   465
                  Top             =   405
                  Width           =   1770
               End
               Begin VB.Label lbl单位 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "收费时按"
                  Height          =   180
                  Index           =   1
                  Left            =   300
                  TabIndex        =   463
                  Top             =   405
                  Width           =   975
               End
            End
            Begin VB.Frame fra 
               Caption         =   "汇总栏显示方式"
               Height          =   1260
               Index           =   2
               Left            =   5190
               TabIndex        =   474
               Top             =   90
               Width           =   4215
               Begin VB.OptionButton optChargeBillTotalShow 
                  Caption         =   "按单据分类汇总显示"
                  Height          =   195
                  Index           =   2
                  Left            =   300
                  TabIndex        =   477
                  Top             =   900
                  Width           =   2280
               End
               Begin VB.OptionButton optChargeBillTotalShow 
                  Caption         =   "以收入项目显示分类合计"
                  Height          =   195
                  Index           =   1
                  Left            =   300
                  TabIndex        =   476
                  Top             =   645
                  Width           =   2280
               End
               Begin VB.OptionButton optChargeBillTotalShow 
                  Caption         =   "以收据费目显示分类合计"
                  Height          =   195
                  Index           =   0
                  Left            =   300
                  TabIndex        =   475
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   2280
               End
            End
            Begin VB.Frame fraBillInputItem 
               Caption         =   "收费时要输入的项目"
               Height          =   1065
               Index           =   1
               Left            =   165
               TabIndex        =   466
               Top             =   6240
               Width           =   4785
               Begin VB.CheckBox chk 
                  Caption         =   "开单人"
                  Height          =   210
                  Index           =   109
                  Left            =   1530
                  TabIndex        =   472
                  Top             =   675
                  Value           =   1  'Checked
                  Width           =   840
               End
               Begin VB.CheckBox chk 
                  Caption         =   "开单日期"
                  Height          =   210
                  Index           =   101
                  Left            =   240
                  TabIndex        =   471
                  Top             =   675
                  Value           =   1  'Checked
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "是否加班"
                  Height          =   210
                  Index           =   100
                  Left            =   1530
                  TabIndex        =   468
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "费别"
                  Height          =   210
                  Index           =   99
                  Left            =   3810
                  TabIndex        =   470
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "年龄"
                  Height          =   210
                  Index           =   98
                  Left            =   2805
                  TabIndex        =   469
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "性别"
                  Height          =   210
                  Index           =   97
                  Left            =   240
                  TabIndex        =   467
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "医疗付款方式"
                  Height          =   210
                  Index           =   96
                  Left            =   2805
                  TabIndex        =   473
                  Top             =   675
                  Value           =   1  'Checked
                  Width           =   1380
               End
            End
            Begin VB.Frame fra库存显示 
               Caption         =   "库存显示"
               Height          =   1320
               Index           =   1
               Left            =   5190
               TabIndex        =   478
               Top             =   1410
               Width           =   4215
               Begin VB.CheckBox chk 
                  Caption         =   "显示其它药库库存"
                  Height          =   195
                  Index           =   81
                  Left            =   2250
                  TabIndex        =   480
                  Top             =   375
                  Width           =   1770
               End
               Begin VB.CheckBox chk 
                  Caption         =   "显示其它药房库存"
                  Height          =   195
                  Index           =   80
                  Left            =   300
                  TabIndex        =   479
                  Top             =   375
                  Width           =   1770
               End
               Begin VB.OptionButton opt收费库存显示方式 
                  Caption         =   "显示库存数"
                  Height          =   180
                  Index           =   0
                  Left            =   1455
                  TabIndex        =   482
                  Top             =   915
                  Width           =   1290
               End
               Begin VB.OptionButton opt收费库存显示方式 
                  Caption         =   "仅显示有无"
                  Height          =   180
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   483
                  Top             =   915
                  Width           =   1215
               End
               Begin VB.Label lbl库存显示方式 
                  AutoSize        =   -1  'True
                  Caption         =   "库存显示方式"
                  Height          =   180
                  Index           =   1
                  Left            =   300
                  TabIndex        =   481
                  Top             =   915
                  Width           =   1080
               End
               Begin VB.Line lnSplit 
                  BorderColor     =   &H80000000&
                  Index           =   3
                  X1              =   15
                  X2              =   4180
                  Y1              =   705
                  Y2              =   705
               End
               Begin VB.Line lnSplit 
                  BorderColor     =   &H00FFFFFF&
                  Index           =   2
                  X1              =   15
                  X2              =   4180
                  Y1              =   720
                  Y2              =   720
               End
            End
            Begin VB.Frame fraRegPrompt 
               Caption         =   "未挂号病人收费"
               Height          =   810
               Left            =   5190
               TabIndex        =   484
               Top             =   2790
               Width           =   4215
               Begin VB.OptionButton optChargeRegPrompt 
                  Caption         =   "允许"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   485
                  Top             =   405
                  Value           =   -1  'True
                  Width           =   795
               End
               Begin VB.OptionButton optChargeRegPrompt 
                  Caption         =   "禁止"
                  Height          =   180
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   487
                  Top             =   405
                  Width           =   795
               End
               Begin VB.OptionButton optChargeRegPrompt 
                  Caption         =   "提醒"
                  Height          =   180
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   486
                  Top             =   405
                  Width           =   795
               End
            End
            Begin VB.Frame fra 
               Caption         =   "药品摆药后退费方式"
               Height          =   810
               Index           =   17
               Left            =   5190
               TabIndex        =   488
               Top             =   3675
               Width           =   4215
               Begin VB.OptionButton optDrug 
                  Caption         =   "不检查"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   489
                  Top             =   420
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton optDrug 
                  Caption         =   "禁止"
                  Height          =   180
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   490
                  Top             =   420
                  Width           =   690
               End
               Begin VB.OptionButton optDrug 
                  Caption         =   "提醒"
                  Height          =   180
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   491
                  Top             =   420
                  Width           =   690
               End
            End
            Begin VB.Frame fra缴款控制 
               Caption         =   "缴款金额输入控制"
               Height          =   1605
               Left            =   5190
               TabIndex        =   492
               Top             =   4545
               Width           =   4215
               Begin VB.OptionButton opt缴款 
                  Caption         =   $"frmParFee.frx":1BB6C
                  Height          =   285
                  Index           =   2
                  Left            =   300
                  TabIndex        =   494
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.OptionButton opt缴款 
                  Caption         =   $"frmParFee.frx":1BB8A
                  Height          =   285
                  Index           =   0
                  Left            =   300
                  TabIndex        =   493
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   3780
               End
               Begin VB.OptionButton opt缴款 
                  Caption         =   "收费时按多病人累计"
                  Height          =   285
                  Index           =   1
                  Left            =   300
                  TabIndex        =   495
                  Top             =   870
                  Width           =   2715
               End
               Begin VB.OptionButton opt缴款 
                  Caption         =   "收费时按单病人累计"
                  Height          =   285
                  Index           =   3
                  Left            =   300
                  TabIndex        =   496
                  Top             =   1155
                  Width           =   2715
               End
            End
         End
         Begin VB.PictureBox picChargePg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   7455
            Index           =   1
            Left            =   450
            ScaleHeight     =   7425
            ScaleWidth      =   9105
            TabIndex        =   501
            TabStop         =   0   'False
            Top             =   240
            Width           =   9135
            Begin VB.Frame fra票据格式 
               Caption         =   "退费票据格式"
               Height          =   1455
               Index           =   5
               Left            =   180
               TabIndex        =   542
               Top             =   5850
               Width           =   6555
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1125
                  Index           =   5
                  Left            =   90
                  TabIndex        =   521
                  Top             =   255
                  Width           =   6375
                  _cx             =   11245
                  _cy             =   1984
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   3
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   3
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1BBA8
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   2
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "按病人补打票据不根据结算次数补打发票"
               Height          =   180
               Index           =   173
               Left            =   3195
               TabIndex        =   671
               Top             =   4290
               Width           =   3540
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   4
               ItemData        =   "frmParFee.frx":1BC3E
               Left            =   1470
               List            =   "frmParFee.frx":1BC40
               Style           =   2  'Dropdown List
               TabIndex        =   502
               Top             =   255
               Width           =   3015
            End
            Begin VB.Frame fraBillSplitRule 
               Caption         =   "票据分配规则 "
               Height          =   3900
               Left            =   195
               TabIndex        =   503
               Top             =   330
               Width           =   6555
               Begin VB.PictureBox picRuleBack 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   2580
                  Index           =   0
                  Left            =   60
                  ScaleHeight     =   2550
                  ScaleWidth      =   6000
                  TabIndex        =   525
                  TabStop         =   0   'False
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   6030
                  Begin VB.CheckBox chk 
                     Caption         =   "门诊收费时自动加收工本费"
                     Height          =   195
                     Index           =   130
                     Left            =   330
                     TabIndex        =   526
                     Top             =   855
                     Width           =   2460
                  End
                  Begin VB.CheckBox chk 
                     Caption         =   "体检病人每张单据分别打印(该参数同时影响工本费数量计算)"
                     Height          =   195
                     Index           =   189
                     Left            =   630
                     TabIndex        =   528
                     Top             =   300
                     Width           =   5190
                  End
                  Begin VB.CheckBox chk 
                     Caption         =   "门诊收费每张单据分别打印(该参数同时影响工本费数量计算)"
                     Height          =   195
                     Index           =   128
                     Left            =   345
                     TabIndex        =   527
                     Top             =   75
                     Width           =   5160
                  End
                  Begin VB.CheckBox chk 
                     Caption         =   "收费每次打印只用一张票据(该参数同时影响工本费数量计算)"
                     Height          =   195
                     Index           =   129
                     Left            =   330
                     TabIndex        =   529
                     Top             =   585
                     Width           =   5160
                  End
                  Begin VB.Frame fraActuallyPrint 
                     Height          =   1695
                     Left            =   105
                     TabIndex        =   530
                     Top             =   855
                     Width           =   5850
                     Begin VB.OptionButton optBillMode 
                        Caption         =   "打印收据费目"
                        Height          =   255
                        Index           =   0
                        Left            =   2505
                        TabIndex        =   533
                        Top             =   795
                        Value           =   -1  'True
                        Width           =   1575
                     End
                     Begin VB.OptionButton optBillMode 
                        Caption         =   "打印收费项目"
                        Height          =   255
                        Index           =   1
                        Left            =   4065
                        TabIndex        =   534
                        Top             =   795
                        Width           =   1455
                     End
                     Begin VB.CheckBox chk 
                        Caption         =   "按执行科室分别打印"
                        Height          =   195
                        Index           =   131
                        Left            =   200
                        TabIndex        =   532
                        Top             =   825
                        Width           =   1980
                     End
                     Begin VB.TextBox txtUD 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   9
                        Left            =   1305
                        Locked          =   -1  'True
                        TabIndex        =   536
                        Text            =   "3"
                        Top             =   1140
                        Width           =   405
                     End
                     Begin MSComCtl2.UpDown ud 
                        Height          =   300
                        Index           =   9
                        Left            =   1680
                        TabIndex        =   537
                        TabStop         =   0   'False
                        Top             =   1140
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   3
                        BuddyControl    =   "txtUD(9)"
                        BuddyDispid     =   196647
                        BuddyIndex      =   9
                        OrigLeft        =   1620
                        OrigTop         =   1140
                        OrigRight       =   1875
                        OrigBottom      =   1440
                        Max             =   100
                        Min             =   3
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin VB.Label lblRows 
                        AutoSize        =   -1  'True
                        Caption         =   "收费收据行次"
                        Height          =   180
                        Left            =   150
                        TabIndex        =   535
                        Top             =   1200
                        Width           =   1080
                     End
                     Begin VB.Label lbl 
                        Caption         =   "工本费数量由票据张数决定,票据张数按以下规则计算.但实际打印张数由票据数据源及票据设计决定,如果两者不一致,工本费数量将不准确."
                        Height          =   495
                        Index           =   25
                        Left            =   120
                        TabIndex        =   531
                        Top             =   285
                        Width           =   5655
                     End
                  End
               End
               Begin VB.PictureBox picRuleBack 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  ForeColor       =   &H80000008&
                  Height          =   3525
                  Index           =   1
                  Left            =   105
                  ScaleHeight     =   3495
                  ScaleWidth      =   6345
                  TabIndex        =   504
                  TabStop         =   0   'False
                  Top             =   285
                  Visible         =   0   'False
                  Width           =   6375
                  Begin VB.Frame fraRuleSystem 
                     Height          =   3345
                     Left            =   150
                     TabIndex        =   505
                     Top             =   75
                     Width           =   6150
                     Begin VB.OptionButton optRuleTotal 
                        Caption         =   "按执行科室分组汇总"
                        Height          =   240
                        Index           =   2
                        Left            =   2985
                        TabIndex        =   523
                        Top             =   2265
                        Width           =   2025
                     End
                     Begin VB.OptionButton optRuleTotal 
                        Caption         =   "首页打印汇总"
                        Height          =   240
                        Index           =   1
                        Left            =   1425
                        TabIndex        =   522
                        Top             =   2265
                        Width           =   1440
                     End
                     Begin VB.OptionButton optRuleTotal 
                        Caption         =   "不汇总"
                        Height          =   240
                        Index           =   0
                        Left            =   330
                        TabIndex        =   520
                        Top             =   2265
                        Value           =   -1  'True
                        Width           =   1005
                     End
                     Begin VB.TextBox txtBillRuleNum 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   2
                        Left            =   2430
                        Locked          =   -1  'True
                        TabIndex        =   518
                        Text            =   "3"
                        Top             =   1875
                        Width           =   315
                     End
                     Begin VB.TextBox txtBillRuleNum 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   1
                        Left            =   2430
                        Locked          =   -1  'True
                        TabIndex        =   514
                        Text            =   "3"
                        Top             =   1530
                        Width           =   315
                     End
                     Begin VB.TextBox txtBillRuleNum 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   0
                        Left            =   2430
                        Locked          =   -1  'True
                        TabIndex        =   510
                        Text            =   "3"
                        Top             =   1155
                        Width           =   315
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "4.按收费细目分页"
                        Height          =   180
                        Index           =   3
                        Left            =   285
                        TabIndex        =   516
                        Top             =   1920
                        Width           =   1770
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "3.按收据费目分页"
                        Height          =   180
                        Index           =   2
                        Left            =   270
                        TabIndex        =   512
                        Top             =   1575
                        Width           =   1770
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "2.按执行科室分页"
                        Height          =   180
                        Index           =   1
                        Left            =   270
                        TabIndex        =   508
                        Top             =   1215
                        Width           =   1770
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "1.按单据分页"
                        Height          =   225
                        Index           =   0
                        Left            =   270
                        TabIndex        =   507
                        Top             =   870
                        Width           =   1635
                     End
                     Begin MSComCtl2.UpDown updBillRuleNum 
                        Height          =   300
                        Index           =   0
                        Left            =   2760
                        TabIndex        =   511
                        TabStop         =   0   'False
                        Top             =   1155
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   1
                        BuddyControl    =   "txtBillRuleNum(0)"
                        BuddyDispid     =   196806
                        BuddyIndex      =   0
                        OrigLeft        =   4440
                        OrigTop         =   825
                        OrigRight       =   4695
                        OrigBottom      =   1125
                        Max             =   100
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin MSComCtl2.UpDown updBillRuleNum 
                        Height          =   300
                        Index           =   1
                        Left            =   2760
                        TabIndex        =   515
                        TabStop         =   0   'False
                        Top             =   1530
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   4
                        BuddyControl    =   "txtBillRuleNum(1)"
                        BuddyDispid     =   196806
                        BuddyIndex      =   1
                        OrigLeft        =   4440
                        OrigTop         =   825
                        OrigRight       =   4695
                        OrigBottom      =   1125
                        Max             =   100
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin MSComCtl2.UpDown updBillRuleNum 
                        Height          =   300
                        Index           =   2
                        Left            =   2760
                        TabIndex        =   519
                        TabStop         =   0   'False
                        Top             =   1875
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   20
                        BuddyControl    =   "txtBillRuleNum(2)"
                        BuddyDispid     =   196806
                        BuddyIndex      =   2
                        OrigLeft        =   4440
                        OrigTop         =   825
                        OrigRight       =   4695
                        OrigBottom      =   1125
                        Max             =   100
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin VB.Label lblBillRuleNum 
                        AutoSize        =   -1  'True
                        Caption         =   "：每　　   个收费细目分一页"
                        Height          =   180
                        Index           =   2
                        Left            =   2025
                        TabIndex        =   517
                        Top             =   1935
                        Width           =   2430
                     End
                     Begin VB.Label lblBillRuleNum 
                        AutoSize        =   -1  'True
                        Caption         =   "：每　　   个收据费目分一页"
                        Height          =   180
                        Index           =   1
                        Left            =   2025
                        TabIndex        =   513
                        Top             =   1590
                        Width           =   2430
                     End
                     Begin VB.Label lblBillRuleNum 
                        AutoSize        =   -1  'True
                        Caption         =   "：每　　   个执行科室分一页"
                        Height          =   180
                        Index           =   0
                        Left            =   2025
                        TabIndex        =   509
                        Top             =   1215
                        Width           =   2430
                     End
                     Begin VB.Label lblInfor 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        ForeColor       =   &H80000008&
                        Height          =   540
                        Left            =   75
                        TabIndex        =   524
                        Top             =   2700
                        Width           =   6030
                     End
                     Begin VB.Label lblRuleSystem 
                        Caption         =   "工本费数量由票据张数决定,票据张数按以下规则计算.但实际打印张数由收费划价单决定,如果手工录入费用单据,工本费的数量计算将不准确."
                        Height          =   585
                        Left            =   180
                        TabIndex        =   506
                        Top             =   330
                        Width           =   5730
                     End
                  End
               End
               Begin VB.PictureBox picRuleBack 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   1050
                  Index           =   2
                  Left            =   150
                  ScaleHeight     =   1020
                  ScaleWidth      =   6210
                  TabIndex        =   538
                  TabStop         =   0   'False
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   6240
                  Begin VB.Label lblCustomInfor 
                     Caption         =   $"frmParFee.frx":1BC42
                     Height          =   570
                     Left            =   60
                     TabIndex        =   539
                     Top             =   330
                     Width           =   6165
                  End
               End
            End
            Begin VB.Frame fra票据格式 
               Caption         =   "收费票据格式"
               Height          =   1455
               Index           =   1
               Left            =   180
               TabIndex        =   540
               Top             =   4305
               Width           =   6555
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1125
                  Index           =   1
                  Left            =   90
                  TabIndex        =   541
                  Top             =   255
                  Width           =   6375
                  _cx             =   11245
                  _cy             =   1984
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   3
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   3
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1BCE1
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   2
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.Frame fraFeeList 
               Caption         =   "收费后费用清单"
               Height          =   1290
               Left            =   7020
               TabIndex        =   543
               Top             =   330
               Width           =   2205
               Begin VB.OptionButton optFeeListPrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   345
                  TabIndex        =   544
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton optFeeListPrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   345
                  TabIndex        =   546
                  Top             =   915
                  Width           =   1455
               End
               Begin VB.OptionButton optFeeListPrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   345
                  TabIndex        =   545
                  Top             =   630
                  Width           =   1065
               End
            End
            Begin VB.Frame fraFeeExe 
               Caption         =   "收费执行单"
               Height          =   1290
               Left            =   7020
               TabIndex        =   551
               Top             =   3180
               Width           =   2205
               Begin VB.OptionButton optChargeExeBillPrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   345
                  TabIndex        =   552
                  Top             =   420
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton optChargeExeBillPrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   345
                  TabIndex        =   554
                  Top             =   930
                  Width           =   1455
               End
               Begin VB.OptionButton optChargeExeBillPrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   345
                  TabIndex        =   553
                  Top             =   660
                  Width           =   1065
               End
            End
            Begin VB.Frame fraRefundReceipt 
               Caption         =   "退费回单控制"
               Height          =   1290
               Left            =   7020
               TabIndex        =   547
               Top             =   1740
               Width           =   2205
               Begin VB.OptionButton optDelFeeRefundPrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   1
                  Left            =   360
                  TabIndex        =   549
                  Top             =   630
                  Width           =   1065
               End
               Begin VB.OptionButton optDelFeeRefundPrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   2
                  Left            =   360
                  TabIndex        =   550
                  Top             =   900
                  Width           =   1455
               End
               Begin VB.OptionButton optDelFeeRefundPrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   360
                  TabIndex        =   548
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
         End
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   975
            Index           =   1
            Left            =   150
            TabIndex        =   428
            TabStop         =   0   'False
            Top             =   105
            Width           =   5025
            _Version        =   589884
            _ExtentX        =   8864
            _ExtentY        =   1720
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7230
         Index           =   14
         Left            =   -74910
         ScaleHeight     =   7200
         ScaleWidth      =   9165
         TabIndex        =   583
         TabStop         =   0   'False
         Top             =   750
         Visible         =   0   'False
         Width           =   9195
         Begin VB.ComboBox cbo 
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   6
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   607
            Top             =   5655
            Width           =   1770
         End
         Begin VB.Frame fraPrint 
            Caption         =   "单据打印"
            Height          =   1815
            Left            =   195
            TabIndex        =   600
            Top             =   3720
            Width           =   4950
            Begin VB.CheckBox chk 
               Caption         =   "医技科室中记帐后打印单据"
               Height          =   195
               Index           =   140
               Left            =   150
               TabIndex        =   603
               Top             =   840
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "科室分散记帐中记帐后打印单据"
               Height          =   240
               Index           =   139
               Left            =   150
               TabIndex        =   602
               Top             =   570
               Width           =   2820
            End
            Begin VB.CheckBox chk 
               Caption         =   "住院记帐管理中记帐后打印单据"
               Height          =   270
               Index           =   138
               Left            =   150
               TabIndex        =   601
               Top             =   285
               Width           =   2835
            End
            Begin VB.CheckBox chk 
               Caption         =   "划价后打印划价记帐单"
               Height          =   195
               Index           =   141
               Left            =   150
               TabIndex        =   604
               Top             =   1095
               Width           =   2325
            End
            Begin VB.CheckBox chk 
               Caption         =   "划价单审核后打印记帐单"
               Height          =   195
               Index           =   142
               Left            =   150
               TabIndex        =   605
               Top             =   1365
               Width           =   2850
            End
         End
         Begin VB.Frame fraJZUnit 
            Caption         =   " 药品单位 "
            Height          =   930
            Left            =   195
            TabIndex        =   596
            Top             =   2640
            Width           =   4950
            Begin VB.OptionButton optJZDrugUnit 
               Caption         =   "住院单位"
               Height          =   180
               Index           =   1
               Left            =   2985
               TabIndex        =   599
               Top             =   435
               Width           =   1020
            End
            Begin VB.OptionButton optJZDrugUnit 
               Caption         =   "售价单位"
               Height          =   180
               Index           =   0
               Left            =   1410
               TabIndex        =   598
               Top             =   435
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl单位 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "记帐时按"
               Height          =   180
               Index           =   4
               Left            =   480
               TabIndex        =   597
               Top             =   435
               Width           =   720
            End
         End
         Begin VB.Frame fra 
            Height          =   2295
            Index           =   19
            Left            =   195
            TabIndex        =   584
            Top             =   195
            Width           =   4950
            Begin VB.CheckBox chk 
               Caption         =   "病人未入科禁止记账操作"
               Height          =   210
               Index           =   84
               Left            =   2460
               TabIndex        =   590
               Top             =   880
               Width           =   2340
            End
            Begin VB.CheckBox chk 
               Caption         =   "必须输入开单人"
               Height          =   210
               Index           =   18
               Left            =   2460
               TabIndex        =   586
               Top             =   262
               Width           =   1590
            End
            Begin VB.CheckBox chk 
               Caption         =   "可以输入其它科室的开单人"
               Height          =   210
               Index           =   19
               Left            =   2460
               TabIndex        =   588
               Top             =   571
               Width           =   2460
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   0
               Left            =   1470
               Style           =   2  'Dropdown List
               TabIndex        =   595
               Top             =   1800
               Width           =   1140
            End
            Begin VB.CheckBox chk 
               Caption         =   "欠费时允许保存为划价单"
               Height          =   195
               Index           =   146
               Left            =   195
               TabIndex        =   593
               Top             =   1498
               Width           =   2400
            End
            Begin VB.CheckBox chk 
               Caption         =   "中药可以输入付数"
               Height          =   195
               Index           =   143
               Left            =   195
               TabIndex        =   585
               Top             =   270
               Value           =   1  'Checked
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "开单人中包含护士"
               Height          =   195
               Index           =   145
               Left            =   195
               TabIndex        =   589
               Top             =   884
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "变价允许输入数次"
               Height          =   195
               Index           =   144
               Left            =   195
               TabIndex        =   587
               Top             =   577
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药库库存"
               Height          =   195
               Index           =   148
               Left            =   2460
               TabIndex        =   592
               Top             =   1191
               Width           =   1850
            End
            Begin VB.CheckBox chk 
               Caption         =   "显示其它药房库存"
               Height          =   195
               Index           =   147
               Left            =   195
               TabIndex        =   591
               Top             =   1191
               Width           =   1845
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "已结记帐单操作"
               Height          =   180
               Index           =   1
               Left            =   195
               TabIndex        =   594
               Top             =   1860
               Width           =   1260
            End
         End
         Begin VB.Label lbl发药 
            AutoSize        =   -1  'True
            Caption         =   "记帐之后"
            Height          =   180
            Left            =   240
            TabIndex        =   606
            Top             =   5715
            Width           =   720
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7680
         Index           =   3
         Left            =   -74850
         ScaleHeight     =   7650
         ScaleWidth      =   9030
         TabIndex        =   377
         Top             =   105
         Visible         =   0   'False
         Width           =   9060
         Begin VB.OptionButton optInExseCharge 
            Caption         =   "所有住院划价费用"
            Height          =   180
            Index           =   1
            Left            =   4725
            TabIndex        =   29
            Top             =   7395
            Width           =   1860
         End
         Begin VB.OptionButton optInExseCharge 
            Caption         =   "仅本次住院划价费用"
            Height          =   180
            Index           =   0
            Left            =   6690
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   7395
            Width           =   2070
         End
         Begin VB.Frame fra 
            Height          =   45
            Index           =   20
            Left            =   -15
            TabIndex        =   669
            Top             =   7260
            Width           =   10000
         End
         Begin VB.CommandButton cmdWarnDel 
            Caption         =   "删除报警方案(&D)"
            Height          =   350
            Left            =   7845
            TabIndex        =   26
            Top             =   6825
            Width           =   1590
         End
         Begin VB.CommandButton cmdWarnNew 
            Caption         =   "增加报警方案(&A)"
            Height          =   350
            Left            =   7845
            TabIndex        =   25
            Top             =   6465
            Width           =   1590
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊或住院记帐报警包含划价费用"
            Height          =   255
            Index           =   41
            Left            =   90
            TabIndex        =   28
            Top             =   7365
            Width           =   3090
         End
         Begin VB.ListBox lst类别 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2130
            Left            =   2895
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   975
            Visible         =   0   'False
            Width           =   1530
         End
         Begin ZL9BillEdit.BillEdit Bill 
            Height          =   5370
            Index           =   1
            Left            =   210
            TabIndex        =   22
            Top             =   765
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   9472
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
         Begin MSComctlLib.TabStrip tab报警 
            Height          =   5955
            Left            =   90
            TabIndex        =   21
            Top             =   405
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   10504
            HotTracking     =   -1  'True
            TabMinWidth     =   0
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "普通病人"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInExseCharge 
            AutoSize        =   -1  'True
            Caption         =   "住院记帐报警"
            Height          =   180
            Left            =   3555
            TabIndex        =   670
            Top             =   7395
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmParFee.frx":1BDAF
            Height          =   555
            Left            =   90
            TabIndex        =   31
            Top             =   6495
            Width           =   7740
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "报警方案:每种方案包括各病区报警线及报警方式，需和 zl_PatiWarnScheme 函数配合使用"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   14
            Left            =   210
            TabIndex        =   27
            Top             =   120
            Width           =   7200
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6795
         Index           =   15
         Left            =   -74760
         ScaleHeight     =   6765
         ScaleWidth      =   8940
         TabIndex        =   686
         TabStop         =   0   'False
         Top             =   720
         Width           =   8970
         Begin VB.PictureBox picSettlePar 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   8415
            Index           =   0
            Left            =   825
            ScaleHeight     =   8385
            ScaleWidth      =   9270
            TabIndex        =   688
            TabStop         =   0   'False
            Top             =   75
            Visible         =   0   'False
            Width           =   9300
            Begin VB.Frame fraOrder 
               Caption         =   "缺省冲预交设置"
               Height          =   2520
               Index           =   0
               Left            =   180
               TabIndex        =   744
               Top             =   4830
               Width           =   9075
               Begin VB.CommandButton cmdDepositUp 
                  Caption         =   "↑"
                  Height          =   510
                  Left            =   8280
                  TabIndex        =   749
                  Top             =   1065
                  Width           =   375
               End
               Begin VB.CommandButton cmdDepositDown 
                  Caption         =   "↓"
                  Height          =   510
                  Left            =   8280
                  TabIndex        =   748
                  Top             =   1665
                  Width           =   375
               End
               Begin VB.OptionButton optOrder 
                  Caption         =   "按结算类别进行冲预交"
                  Height          =   300
                  Index           =   1
                  Left            =   135
                  TabIndex        =   746
                  Top             =   495
                  Width           =   3585
               End
               Begin VB.OptionButton optOrder 
                  Caption         =   "按缴款时间先后冲预交"
                  Height          =   300
                  Index           =   0
                  Left            =   135
                  TabIndex        =   745
                  Top             =   225
                  Width           =   3645
               End
               Begin VSFlex8Ctl.VSFlexGrid vsDepositSort 
                  Height          =   1635
                  Left            =   165
                  TabIndex        =   747
                  Top             =   810
                  Width           =   8010
                  _cx             =   14129
                  _cy             =   2884
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483634
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   2
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   5
                  Cols            =   5
                  FixedRows       =   2
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   300
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1BE90
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   4
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   8
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   2
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.Frame fraBalanceFeeDate 
               Caption         =   "结帐费用期间设置"
               ForeColor       =   &H00000000&
               Height          =   930
               Left            =   5205
               TabIndex        =   719
               Top             =   2145
               Width           =   2025
               Begin VB.OptionButton optBalanceTime 
                  Caption         =   "按登记时间"
                  Height          =   195
                  Index           =   0
                  Left            =   195
                  TabIndex        =   721
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1320
               End
               Begin VB.OptionButton optBalanceTime 
                  Caption         =   "按发生时间"
                  Height          =   195
                  Index           =   1
                  Left            =   195
                  TabIndex        =   720
                  Top             =   600
                  Width           =   1320
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "出院结帐代收款检查"
               Height          =   900
               Index           =   0
               Left            =   5205
               TabIndex        =   716
               Top             =   120
               Width           =   2025
               Begin VB.OptionButton optBalanceDSCheck 
                  Caption         =   "禁止"
                  Height          =   195
                  Index           =   0
                  Left            =   195
                  TabIndex        =   718
                  Top             =   285
                  Width           =   870
               End
               Begin VB.OptionButton optBalanceDSCheck 
                  Caption         =   "提示"
                  Height          =   195
                  Index           =   1
                  Left            =   195
                  TabIndex        =   717
                  Top             =   585
                  Value           =   -1  'True
                  Width           =   870
               End
            End
            Begin VB.Frame fraBalaceBlood 
               Caption         =   "结帐时输血费检查"
               Height          =   900
               Left            =   5205
               TabIndex        =   713
               Top             =   1125
               Width           =   2025
               Begin VB.OptionButton optBlood 
                  Caption         =   "检查并提示"
                  Height          =   210
                  Index           =   1
                  Left            =   195
                  TabIndex        =   715
                  Top             =   570
                  Width           =   1305
               End
               Begin VB.OptionButton optBlood 
                  Caption         =   "不检查"
                  Height          =   210
                  Index           =   0
                  Left            =   195
                  TabIndex        =   714
                  Top             =   300
                  Value           =   -1  'True
                  Width           =   945
               End
            End
            Begin VB.Frame fraMzDepositDefaultUse 
               Caption         =   "门诊预交缺省使用方式"
               Height          =   1275
               Left            =   5205
               TabIndex        =   709
               Top             =   3120
               Width           =   4065
               Begin VB.OptionButton optMzDeposit 
                  Caption         =   "按结帐金额使用预交"
                  Height          =   350
                  Index           =   1
                  Left            =   195
                  TabIndex        =   712
                  Top             =   540
                  Width           =   2256
               End
               Begin VB.OptionButton optMzDeposit 
                  Caption         =   "不使用预交款"
                  Height          =   350
                  Index           =   0
                  Left            =   195
                  TabIndex        =   711
                  Top             =   255
                  Width           =   1524
               End
               Begin VB.OptionButton optMzDeposit 
                  Caption         =   "使用剩余所有预交款"
                  Height          =   350
                  Index           =   2
                  Left            =   195
                  TabIndex        =   710
                  Top             =   840
                  Value           =   -1  'True
                  Width           =   2028
               End
            End
            Begin VB.Frame fraOwnFeeType 
               Caption         =   "结帐先结自费类别"
               Height          =   2940
               Left            =   7335
               TabIndex        =   707
               Top             =   150
               Width           =   1950
               Begin VB.ListBox lst 
                  Height          =   2580
                  Index           =   4
                  Left            =   75
                  Style           =   1  'Checkbox
                  TabIndex        =   708
                  Top             =   315
                  Width           =   1770
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "结帐缴款控制"
               Height          =   1245
               Index           =   3
               Left            =   180
               TabIndex        =   703
               Top             =   3495
               Width           =   4800
               Begin VB.OptionButton optBalancePayin 
                  Caption         =   "不进行缴款控制"
                  Height          =   180
                  Index           =   0
                  Left            =   255
                  TabIndex        =   706
                  Top             =   330
                  Value           =   -1  'True
                  Width           =   1770
               End
               Begin VB.OptionButton optBalancePayin 
                  Caption         =   "存在收取现金时,必须输入缴款"
                  Height          =   180
                  Index           =   1
                  Left            =   255
                  TabIndex        =   705
                  Top             =   615
                  Width           =   2835
               End
               Begin VB.OptionButton optBalancePayin 
                  Caption         =   "按单病人累计"
                  Height          =   180
                  Index           =   2
                  Left            =   255
                  TabIndex        =   704
                  Top             =   900
                  Width           =   1470
               End
            End
            Begin VB.Frame frabalance 
               Height          =   2970
               Index           =   4
               Left            =   180
               TabIndex        =   690
               Top             =   435
               Width           =   4800
               Begin VB.CheckBox chk 
                  Caption         =   "三方卡结帐退款默认退现"
                  Height          =   225
                  Index           =   200
                  Left            =   120
                  TabIndex        =   788
                  Top             =   2265
                  Width           =   3195
               End
               Begin VB.CheckBox chk 
                  Caption         =   "自费费用结帐缺省使用预交"
                  Height          =   255
                  Index           =   157
                  Left            =   120
                  TabIndex        =   701
                  Top             =   1660
                  Width           =   2640
               End
               Begin VB.CheckBox chk 
                  Caption         =   "医保先结自费费用不打印结帐票据"
                  Height          =   210
                  Index           =   156
                  Left            =   120
                  TabIndex        =   700
                  Top             =   1367
                  Width           =   3105
               End
               Begin VB.CheckBox chk 
                  Caption         =   "结帐检查病历接收情况"
                  Height          =   255
                  Index           =   155
                  Left            =   120
                  TabIndex        =   699
                  Top             =   751
                  Width           =   2190
               End
               Begin VB.CheckBox chk 
                  Caption         =   "结帐后不清除界面信息"
                  Height          =   225
                  Index           =   154
                  Left            =   2490
                  TabIndex        =   698
                  Top             =   458
                  Width           =   2175
               End
               Begin VB.CheckBox chk 
                  Caption         =   "中途结帐缺省退预交款"
                  Height          =   195
                  Index           =   153
                  Left            =   120
                  TabIndex        =   697
                  Top             =   473
                  Width           =   2160
               End
               Begin VB.CheckBox chk 
                  Caption         =   "仅使用指定住院次数的预交款"
                  Height          =   195
                  Index           =   152
                  Left            =   120
                  TabIndex        =   696
                  Top             =   1089
                  Width           =   2760
               End
               Begin VB.CheckBox chk 
                  Caption         =   "对病人的零费用进行结帐"
                  Height          =   195
                  Index           =   151
                  Left            =   2490
                  TabIndex        =   695
                  Top             =   195
                  Width           =   2280
               End
               Begin VB.CheckBox chk 
                  Caption         =   "病人出院结帐后自动出院"
                  Height          =   195
                  Index           =   150
                  Left            =   120
                  TabIndex        =   694
                  Top             =   195
                  Width           =   2280
               End
               Begin VB.CheckBox chk 
                  Caption         =   "合约单位结帐每位病人分别打印票据"
                  Height          =   225
                  Index           =   149
                  Left            =   120
                  TabIndex        =   693
                  Top             =   1998
                  Width           =   3195
               End
               Begin VB.CheckBox chk 
                  Caption         =   "在院病人不允许出院结帐"
                  Height          =   210
                  Index           =   55
                  Left            =   2490
                  TabIndex        =   692
                  Top             =   773
                  Width           =   2280
               End
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   5
                  ItemData        =   "frmParFee.frx":1BFB6
                  Left            =   1320
                  List            =   "frmParFee.frx":1BFB8
                  Style           =   2  'Dropdown List
                  TabIndex        =   691
                  Top             =   2565
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "未审单据结帐"
                  Height          =   180
                  Index           =   24
                  Left            =   120
                  TabIndex        =   702
                  Top             =   2625
                  Width           =   1080
               End
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   7
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   689
               Top             =   135
               Width           =   2550
            End
            Begin VB.Label lblUnit 
               AutoSize        =   -1  'True
               Caption         =   "合约单位结帐使用                             的票据"
               Height          =   180
               Left            =   210
               TabIndex        =   722
               Top             =   195
               Width           =   4590
            End
         End
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   975
            Index           =   2
            Left            =   0
            TabIndex        =   687
            TabStop         =   0   'False
            Top             =   0
            Width           =   5025
            _Version        =   589884
            _ExtentX        =   8864
            _ExtentY        =   1720
            _StockProps     =   64
         End
         Begin VB.PictureBox picSettlePar 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7275
            Index           =   2
            Left            =   -360
            ScaleHeight     =   7245
            ScaleWidth      =   9780
            TabIndex        =   760
            TabStop         =   0   'False
            Top             =   3720
            Visible         =   0   'False
            Width           =   9810
            Begin VB.Frame fraColor 
               Caption         =   "结帐缴款界面字体颜色"
               Height          =   2085
               Left            =   6270
               TabIndex        =   767
               Top             =   3885
               Width           =   3285
               Begin VB.PictureBox pic缴款栏退款色 
                  BackColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   772
                  Top             =   1665
                  Width           =   705
               End
               Begin VB.PictureBox pic当前付款未退色 
                  BackColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   771
                  Top             =   915
                  Width           =   705
               End
               Begin VB.PictureBox pic缴款栏缴款色 
                  BackColor       =   &H00FF0000&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   770
                  Top             =   1305
                  Width           =   705
               End
               Begin VB.PictureBox pic当前付款未付色 
                  BackColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   769
                  Top             =   555
                  Width           =   705
               End
               Begin VB.PictureBox pic自付合计色 
                  BackColor       =   &H00FF0000&
                  Height          =   300
                  Left            =   1635
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   768
                  Top             =   225
                  Width           =   705
               End
               Begin VB.Label lbl结帐Color 
                  AutoSize        =   -1  'True
                  Caption         =   "退款颜色         "
                  Height          =   180
                  Index           =   4
                  Left            =   1665
                  TabIndex        =   777
                  Top             =   1725
                  Width           =   1530
               End
               Begin VB.Label lbl结帐Color 
                  AutoSize        =   -1  'True
                  Caption         =   "未退颜色         "
                  Height          =   180
                  Index           =   2
                  Left            =   1665
                  TabIndex        =   776
                  Top             =   975
                  Width           =   1530
               End
               Begin VB.Label lbl结帐Color 
                  AutoSize        =   -1  'True
                  Caption         =   "缴款栏字体颜色:缴款颜色         "
                  Height          =   180
                  Index           =   3
                  Left            =   315
                  TabIndex        =   775
                  Top             =   1365
                  Width           =   2880
               End
               Begin VB.Label lbl结帐Color 
                  AutoSize        =   -1  'True
                  Caption         =   "当前付款字体颜色:未付颜色         "
                  Height          =   180
                  Index           =   0
                  Left            =   135
                  TabIndex        =   774
                  Top             =   600
                  Width           =   3060
               End
               Begin VB.Label lbl结帐Color 
                  AutoSize        =   -1  'True
                  Caption         =   "自付合计字体颜色"
                  Height          =   180
                  Index           =   1
                  Left            =   135
                  TabIndex        =   773
                  Top             =   270
                  Width           =   1440
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "病人多次结帐弹出结帐条件窗体"
               Height          =   195
               Index           =   191
               Left            =   2010
               TabIndex        =   765
               Top             =   3615
               Width           =   4125
            End
            Begin VB.PictureBox picDisplay 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   3240
               Index           =   0
               Left            =   1590
               Picture         =   "frmParFee.frx":1BFBA
               ScaleHeight     =   3210
               ScaleWidth      =   4515
               TabIndex        =   763
               Top             =   3975
               Width           =   4545
            End
            Begin VB.OptionButton opt界面风格 
               Caption         =   "结帐缴款风格"
               Height          =   525
               Index           =   1
               Left            =   75
               TabIndex        =   762
               Top             =   5333
               Width           =   1515
            End
            Begin VB.OptionButton opt界面风格 
               Caption         =   "传统结帐风格"
               Height          =   525
               Index           =   0
               Left            =   75
               TabIndex        =   761
               Top             =   1560
               Width           =   1455
            End
            Begin VB.PictureBox picDisplay 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   3240
               Index           =   1
               Left            =   1590
               Picture         =   "frmParFee.frx":22690
               ScaleHeight     =   3210
               ScaleWidth      =   4515
               TabIndex        =   764
               Top             =   195
               Width           =   4545
            End
         End
         Begin VB.PictureBox picSettlePar 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6570
            Index           =   1
            Left            =   4500
            ScaleHeight     =   6540
            ScaleWidth      =   9780
            TabIndex        =   723
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   9810
            Begin VB.Frame fra票据格式 
               Caption         =   "结帐红票打印设置"
               Height          =   1620
               Index           =   7
               Left            =   4875
               TabIndex        =   742
               Top             =   2235
               Width           =   4515
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1275
                  Index           =   7
                  Left            =   105
                  TabIndex        =   743
                  Top             =   255
                  Width           =   4320
                  _cx             =   7620
                  _cy             =   2249
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   3
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   3
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":27BC0
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   2
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.Frame fra票据格式 
               Caption         =   "结帐票据打印设置"
               Height          =   1740
               Index           =   3
               Left            =   4875
               TabIndex        =   740
               Top             =   240
               Width           =   4515
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1350
                  Index           =   3
                  Left            =   105
                  TabIndex        =   741
                  Top             =   255
                  Width           =   4305
                  _cx             =   7594
                  _cy             =   2381
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   3
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   3
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":27C2D
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   0
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   0
                  OutlineCol      =   0
                  Ellipsis        =   0
                  ExplorerBar     =   2
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   0
                  ShowComboButton =   1
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   3
                  AutoSizeMouse   =   -1  'True
                  FrozenRows      =   0
                  FrozenCols      =   0
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "病人退款收据打印方式"
               Height          =   510
               Index           =   1
               Left            =   75
               TabIndex        =   736
               Top             =   2235
               Width           =   4725
               Begin VB.OptionButton optDelBalancePrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   2
                  Left            =   1500
                  TabIndex        =   739
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.OptionButton optDelBalancePrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   1
                  Left            =   2850
                  TabIndex        =   738
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.OptionButton optDelBalancePrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   255
                  TabIndex        =   737
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   900
               End
            End
            Begin VB.Frame fraOwnFee 
               Caption         =   "自费费用清单打印"
               Height          =   480
               Left            =   75
               TabIndex        =   732
               Top             =   240
               Width           =   4725
               Begin VB.OptionButton optOwnFee 
                  Caption         =   "不打印"
                  Height          =   255
                  Index           =   0
                  Left            =   285
                  TabIndex        =   735
                  Top             =   195
                  Width           =   1050
               End
               Begin VB.OptionButton optOwnFee 
                  Caption         =   "自动打印"
                  Height          =   255
                  Index           =   1
                  Left            =   1500
                  TabIndex        =   734
                  Top             =   195
                  Width           =   1050
               End
               Begin VB.OptionButton optOwnFee 
                  Caption         =   "选择是否打印"
                  Height          =   255
                  Index           =   2
                  Left            =   2850
                  TabIndex        =   733
                  Top             =   195
                  Width           =   1395
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "结帐费用明细打印方式"
               Height          =   495
               Index           =   2
               Left            =   75
               TabIndex        =   728
               Top             =   885
               Width           =   4725
               Begin VB.OptionButton optBalanceFeeListPrint 
                  Caption         =   "自动打印"
                  Height          =   180
                  Index           =   2
                  Left            =   1500
                  TabIndex        =   731
                  Top             =   225
                  Width           =   1065
               End
               Begin VB.OptionButton optBalanceFeeListPrint 
                  Caption         =   "选择是否打印"
                  Height          =   180
                  Index           =   1
                  Left            =   2850
                  TabIndex        =   730
                  Top             =   225
                  Width           =   1455
               End
               Begin VB.OptionButton optBalanceFeeListPrint 
                  Caption         =   "不打印"
                  Height          =   180
                  Index           =   0
                  Left            =   255
                  TabIndex        =   729
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   900
               End
            End
            Begin VB.Frame fraDeposit 
               Caption         =   "预交票据打印"
               Height          =   510
               Left            =   75
               TabIndex        =   724
               Top             =   1545
               Width           =   4725
               Begin VB.OptionButton optBalanceDepositPrint 
                  Caption         =   "不打印"
                  Height          =   255
                  Index           =   0
                  Left            =   255
                  TabIndex        =   727
                  Top             =   210
                  Width           =   1110
               End
               Begin VB.OptionButton optBalanceDepositPrint 
                  Caption         =   "自动打印"
                  Height          =   255
                  Index           =   1
                  Left            =   1500
                  TabIndex        =   726
                  Top             =   210
                  Width           =   1335
               End
               Begin VB.OptionButton optBalanceDepositPrint 
                  Caption         =   "选择是否打印"
                  Height          =   255
                  Index           =   2
                  Left            =   2850
                  TabIndex        =   725
                  Top             =   210
                  Width           =   1395
               End
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmParFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '参数与控件对应记录集（同一个参数可能对应一组多个控件）
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnInstantActive As Boolean
Private mblnNotChange As Boolean
Private mblnExistPrintData As Boolean '收费存在打印数据
Private mint原票据分配规则 As Integer
Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_在院病人不准出院结帐 = 55
    
    chk_输入开单人 = 18
    chk_它科开单人 = 19
    chk_未入科禁止记账 = 84
    
    chk_病人姓名 = 7
    chk_病人ID = 8
    chk_刷就诊卡 = 9
    chk_挂号单号 = 10
    
    chk_出院病人允许门诊转住院 = 183
    chk_门诊费用转住院预交票据打印控制 = 197
    
    chk_专家号挂号限制 = 178
    chk_专家号预约限制 = 179
    chk_启用免挂号模式 = 290
    
    chk_预约排队按时点 = 186
        
    chk_收费项目首位当类别简码 = 56
    chk_首先输入收费类别 = 25
    chk_门诊退费须先申请 = 16
    chk_从属项目汇总计算折扣 = 39
    chk_指定发料部门时不显示无库存卫材 = 204
    
    '一卡通
    chk_项目执行前必须收费或审核 = 67
    chk_项目开单后立即收费或记帐审核 = 90
    
    
    chk_票号控制 = 13
    
    chk_记帐报警包含划价费用 = 41
    
    chk_自动修正 = 12
    chk_下午算半天模式 = 43
    
    chk_挂号必须刷卡 = 102
    chk_优先使用预交 = 103
    chk_允许住院病人挂号 = 104
    chk_随机序号选择 = 105
    chk_姓名模糊查找 = 106
    chk_挂号包含科室安排 = 107
    chk_预约包含科室安排 = 184
    chk_预约时收款 = 108
    chk_医生站输入医生 = 169
    
    chk_附费_备货卫材输入条码 = 187
    
    
    '预交款管理
    chk_缴预交后不清除界面信息 = 0
    chk_只显示有剩余的历史缴款 = 1
    chk_在院病人未入科不准收预交 = 2
    chk_允许更改病人的缴款科室 = 3
    chk_允许出院病人缴住院预交 = 4
    chk_禁止在院病人缴门诊预交 = 57
    chk_允许通过姓名模糊查找病人 = 5
    chk_退住院预交刷卡验证 = 6
    chk_票据剩余N张提醒操作员 = 11
    chk_允许在院病人余额退款 = 188
    chk_预交款分站点显示 = 202
    
    '医疗卡管理
    chk_卡费以记账方式收取 = 14
    chk_姓名允许模糊查找 = 15
    chk_卡费使用门诊收费医疗收据 = 192
    chk_收取病历费 = 199
    chk_医疗卡_自动生成门诊号 = 201
    
    '挂号安排
    chk_只对医内医生进行挂号安排 = 17
    chk_存在预约挂号单禁止删除安排 = 21
    '挂号管理
    chk_挂号_自动刷新挂号安排 = 22
    chk_挂号_自动生成门诊号 = 23
    chk_挂号_存为划价单 = 24
    chk_挂号_卡费与挂号一起收 = 26
    chk_挂号_无医生必须输入医生 = 27
    chk_挂号_发卡自动生成新病人 = 29
    chk_挂号_优先使用预交款缴费 = 30
    chk_挂号_姓名模糊查找 = 31
    chk_挂号_零费用打印票据 = 32
    chk_挂号_输入_姓名 = 33
    chk_挂号_输入_性别 = 34
    chk_挂号_输入_年龄 = 35
    chk_挂号_输入_家庭地址 = 36
    chk_挂号_输入_付款方式 = 37
    chk_挂号_输入_费别 = 38
    chk_挂号_输入_结算方式 = 40
    chk_挂号_输入_联系电话 = 20
    chk_挂号_发卡不弹病人登记窗口 = 42
    chk_挂号_退号不退卡重打票据 = 44
    chk_挂号_挂号后打印病历标签 = 45
    chk_挂号_预约显示所有号别 = 46
    chk_挂号_预约接收确定挂号费 = 48
    chk_挂号_允许住院病人挂号 = 49
    chk_挂号_预约不生成门诊号 = 50
    chk_挂号_随机序号选择 = 51
    chk_挂号_N天内退号需审核 = 52
    chk_挂号_预约失约用于挂号 = 53
    chk_挂号_家庭地址联想输入 = 54
    chk_挂号_扫描身份证签约 = 58
    chk_挂号_已退序号允许挂号 = 60
    chk_挂号_严格按时段挂号 = 61
    chk_挂号_默认勾选病历选项 = 62
    chk_挂号_门诊号有效性检查 = 63
    chk_挂号_非严格控制为发卡 = 64
    chk_挂号_病人预约科室数 = 47
    chk_挂号_病人同科限约N个号 = 172
    chk_挂号_病人挂号科室限制 = 176
    chk_挂号_病人同科限挂N个号 = 174
    chk_挂号_病人同科限挂N个号_急诊 = 175
    chk_挂号_计划排班模式精简界面 = 185
    chk_挂号_禁止输入年龄 = 190
    chk_挂号_同一身份证只能对应一个建档病人 = 28
    chk_挂号_病人同一号源限挂N个号 = 203
    
    '分诊
    chk_分诊_分诊台签到开始排队 = 65
    chk_分诊_预约挂号进入队列 = 66
    chk_分诊_诊室忙允许分诊 = 68
    chk_分诊_回诊签到 = 171
    chk_分诊_重新签到 = 177
    
    '临床出诊安排
    chk_出诊安排_替诊医生级别检查 = 182
    chk_出诊安排_按替诊医生同步更新预约挂号单 = 180
    chk_出诊安排_只允许选院内医生 = 181
    
    '门诊划价
    chk_划价_门诊输入中药付数 = 69
    chk_划价_门诊变价输入数次 = 74
    chk_划价_门诊开单人含护士 = 75
    chk_划价_门诊显示其它药房库存 = 78
    chk_划价_门诊显示其它药库库存 = 79
    chk_划价_门诊姓名模糊查找 = 85
    chk_划价_门诊输入_性别 = 88
    chk_划价_门诊输入_是否加班 = 89
    chk_划价_门诊输入_年龄 = 91
    chk_划价_门诊输入_费别 = 92
    chk_划价_门诊输入_开单日期 = 93
    chk_划价_门诊输入_开单人 = 94
    chk_划价_门诊输入_医疗付款 = 95
    chk_划价_门诊不缺省开单人 = 110
    chk_划价_门诊必须要输入开单人 = 111
    chk_划价_门诊缺省科室优先 = 112
    chk_划价_住院病人按门诊收费 = 168
    chk_划价_允许录入特殊使用的抗生素 = 194
    
    
    '门诊收费
    chk_收费_门诊输入中药付数 = 70
    chk_收费_门诊变价输入数次 = 73
    chk_收费_门诊开单人含护士 = 76
    chk_收费_门诊显示其它药房库存 = 80
    chk_收费_门诊显示其它药库库存 = 81
    chk_收费_门诊姓名模糊查找 = 86
    chk_收费_门诊输入_性别 = 97
    chk_收费_门诊输入_是否加班 = 100
    chk_收费_门诊输入_年龄 = 98
    chk_收费_门诊输入_费别 = 99
    chk_收费_门诊输入_开单日期 = 101
    chk_收费_门诊输入_开单人 = 109
    chk_收费_门诊输入_医疗付款 = 96
    chk_收费_门诊不缺省开单人 = 113
    chk_收费_门诊必须要输入开单人 = 114
    chk_收费_门诊缺省科室优先 = 115
    chk_收费_门诊显示收款累计 = 116
    chk_收费_提划价单检查皮试结果 = 117
    chk_收费_优先使用预交款缴费 = 118
    chk_收费_未挂号自动加收挂号费 = 119
    chk_收费_允许输入多张单据 = 120
    chk_收费_搜寻划价单据 = 121
    chk_收费_不弹出划价单选择窗口 = 122
    chk_收费_检查病人挂号科室 = 123
    chk_收费_自动组合单据 = 124
    chk_收费_票据剩余X张开始提醒 = 125
    chk_收费_住院按门诊收费 = 126
    chk_收费_提取划价立即缴款 = 127
    chk_收费_多张单据收费分别打印 = 128
    chk_收费_每次只用一张票据 = 129
    chk_收费_自动加收工本费 = 130
    chk_收费_票据生成方式 = 131
    chk_收费_只对医保结算成功单据收费 = 170
    chk_收费_按病人不打票据不分次数 = 173
    chk_收费_体检病人按单据分别打印 = 189
    chk_收费_允许录入特殊使用的抗生素 = 196
    
    '门诊记帐
    chk_记帐_门诊输入中药付数 = 71
    chk_记帐_门诊变价输入数次 = 72
    chk_记帐_门诊开单人含护士 = 77
    chk_记帐_门诊显示其它药房库存 = 82
    chk_记帐_门诊显示其它药库库存 = 83
    chk_记帐_门诊姓名模糊查找 = 87
    chk_记帐_只查找合约单位病人 = 132
    chk_记帐_记帐打印 = 133
    chk_记帐_划价打印 = 134
    chk_记帐_审核打印 = 135
    chk_记帐_允许录入特殊使用的抗生素 = 195
    
    '补结算
    chk_补结算_姓名模糊查找 = 136
    chk_补结算_票据剩余X张开始提醒 = 137
    '住院记帐相关
    chk_住院记帐_记帐打印 = 138
    chk_分散记帐_记帐打印 = 139
    chk_医技科室_记帐打印 = 140
    chk_记帐操作_划价打印 = 141
    chk_记帐操作_审核打印 = 142
    chk_记帐操作_中药输入付数 = 143
    chk_记帐操作_开单人包含护士 = 145
    chk_记帐操作_变价输入数次 = 144
    chk_记帐操作_欠费保存划价单 = 146
    chk_记帐操作_显示其它药房库存 = 147
    chk_记帐操作_显示其它药库库存 = 148

    
    '病人结帐管理
    chk_结帐_合约单位按病人打印 = 149
    chk_结帐_出院结帐后自动出院 = 150
    chk_结帐_处理零费用 = 151
    chk_结帐_使用指定次数预交款 = 152
    chk_结帐_中途结帐缺省退预交款 = 153
    chk_结帐_结帐后不清除界面信息 = 154
    chk_结帐_结帐检查病历接收情况 = 155
    chk_结帐_自费费用不打印结帐票据 = 156
    chk_结帐_自费费用缺省使用预交 = 157
    chk_结帐_病人多次结帐弹出结帐条件窗体 = 191
    chk_结帐_三方卡结帐退款控制 = 200
    
    '执行登记管理
    chk_执行登记_显示医嘱发送的单据 = 158   '公共费用
    
    '执行登记管理
    chk_费用审核_门诊转住院先审核 = 159   '公共费用
    '票据使用监控
    chk_票据监控_领用票据签字确认 = 160
    '人员借款管理
    chk_借款_申请打印 = 161
    chk_借款_借出打印 = 162
    '消费卡管理
    chk_消费卡_缴款单打印 = 163
    chk_消费卡_刷卡消费须定位到密码框 = 193
    chk_消费卡_消费卡退费刷卡控制 = 59
    
    '医嘱附费管理
    chk_附费_中药输入付数 = 164
    chk_附费_变价输入数次 = 165
    chk_附费_显示其它药房库存 = 166
    chk_附费_显示其它药库库存 = 167
    
    '收费员轧账管理
    chk_轧账_预交轧帐按门诊或住院分别轧帐 = 198
    
End Enum


Private Enum constCbo
       
    cbo_已结单据 = 0
    cbo_病人审核方式 = 1
    cbo_挂号_缺省排序方式 = 2
    cbo_收费_自动组合单据 = 3
    cbo_收费_票据分配规则 = 4
    cbo_未审单据结帐 = 5
    cbo_记帐操作_记帐后发药 = 6
    cbo_结帐_合约单位结帐打印 = 7
    cbo_一卡通_收票票据格式 = 8
    cbo_一卡通_记帐票据格式 = 9
    
    cbo_临床安排_全院通用号源安排站点 = 18
    cbo_临床安排_号码比较方式 = 19
    cbo_挂号_缺省预约方式 = 10
    cbo_挂号_预约有效时间 = 11
    
    cbo_挂号零钱处理 = 12
    cbo_收费零钱处理 = 13
    cbo_结帐零钱处理 = 14
    cbo_消费卡零钱处理 = 17
    cbo_自动记帐模式 = 15
    cbo_门诊费用转住院预交发票格式 = 16
End Enum

Private Enum constUpDown
    ud_挂号单 = 1
    ud_急诊挂号单 = 10
    
    ud_挂号预约天数 = 0
    
    ud_费用金额保留位数 = 5
    ud_费用单价保留位数 = 6
    
    ud_号码长度 = 4
    
    '预交相关
    ud_票据张数 = 2
    
    '分诊
    ud_分诊有效天数 = 3
    
    '划价
    ud_划价_取消N天的划价单 = 7
    
    '收费
    ud_收费_票据张数 = 8
    ud_收费_收费收据总行次 = 9
    
    '补结算
    ud_补结算_票据张数 = 11
    ud_补结算_有效天数 = 12
    
    
End Enum

Private Enum constOpt

    '预交款管理:退款设置
    opt_余额不足时提醒退款 = 0
    opt_余额不足时禁止退款 = 1

End Enum

Private Enum constBill
    bill_自动计算 = 0
    bill_记帐报警 = 1
End Enum

Private Enum constLvw
    lvw_票据 = 0
    lvw_单据 = 1
    lvw_一卡通 = 3
End Enum

Private Enum constListBox
    lst_医保病人 = 0
    lst_公费病人 = 1
    lst_刷卡密码 = 3
    lst_补结算_结算方式 = 2
    lst_结帐_自费费用类别 = 4
    lst_脱机医保结算方式 = 5
End Enum

Private Enum constTxt
    txt_单笔最大金额 = 0
    txt_姓名查找天数 = 106
    txt_专家号挂号限制 = 4
    txt_专家号预约限制 = 24
    txt_姓名模糊查找天数 = 1
    txt_挂号_刷新时间 = 2
    txt_挂号_姓名查找天数 = 3
    txt_挂号_预约有效时间 = 5
    txt_挂号_预约失效次数 = 6
    txt_挂号_N天内不能取消预约号 = 7
    txt_挂号_预约限制时间_分钟 = 8
    txt_挂号_预约限制时间_天数 = 9
    txt_挂号_N岁以下输入监护人 = 10
    txt_分诊_提前N小时分诊 = 11
    
    txt_划价_门诊单据最大金额 = 12
    txt_划价_门诊姓名模糊查找天数 = 14
    txt_收费_门诊单据最大金额 = 13
    txt_收费_门诊姓名模糊查找天数 = 15
    txt_记帐_门诊姓名模糊查找天数 = 16
    txt_收费_自助加收挂号费 = 17
    txt_收费_搜寻划价单据天数 = 18
    txt_补结算_姓名模糊查找天数 = 19
    
    txt_挂号_病人同科限挂N个号 = 23
    txt_挂号_病人挂号科室限制 = 20
    txt_挂号_病人同科限约N个号 = 22
    txt_挂号_病人预约科室数 = 21
    txt_挂号_病人同一号源限挂N个号 = 25
    
    txt_免密支付 = 26
    
End Enum
Private Enum constVsGridBill
    vsGrid_预交票据格式 = 0
    vsGrid_收费票据格式 = 1
    vsGrid_补结算票据格式 = 2
    vsGrid_结帐票据格式 = 3
    vsGrid_预交红票格式 = 4
    vsGrid_退费票据格式 = 5
    vsGrid_补结算退费票据格式 = 6
    vsGrid_结帐红票格式 = 7
    vsGrid_发卡预交票据格式 = 8
    vsGrid_医疗卡收据格式 = 9
End Enum
Private Enum constVsGridInputItemSet
    vsGrid_发卡输入项设置 = 0
End Enum

Private Enum constTbPage
    Pg_挂号业务 = 0
    Pg_门诊收费 = 1
    Pg_结帐业务 = 2
End Enum
Private Enum constTbPageItemID
    Pg_挂号_安排 = 100
    Pg_挂号_挂号 = 101
    Pg_挂号_预约 = 102
    Pg_挂号_其他 = 103
    Pg_收费_单据控制 = 204
    Pg_收费_票据控制 = 205
    Pg_结帐_结帐参数 = 300
    Pg_结帐_票据控制 = 301
    Pg_结帐_界面风格 = 302
End Enum


'自动计算设置中保存当前行和列
Private mintCurRow As Integer
Private mintCurCol As Integer
Private mblnJRaiseByDate As Boolean     '判断床位类项目及从属项目是否按日调价
Private mblnHRaiseByDate As Boolean     '判断护理类项目及从属项目是否按日调价
Private mstrDel适用病人 As String           '记录记帐报警中删除的适用病人类型

Private mintColumn As Integer '

Private mrsWarn As ADODB.Recordset
Private mrs类别 As ADODB.Recordset
Private mrsBillUseType As ADODB.Recordset
Private mblnOK As Boolean


Private Sub chkBillRule_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkBillRule_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chkBillRule, Index, mrsPar)
End Sub

 

Private Sub cmdDepClearAll_Click()
    Dim i As Integer
    If chk(chk_分诊_分诊台签到开始排队).value = 0 Then Exit Sub
    With vsfTriageQueuingDep
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("启用")) = 0
            If .RowData(i) <> .TextMatrix(i, .ColIndex("启用")) Then
                .Cell(flexcpForeColor, i, .ColIndex("科室")) = vbRed
            Else
                .Cell(flexcpForeColor, i, .ColIndex("科室")) = &H80000008
            End If
        Next
    End With
End Sub

Private Sub cmdDepSelectAll_Click()
    Dim i As Integer
    If chk(chk_分诊_分诊台签到开始排队).value = 0 Then Exit Sub
    With vsfTriageQueuingDep
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("启用")) = 1
            If .RowData(i) <> .TextMatrix(i, .ColIndex("启用")) Then
                .Cell(flexcpForeColor, i, .ColIndex("科室")) = vbRed
            Else
                .Cell(flexcpForeColor, i, .ColIndex("科室")) = &H80000008
            End If
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1802", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    End Select
End Sub

Private Sub cmdInstantActive_Click()
    Dim datTime As Date, rsCheck As ADODB.Recordset
    Dim strSQL As String, strValue As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    strSQL = "Select 1 From 临床出诊表 Where 发布时间 Is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        MsgBox "不存在任何临床出诊数据,不能切换为出诊表排班模式!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    datTime = zlDatabase.Currentdate
    If MsgBox("注意:" & vbCrLf & "将要立即切换至出诊表排班模式,已经存在的预约记录将会自动修正至对应的新的出诊表排班中,修正数据需要一定时间,请耐心等待" & _
                vbCrLf & "请确保已经按照计划排班数据生成了出诊表排班数据,如果预约记录没有找到对应的出诊表排班,立即启用将会失败!" & _
                vbCrLf & "是否立即启用出诊表排班模式?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    strSQL = "Zl_出诊表挂号_Turn(To_Date('" & Format(datTime, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ")"
    Me.MousePointer = 11
    zlDatabase.ExecuteProcedure strSQL, "修正预约记录到出诊表排班"
    Me.MousePointer = 0
    
    mblnInstantActive = True
    optRegistPlanMode(0).value = 0
    optRegistPlanMode(1).value = 1
    optRegistPlanMode(0).Enabled = False
    optRegistPlanMode(1).Enabled = False
    dtpRegistPlanMode.Enabled = False
    dtpRegistPlanMode.value = datTime
    dtpRegistPlanMode.Enabled = False
    cmdInstantActive.Enabled = False
    mblnInstantActive = False
    
    strValue = "1|" & Format(dtpRegistPlanMode.value, "yyyy-mm-dd hh:mm:ss")
    zlDatabase.SetPara "挂号排班模式", strValue, glngSys
    
    fraNewPaln.Visible = True
    chk(chk_只对医内医生进行挂号安排).Visible = False
    chk(chk_存在预约挂号单禁止删除安排).Visible = False
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdStationRegOrder_Click(Index As Integer)
    Dim strValue As String, i As Integer
    With vsStationRegSort
        If Index = 0 Then
            If .Row <= 1 Then Exit Sub
            .RowPosition(.Row) = .Row - 1
            .Row = .Row - 1
        Else
            If .Row >= .Rows - 1 Then Exit Sub
            .RowPosition(.Row) = .Row + 1
            .Row = .Row + 1
        End If
    End With
    With vsStationRegSort
        For i = 1 To 5
            strValue = strValue & "|" & .TextMatrix(i, .ColIndex("排序字段")) & "," & IIF(.TextMatrix(i, .ColIndex("是否升序")) = -1, 1, 0)
        Next i
        strValue = Mid(strValue, 2)
    End With
    Call SetParChange(vsStationRegSort, 0, mrsPar, True, strValue)
End Sub

Private Sub dtpRegistTime_LostFocus()
    Call SetParChange(dtpRegistTime, 0, mrsPar, True, Format(dtpRegistTime.value, "HH:mm:ss"))
End Sub

Private Sub dtpRegistTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(dtpRegistTime, 0, mrsPar)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "初始成功" Then
        '不知道为何，使用源代码的情况下将页签先缺省后不会触发Form_Activate事件，编译后是会触发的
        '为了在源代码的情况下也能触发Form_Activate事件，将页签缺省调整到Form_Activate事件中
        tbPage(Pg_挂号业务).Item(1).Selected = True   '缺省为第一个,主要是常针对挂号窗口进行设置
        tbPage(Pg_门诊收费).Item(0).Selected = True
        tbPage(Pg_结帐业务).Item(0).Selected = True
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    Dim objPic As PictureBox
    
    For Each objPic In picPar
        Set objPic.Container = Me
    Next
    
    strCategory = "参数设置,基础项目"
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "100,0,费用公共参数"
    marrFunc(0) = marrFunc(0) & ";108,1,一卡通业务"
    marrFunc(0) = marrFunc(0) & ";112,8,医疗卡业务"
    marrFunc(0) = marrFunc(0) & ";107,7,预交款业务"
    marrFunc(0) = marrFunc(0) & ";106,6,病人挂号业务"
    marrFunc(0) = marrFunc(0) & ";115,9,分诊台业务"
    marrFunc(0) = marrFunc(0) & ";101,10,门诊划价管理"
    marrFunc(0) = marrFunc(0) & ";113,11,门诊收费管理"
    marrFunc(0) = marrFunc(0) & ";109,12,门诊记帐管理"
    marrFunc(0) = marrFunc(0) & ";114,13,保险补充结算"
    marrFunc(0) = marrFunc(0) & ";110,14,住院记帐业务"
    marrFunc(0) = marrFunc(0) & ";115,15,病人结帐管理"
    marrFunc(0) = marrFunc(0) & ";111,16,财务监控业务"
    marrFunc(0) = marrFunc(0) & ";116,17,医嘱附费管理"
    
    marrFunc(1) = "105,5,费用基础配置;102,2,操作员限制;103,3,记帐报警;104,4,费用自动计算"
    
    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    mblnOK = False
    Me.Tag = "初始成功"
End Sub

Private Sub lblAvailabilityTimes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_预约有效时间, mrsPar)
End Sub

Private Sub lblBespeakDefaultDays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_预约限制时间_天数, mrsPar)
End Sub
Private Sub lblBespeakMinTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_预约限制时间_分钟, mrsPar)
End Sub

 

Private Sub lblBreakAnAppointmentNums_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_预约失效次数, mrsPar)
End Sub

Private Sub lblCancelBespeak_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_N天内不能取消预约号, mrsPar)
End Sub

Private Sub lblDeptNums_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_病人预约科室数, mrsPar)
End Sub


Private Sub lblGuardian_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_挂号_N岁以下输入监护人, mrsPar)
End Sub
Private Sub optBalanceDepositPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceDepositPrint, Index, mrsPar)
End Sub

Private Sub optBalanceDepositPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalanceDepositPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceDepositPrint, Index, mrsPar)
End Sub
Private Sub optBalanceDSCheck_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceDSCheck, Index, mrsPar)
End Sub

Private Sub optBalanceDSCheck_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalanceDSCheck_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceDSCheck, Index, mrsPar)
End Sub

Private Sub optBalancePayin_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalancePayin, Index, mrsPar)
End Sub

Private Sub optBalancePayin_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalancePayin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalancePayin, Index, mrsPar)
End Sub
Private Sub optBalanceTime_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceTime, Index, mrsPar)
End Sub

Private Sub optBalanceTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalanceTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceTime, Index, mrsPar)
End Sub


Private Sub optBillMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBillMode, Index, mrsPar)
End Sub

Private Sub optBillMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBillMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBillMode, Index, mrsPar)
End Sub

 
 
Private Sub optBalanceFeeListPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceFeeListPrint, Index, mrsPar)
End Sub

Private Sub optBalanceFeeListPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub optBalanceFeeListPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceFeeListPrint, Index, mrsPar)
End Sub


Private Sub optBlood_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBlood, Index, mrsPar)
End Sub

Private Sub optBlood_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub optBlood_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBlood, Index, mrsPar)
End Sub


Private Sub optBrushCard_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Dim strValue As String
    txt(txt_免密支付).Enabled = optBrushCard(3).value = True
    If optBrushCard(1).value Then
        strValue = 1
    ElseIf optBrushCard(2).value Then
        strValue = 2
    ElseIf optBrushCard(3).value Then
        If Val(txt(txt_免密支付).Text) = 0 Then Exit Sub
        strValue = -1 * Val(txt(txt_免密支付).Text)
    Else
        strValue = 0
    End If
    strValue = strValue & "|" & IIF(optBrushCard(11).value, 1, IIF(optBrushCard(12).value, 2, 0))
    Call SetParChange(optBrushCard, Index, mrsPar, True, strValue)
End Sub

Private Sub optBrushCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBrushCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBrushCard, Index, mrsPar)
End Sub

Private Sub optChargeExeBillPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optChargeExeBillPrint, Index, mrsPar)
End Sub

Private Sub optChargeExeBillPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    'zlCommFun.PressKey vbKeyTab
End Sub
Private Sub optChargeExeBillPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optChargeExeBillPrint, Index, mrsPar)
End Sub
Private Sub optChargeRegPrompt_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optChargeRegPrompt, Index, mrsPar)
End Sub

Private Sub optChargeRegPrompt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optChargeRegPrompt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optChargeRegPrompt, Index, mrsPar)
End Sub

 
Private Sub optDelBalancePrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDelBalancePrint, Index, mrsPar)
End Sub

Private Sub optDelBalancePrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDelBalancePrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDelBalancePrint, Index, mrsPar)
End Sub

Private Sub optDelFeeRefundPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDelFeeRefundPrint, Index, mrsPar)
End Sub

Private Sub optDelFeeRefundPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDelFeeRefundPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDelFeeRefundPrint, Index, mrsPar)
End Sub

 

Private Sub optDrugSupplementary_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDrugSupplementary, Index, mrsPar)
End Sub

Private Sub optDrugSupplementary_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDrugSupplementary_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDrugSupplementary, Index, mrsPar)
End Sub
Private Sub optDrugUnitFF_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDrugUnitFF, Index, mrsPar)
End Sub

Private Sub optDrugUnitFF_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDrugUnitFF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDrugUnitFF, Index, mrsPar)
End Sub

Private Sub optFeeListPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optFeeListPrint, Index, mrsPar)
End Sub

Private Sub optFeeListPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optFeeListPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optFeeListPrint, Index, mrsPar)
End Sub

Private Sub optJZDrugUnit_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optJZDrugUnit, Index, mrsPar)
End Sub

Private Sub optJZDrugUnit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optJZDrugUnit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optJZDrugUnit, Index, mrsPar)
End Sub
Private Sub optMzDeposit_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optMzDeposit, Index, mrsPar)
End Sub

Private Sub optMzDeposit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optMzDeposit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optMzDeposit, Index, mrsPar)
End Sub

Private Sub optOrder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optOrder, Index, mrsPar)
End Sub

Private Sub opt界面风格_Click(Index As Integer)
    Dim i As Integer
    If opt界面风格(0).value = True Then
        chk(chk_结帐_病人多次结帐弹出结帐条件窗体).Enabled = True
        fraColor.Enabled = False
        For i = 0 To 4
            lbl结帐Color(i).Enabled = False
        Next i
        Call SetParChange(opt界面风格, Index, mrsPar, True, "0")
    Else
        chk(chk_结帐_病人多次结帐弹出结帐条件窗体).Enabled = False
        fraColor.Enabled = True
        For i = 0 To 4
            lbl结帐Color(i).Enabled = True
        Next i
        Call SetParChange(opt界面风格, Index, mrsPar, True, "1")
    End If
End Sub

Private Sub opt界面风格_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt界面风格, Index, mrsPar)
End Sub

Private Sub optOrder_Click(Index As Integer)
    Dim strValue As String, i As Integer
    Dim intIndex As Integer
    If optOrder(0).value = True Then
        vsDepositSort.Enabled = False
        cmdDepositDown.Enabled = False
        cmdDepositUp.Enabled = False
        Call SetParChange(optOrder, Index, mrsPar, True, "0")
    Else
        vsDepositSort.Enabled = True
        cmdDepositDown.Enabled = True
        cmdDepositUp.Enabled = True
        strValue = "1|"
        With vsDepositSort
            For i = 2 To 4
                If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
                If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
                If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
                If i <> 4 Then
                    strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
                Else
                    strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
                End If
            Next i
        End With
        Call SetParChange(optOrder, Index, mrsPar, True, strValue)
    End If
End Sub

Private Sub optOwnFee_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optOwnFee, Index, mrsPar)
End Sub

Private Sub optOwnFee_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optOwnFee_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optOwnFee, Index, mrsPar)
End Sub

Private Sub optPrintMode_SendCard_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintMode_SendCard, Index, mrsPar)
End Sub

Private Sub optPrintMode_SendCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintMode_SendCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintMode_SendCard, Index, mrsPar)
End Sub
Private Sub optPrintModeDraw_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintModeDraw, Index, mrsPar)
End Sub

Private Sub optPrintModeDraw_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintModeDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintModeDraw, Index, mrsPar)
End Sub
Private Sub optPrintModeSJ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintModeSJ, Index, mrsPar)
End Sub

Private Sub optPrintModeSJ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintModeSJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintModeSJ, Index, mrsPar)
End Sub

Private Sub optPrintRequisition_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintRequisition, Index, mrsPar)
End Sub

Private Sub optPrintRequisition_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintRequisition_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintRequisition, Index, mrsPar)
End Sub


Private Sub optRegistClearMzInfor_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optRegistClearMzInfor, Index, mrsPar)
End Sub

Private Sub optRegistClearMzInfor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegistClearMzInfor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegistClearMzInfor, Index, mrsPar)
End Sub

Private Sub optDelCardMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDelCardMode, Index, mrsPar)
End Sub

Private Sub optDelCardMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDelCardMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDelCardMode, Index, mrsPar)
End Sub
 
Private Sub optRegistPrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optRegistPrintMode, Index, mrsPar)
End Sub

Private Sub optRegPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optRegPrint, Index, mrsPar)
End Sub

Private Sub optRegistPrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegistPrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegistPrintMode, Index, mrsPar)
End Sub

Private Sub optRegPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegPrint, Index, mrsPar)
End Sub
 

Private Sub optRuleTotal_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SaveBillRuleChange
End Sub

Private Sub optRuleTotal_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRuleTotal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRuleTotal, Index, mrsPar)
End Sub
Private Sub optSendDrugFF_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSendDrugFF, Index, mrsPar)
End Sub

Private Sub optSendDrugFF_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
End Sub

Private Sub optSendDrugFF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSendDrugFF, Index, mrsPar)
End Sub

Private Sub optSetMoneyMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSetMoneyMode, Index, mrsPar)
End Sub

Private Sub optSetMoneyMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSetMoneyMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSetMoneyMode, Index, mrsPar)
End Sub

Private Sub optSlipPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSlipPrint, Index, mrsPar)
End Sub

Private Sub optSlipPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSlipPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSlipPrint, Index, mrsPar)
End Sub

Private Sub optMoneyControl_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optMoneyControl, Index, mrsPar)
End Sub

Private Sub optMoneyControl_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optMoneyControl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optMoneyControl, Index, mrsPar)
End Sub

Private Sub optBarCodePrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBarCodePrint, Index, mrsPar)
End Sub

Private Sub optBarCodePrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBarCodePrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBarCodePrint, Index, mrsPar)
End Sub


Private Sub optPrintBespeak_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintBespeak, Index, mrsPar)
End Sub

Private Sub optPrintBespeak_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintBespeak_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintBespeak, Index, mrsPar)
End Sub

Private Sub optReceiveMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optReceiveMode, Index, mrsPar)
End Sub

Private Sub optReceiveMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optReceiveMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optReceiveMode, Index, mrsPar)
End Sub


Private Sub optSupplementaryPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSupplementaryPrint, Index, mrsPar)
End Sub

Private Sub optSupplementaryPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSupplementaryPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSupplementaryPrint, Index, mrsPar)
End Sub


Private Sub optSupplementaryUnit_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSupplementaryUnit, Index, mrsPar)
End Sub

Private Sub optSupplementaryUnit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSupplementaryUnit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSupplementaryUnit, Index, mrsPar)
End Sub



Private Sub optTriageBarcodePrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriageBarcodePrintMode, Index, mrsPar)
End Sub

Private Sub optTriageBarcodePrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriageBarcodePrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriageBarcodePrintMode, Index, mrsPar)
End Sub

Private Sub optTriageQueuingMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriageQueuingMode, Index, mrsPar)
End Sub

Private Sub optTriageQueuingMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 1 Then
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriageQueuingMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriageQueuingMode, Index, mrsPar)
End Sub
Private Sub optTriageSort_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriageSort, Index, mrsPar)
End Sub

Private Sub optTriageSort_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriageSort_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriageSort, Index, mrsPar)
End Sub
Private Sub optTriagePrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriagePrintMode, Index, mrsPar)
End Sub

Private Sub optTriagePrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriagePrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriagePrintMode, Index, mrsPar)
End Sub
Private Sub opt划价单位_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt划价单位, Index, mrsPar)
End Sub

Private Sub opt划价单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt划价单位_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt划价单位, Index, mrsPar)
End Sub

Private Sub optBillTotalShow_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBillTotalShow, Index, mrsPar)
End Sub

Private Sub optBillTotalShow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    
    'zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBillTotalShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBillTotalShow, Index, mrsPar)
End Sub
Private Sub optChargeBillTotalShow_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optChargeBillTotalShow, Index, mrsPar)
End Sub

Private Sub optChargeBillTotalShow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optChargeBillTotalShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optChargeBillTotalShow, Index, mrsPar)
End Sub
Private Sub optDrug_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDrug, Index, mrsPar)
End Sub
Private Sub optDrug_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDrug_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDrug, Index, mrsPar)
End Sub

Private Sub opt缴款_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt缴款, Index, mrsPar)
End Sub

Private Sub opt缴款_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    tbPage(Pg_门诊收费).Item(1).Selected = True
    'zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt缴款_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt缴款, Index, mrsPar)
End Sub

Private Sub opt收费库存显示方式_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt收费库存显示方式, Index, mrsPar)
End Sub

Private Sub opt收费库存显示方式_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt收费库存显示方式_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt收费库存显示方式, Index, mrsPar)
End Sub
Private Sub opt记帐库存显示方式_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt记帐库存显示方式, Index, mrsPar)
End Sub

Private Sub opt记帐库存显示方式_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    'zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt记帐库存显示方式_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt记帐库存显示方式, Index, mrsPar)
End Sub

Private Sub opt划价库存显示方式_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt划价库存显示方式, Index, mrsPar)
End Sub

Private Sub opt划价库存显示方式_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt划价库存显示方式_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt划价库存显示方式, Index, mrsPar)
End Sub


Private Sub opt收费单位_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt收费单位, Index, mrsPar)
End Sub

Private Sub opt收费单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt收费单位_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt收费单位, Index, mrsPar)
End Sub

Private Sub opt记帐单位_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt记帐单位, Index, mrsPar)
End Sub

Private Sub opt记帐单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt记帐单位_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt记帐单位, Index, mrsPar)
End Sub

Private Sub picDisplay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(picDisplay, Index, mrsPar)
End Sub

Private Sub picPar_Resize(Index As Integer)
    If Index = 6 Then
        'index=6:挂号业务
        '挂号业务
        With tbPage(Pg_挂号业务)
            .Left = picPar(Index).ScaleLeft
            .Top = picPar(Index).ScaleTop
            .Height = picPar(Index).ScaleHeight
            .Width = picPar(Index).ScaleWidth
        End With
    End If
    If Index = 11 Then
        With tbPage(Pg_门诊收费)
            .Left = picPar(Index).ScaleLeft
            .Top = picPar(Index).ScaleTop
            .Height = picPar(Index).ScaleHeight
            .Width = picPar(Index).ScaleWidth
        End With
    End If
    If Index = 15 Then
        With tbPage(Pg_结帐业务)
            .Left = picPar(Index).ScaleLeft
            .Top = picPar(Index).ScaleTop
            .Height = picPar(Index).ScaleHeight
            .Width = picPar(Index).ScaleWidth
        End With
    End If
End Sub

Private Sub tbPage_SelectedChanged(Index As Integer, ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub
    Dim objTemp As Object
    Select Case Index
    Case Pg_挂号业务
        With tbPage(Index)
            Select Case Val(.Selected.Tag)
            Case Pg_挂号_安排
                If optRegistPlanMode(0).value = True Then
                    Set objTemp = chk(chk_只对医内医生进行挂号安排)
                Else
                    Set objTemp = chk(chk_出诊安排_只允许选院内医生)
                End If
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_挂号_挂号
                Set objTemp = chk(chk_挂号_自动刷新挂号安排)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_挂号_其他
                Set objTemp = chk(chk_挂号必须刷卡)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_挂号_预约
                Set objTemp = chk(chk_挂号_预约显示所有号别)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            End Select
        End With
    Case Pg_门诊收费
        With tbPage(Index)
            Select Case Val(.Selected.Tag)
            Case Pg_收费_单据控制
               Set objTemp = chk(chk_收费_门诊输入中药付数)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_收费_票据控制
                Set objTemp = cbo(cbo_收费_票据分配规则)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            End Select
        End With
    Case Else
    End Select
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("记帐报警", marrFunc) Or Item.ID = GetFuncID("费用自动计算", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '用于获取当前选中的TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For i = 0 To picPar.UBound
        picPar(i).Top = Me.ScaleTop
        picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
        picPar(i).Width = Me.ScaleWidth - picPar(i).Left
        picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub

Private Sub mshAutoCalc_GotFocus()
    If lblLocate(txt_Dept).Tag <> "mshAutoCalc" Then
        lblLocate(txt_Dept).Tag = "mshAutoCalc"
        mlngPreFind = 1
    End If
End Sub

Private Sub Bill_GotFocus(Index As Integer)
    If Index = bill_自动计算 Then
        If lblLocate(txt_Dept).Tag <> "Bill" Then
            lblLocate(txt_Dept).Tag = "Bill"
            mlngPreFind = 1
        End If
    End If
End Sub

Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'功能：根据ID选中一级和二级分类
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
                Next
            End If
        Next
    Next
End Sub

Private Sub InitData()
'功能：初始化界面控件,读取并加载数据

    '1.初始化变量
    
    mlngPreFind = 1
    mblnJRaiseByDate = IsRaiseByDate("J")
    mblnHRaiseByDate = IsRaiseByDate("H")
        
    Call InitSystemPara
    
    
    
    '2.初始化界面控件
    Call InitEnv
        
    Call Load费用类型
    Call LoadOneCard
    Call Load病区
    Call Load单据操作
    
    Call LoadOther
    
    RestoreFlexState mshAutoCalc, App.ProductName & "\" & Me.Name
    RestoreFlexState Bill(bill_自动计算), App.ProductName & "\" & Me.Name & bill_自动计算
    RestoreFlexState Bill(bill_记帐报警), App.ProductName & "\" & Me.Name & bill_记帐报警
    
    
    '3.加载系统参数
    Call LoadPar
    
    
End Sub

Private Sub LoadPar()
'功能：读取并加载参数到界面控件
    Dim strValue As String, strTmp As String
    Dim i As Long, n As Long, blnFind As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String, rsTemp As ADODB.Recordset
    Dim arrObj As Variant  '数组对象：模块1,参数号1,控件对象1,模块2,参数号2,控件对象2,......
    Dim varData As Variant, varTemp As Variant
    Dim strArr() As String, strStyle() As String
    Dim intTemp As Integer
    
    Set rsTmp = GetPar(mrsPar, "9000,1103,1107,1110,1111,1113,1114,1120,1121,1122,1124,1133,1134,1135,1150,1137,1142,1143,1151,1500,1501,1502,1503,1504,1506,1257")
     '1.设置CheckBox类参数
    strTmp = "0:7:" & chk_自动修正 & _
            ",0:31:" & chk_在院病人不准出院结帐 & _
            ",0:52:" & chk_输入开单人 & _
            ",0:53:" & chk_它科开单人 & _
            ",0:72:" & chk_首先输入收费类别 & _
            ",0:93:" & chk_从属项目汇总计算折扣 & _
            ",0:270:" & chk_预约排队按时点 & _
            ",0:98:" & chk_记帐报警包含划价费用 & _
            ",0:276:" & chk_消费卡_刷卡消费须定位到密码框 & _
            ",0:282:" & chk_消费卡_消费卡退费刷卡控制 & _
            ",0:279:" & chk_挂号_同一身份证只能对应一个建档病人 & _
            ",0:290:" & chk_启用免挂号模式 & _
            ",0:316:" & chk_指定发料部门时不显示无库存卫材
            
    strTmp = strTmp & _
            ",0:100:" & chk_下午算半天模式 & _
            ",0:144:" & chk_收费项目首位当类别简码 & _
            ",0:151:" & chk_门诊退费须先申请 & _
            ",0:163:" & chk_项目执行前必须收费或审核 & _
            ",0:232:" & chk_项目开单后立即收费或记帐审核 & _
            ",0:215:" & chk_未入科禁止记账 & _
            ",0:283:" & chk_门诊费用转住院预交票据打印控制
    
    strTmp = strTmp & _
            "," & p费用虚拟模块 & ":挂号必须刷卡:" & chk_挂号必须刷卡 & _
            "," & p费用虚拟模块 & ":优先使用预交:" & chk_优先使用预交 & _
            "," & p费用虚拟模块 & ":允许住院病人挂号:" & chk_允许住院病人挂号 & _
            "," & p费用虚拟模块 & ":随机序号选择:" & chk_随机序号选择 & _
            "," & p费用虚拟模块 & ":预约时收款:" & chk_预约时收款 & _
            "," & p费用虚拟模块 & ":姓名模糊查找:" & chk_姓名模糊查找 & _
            "," & p费用虚拟模块 & ":输入医生:" & chk_医生站输入医生
    
    '预交相关Check参数
    strTmp = strTmp & _
        "," & p预交款管理 & ":仅显有余款的缴款单:" & chk_只显示有剩余的历史缴款 & _
        "," & p预交款管理 & ":允许更改缴款科室:" & chk_允许更改病人的缴款科室 & _
        "," & p预交款管理 & ":缴预交后不清除信息:" & chk_缴预交后不清除界面信息 & _
        "," & p预交款管理 & ":病人未入科不准收预交:" & chk_在院病人未入科不准收预交 & _
        "," & p预交款管理 & ":允许出院病人缴住院预交:" & chk_允许出院病人缴住院预交 & _
        "," & p预交款管理 & ":姓名模糊查找:" & chk_允许通过姓名模糊查找病人 & _
        "," & p预交款管理 & ":住院退预交验证:" & chk_退住院预交刷卡验证 & _
        "," & p预交款管理 & ":允许在院病人余额退款:" & chk_允许在院病人余额退款 & _
        "," & p预交款管理 & ":禁止在院病人缴门诊预交:" & chk_禁止在院病人缴门诊预交 & _
        "," & p预交款管理 & ":预交款分站点显示:" & chk_预交款分站点显示

    '医疗卡相关Check参数
    strTmp = strTmp & _
        "," & p医疗卡管理 & ":卡费记帐:" & chk_卡费以记账方式收取 & _
        "," & p医疗卡管理 & ":姓名模糊查找:" & chk_姓名允许模糊查找 & _
        "," & p医疗卡管理 & ":卡费使用门诊收费医疗收据:" & chk_卡费使用门诊收费医疗收据 & _
        "," & p医疗卡管理 & ":收取病历费:" & chk_收取病历费 & _
        "," & p医疗卡管理 & ":自动门诊号:" & chk_医疗卡_自动生成门诊号 & _
        ""
    
    '挂号安排相关Check参数
    strTmp = strTmp & _
    "," & p挂号安排 & ":只允许选院内医生:" & chk_只对医内医生进行挂号安排 & _
    "," & p挂号安排 & ":预约单存在禁止删除:" & chk_存在预约挂号单禁止删除安排 & _
    ""
    
    '挂号管理相关Check参数
    strTmp = strTmp & _
        "," & p挂号管理 & ":自动门诊号:" & chk_挂号_自动生成门诊号 & _
        "," & p挂号管理 & ":存为划价单:" & chk_挂号_存为划价单 & _
        "," & p挂号管理 & ":收取卡费:" & chk_挂号_卡费与挂号一起收 & _
        "," & p挂号管理 & ":输入医生:" & chk_挂号_无医生必须输入医生 & _
        "," & p挂号管理 & ":自动产生姓名:" & chk_挂号_发卡自动生成新病人 & _
        "," & p挂号管理 & ":优先使用预交款:" & chk_挂号_优先使用预交款缴费 & _
        "," & p挂号管理 & ":姓名模糊查找:" & chk_挂号_姓名模糊查找 & _
        "," & p挂号管理 & ":零费用打印:" & chk_挂号_零费用打印票据 & _
        "," & p挂号管理 & ":输入姓名:" & chk_挂号_输入_姓名 & _
        "," & p挂号管理 & ":输入性别:" & chk_挂号_输入_性别 & _
        "," & p挂号管理 & ":输入年龄:" & chk_挂号_输入_年龄 & _
        "," & p挂号管理 & ":输入家庭地址:" & chk_挂号_输入_家庭地址 & _
        "," & p挂号管理 & ":输入付款方式:" & chk_挂号_输入_付款方式 & _
        "," & p挂号管理 & ":输入费别:" & chk_挂号_输入_费别 & _
        "," & p挂号管理 & ":输入结算方式:" & chk_挂号_输入_结算方式 & _
        "," & p挂号管理 & ":发卡不弹窗口:" & chk_挂号_发卡不弹病人登记窗口 & _
        "," & p挂号管理 & ":退费重打:" & chk_挂号_退号不退卡重打票据 & _
        "," & p挂号管理 & ":打印病历标签:" & chk_挂号_挂号后打印病历标签 & _
        "," & p挂号管理 & ":预约显示所有号别:" & chk_挂号_预约显示所有号别 & _
        "," & p挂号管理 & ":预约接收确定挂号费:" & chk_挂号_预约接收确定挂号费 & _
        "," & p挂号管理 & ":允许住院病人挂号:" & chk_挂号_允许住院病人挂号 & _
        ""
    
    strTmp = strTmp & _
        "," & p挂号管理 & ":预约不生成门诊号:" & chk_挂号_预约不生成门诊号 & _
        "," & p挂号管理 & ":随机序号选择:" & chk_挂号_随机序号选择 & _
        "," & p挂号管理 & ":退号审核:" & chk_挂号_N天内退号需审核 & _
        "," & p挂号管理 & ":失约用于挂号:" & chk_挂号_预约失约用于挂号 & _
        "," & p挂号管理 & ":家庭地址输入方式:" & chk_挂号_家庭地址联想输入 & _
        "," & p挂号管理 & ":扫描身份证签约:" & chk_挂号_扫描身份证签约 & _
        "," & p挂号管理 & ":计划排班挂号默认界面:" & chk_挂号_计划排班模式精简界面 & _
        "," & p挂号管理 & ":已退序号允许挂号:" & chk_挂号_已退序号允许挂号 & _
        "," & p挂号管理 & ":严格按时段挂号:" & chk_挂号_严格按时段挂号 & _
        "," & p挂号管理 & ":默认购买病历:" & chk_挂号_默认勾选病历选项 & _
        "," & p挂号管理 & ":门诊号有效性检查:" & chk_挂号_门诊号有效性检查 & _
        "," & p挂号管理 & ":非严格控制时始终发卡:" & chk_挂号_非严格控制为发卡 & _
        "," & p挂号管理 & ":输入联系电话:" & chk_挂号_输入_联系电话 & _
        "," & p挂号管理 & ":禁止输入年龄:" & chk_挂号_禁止输入年龄 & _
        ""
    '分诊
        
    strTmp = strTmp & _
        "," & p分诊管理 & ":分诊台签到排队:" & chk_分诊_分诊台签到开始排队 & _
        "," & p分诊管理 & ":预约生成队列:" & chk_分诊_预约挂号进入队列 & _
        "," & p分诊管理 & ":诊室忙时允许分诊:" & chk_分诊_诊室忙允许分诊 & _
        "," & p分诊管理 & ":回诊病人需重新排队:" & chk_分诊_回诊签到 & _
        "," & p分诊管理 & ":再次签到需重新排队:" & chk_分诊_重新签到 & _
        ""
        
    '临床出诊安排
    strTmp = strTmp & _
        "," & p临床出诊安排 & ":替诊医生级别检查:" & chk_出诊安排_替诊医生级别检查 & _
        "," & p临床出诊安排 & ":按替诊医生同步更新预约挂号单:" & chk_出诊安排_按替诊医生同步更新预约挂号单 & _
        "," & p临床出诊安排 & ":只允许选院内医生:" & chk_出诊安排_只允许选院内医生
    
    '门诊划价
    strTmp = strTmp & _
        "," & p门诊划价管理 & ":中药付数:" & chk_划价_门诊输入中药付数 & _
        "," & p门诊划价管理 & ":变价数次:" & chk_划价_门诊变价输入数次 & _
        "," & p门诊划价管理 & ":显示护士:" & chk_划价_门诊开单人含护士 & _
        "," & p门诊划价管理 & ":显示其它药房库存:" & chk_划价_门诊显示其它药房库存 & _
        "," & p门诊划价管理 & ":显示其它药库库存:" & chk_划价_门诊显示其它药库库存 & _
        "," & p门诊划价管理 & ":姓名模糊查找:" & chk_划价_门诊姓名模糊查找 & _
        "," & p门诊划价管理 & ":性别:" & chk_划价_门诊输入_性别 & _
        "," & p门诊划价管理 & ":年龄:" & chk_划价_门诊输入_年龄 & _
        "," & p门诊划价管理 & ":费别:" & chk_划价_门诊输入_费别 & _
        "," & p门诊划价管理 & ":医疗付款:" & chk_划价_门诊输入_医疗付款 & _
        "," & p门诊划价管理 & ":加班:" & chk_划价_门诊输入_是否加班 & _
        "," & p门诊划价管理 & ":开单日期:" & chk_划价_门诊输入_开单日期 & _
        "," & p门诊划价管理 & ":开单人:" & chk_划价_门诊输入_开单人 & _
        "," & p门诊划价管理 & ":不使用缺省开单人:" & chk_划价_门诊不缺省开单人 & _
        "," & p门诊划价管理 & ":必须要输入开单人:" & chk_划价_门诊必须要输入开单人 & _
        "," & p门诊划价管理 & ":缺省科室优先:" & chk_划价_门诊缺省科室优先 & _
        "," & p门诊划价管理 & ":住院病人按门诊收费:" & chk_划价_住院病人按门诊收费 & _
        "," & p门诊划价管理 & ":允许录入特殊使用的抗生素:" & chk_划价_允许录入特殊使用的抗生素
   

    '门诊收费
    strTmp = strTmp & _
        "," & p门诊收费管理 & ":中药付数:" & chk_收费_门诊输入中药付数 & _
        "," & p门诊收费管理 & ":变价数次:" & chk_收费_门诊变价输入数次 & _
        "," & p门诊收费管理 & ":显示护士:" & chk_收费_门诊开单人含护士 & _
        "," & p门诊收费管理 & ":显示其它药房库存:" & chk_收费_门诊显示其它药房库存 & _
        "," & p门诊收费管理 & ":显示其它药库库存:" & chk_收费_门诊显示其它药库库存 & _
        "," & p门诊收费管理 & ":姓名模糊查找:" & chk_收费_门诊姓名模糊查找 & _
        "," & p门诊收费管理 & ":性别:" & chk_收费_门诊输入_性别 & _
        "," & p门诊收费管理 & ":年龄:" & chk_收费_门诊输入_年龄 & _
        "," & p门诊收费管理 & ":费别:" & chk_收费_门诊输入_费别 & _
        "," & p门诊收费管理 & ":医疗付款:" & chk_收费_门诊输入_医疗付款 & _
        "," & p门诊收费管理 & ":加班:" & chk_收费_门诊输入_是否加班 & _
        "," & p门诊收费管理 & ":开单日期:" & chk_收费_门诊输入_开单日期 & _
        "," & p门诊收费管理 & ":开单人:" & chk_收费_门诊输入_开单人 & _
        "," & p门诊收费管理 & ":不使用缺省开单人:" & chk_收费_门诊不缺省开单人 & _
        "," & p门诊收费管理 & ":必须要输入开单人:" & chk_收费_门诊必须要输入开单人 & _
        "," & p门诊收费管理 & ":缺省科室优先:" & chk_收费_门诊缺省科室优先 & _
        "," & p门诊收费管理 & ":显示累计:" & chk_收费_门诊显示收款累计 & _
        "," & p门诊收费管理 & ":检查皮试结果:" & chk_收费_提划价单检查皮试结果 & _
        "," & p门诊收费管理 & ":优先使用预交款:" & chk_收费_优先使用预交款缴费 & _
        "," & p门诊收费管理 & ":多单据收费:" & chk_收费_允许输入多张单据 & _
        "," & p门诊收费管理 & ":搜寻划价单据:" & chk_收费_搜寻划价单据 & _
        "," & p门诊收费管理 & ":不弹出划价单选择:" & chk_收费_不弹出划价单选择窗口 & _
        "," & p门诊收费管理 & ":检查病人挂号科室:" & chk_收费_检查病人挂号科室

  strTmp = strTmp & _
        "," & p门诊收费管理 & ":住院病人按门诊收费:" & chk_收费_住院按门诊收费 & _
        "," & p门诊收费管理 & ":提取划价后立即缴款:" & chk_收费_提取划价立即缴款 & _
        "," & p门诊收费管理 & ":收据加收工本费:" & chk_收费_自动加收工本费 & _
        "," & p门诊收费管理 & ":多张单据收费分别打印:" & chk_收费_多张单据收费分别打印 & _
        "," & p门诊收费管理 & ":收费每次只用一张票据:" & chk_收费_每次只用一张票据 & _
        "," & p门诊收费管理 & ":只对医保结算成功单据收费:" & chk_收费_只对医保结算成功单据收费 & _
        "," & p门诊收费管理 & ":按病人补打发票不区分结算次数:" & chk_收费_按病人不打票据不分次数 & _
        "," & p门诊收费管理 & ":体检病人分单据打印:" & chk_收费_体检病人按单据分别打印 & _
        "," & p门诊收费管理 & ":允许录入特殊使用的抗生素:" & chk_收费_允许录入特殊使用的抗生素

    '门诊记帐
    strTmp = strTmp & _
        "," & p门诊记帐管理 & ":只查找合约单位病人:" & chk_记帐_只查找合约单位病人 & _
        "," & p门诊记帐管理 & ":中药付数:" & chk_记帐_门诊输入中药付数 & _
        "," & p门诊记帐管理 & ":变价数次:" & chk_记帐_门诊变价输入数次 & _
        "," & p门诊记帐管理 & ":显示护士:" & chk_记帐_门诊开单人含护士 & _
        "," & p门诊记帐管理 & ":记帐打印:" & chk_记帐_记帐打印 & _
        "," & p门诊记帐管理 & ":划价打印:" & chk_记帐_划价打印 & _
        "," & p门诊记帐管理 & ":审核打印:" & chk_记帐_审核打印 & _
        "," & p门诊记帐管理 & ":显示其它药房库存:" & chk_记帐_门诊显示其它药房库存 & _
        "," & p门诊记帐管理 & ":显示其它药库库存:" & chk_记帐_门诊显示其它药库库存 & _
        "," & p门诊记帐管理 & ":姓名模糊查找:" & chk_记帐_门诊姓名模糊查找 & _
        "," & p门诊记帐管理 & ":允许录入特殊使用的抗生素:" & chk_记帐_允许录入特殊使用的抗生素
      
    '住院记帐业务
    strTmp = strTmp & _
        "," & p住院记帐管理 & ":记帐打印:" & chk_住院记帐_记帐打印 & _
        "," & p科室分散记帐 & ":记帐打印:" & chk_分散记帐_记帐打印 & _
        "," & p医技科室记帐 & ":记帐打印:" & chk_医技科室_记帐打印 & _
        "," & p住院记帐操作 & ":中药付数:" & chk_记帐操作_中药输入付数 & _
        "," & p住院记帐操作 & ":变价数次:" & chk_记帐操作_变价输入数次 & _
        "," & p住院记帐操作 & ":显示护士:" & chk_记帐操作_开单人包含护士 & _
        "," & p住院记帐操作 & ":允许保存为划价单:" & chk_记帐操作_欠费保存划价单 & _
        "," & p住院记帐操作 & ":显示其它药房库存:" & chk_记帐操作_显示其它药房库存 & _
        "," & p住院记帐操作 & ":显示其它药库库存:" & chk_记帐操作_显示其它药库库存 & _
        "," & p住院记帐操作 & ":划价打印:" & chk_记帐操作_划价打印 & _
        "," & p住院记帐操作 & ":审核打印:" & chk_记帐操作_审核打印 & _
        ""
    
    '病人结帐
    strTmp = strTmp & _
        "," & p病人结帐管理 & ":合约单位按病人打印:" & chk_结帐_合约单位按病人打印 & _
        "," & p病人结帐管理 & ":在院病人结帐后自动出院:" & chk_结帐_出院结帐后自动出院 & _
        "," & p病人结帐管理 & ":处理零费用:" & chk_结帐_处理零费用 & _
        "," & p病人结帐管理 & ":仅用指定预交款:" & chk_结帐_使用指定次数预交款 & _
        "," & p病人结帐管理 & ":中途结帐退预交:" & chk_结帐_中途结帐缺省退预交款 & _
        "," & p病人结帐管理 & ":结帐后不清除信息:" & chk_结帐_结帐后不清除界面信息 & _
        "," & p病人结帐管理 & ":结帐检查病历接收:" & chk_结帐_结帐检查病历接收情况 & _
        "," & p病人结帐管理 & ":先结自费费用不打印结帐票据:" & chk_结帐_自费费用不打印结帐票据 & _
        "," & p病人结帐管理 & ":自费缺省使用预交:" & chk_结帐_自费费用缺省使用预交 & _
        "," & p病人结帐管理 & ":出院病人允许门诊转住院:" & chk_出院病人允许门诊转住院 & _
        "," & p病人结帐管理 & ":病人多次结帐弹出结帐条件窗体:" & chk_结帐_病人多次结帐弹出结帐条件窗体 & _
        "," & p病人结帐管理 & ":三方卡结帐退款控制:" & chk_结帐_三方卡结帐退款控制 & _
        ""

    '执行登记管理
    strTmp = strTmp & _
        "," & p执行登记管理 & ":医技医嘱发送:" & chk_执行登记_显示医嘱发送的单据 & _
        ""
        
    '费用审核管理
    strTmp = strTmp & _
        "," & p费用审核管理 & ":门诊转住院先审核:" & chk_费用审核_门诊转住院先审核 & _
        ""
    '票据使用监控
    strTmp = strTmp & _
        "," & p票据使用监控 & ":领用票据签字确认:" & chk_票据监控_领用票据签字确认 & _
        ""
    '人员借款管理
    strTmp = strTmp & _
        "," & p人员借款管理 & ":申请打印:" & chk_借款_申请打印 & _
        "," & p人员借款管理 & ":借出打印:" & chk_借款_借出打印 & _
        ""
    
    '消费卡管理
    strTmp = strTmp & _
        "," & p消费卡管理 & ":缴款单打印:" & chk_消费卡_缴款单打印 & _
        ""
    
    '医嘱附费管理
    strTmp = strTmp & _
        "," & p医嘱附费管理 & ":中药输入付数:" & chk_附费_中药输入付数 & _
        "," & p医嘱附费管理 & ":变价输入数次:" & chk_附费_变价输入数次 & _
        "," & p医嘱附费管理 & ":显示其它药房库存:" & chk_附费_显示其它药房库存 & _
        "," & p医嘱附费管理 & ":显示其它药库库存:" & chk_附费_显示其它药库库存 & _
        "," & p医嘱附费管理 & ":备货卫材只能输内部条码:" & chk_附费_备货卫材输入条码 & _
        ""
        
    '收费轧账管理
    strTmp = strTmp & _
        "," & p收费轧帐管理 & ":预交轧帐按门诊或住院分别轧帐:" & chk_轧账_预交轧帐按门诊或住院分别轧帐 & _
        ""
    
    Call SetParToControl(strTmp, mrsPar, chk)
    Call LoadTriageQueuingDep
    Call SetTriageQueuingEnalbe(chk(chk_分诊_分诊台签到开始排队).value)
    
    chk(chk_记帐_只查找合约单位病人).Enabled = chk(chk_记帐_门诊姓名模糊查找).value = 1
    
    With vsBillFormat(vsGrid_收费票据格式)
        .ColHidden(.ColIndex("按病人补打票据格式")) = chk(chk_收费_按病人不打票据不分次数).value <> 1
    End With
    chk(chk_收费_体检病人按单据分别打印).Enabled = (chk(chk_收费_多张单据收费分别打印).value = vbChecked)


    '设置相关参数
    '2.设置ComboBox类参数
    strTmp = "0:23:" & cbo_已结单据 & _
            ",0:58:" & cbo_未审单据结帐
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = "0:185:" & cbo_病人审核方式 & _
            ",0:284:" & cbo_门诊费用转住院预交发票格式
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '临床出诊安排
    strTmp = p临床出诊安排 & ":号码排序比较方式:" & cbo_临床安排_号码比较方式
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '挂号相关
    strTmp = p挂号管理 & ":缺省排序方式:" & cbo_挂号_缺省排序方式
        
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '住院记帐操作
    strTmp = p住院记帐操作 & ":记帐后发药:" & cbo_记帐操作_记帐后发药
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    
    '病人结帐管理
    strTmp = p病人结帐管理 & ":合约单位结帐打印:" & cbo_结帐_合约单位结帐打印
    Call SetParToControl(strTmp, mrsPar, cbo, 3)  '3-按文本直接比较
    
    '一卡通消费操作
    strTmp = p一卡通消费操作 & ":收费收据格式:" & cbo_一卡通_收票票据格式
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    strTmp = p一卡通消费操作 & ":审核收据格式:" & cbo_一卡通_记帐票据格式
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '3.设置UpDown类参数
    strTmp = "0:9:" & ud_费用金额保留位数 & _
            ",0:66:" & ud_挂号预约天数 & _
            ",0:157:" & ud_费用单价保留位数
    
    strTmp = strTmp & "," & p分诊管理 & ":分诊有效天数:" & ud_分诊有效天数
    '划价
    strTmp = strTmp & "," & p门诊划价管理 & ":取消划价单:" & ud_划价_取消N天的划价单
    
    '收费
    strTmp = strTmp & "," & p门诊收费管理 & ":收费收据总行次:" & ud_收费_收费收据总行次
    '补结算
    strTmp = strTmp & "," & p门诊补结算 & ":补结算有效天数:" & ud_补结算_有效天数
    
    Call SetParToControl(strTmp, mrsPar, ud) 'mrsPar存储的控件名是txtUD


    '4.设置TextBox类参数
    
    strTmp = "" & p费用虚拟模块 & ":姓名查找天数:" & txt_姓名查找天数
    
    '   医疗卡管理
    strTmp = strTmp & "," & p医疗卡管理 & ":姓名查找天数:" & txt_姓名模糊查找天数
    txt(txt_姓名模糊查找天数).Enabled = chk(chk_姓名允许模糊查找).value = 1
    
    '   挂号安排
    
    '   挂号管理
    strTmp = strTmp & "," & p挂号管理 & ":自动刷新间隔:" & txt_挂号_刷新时间
    strTmp = strTmp & "," & p挂号管理 & ":姓名查找天数:" & txt_挂号_姓名查找天数
    strTmp = strTmp & "," & p挂号管理 & ":预约失约次数:" & txt_挂号_预约失效次数
    strTmp = strTmp & "," & p挂号管理 & ":N天内不能取消预约号:" & txt_挂号_N天内不能取消预约号
    strTmp = strTmp & "," & p挂号管理 & ":N岁以下必须录入监护人:" & txt_挂号_N岁以下输入监护人
    
    '分诊管理
    strTmp = strTmp & "," & p分诊管理 & ":提前N小时分诊:" & txt_分诊_提前N小时分诊
    
    '门诊划价
    strTmp = strTmp & "," & p门诊划价管理 & ":最大金额:" & txt_划价_门诊单据最大金额
    strTmp = strTmp & "," & p门诊划价管理 & ":姓名查找天数:" & txt_划价_门诊姓名模糊查找天数
    '门诊收费
    strTmp = strTmp & "," & p门诊收费管理 & ":最大金额:" & txt_收费_门诊单据最大金额
    strTmp = strTmp & "," & p门诊收费管理 & ":姓名查找天数:" & txt_收费_门诊姓名模糊查找天数
    strTmp = strTmp & "," & p门诊收费管理 & ":搜寻单据天数:" & txt_收费_搜寻划价单据天数
    
    '门诊记帐
    strTmp = strTmp & "," & p门诊记帐管理 & ":姓名查找天数:" & txt_记帐_门诊姓名模糊查找天数
    
    Call SetParToControl(strTmp, mrsPar, txt)
    chk(chk_挂号_自动刷新挂号安排).value = IIF(Val(txt(txt_挂号_刷新时间)) > 0, 1, 0)
    txt(txt_挂号_刷新时间).Enabled = Val(txt(txt_挂号_刷新时间).Text) > 0
    txt(txt_挂号_姓名查找天数).Enabled = chk(chk_挂号_姓名模糊查找).value = 1
    
    txt(txt_划价_门诊姓名模糊查找天数).Enabled = chk(chk_划价_门诊姓名模糊查找).value = 1
    txt(txt_收费_门诊姓名模糊查找天数).Enabled = chk(chk_收费_门诊姓名模糊查找).value = 1
    txt(txt_收费_搜寻划价单据天数).Enabled = chk(chk_收费_搜寻划价单据).value = 1
    txt(txt_记帐_门诊姓名模糊查找天数).Enabled = chk(chk_记帐_门诊姓名模糊查找).value = 1
    
    '5.设置ListBox类参数
    strTmp = ""
    'Call SetParToControl(strTmp, mrsPar, lst)

    '6.设置OptionButton类参数
    arrObj = Array(0, 160, opt护理)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p费用虚拟模块, "挂号模式", optRegist)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p费用虚拟模块, "挂号发票打印方式", optPrintFact)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p费用虚拟模块, "挂号凭条打印方式", optPrintSlip)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p费用虚拟模块, "预约挂号单打印方式", optPrintAppoint)
    Call SetParToControl("", mrsPar, arrObj)
    
    '预交管理
    arrObj = Array(p预交款管理, "退款禁止方式", optDepsoitDelSet)
    Call SetParToControl("", mrsPar, arrObj)
    
    '一卡通消费管理
    arrObj = Array(p一卡通消费操作, "刷卡缺省金额操作", optSetMoneyMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    '医疗卡管理
    arrObj = Array(p医疗卡管理, "退卡刷卡", optDelCardMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p医疗卡管理, "发卡打印方式", optPrintMode_SendCard)
    Call SetParToControl("", mrsPar, arrObj)
    
    '挂号管理
    arrObj = Array(p挂号管理, "挂号发票打印方式", optRegistPrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p挂号管理, "退号清除门诊信息", optRegistClearMzInfor)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p挂号管理, "病人条码打印方式", optBarCodePrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p挂号管理, "预约挂号单打印方式", optPrintBespeak)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p挂号管理, "预约接收模式", optReceiveMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p挂号管理, "挂号凭条打印方式", optSlipPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p挂号管理, "挂号缴款输入控制", optMoneyControl)
    Call SetParToControl("", mrsPar, arrObj)
    
    '分诊台
    arrObj = Array(p分诊管理, "排队叫号模式", optTriageQueuingMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p分诊管理, "排队单打印", optTriagePrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p分诊管理, "候诊排序方式", optTriageSort)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p分诊管理, "条码打印方式", optTriageBarcodePrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    '临床出诊安排
    arrObj = Array(p临床出诊安排, "预约清单控制方式", optToExcelMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p临床出诊安排, "预约清单打印方式", optPrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p临床出诊安排, "出诊表打印方式", optVisitTablePrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    '门诊划价
    arrObj = Array(p门诊划价管理, "药品单位", opt划价单位)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊划价管理, "库存显示方式", opt划价库存显示方式)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊划价管理, "分类合计方式", optBillTotalShow)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊划价管理, "划价通知单打印方式", optPrintRequisition)
    Call SetParToControl("", mrsPar, arrObj)
    
     
    
    '门诊收费
    arrObj = Array(p门诊收费管理, "药品单位", opt收费单位)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p门诊收费管理, "库存显示方式", opt收费库存显示方式)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊收费管理, "分类合计方式", optChargeBillTotalShow)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p门诊收费管理, "收费缴款输入控制", opt缴款)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊收费管理, "未挂号病人收费", optChargeRegPrompt)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊收费管理, "收费清单打印方式", optFeeListPrint)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p门诊收费管理, "药品摆药退费方式", optDrug)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊收费管理, "退费回单打印方式", optDelFeeRefundPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊收费管理, "收费执行单打印方式", optChargeExeBillPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    
    '门诊记帐
    arrObj = Array(p门诊记帐管理, "药品单位", opt记帐单位)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p门诊记帐管理, "库存显示方式", opt记帐库存显示方式)
    Call SetParToControl("", mrsPar, arrObj)
    
    '补结算
    arrObj = Array(p门诊补结算, "药品单位显示", optSupplementaryUnit)   '显示单位
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊补结算, "结算清单打印方式", optSupplementaryPrint)   '结算清单打印方式
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊补结算, "药品摆药退费方式", optDrugSupplementary)   '药品摆药后退费方式
    Call SetParToControl("", mrsPar, arrObj)
    
    '住院记帐业务
    arrObj = Array(p住院记帐操作, "记帐药品单位", optJZDrugUnit)    '显示单位
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p住院记帐操作, "记帐报警包含所有住院划价费用", optInExseCharge)
    Call SetParToControl("", mrsPar, arrObj)
    
    '病人结帐管理
    arrObj = Array(p病人结帐管理, "结帐费用时间", optBalanceTime) '结帐费用按哪种(登记或发生时间)时间结帐
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "结帐检查代收款项", optBalanceDSCheck) '结帐检查代收款项
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "退款收据打印", optDelBalancePrint) '病人退款收据打印方式
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "结帐时输血费检查", optBlood)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "结帐明细打印", optBalanceFeeListPrint) '结帐明细打印
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p病人结帐管理, "结帐缴款输入控制", optBalancePayin) '缴款控制
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "门诊预交缺省使用方式", optMzDeposit) '门诊预交缺省打印方式
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "自费费用打印方式", optOwnFee) '自费清单打印方式
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p病人结帐管理, "预交票据打印方式", optBalanceDepositPrint) '预交款打印方式
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p收费财务监控, "收款收据打印方式", optPrintModeSJ)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p收费财务监控, "备用金领用单打印方式", optPrintModeDraw)
    Call SetParToControl("", mrsPar, arrObj)
    
    '执行登记管理
    arrObj = Array(p执行登记管理, "执行登记单打印方式", optRegPrint)
    Call SetParToControl("", mrsPar, arrObj)

    '医嘱附费管理
    arrObj = Array(p医嘱附费管理, "药品单位", optDrugUnitFF)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p医嘱附费管理, "记帐后发药", optSendDrugFF)
    Call SetParToControl("", mrsPar, arrObj)

    
    '7.其他系统参数
    rsTmp.Filter = "模块=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
    
        Select Case rsTmp!参数号
        Case 28 '一卡通消费刷卡控制
            strTmp = NVL(strValue, "1|0")
            If InStr(strTmp, "|") = 0 Then strTmp = "1|0"
            
            intTemp = Val(Split(strTmp, "|")(0))
            If intTemp > 2 Then intTemp = 1
            If intTemp < 0 Then txt(txt_免密支付).Text = -1 * intTemp: intTemp = 3
            optBrushCard(intTemp).value = True
            txt(txt_免密支付).Enabled = intTemp = 3
            intTemp = Val(Split(strTmp, "|")(1))
            If intTemp < 0 Or intTemp > 2 Then intTemp = 0
            optBrushCard(10 + intTemp).value = True
            
            Call SetParRelations(Array(optBrushCard(0), optBrushCard(1), optBrushCard(2), optBrushCard(3), _
                optBrushCard(10), optBrushCard(11), optBrushCard(12)), mrsPar, rsTmp!参数号)
        Case 14    '零钱处理
            strTmp = IIF(IsNull(strValue), "0000", strValue)
            n = Val(Mid(strTmp, 1, 1))
            For i = 0 To cbo(cbo_挂号零钱处理).ListCount
                If Val(Split(cbo(cbo_挂号零钱处理).List(i) & "-", "-")(0)) = n Then cbo(cbo_挂号零钱处理).ListIndex = i: Exit For
            Next
            cbo(cbo_收费零钱处理).ListIndex = Val(Mid(strTmp, 2, 1))
            cbo(cbo_结帐零钱处理).ListIndex = Val(Mid(strTmp, 3, 1))
            cbo(cbo_消费卡零钱处理).ListIndex = Val(Mid(strTmp, 4, 1))
        
            Call SetParRelation(cbo, cbo_挂号零钱处理, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_收费零钱处理, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_结帐零钱处理, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_消费卡零钱处理, mrsPar)
        Case 278    '自动记帐模式
            n = Val(strValue)
            With cbo(cbo_自动记帐模式)
                .ListIndex = -1
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = n Then .ListIndex = i: Exit For
                Next
                If .ListIndex < 0 And n = 2 Then
                    .AddItem "2-临时试点记帐"
                    .ItemData(.NewIndex) = 2: .ListIndex = .NewIndex
                ElseIf .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                chk(chk_下午算半天模式).Visible = InStr(1, ",0,2,", "," & .ItemData(.ListIndex) & ",") > 0
                opt护理(0).Enabled = InStr(1, ",0,2,", "," & .ItemData(.ListIndex) & ",") > 0
                opt护理(1).Enabled = InStr(1, ",0,2,", "," & .ItemData(.ListIndex) & ",") > 0
                
                lblAutoChargeNM.Visible = .ItemData(.ListIndex) = 1
            End With
            Call SetParRelation(cbo, cbo_自动记帐模式, mrsPar, rsTmp!参数号)

        Case 17    '病人输入方式，分别为姓名、就诊卡、挂号单、病人ID
            strTmp = NVL(strValue, "1111")
            chk(chk_病人姓名).value = IIF(Val(Mid(strTmp, 1, 1)) = 0, 0, 1)
            chk(chk_刷就诊卡).value = IIF(Val(Mid(strTmp, 2, 1)) = 0, 0, 1)
            chk(chk_挂号单号).value = IIF(Val(Mid(strTmp, 3, 1)) = 0, 0, 1)
            chk(chk_病人ID).value = IIF(Val(Mid(strTmp, 4, 1)) = 0, 0, 1)
            
            Call SetParRelation(chk, chk_病人姓名, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_刷就诊卡, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_挂号单号, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_病人ID, mrsPar)
            
        Case 20    '表示各种票据的号码长度，各位分别为1-收费,2-预交,3-结帐,4-挂号
            strTmp = IIF(strValue = "", "7|7|7|7", strValue)
            lvw(lvw_票据).ListItems("C1").SubItems(1) = Split(strTmp, "|")(0)
            lvw(lvw_票据).ListItems("C2").SubItems(1) = Split(strTmp, "|")(1)
            lvw(lvw_票据).ListItems("C3").SubItems(1) = Split(strTmp, "|")(2)
            lvw(lvw_票据).ListItems("C4").SubItems(1) = Split(strTmp, "|")(3)
            
            
            varData = Array(lvw(lvw_票据), txtUD(ud_号码长度))
            Call SetParRelations(varData, rsTmp, Val(NVL(rsTmp!参数号)))
            Call lvw_ItemClick(lvw_票据, lvw(lvw_票据).SelectedItem)
            
            'Call SetParRelation(lvw, lvw_票据, mrsPar, rsTmp!参数号)
        Case 21  '挂号有效天数
            '普通号
            ud(ud_挂号单).value = IIF(Left(strValue, 1) = 0, 1, Left(strValue, 1))
            '急诊号
            ud(ud_急诊挂号单).value = IIF(Mid(strValue, 2, 1) = 0, 1, Mid(strValue, 2, 1))
        
            Call SetParRelation(txtUD, ud_挂号单, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txtUD, ud_急诊挂号单, mrsPar)
            
        Case 24    '表示是否严格控制管理对票据的使用，各位分别为1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
            strTmp = NVL(strValue, "1111")
            lvw(lvw_票据).ListItems("C1").SubItems(2) = IIF(Mid(strTmp, 1, 1) = "1", "√", "")
            lvw(lvw_票据).ListItems("C2").SubItems(2) = IIF(Mid(strTmp, 2, 1) = "1", "√", "")
            lvw(lvw_票据).ListItems("C3").SubItems(2) = IIF(Mid(strTmp, 3, 1) = "1", "√", "")
            lvw(lvw_票据).ListItems("C4").SubItems(2) = IIF(Mid(strTmp, 4, 1) = "1", "√", "")
            
            lvw(lvw_票据).ListItems("C1").Selected = True
            Call lvw_ItemClick(lvw_票据, lvw(lvw_票据).SelectedItem)
            
            Call SetParRelation(chk, chk_票号控制, mrsPar, rsTmp!参数号)
            
        Case 41    '医保病人适用费用类型
            SetListByText lst(lst_医保病人), Replace(strValue, "|", ",")
            Call SetParRelation(lst, lst_医保病人, mrsPar, rsTmp!参数号)
            
        Case 42    '公费病人适用费用类型
            SetListByText lst(lst_公费病人), Replace(strValue, "|", ",")
            Call SetParRelation(lst, lst_公费病人, mrsPar, rsTmp!参数号)
            
        Case 46    '刷卡要求输入密码
            With lst(lst_刷卡密码)
                For i = 1 To Len(NVL(strValue))
                    If Mid(strValue, i, 1) = "1" And i - 1 <= .ListCount - 1 Then
                        .Selected(i - 1) = True
                    End If
                Next
            End With
            Call SetParRelation(lst, lst_刷卡密码, mrsPar, rsTmp!参数号)
            
        Case 60    '单笔费用最大提醒金额
            txt(txt_单笔最大金额).Text = strValue
            Call txt_Validate(txt_单笔最大金额, False)
            Call SetParRelation(txt, txt_单笔最大金额, mrsPar, rsTmp!参数号)
        Case 98   '记帐报警包含划价费用
            If Val(strValue) = 1 Then
                lblInExseCharge.Enabled = True
                optInExseCharge(0).Enabled = True
                optInExseCharge(1).Enabled = True
            Else
                lblInExseCharge.Enabled = False
                optInExseCharge(0).Enabled = False
                optInExseCharge(1).Enabled = False
            End If
        Case 256    '挂号排班模式
            fraNewPaln.Move chk(chk_只对医内医生进行挂号安排).Left, chk(chk_只对医内医生进行挂号安排).Top
            If Val(Split(strValue & "|", "|")(0)) = 0 Then
                mblnNotChange = True
                optRegistPlanMode(0).value = 1
                optRegistPlanMode(1).value = 0
                mblnNotChange = False
                dtpRegistPlanMode.Enabled = False
                strSQL = "Select Max(发生时间) As 最大时间 From 病人挂号记录 Where 记录状态 =1 And 发生时间 >= [1] And 发生时间 Is Not Null"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, zlDatabase.Currentdate)
                If IsNull(rsTemp!最大时间) Then
                    dtpRegistPlanMode.value = zlDatabase.Currentdate
                Else
                    dtpRegistPlanMode.value = CDate(rsTemp!最大时间)
                End If
                cmdInstantActive.Enabled = True
                
                fraNewPaln.Visible = False
                chk(chk_只对医内医生进行挂号安排).Visible = True
                chk(chk_存在预约挂号单禁止删除安排).Visible = True
            Else
                mblnNotChange = True
                optRegistPlanMode(0).value = 0
                optRegistPlanMode(1).value = 1
                mblnNotChange = False
                dtpRegistPlanMode.Enabled = True
                If Split(strValue & "|", "|")(1) = "" Then
                    strSQL = "Select Max(发生时间) As 最大时间 From 病人挂号记录 Where 记录状态 =1 And 发生时间 >= [1] And 发生时间 Is Not Null"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, zlDatabase.Currentdate)
                    If IsNull(rsTemp!最大时间) Then
                        dtpRegistPlanMode.value = zlDatabase.Currentdate
                    Else
                        dtpRegistPlanMode.value = CDate(rsTemp!最大时间)
                    End If
                Else
                    dtpRegistPlanMode.value = CDate(Split(strValue & "|", "|")(1))
                End If
                cmdInstantActive.Enabled = False
                
                fraNewPaln.Visible = True
                chk(chk_只对医内医生进行挂号安排).Visible = False
                chk(chk_存在预约挂号单禁止删除安排).Visible = False
            End If
            strSQL = "Select 1 From 病人挂号记录 Where 记录状态 =1 And 出诊记录ID Is Not Null And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If Not rsTemp.EOF Then
                dtpRegistPlanMode.Enabled = False
                dtpRegistPlanMode.ToolTipText = "已经存在出诊表排班下的挂号记录,不能再修改启用时间!"
            End If
            Call SetParRelations(Array(optRegistPlanMode(0), optRegistPlanMode(1), dtpRegistPlanMode), mrsPar, rsTmp!参数名)
        Case 261
            strValue = Replace(strValue, " ", "")
            With lst(lst_脱机医保结算方式)
                For i = 0 To .ListCount - 1
                    If InStr("|" & strValue & "|", "|" & zlCommFun.GetNeedName(.List(i), "-") & "|") > 0 Then
                        .Selected(i) = True
                    End If
                Next
            End With
            Call SetParRelation(lst, lst_脱机医保结算方式, mrsPar, rsTmp!参数号)
        Case 263
            txt(txt_专家号挂号限制).Text = Val(strValue)
            chk(chk_专家号挂号限制).value = IIF(Val(strValue) <> 0, 1, 0)
            If chk(chk_专家号挂号限制).value = 1 Then
                txt(txt_专家号挂号限制).Enabled = True
            Else
                txt(txt_专家号挂号限制).Enabled = False
                txt(txt_专家号挂号限制).Text = ""
            End If
            Call SetParRelations(Array(txt(txt_专家号挂号限制), chk(chk_专家号挂号限制)), mrsPar, rsTmp!参数号)
        Case 264
            txt(txt_专家号预约限制).Text = Val(strValue)
            chk(chk_专家号预约限制).value = IIF(Val(strValue) <> 0, 1, 0)
            If chk(chk_专家号预约限制).value = 1 Then
                txt(txt_专家号预约限制).Enabled = True
            Else
                txt(txt_专家号预约限制).Enabled = False
                txt(txt_专家号预约限制).Text = ""
            End If
            Call SetParRelations(Array(txt(txt_专家号预约限制), chk(chk_专家号预约限制)), mrsPar, rsTmp!参数号)
        Case 277 '号源开放时间
            dtpRegistTime.value = strValue
            Call SetParRelations(Array(dtpRegistTime), mrsPar, rsTmp!参数号)
        End Select
        rsTmp.MoveNext
    Loop
    
    
    '8.其他模块参数
    rsTmp.Filter = "模块=" & p费用虚拟模块
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "姓名模糊查找" '15
            If Val(strValue) = 0 Then
                txt(txt_姓名查找天数).Enabled = False
            Else
                txt(txt_姓名查找天数).Enabled = True
            End If
        Case "包含科室安排"
            strValue = strValue & "|"
            chk(chk_挂号包含科室安排).value = Val(Split(strValue, "|")(0))
            chk(chk_预约包含科室安排).value = Val(Split(strValue, "|")(1))
            Call SetParRelations(Array(chk(chk_挂号包含科室安排), chk(chk_预约包含科室安排)), rsTmp, CStr(NVL(rsTmp!参数名)), p费用虚拟模块)
        Case "医生站挂号排序控制"
            fraOrder(1).Visible = True
            Call LoadStationRegOrder(strValue)
            Call SetParRelation(vsStationRegSort, 0, mrsPar, CStr(NVL(rsTmp!参数名)), p费用虚拟模块)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p预交款管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "票据剩余X张时开始提醒收费员"
            If strValue = "" Then strValue = "0|10"
            varData = Split(strValue & "|", "|")
            chk(chk_票据剩余N张提醒操作员).value = IIF(Val(varData(0)) = 1, 1, 0)
            txtUD(ud_票据张数).Text = Val(varData(1))
            ud(ud_票据张数).value = Val(varData(1))
            txtUD(ud_票据张数).Enabled = Val(varData(0)) = 1
            ud(ud_票据张数).Enabled = Val(varData(0)) = 1
            
            
            varData = Array(chk(chk_票据剩余N张提醒操作员), txtUD(ud_票据张数), ud(ud_票据张数))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!参数名)), p预交款管理)
            
        Case "代收款设置"
            Call SetParRelation(vs代收, 0, mrsPar, CStr(NVL(rsTmp!参数名)), p预交款管理)
            Call Load代收款(strValue)
        End Select
        rsTmp.MoveNext
    Loop
    Call Load预交票据格式(rsTmp)
    Call Load预交红票格式(rsTmp)
    
    Call SetDrugStore
    
    rsTmp.Filter = "模块=" & p一卡通消费操作
    Call Load药房(rsTmp)
    
    '医疗卡管理
    rsTmp.Filter = "模块=" & p医疗卡管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "输入项控制"
            If strValue = "" Then strValue = "户口地址邮编|医保号|费别|国籍|医疗付款|学历|其他证件|出生地点|工作单位|单位电话|单位邮编|单位帐户|单位开户行|联系人身份证号|联系人地址|联系人姓名|联系人电话|联系人关系"
            Call LoadInputItem(vsGrid_发卡输入项设置, strValue)
            Call SetParRelation(vsInputItemSet, vsGrid_发卡输入项设置, mrsPar, CStr(NVL(rsTmp!参数名)), p医疗卡管理)
        Case 1
            
        End Select
        rsTmp.MoveNext
    Loop
    
    Call Load发卡预交票据格式(rsTmp)
    Call Load医疗卡票据格式(rsTmp)
    
    '临床出诊安排
    rsTmp.Filter = "模块=" & p临床出诊安排
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "未区分站点的号源的维护站点"
            With cbo(cbo_临床安排_全院通用号源安排站点)
                For i = 0 To .ListCount - 1
                    If zlStr.NeedCode(.List(i), "-") = strValue Then .ListIndex = i: Exit For
                Next
            End With
            Call SetParRelations(Array(cbo(cbo_临床安排_全院通用号源安排站点)), rsTmp, CStr(NVL(rsTmp!参数名)), p临床出诊安排)
        End Select
        rsTmp.MoveNext
    Loop
    
    '挂号管理
    rsTmp.Filter = "模块=" & p挂号管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "预约限制时间"
            varData = Split(strValue & "|", "|")
            txt(txt_挂号_预约限制时间_分钟) = Val(varData(1))
            txt(txt_挂号_预约限制时间_天数) = Val(varData(0))
            Call SetParRelations(Array(txt(txt_挂号_预约限制时间_分钟), txt(txt_挂号_预约限制时间_天数)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "病人挂号科室限制"
            txt(txt_挂号_病人挂号科室限制).Text = Val(strValue)
            chk(chk_挂号_病人挂号科室限制).value = IIF(Val(strValue) <> 0, 1, 0)
            Call SetParRelations(Array(txt(txt_挂号_病人挂号科室限制), chk(chk_挂号_病人挂号科室限制)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "病人预约科室数"
            txt(txt_挂号_病人预约科室数).Text = Val(strValue)
            chk(chk_挂号_病人预约科室数).value = IIF(Val(strValue) <> 0, 1, 0)
            Call SetParRelations(Array(txt(txt_挂号_病人预约科室数), chk(chk_挂号_病人预约科室数)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "病人同科限约N个号"
            txt(txt_挂号_病人同科限约N个号).Text = Val(strValue)
            chk(chk_挂号_病人同科限约N个号).value = IIF(Val(strValue) <> 0, 1, 0)
            If chk(chk_挂号_病人同科限约N个号).value = 1 Then
                txt(txt_挂号_病人同科限约N个号).Enabled = True
            Else
                txt(txt_挂号_病人同科限约N个号).Enabled = False
                txt(txt_挂号_病人同科限约N个号).Text = ""
            End If
            Call SetParRelations(Array(txt(txt_挂号_病人同科限约N个号), chk(chk_挂号_病人同科限约N个号)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "病人同科限挂N个号"
            varData = Split(strValue & "|", "|")
            txt(txt_挂号_病人同科限挂N个号).Text = Val(varData(0))
            chk(chk_挂号_病人同科限挂N个号).value = IIF(Val(varData(0)) <> 0, 1, 0)
            If chk(chk_挂号_病人同科限挂N个号).value = 0 Then
                chk(chk_挂号_病人同科限挂N个号_急诊).value = 0
                chk(chk_挂号_病人同科限挂N个号_急诊).Enabled = False
                txt(txt_挂号_病人同科限挂N个号).Enabled = False
                txt(txt_挂号_病人同科限挂N个号).Text = ""
            Else
                txt(txt_挂号_病人同科限挂N个号).Enabled = True
                chk(chk_挂号_病人同科限挂N个号_急诊).value = IIF(Val(varData(1)) <> 0, 1, 0)
            End If
            Call SetParRelations(Array(txt(txt_挂号_病人同科限挂N个号), chk(chk_挂号_病人同科限挂N个号), chk(chk_挂号_病人同科限挂N个号_急诊)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "缺省预约方式"
            With cbo(cbo_挂号_缺省预约方式)
                If .ListCount > 0 Then
                    For i = 0 To .ListCount - 1
                        If zlCommFun.GetNeedName(.List(i), "-") = strValue Then
                            .ListIndex = i: Exit For
                        End If
                    Next i
                    If .ListIndex < 0 Then .ListIndex = 0
                End If
            End With
            Call SetParRelations(Array(cbo(cbo_挂号_缺省预约方式)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "预约有效时间"
            If Val(strValue) >= 0 Then
                '提前
                cbo(cbo_挂号_预约有效时间).ListIndex = 0
            Else
                '延后
                cbo(cbo_挂号_预约有效时间).ListIndex = 1
            End If
            txt(txt_挂号_预约有效时间).Text = Abs(strValue)
            Call SetParRelations(Array(cbo(cbo_挂号_预约有效时间), txt(txt_挂号_预约有效时间)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "提前挂号颜色"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "0"
            pic提前颜色.BackColor = strValue
            Call SetParRelations(Array(pic提前颜色), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        Case "病人同一号源限挂N个号"
            txt(txt_挂号_病人同一号源限挂N个号).Text = Val(strValue)
            chk(chk_挂号_病人同一号源限挂N个号).value = IIF(Val(strValue) <> 0, 1, 0)
            Call SetParRelations(Array(txt(txt_挂号_病人同一号源限挂N个号), chk(chk_挂号_病人同一号源限挂N个号)), rsTmp, CStr(NVL(rsTmp!参数名)), p挂号管理)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p门诊收费管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "自动加收挂号费"
            If InStr(1, strValue, ";") > 0 Then
                chk(chk_收费_未挂号自动加收挂号费).value = 1   '会调用click事件,须先加载收费类别
                txt(txt_收费_自助加收挂号费).Tag = Split(strValue, ";")(0)
                txt(txt_收费_自助加收挂号费).Text = Split(strValue, ";")(1)
            Else
                chk(chk_收费_未挂号自动加收挂号费).value = 0
                txt(txt_收费_自助加收挂号费).Text = ""
                txt(txt_收费_自助加收挂号费).Tag = ""
            End If
            Call SetParRelations(Array(txt(txt_收费_自助加收挂号费), chk(chk_收费_未挂号自动加收挂号费)), rsTmp, CStr(NVL(rsTmp!参数名)), p门诊收费管理)
        Case "自动组合单据"
            If Val(strValue) <> 0 Then
                chk(chk_收费_自动组合单据).value = 1
                cbo(cbo_收费_自动组合单据).ListIndex = IIF(Val(strValue) = 1, 0, 1)
                cbo(cbo_收费_自动组合单据).Enabled = True
            Else
                chk(chk_收费_自动组合单据).value = 0
                cbo(cbo_收费_自动组合单据).ListIndex = 0
                cbo(cbo_收费_自动组合单据).Enabled = False
            End If
            Call SetParRelations(Array(cbo(cbo_收费_自动组合单据), chk(chk_收费_自动组合单据)), rsTmp, CStr(NVL(rsTmp!参数名)), p门诊收费管理)
        Case "票据剩余X张时开始提醒收费员"
            If strValue = "" Then strValue = "0|10"
            varData = Split(strValue & "|", "|")
            chk(chk_收费_票据剩余X张开始提醒).value = IIF(Val(varData(0)) = 1, 1, 0)
            txtUD(ud_收费_票据张数).Text = Val(varData(1))
            ud(ud_收费_票据张数).value = Val(varData(1))
            txtUD(ud_收费_票据张数).Enabled = Val(varData(0)) = 1
            ud(ud_收费_票据张数).Enabled = Val(varData(0)) = 1
            
            varData = Array(chk(chk_收费_票据剩余X张开始提醒), txtUD(ud_收费_票据张数), ud(ud_收费_票据张数))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!参数名)), p门诊收费管理)
        Case "票据分配规则"
            '启用标志||NO;执行科室(条数);收据费目;收费细目(条数);汇总条件(0-不汇总;1-首页汇总(按第1页汇总),2-分组汇总(选择明细时有效)).
            If strValue = "" Then strValue = "0||0;0;0;0;0;0"
            varData = Split(strValue & "||", "||")
            If Val(varData(0)) = 0 Then cbo(cbo_收费_票据分配规则).ListIndex = 0
            If Val(varData(0)) = 1 Then cbo(cbo_收费_票据分配规则).ListIndex = 1
            If Val(varData(0)) = 2 Then cbo(cbo_收费_票据分配规则).ListIndex = 2
            If cbo(cbo_收费_票据分配规则).ListIndex < 0 Then cbo(cbo_收费_票据分配规则).ListIndex = 0
            
            mint原票据分配规则 = cbo(cbo_收费_票据分配规则).ListIndex
            Call SetBillRuleParaLocale
            
            varTemp = Split(varData(1) & ";;;;", ";")
            
            varData = Array(cbo(cbo_收费_票据分配规则), chkBillRule(0), chkBillRule(1), chkBillRule(2), chkBillRule(3), optRuleTotal(0), optRuleTotal(1), optRuleTotal(2), _
                            lblBillRuleNum(0), updBillRuleNum(0), txtBillRuleNum(0), lblBillRuleNum(1), updBillRuleNum(1), lblBillRuleNum(2), txtBillRuleNum(1), updBillRuleNum(2), txtBillRuleNum(2))
            
'            Call SetCtlsEnabled(varData, Not mblnExistPrintData)
            
            '2.根据预定规则分配票号
            '2.1按单据分
            i = Val(varTemp(0))
            chkBillRule(0).value = IIF(i = 1, 1, 0)
            '2.2按执行科室分
            i = Val(varTemp(1))
            chkBillRule(1).value = IIF(i >= 1, 1, 0)
            updBillRuleNum(0).value = IIF(i < 0 Or i > 100, 0, i)
            txtBillRuleNum(0).Text = updBillRuleNum(0).value
            txtBillRuleNum(0).Tag = IIF(updBillRuleNum(0).value = 0, 1, updBillRuleNum(0).value)
            '2.3 按收据费目
            i = Val(varTemp(2))
            chkBillRule(2).value = IIF(i >= 1, 1, 0)
            updBillRuleNum(1).value = IIF(i < 0 Or i > 100, 0, i)
            txtBillRuleNum(1).Text = updBillRuleNum(1).value
            txtBillRuleNum(1).Tag = IIF(updBillRuleNum(1).value = 0, 1, updBillRuleNum(1).value)
            '2.4 按收费细目(先处理收费细目，不然会触发Click事件，将首页汇总执行为空了
            i = Val(varTemp(3))
            chkBillRule(3).value = IIF(i >= 1, 1, 0)
            updBillRuleNum(2).value = IIF(i < 0 Or i > 100, 0, i)
            txtBillRuleNum(2).Text = updBillRuleNum(2).value
            txtBillRuleNum(2).Tag = IIF(updBillRuleNum(2).value = 0, 20, updBillRuleNum(2).value)
            '2.5 分组汇总
            i = Val(varTemp(4)): i = IIF(i > 3 Or i < 0, 0, i)
            optRuleTotal(i).value = True
            
            varData = Array(cbo(cbo_收费_票据分配规则), chkBillRule(0), chkBillRule(1), chkBillRule(2), chkBillRule(3), optRuleTotal(0), optRuleTotal(1), optRuleTotal(2), _
                             updBillRuleNum(0), txtBillRuleNum(0), updBillRuleNum(1), txtBillRuleNum(1), updBillRuleNum(2), txtBillRuleNum(2))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!参数名)), p门诊收费管理)
            Call ShowRuleInfor
            
        Case "收费票据生成方式"
            i = Val(strValue)
            chk(chk_收费_票据生成方式).value = IIF(i >= 10, 1, 0)
            optBillMode(i Mod 10).value = True
            varData = Array(chk(chk_收费_票据生成方式), optBillMode)
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!参数名)), p门诊收费管理)
        
        End Select
        rsTmp.MoveNext
    Loop
    Call Load收费票据格式(rsTmp)
    Call Load退费票据格式(rsTmp)
    Call LoadDelFeeDefaultType
    
    rsTmp.Filter = "模块=" & p门诊补结算
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "姓名模糊查找方式"
            varData = Split(strValue & "|", "|")
            '启用标志|天数
            chk(chk_补结算_姓名模糊查找).value = IIF(Val(varData(0)) = 1, 1, 0)
            txt(txt_补结算_姓名模糊查找天数).Text = Val(varData(1))
            txt(txt_补结算_姓名模糊查找天数).Enabled = chk(chk_补结算_姓名模糊查找).value = 1
            varData = Array(chk(chk_补结算_姓名模糊查找), txt(txt_补结算_姓名模糊查找天数))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!参数名)), p门诊补结算)
        Case "票据剩余X张时开始提醒收费员"
            If strValue = "" Then strValue = "0|10"
            varData = Split(strValue & "|", "|")
            chk(chk_补结算_票据剩余X张开始提醒).value = IIF(Val(varData(0)) = 1, 1, 0)
            txtUD(ud_补结算_票据张数).Text = Val(varData(1))
            ud(ud_补结算_票据张数).value = Val(varData(1))
            txtUD(ud_补结算_票据张数).Enabled = Val(varData(0)) = 1
            ud(ud_补结算_票据张数).Enabled = Val(varData(0)) = 1
            
            varData = Array(chk(chk_补结算_票据剩余X张开始提醒), txtUD(ud_补结算_票据张数), ud(ud_补结算_票据张数))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!参数名)), p门诊补结算)
        Case "允许补结算的收费结算方式"
            SetListByText lst(lst_补结算_结算方式), Replace(strValue, "|", ",")
            Call SetParRelation(lst, lst_补结算_结算方式, mrsPar, CStr(NVL(rsTmp!参数名)), p门诊补结算)
        End Select
        rsTmp.MoveNext
    Loop
    Call Load补结算票据格式(rsTmp)
    Call Load补结算退费票据格式(rsTmp)
    
    rsTmp.Filter = "模块=" & p病人结帐管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "结算前先结自费费用"
            strValue = Replace(strValue, " ", "")
            With lst(lst_结帐_自费费用类别)
                For i = 0 To .ListCount - 1
                    If InStr("," & strValue & ",", "," & Chr(.ItemData(i)) & ",") > 0 Then
                        .Selected(i) = True
                    End If
                Next
            End With
            Call SetParRelation(lst, lst_结帐_自费费用类别, mrsPar, CStr(NVL(rsTmp!参数名)), p病人结帐管理)
        Case "自付合计栏字体色"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "16711680"
            pic自付合计色.BackColor = strValue
            Call SetParRelations(Array(pic自付合计色), rsTmp, CStr(NVL(rsTmp!参数名)), p病人结帐管理)
        Case "当前付款栏字体色"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "255|255"
            pic当前付款未付色.BackColor = Mid(strValue, 1, InStr(strValue, "|") - 1)
            pic当前付款未退色.BackColor = Mid(strValue, InStr(strValue, "|") + 1)
            Call SetParRelations(Array(pic当前付款未付色, pic当前付款未退色), rsTmp, CStr(NVL(rsTmp!参数名)), p病人结帐管理)
        Case "缴款栏字体色"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "16711680|255"
            pic缴款栏缴款色.BackColor = Mid(strValue, 1, InStr(strValue, "|") - 1)
            pic缴款栏退款色.BackColor = Mid(strValue, InStr(strValue, "|") + 1)
            Call SetParRelations(Array(pic缴款栏缴款色, pic缴款栏退款色), rsTmp, CStr(NVL(rsTmp!参数名)), p病人结帐管理)
        Case "冲预交缺省顺序"
            strValue = Replace(strValue, " ", "")
            strArr = Split(strValue & "|", "|")
            vsDepositSort.MergeRow(0) = True
            vsDepositSort.MergeCol(0) = True
            vsDepositSort.MergeCol(1) = True
            vsDepositSort.MergeCol(2) = True
            If Val(strArr(0)) = 0 Then
                optOrder(0).value = True
                vsDepositSort.Enabled = False
                cmdDepositDown.Enabled = False
                cmdDepositUp.Enabled = False
            Else
                strStyle = Split(strArr(1), ",")
                For i = 0 To UBound(strStyle)
                    With vsDepositSort
                        .TextMatrix(i + 2, .ColIndex("结算类别")) = Split(strStyle(i), ":")(0)
                        intTemp = Val(Split(strStyle(i), ":")(1))
                        Select Case intTemp
                            Case 0
                                .TextMatrix(i + 2, 2) = -1
                                .TextMatrix(i + 2, 3) = 0
                                .TextMatrix(i + 2, 4) = 0
                            Case 1
                                .TextMatrix(i + 2, 2) = 0
                                .TextMatrix(i + 2, 3) = -1
                                .TextMatrix(i + 2, 4) = 0
                            Case 2
                                .TextMatrix(i + 2, 2) = 0
                                .TextMatrix(i + 2, 3) = 0
                                .TextMatrix(i + 2, 4) = -1
                        End Select
                    End With
                Next i
                optOrder(1).value = True
                vsDepositSort.Enabled = True
                cmdDepositDown.Enabled = True
                cmdDepositUp.Enabled = True
            End If
            Call SetParRelations(Array(optOrder(0), optOrder(1), vsDepositSort, cmdDepositDown, cmdDepositUp), rsTmp, CStr(NVL(rsTmp!参数名)), p病人结帐管理)
        Case "结帐界面风格"
            If Val(strValue) = 1 Then
                opt界面风格(0).value = False
                opt界面风格(1).value = True
                chk(chk_结帐_病人多次结帐弹出结帐条件窗体).Enabled = False
                fraColor.Enabled = True
                For i = 0 To 4
                    lbl结帐Color(i).Enabled = True
                Next i
            Else
                opt界面风格(0).value = True
                opt界面风格(1).value = False
                chk(chk_结帐_病人多次结帐弹出结帐条件窗体).Enabled = True
                fraColor.Enabled = False
                For i = 0 To 4
                    lbl结帐Color(i).Enabled = False
                Next i
            End If
            Call SetParRelations(Array(opt界面风格(0), opt界面风格(1), picDisplay(0), picDisplay(1)), rsTmp, CStr(NVL(rsTmp!参数名)), p病人结帐管理)
        End Select
        rsTmp.MoveNext
    Loop
    Call Load结帐票据格式(rsTmp)
    Call Load结帐红票格式(rsTmp)
End Sub


Private Sub cmdDepositUp_Click()
    Dim strValue As String, i As Integer, intIndex As Integer
    With vsDepositSort
        If .Row <= 2 Then Exit Sub
        .RowPosition(.Row) = .Row - 1
        .Row = .Row - 1
    End With
    strValue = "1|"
    With vsDepositSort
        For i = 2 To 4
            If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
            If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
            If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
            If i <> 4 Then
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
            Else
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
            End If
        Next i
    End With
    Call SetParChange(optOrder, 1, mrsPar, True, strValue)
End Sub

Private Sub cmdDepositDown_Click()
    Dim strValue As String, i As Integer, intIndex As Integer
    With vsDepositSort
        If .Row >= .Rows - 1 Then Exit Sub
        .RowPosition(.Row) = .Row + 1
        .Row = .Row + 1
    End With
    strValue = "1|"
    With vsDepositSort
        For i = 2 To 4
            If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
            If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
            If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
            If i <> 4 Then
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
            Else
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
            End If
        Next i
    End With
    Call SetParChange(optOrder, 1, mrsPar, True, strValue)
End Sub

Private Sub SetCtlsEnabled(ByVal varDara As Variant, blnEnabled)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enabled属性
    '编制:刘兴洪
    '日期:2015-06-17 17:46:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngIndex As Long, blnNotClear As Boolean
    On Error GoTo ErrHandle
    For i = 0 To UBound(varDara)
        varDara(i).Enabled = blnEnabled
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub LoadInputItem(ByVal intIndex As Integer, ByVal strValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载输入项控制
    '入参:intIndex-索引值
    '     strValue-缺省参数值和输入项,格式:输入项目,禁止录入,光标是否跳过,必输项|....
    '编制:刘兴洪
    '日期:2015-06-11 17:32:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim intRow As Integer, i As Integer
    
    On Error GoTo ErrHandle
    varData = Split(strValue, "|")
    With vsInputItemSet(intIndex)
        .redraw = flexRDNone
        .Clear 1
        If strValue = "" Then .Rows = 2: Exit Sub
        .Rows = 2: intRow = 1
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,,,", ",")
            If varTemp(0) <> "" Then
                .TextMatrix(intRow, .ColIndex("输入项目")) = varTemp(0)
                .TextMatrix(intRow, .ColIndex("禁止录入")) = IIF(Val(varTemp(1)) = 1, "√", "")
                .TextMatrix(intRow, .ColIndex("必输项")) = IIF(Val(varTemp(2)) = 1, "√", "")
                .TextMatrix(intRow, .ColIndex("光标进入")) = IIF(Val(varTemp(3)) = 1, "√", "")
                If .TextMatrix(intRow, .ColIndex("禁止录入")) = "√" Then
                    .Cell(flexcpBackColor, intRow, .ColIndex("必输项"), intRow, .ColIndex("光标进入")) = &H8000000F
                ElseIf .TextMatrix(intRow, .ColIndex("必输项")) = "√" _
                    Or .TextMatrix(intRow, .ColIndex("光标进入")) = "√" Then
                    .Cell(flexcpBackColor, intRow, .ColIndex("禁止录入")) = &H8000000F
                End If
                .Rows = .Rows + 1: intRow = intRow + 1
            End If
        Next
        If .Rows > 2 And Trim(.TextMatrix(.Rows - 1, .ColIndex("输入项目"))) = "" Then
            .Rows = .Rows - 1
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandle:
    vsInputItemSet(intIndex).redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
'功能：初始化界面控件，加载基础数据
    Dim strTmp As String, rsTmp As ADODB.Recordset
    Dim i As Long, rsTemp As ADODB.Recordset

    cbo(cbo_已结单据).AddItem "0-允许"
    cbo(cbo_已结单据).AddItem "1-提示"
    cbo(cbo_已结单据).AddItem "2-禁止"
    cbo(cbo_已结单据).ListIndex = 0
        

    '6-分币五舍六入:34519
    strTmp = "0-不处理|1-分币四舍五入|2-分币补整收取|3-分币舍分收取|4-分币四舍六入五成双|5-角币三七作五、二舍八入|6-分币五舍六入"
    For i = 0 To UBound(Split(strTmp, "|"))
        '挂号不支持四舍六入五成双,因挂号是使用医保的结算修正过程处理分币,Oracle中没有四舍六入五成双函数
        If i <> 4 Then cbo(cbo_挂号零钱处理).AddItem Split(strTmp, "|")(i)
        cbo(cbo_收费零钱处理).AddItem Split(strTmp, "|")(i)
        cbo(cbo_结帐零钱处理).AddItem Split(strTmp, "|")(i)
        cbo(cbo_消费卡零钱处理).AddItem Split(strTmp, "|")(i)
    Next
    cbo(cbo_挂号零钱处理).ListIndex = 0
    cbo(cbo_收费零钱处理).ListIndex = 0
    cbo(cbo_结帐零钱处理).ListIndex = 0
    cbo(cbo_消费卡零钱处理).ListIndex = 0
    zlControl.CboSetWidth cbo(cbo_挂号零钱处理).hwnd, 2300
    zlControl.CboSetWidth cbo(cbo_收费零钱处理).hwnd, 2300
    zlControl.CboSetWidth cbo(cbo_结帐零钱处理).hwnd, 2300
    zlControl.CboSetWidth cbo(cbo_消费卡零钱处理).hwnd, 2300
    
    '自动记帐
    With cbo(cbo_自动记帐模式)
        .AddItem "0-标准记帐模式": .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "1-适用内蒙片区记帐模式": .ItemData(.NewIndex) = 1
    End With
    lblAutoChargeNM.Visible = False
    lblAutoChargeNM.Caption = "" & "" & _
        "1.床位:  计入不计出" & vbCrLf & _
        " 2.护理及其他费用:  入院当天按一天计算,出院当天中午12点之前算半天，12点之后算一天" & vbCrLf & _
        " 3.存在中途调整的(转科，转病区，等级变动等),12点以前，按转入科室为准;12点以后以转出科室为准"
    '病人审核方式:49501
    With cbo(cbo_病人审核方式)
        .Clear
        .AddItem "0-未审核不允许结帐": .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "1-审核时不许调整费用和医嘱": .ItemData(.NewIndex) = 1
    End With
        
    cbo(cbo_未审单据结帐).AddItem "0-不检查"
    cbo(cbo_未审单据结帐).AddItem "1-检查并提示"
    cbo(cbo_未审单据结帐).AddItem "2-检查并禁止"
    cbo(cbo_未审单据结帐).ListIndex = 0
    zlControl.CboSetWidth cbo(cbo_未审单据结帐).hwnd, 2000
    
    '临床出诊安排
    With cbo(cbo_临床安排_全院通用号源安排站点)
        .Clear
        .AddItem ""
        strTmp = "Select Distinct b.编号,b.名称 From 部门表 A,Zlnodelist B Where a.站点=b.编号 Order By b.编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption)
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编号 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    With cbo(cbo_临床安排_号码比较方式)
        .Clear
        .AddItem "0-按字符比较":  .ItemData(.NewIndex) = 0
        .AddItem "1-按数值比较": .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    zlControl.CboSetWidth cbo(cbo_临床安排_号码比较方式).hwnd, cbo(cbo_临床安排_号码比较方式).Width * 4 / 3
    
    '挂号相关
    With cbo(cbo_挂号_缺省排序方式)
        .Clear
        .AddItem "0.号别"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
        .AddItem "1.科室-项目"
        .ItemData(.NewIndex) = 1
        .AddItem "2.科室"
        .ItemData(.NewIndex) = 2
    End With
    
    With cbo(cbo_挂号_缺省预约方式)
        .Clear
        strTmp = "Select 编码,名称 From 预约方式"
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption)
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Loop
    End With
    
    With cbo(cbo_挂号_预约有效时间)
        .Clear
        .AddItem "提前"
        .AddItem "延后"
    End With
    
    '票据种类
    lvw(lvw_票据).ListItems.Add , "C1", "收费收据"
    lvw(lvw_票据).ListItems.Add , "C2", "预交收据"
    lvw(lvw_票据).ListItems.Add , "C3", "结帐收据"
    lvw(lvw_票据).ListItems.Add , "C4", "挂号收据"
    
    '刷卡要求输入密码的场合
    With lst(lst_刷卡密码)
        .AddItem "门诊挂号"
        .AddItem "门诊划价"
        .AddItem "门诊收费"
        .AddItem "门诊记帐"
        .AddItem "入院登记"
        .AddItem "住院记帐"
        .AddItem "病人结帐"
        .AddItem "病人预交款"
        .AddItem "检验技师站"
        .AddItem "影像医技站"
        .ListIndex = 0
    End With
    
    strTmp = "病区|1400|1,床位费|630|4,启用日期|1000|1,护理费|630|4,启用日期|1000|1,床位费原始启用日期|0|4,护理费原始启用日期|0|4"
    Call zlControl.MshSetFormat(mshAutoCalc, strTmp, Me.Caption)
    With mshAutoCalc
        .ColAlignmentFixed(0) = 1
        
        '列标题居中对齐
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
    End With
    
    strTmp = "病区|1400|1,收费细目ID|0|1,收费项目|1000|4,计算方式|1000|1,启用日期|1000|1"
    Call zlControl.MshSetFormat(Bill(bill_自动计算), strTmp, Me.Caption)
    With Bill(bill_自动计算)
        .ColData(0) = 3 '设置列的可操作类型
        .ColData(1) = 5
        .ColData(2) = 1
        .ColData(3) = 0
        .ColData(4) = 4
        
        .PrimaryCol = 0
        .Active = True
    End With
    
    strTmp = "病区|1500|1,报警方法|1000|1,报警值|800|7,报警方式1|1300|1,报警方式2|1300|1,报警方式3|1300|1,催款下限|1000|7,催款标准|1000|7"
    Call zlControl.MshSetFormat(Bill(bill_记帐报警), strTmp, Me.Caption)
    With Bill(bill_记帐报警)
        .ColData(0) = 3
        .ColData(1) = 0
        .ColData(2) = 4
        .ColData(3) = 1
        .ColData(4) = 1
        .ColData(5) = 1
        .ColData(6) = 4
        .ColData(7) = 4
        
        .PrimaryCol = 0
        .Active = True
    End With
    With cbo(cbo_收费_自动组合单据)
        .Clear
        .AddItem "收费类别": .ListIndex = 0
        .AddItem "执行科室"
    End With
    Call InitPage(Pg_挂号业务)  '初始页面
    Call InitPage(Pg_门诊收费)  '初始页面
    Call InitPage(Pg_结帐业务)  '初始页面
    
    mblnExistPrintData = GetPrintListHaveData
    With cbo(cbo_收费_票据分配规则)
        .Clear
        .AddItem "1-根据实际打印分配票号"
        .ListIndex = .NewIndex
        .AddItem "2-根据预定规则分配票号"
        .AddItem "3-根据自定义规则分配票号"
        .Enabled = Not mblnExistPrintData
    End With
    Call InitBillRuleCtrl
    Call SetBillRuleParaLocale
    
    With cbo(cbo_记帐操作_记帐后发药)
        .AddItem "0-不发药": .ListIndex = .NewIndex
        .AddItem "1-自动发药"
        .AddItem "2-提示发药"
    End With
    
    With cbo(cbo_结帐_合约单位结帐打印)
        .Clear
        If GetBillUseTypeRec(rsTemp) Then
            rsTemp.Filter = "ID<>0"
            If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!名称)
                rsTemp.MoveNext
            Loop
        End If
    End With
    Call Load补充结算方式
    Call Load自费费用类别
    Call Load脱机医保结算方式
    Call Load一卡通票据格式
    Call Load门诊费用转住院预交票格式
End Sub

Private Sub pic提前颜色_Click()
    Dim strColor As String
    dlgColor.Color = pic提前颜色.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic提前颜色.BackColor = strColor
    Call SetParChange(pic提前颜色, 0, mrsPar, True, strColor)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK Then
        mrsPar.Filter = "(修改状态=1 ANd ErrType =Null) OR  (修改状态=1 And ErrType=" & PET_值超限 & ")"
        If mrsPar.RecordCount > 0 Or mshAutoCalc.Tag = "已修改" Or Bill(bill_自动计算).Tag = "已修改" _
            Or Bill(bill_记帐报警).Tag = "已修改" Or lvw(lvw_单据).Tag = "已修改" Or fra特定收费项目.Tag = "已修改" Then
            
            If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    
    SaveFlexState Bill(bill_自动计算), App.ProductName & "\" & Me.Name & bill_自动计算
    SaveFlexState Bill(bill_记帐报警), App.ProductName & "\" & Me.Name & bill_记帐报警
    SaveFlexState mshAutoCalc, App.ProductName & "\" & Me.Name
    
    Set mrsWarn = Nothing
    Set mrs类别 = Nothing
    Set mrsPar = Nothing
    Set mrsBillUseType = Nothing
End Sub

Private Sub cmdOK_Click()
    If ValidateData() = False Then Exit Sub

    Call Save自动计价项目
    Call Save记帐报警线
    
    Call Save单据操作
    Call Save收费特定项目
    If Save票据分配规则 = False Then Exit Sub
    Call SaveTriageQueuingDep
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lst类别.Visible Then
            lst类别.Visible = False
            Bill(bill_记帐报警).SetFocus
        End If
    End If
End Sub



Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    If Not Me.Visible Then Exit Sub
    
    Select Case Index
    Case cbo_挂号零钱处理, cbo_收费零钱处理, cbo_结帐零钱处理, cbo_消费卡零钱处理
        blnValue = True
        strValue = Split(cbo(cbo_挂号零钱处理).Text & "-", "-")(0) & cbo(cbo_收费零钱处理).ListIndex & _
            cbo(cbo_结帐零钱处理).ListIndex & cbo(cbo_消费卡零钱处理).ListIndex
        Call SetParChange(cbo, cbo_挂号零钱处理, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_收费零钱处理, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_结帐零钱处理, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_消费卡零钱处理, mrsPar, blnValue, strValue)
        Exit Sub
    Case cbo_自动记帐模式
        blnValue = True
        With cbo(cbo_自动记帐模式)
            If .ListIndex >= 0 Then
               strValue = .ItemData(.ListIndex)
            Else
                strValue = "0"
            End If
        End With
        Call SetParChange(cbo, cbo_自动记帐模式, mrsPar, blnValue, strValue)
        chk(chk_下午算半天模式).Visible = InStr(1, ",0,2,", "," & strValue & ",") > 0
        opt护理(0).Enabled = InStr(1, ",0,2,", "," & strValue & ",") > 0
        opt护理(1).Enabled = InStr(1, ",0,2,", "," & strValue & ",") > 0
        lblAutoChargeNM.Visible = Val(strValue) = 1

    Case cbo_收费_自动组合单据
        blnValue = True
        strValue = "0"
        If chk(chk_收费_自动组合单据).value = 1 Then
            strValue = cbo(cbo_收费_自动组合单据).ListIndex + 1
        End If
        Call SetParChange(chk, chk_收费_自动组合单据, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_收费_自动组合单据, mrsPar, blnValue, strValue)
        Exit Sub
    Case cbo_收费_票据分配规则
        Call SetBillRuleParaLocale
        Call SaveBillRuleChange
        Exit Sub
    Case cbo_结帐_合约单位结帐打印
        Call SetParChange(cbo, cbo_结帐_合约单位结帐打印, mrsPar, True, Trim(cbo(Index).Text))
        lblUnit.ForeColor = cbo(Index).ForeColor
        Exit Sub
    Case cbo_挂号_缺省预约方式
        strValue = zlCommFun.GetNeedName(cbo(cbo_挂号_缺省预约方式).Text, "-")
        Call SetParChange(cbo, cbo_挂号_缺省预约方式, mrsPar, True, strValue)
        Exit Sub
    Case cbo_挂号_预约有效时间
        blnValue = True
        strValue = IIF(cbo(cbo_挂号_预约有效时间).ListIndex = 0, 1, -1) * Val(txt(txt_挂号_预约有效时间))
        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
        txt(txt_挂号_预约有效时间).ForeColor = cbo(Index).ForeColor
        lblAvailabilityTimes.ForeColor = cbo(Index).ForeColor
        Exit Sub
    Case cbo_临床安排_全院通用号源安排站点
        strValue = zlStr.NeedCode(cbo(Index).Text, "-")
        Call SetParChange(cbo, Index, mrsPar, True, strValue)
        Exit Sub
    End Select
    Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
End Sub

Private Sub SaveBillRuleChange()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存规则相关参数发生的改变
    '编制:刘兴洪
    '日期:2015-06-18 10:22:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValue As Boolean, strValue As String, intBillRull As Integer
    
    On Error GoTo ErrHandle
    If Not Me.Visible Then Exit Sub
    
    With cbo(cbo_收费_票据分配规则)
        intBillRull = IIF(.ListIndex < 0, 0, .ListIndex)
    End With
    
    strValue = intBillRull & "||"
    '分单据
    strValue = strValue & IIF(chkBillRule(0).value = 1, 1, 0)
    '执行科室
    strValue = strValue & ";" & IIF(chkBillRule(1).value = 1, Val(txtBillRuleNum(0).Text), 0)
    '收据费目
    strValue = strValue & ";" & IIF(chkBillRule(2).value = 1, Val(txtBillRuleNum(1).Text), 0)
    '收费细目
    strValue = strValue & ";" & IIF(chkBillRule(3).value = 1, Val(txtBillRuleNum(2).Text), 0)
    '汇总条件
    strValue = strValue & ";" & IIF(optRuleTotal(0).value, 0, IIF(optRuleTotal(1).value, 1, 2))
    blnValue = True
    Call SetParChange(cbo, cbo_收费_票据分配规则, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 0, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 1, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 2, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 3, mrsPar, blnValue, strValue)
    Call SetParChange(txtBillRuleNum, 0, mrsPar, blnValue, strValue)
    Call SetParChange(txtBillRuleNum, 1, mrsPar, blnValue, strValue)
    Call SetParChange(txtBillRuleNum, 2, mrsPar, blnValue, strValue)
    Call SetParChange(optRuleTotal, 0, mrsPar, blnValue, strValue)
    optRuleTotal(0).ForeColor = optRuleTotal(0).ForeColor
    optRuleTotal(1).ForeColor = optRuleTotal(0).ForeColor
    optRuleTotal(2).ForeColor = optRuleTotal(0).ForeColor
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case chk_挂号_自动刷新挂号安排
        Call SetParTip(txt, txt_挂号_刷新时间, mrsPar)
    Case chk_专家号挂号限制
        Call SetParTip(txt, txt_专家号挂号限制, mrsPar)
    Case chk_专家号预约限制
        Call SetParTip(txt, txt_专家号预约限制, mrsPar)
    Case Else
        Call SetParTip(chk, Index, mrsPar)
    End Select
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub lvw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lvw, Index, mrsPar)
End Sub

Private Sub opt护理_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt护理, Index, mrsPar)
    End If
End Sub

Private Sub opt护理_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt护理_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt护理, Index, mrsPar)
End Sub

Private Sub optRegist_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optRegist, Index, mrsPar)
    End If
End Sub

Private Sub optRegist_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegist_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegist, Index, mrsPar)
End Sub

Private Sub optInExseCharge_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optInExseCharge, Index, mrsPar)
    End If
End Sub

Private Sub optInExseCharge_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optInExseCharge_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optInExseCharge, Index, mrsPar)
End Sub

Private Sub optPrintFact_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optPrintFact, Index, mrsPar)
    End If
End Sub

Private Sub optPrintFact_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintFact_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintFact, Index, mrsPar)
End Sub

Private Sub optPrintSlip_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optPrintSlip, Index, mrsPar)
    End If
End Sub

Private Sub optPrintSlip_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 2 Then
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub pic自付合计色_Click()
    Dim strColor As String
    dlgColor.Color = pic自付合计色.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic自付合计色.BackColor = strColor
    Call SetParChange(pic自付合计色, 0, mrsPar, True, strColor)
End Sub

Private Sub pic缴款栏缴款色_Click()
    Dim strColor As String
    dlgColor.Color = pic缴款栏缴款色.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic缴款栏缴款色.BackColor = strColor
    strColor = strColor & "|" & pic缴款栏退款色.BackColor
    Call SetParChange(pic缴款栏缴款色, 0, mrsPar, True, strColor)
End Sub

Private Sub pic缴款栏退款色_Click()
    Dim strColor As String
    dlgColor.Color = pic缴款栏退款色.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic缴款栏退款色.BackColor = strColor
    strColor = pic缴款栏缴款色.BackColor & "|" & strColor
    Call SetParChange(pic缴款栏退款色, 0, mrsPar, True, strColor)
End Sub

Private Sub pic当前付款未付色_Click()
    Dim strColor As String
    dlgColor.Color = pic当前付款未付色.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic当前付款未付色.BackColor = strColor
    strColor = strColor & "|" & pic当前付款未退色.BackColor
    Call SetParChange(pic当前付款未付色, 0, mrsPar, True, strColor)
End Sub

Private Sub pic当前付款未退色_Click()
    Dim strColor As String
    dlgColor.Color = pic当前付款未退色.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic当前付款未退色.BackColor = strColor
    strColor = pic当前付款未付色.BackColor & "|" & strColor
    Call SetParChange(pic当前付款未退色, 0, mrsPar, True, strColor)
End Sub

Private Sub pic自付合计色_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic自付合计色, 0, mrsPar)
End Sub

Private Sub pic缴款栏缴款色_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic缴款栏缴款色, 0, mrsPar)
End Sub

Private Sub pic缴款栏退款色_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic缴款栏退款色, 0, mrsPar)
End Sub

Private Sub pic当前付款未付色_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic当前付款未付色, 0, mrsPar)
End Sub

Private Sub pic当前付款未退色_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic当前付款未退色, 0, mrsPar)
End Sub

Private Sub optPrintSlip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintSlip, Index, mrsPar)
End Sub

Private Sub optPrintAppoint_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optPrintAppoint, Index, mrsPar)
    End If
End Sub

Private Sub optPrintAppoint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintAppoint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintAppoint, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    If mblnNotChange Then Exit Sub
    If Not Me.Visible Then Exit Sub
    
    Select Case Index
    Case txt_单笔最大金额
        blnValue = True
        strValue = IIF(Val(txt(Index).Text) = 0, "", Val(txt(Index).Text))
    Case txt_姓名模糊查找天数
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_挂号_刷新时间
        strValue = Val(txt(Index).Text)
        Call SetParChange(txt, Index, mrsPar, True, strValue)
        chk(chk_挂号_自动刷新挂号安排).ForeColor = txt(Index).ForeColor
        Exit Sub
    Case txt_挂号_姓名查找天数
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_挂号_预约限制时间_分钟
        blnValue = True
        strValue = Val(txt(txt_挂号_预约限制时间_天数).Text) & "|" & Val(txt(Index).Text)
    Case txt_挂号_病人挂号科室限制
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_挂号_病人预约科室数
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_挂号_病人同科限约N个号
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_专家号挂号限制
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_专家号预约限制
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_挂号_病人同科限挂N个号
        blnValue = True
        strValue = Val(txt(Index).Text) & "|" & IIF(chk(chk_挂号_病人同科限挂N个号_急诊).value = 1, "1", "0")
    Case txt_挂号_预约限制时间_天数
        blnValue = True
        strValue = Val(txt(Index).Text) & "|" & Val(txt(txt_挂号_预约限制时间_分钟).Text)
    Case txt_挂号_N天内不能取消预约号
        If Val(txt(Index).Text) = 0 Then
            chk(chk_挂号_N天内退号需审核).Caption = "当天内取消预约需要通过审核"
        Else
            chk(chk_挂号_N天内退号需审核).Caption = "在" & txt(Index).Text & "天内取消预约需要通过审核"
        End If
    Case txt_收费_自助加收挂号费
        blnValue = True: strValue = ""
        cmdAddedItem.Tag = ""
        If chk(chk_收费_未挂号自动加收挂号费).value = 1 Then strValue = cmdAddedItem.Tag & ";" & txt(Index).Text
        Call SetParChange(txt, Index, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_收费_未挂号自动加收挂号费, mrsPar, blnValue, strValue)
        Exit Sub
    Case txt_补结算_姓名模糊查找天数
        strValue = IIF(chk(chk_补结算_姓名模糊查找).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txt(txt_补结算_姓名模糊查找天数).Text)
        Call SetParChange(chk, chk_补结算_姓名模糊查找, mrsPar, True, strValue)
        Call SetParChange(txt, txt_补结算_姓名模糊查找天数, mrsPar, True, strValue)
        Exit Sub
    Case txt_挂号_预约有效时间
        blnValue = True
        strValue = IIF(cbo(cbo_挂号_预约有效时间).ListIndex = 0, 1, -1) * Val(txt(txt_挂号_预约有效时间))
    Case txt_挂号_病人同一号源限挂N个号
        blnValue = True
        strValue = Val(txt(Index).Text)
    End Select
    
    Call SetParChange(txt, Index, mrsPar, blnValue, strValue)
    
    '设置标签颜色
    Select Case Index
    Case txt_挂号_N天内不能取消预约号
        lblCancelBespeak.ForeColor = txt(Index).ForeColor
    Case txt_挂号_预约限制时间_天数
        lblBespeakDefaultDays.ForeColor = txt(Index).ForeColor
    Case txt_挂号_预约限制时间_分钟
        lblBespeakMinTime.ForeColor = txt(Index).ForeColor
    Case txt_挂号_预约限制时间_分钟
        lblBespeakMinTime.ForeColor = txt(Index).ForeColor
    Case txt_挂号_预约有效时间
        lblAvailabilityTimes.ForeColor = txt(Index).ForeColor
        cbo(cbo_挂号_预约有效时间).ForeColor = txt(Index).ForeColor
    Case txt_挂号_预约失效次数
        lblBreakAnAppointmentNums.ForeColor = txt(Index).ForeColor
    Case txt_挂号_N岁以下输入监护人
        lblGuardian.ForeColor = txt(Index).ForeColor
    Case Else
    End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Select Case Index
    Case txt_划价_门诊单据最大金额, txt_收费_门诊单据最大金额
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
    End Select
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txtBillRuleNum_Change(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SaveBillRuleChange
End Sub

Private Sub txtBillRuleNum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtBillRuleNum_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtBillRuleNum, Index, mrsPar)
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub

Private Sub dtpRegistPlanMode_Validate(Cancel As Boolean)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strValue As String
    If mblnInstantActive Then Exit Sub
    strSQL = "Select 1 From 病人挂号记录 Where 发生时间 > [1] And 记录状态=1 And 出诊记录ID Is Null And Rownum <2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(dtpRegistPlanMode.value))
    If Not rsTmp.EOF Then
        MsgBox "启用日期之后存在计划排班模式的挂号或者预约记录,请调整启用日期!", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    strValue = "1|" & Format(dtpRegistPlanMode.value, "yyyy-mm-dd hh:mm:ss")
    Call SetParChange(optRegistPlanMode, 0, mrsPar, True, strValue)
End Sub

Private Sub optRegistPlanMode_Click(Index As Integer)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strValue As String
    If mblnInstantActive Then Exit Sub
    If mblnNotChange Then Exit Sub
    If optRegistPlanMode(0).value = 1 Or optRegistPlanMode(0).value = True Then
        strSQL = "Select 1 From 病人挂号记录 Where 出诊记录ID Is Not Null And 记录状态=1 And Rownum < 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "已经存在出诊表排班模式下的挂号记录,不能再切换回计划排班模式!", vbInformation, gstrSysName
            mblnNotChange = True
            optRegistPlanMode(1).value = 1
            mblnNotChange = False
            Exit Sub
        End If
        strValue = 0
        Call SetParChange(optRegistPlanMode, 0, mrsPar, True, strValue)
        
        fraNewPaln.Visible = False
        chk(chk_只对医内医生进行挂号安排).Visible = True
        chk(chk_存在预约挂号单禁止删除安排).Visible = True
    Else
        '出诊表排班模式
        strSQL = "Select 1 From 临床出诊表 Where 发布时间 Is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.EOF Then
            MsgBox "不存在任何临床出诊数据,不能切换为出诊表排班模式!", vbInformation, gstrSysName
            mblnNotChange = True
            optRegistPlanMode(0).value = 1
            mblnNotChange = False
            Exit Sub
        End If
        dtpRegistPlanMode.Enabled = True
        strValue = "1|" & Format(dtpRegistPlanMode.value, "yyyy-mm-dd hh:mm:ss")
        Call SetParChange(optRegistPlanMode, 0, mrsPar, True, strValue)
        
        fraNewPaln.Visible = True
        chk(chk_只对医内医生进行挂号安排).Visible = False
        chk(chk_存在预约挂号单禁止删除安排).Visible = False
    End If
End Sub

Private Sub Save记帐报警线()
    Dim strTmp As String
    Dim i As Integer
    Dim strArr
    Dim str适用病人 As String
    
    If Bill(bill_记帐报警).Tag = "已修改" Then
    
        '先处理删除的适用病人记帐报警
        On Error GoTo ErrHandle
        If mstrDel适用病人 <> "" Then
            mstrDel适用病人 = mstrDel适用病人 & ";"
            strArr = Split(mstrDel适用病人, ";")
            For i = 0 To UBound(strArr) - 1
                If strArr(i) <> "" Then
                    str适用病人 = strArr(i)
                    strTmp = str适用病人 & "|"
                    gstrSQL = "zl_记帐报警线_Modify('" & strTmp & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            Next
        End If
        
        '按适用病人分批保存
        mrsWarn.Filter = 0
        For i = 1 To tab报警.Tabs.Count
            strTmp = ""
            str适用病人 = tab报警.Tabs.Item(i).Caption
            
            mrsWarn.Filter = "适用病人='" & str适用病人 & "'"
            Do While Not mrsWarn.EOF
                strTmp = strTmp & NVL(mrsWarn!病区ID) & "," & mrsWarn!报警方法 & "," & _
                    mrsWarn!报警值 & "," & NVL(mrsWarn!报警标志1) & "," & NVL(mrsWarn!报警标志2) & "," & NVL(mrsWarn!报警标志3) & "," & NVL(mrsWarn!催款下限) & "," & NVL(mrsWarn!催款标准) & ","
                mrsWarn.MoveNext
            Loop
            
            strTmp = str适用病人 & "|" & strTmp
            
            gstrSQL = "zl_记帐报警线_Modify('" & strTmp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
        
        Bill(bill_记帐报警).Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save自动计价项目()
    Dim str病区ID As String
    Dim str细目ID As String
    Dim str计算标志 As String
    Dim str启用日期 As String
    Dim lngTemp As Long, i As Long, blnTrans As Boolean
    
    On Error GoTo ErrHandle
    If mshAutoCalc.Tag = "已修改" Or Bill(bill_自动计算).Tag = "已修改" Then
    
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "Zl_自动计价项目_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "删除自动计价项目")
        
        '按床位
        For i = 1 To mshAutoCalc.Rows - 1
            lngTemp = mshAutoCalc.RowData(i)
            If lngTemp <> 0 Then
                If mshAutoCalc.TextMatrix(i, 1) <> "" Then
                    str病区ID = str病区ID & lngTemp & ","
                    str细目ID = str细目ID & ","
                    str计算标志 = str计算标志 & "1,"
                    str启用日期 = str启用日期 & mshAutoCalc.TextMatrix(i, 2) & ","
                End If
                If mshAutoCalc.TextMatrix(i, 3) <> "" Then
                    str病区ID = str病区ID & lngTemp & ","
                    str细目ID = str细目ID & ","
                    str计算标志 = str计算标志 & "2,"
                    str启用日期 = str启用日期 & mshAutoCalc.TextMatrix(i, 4) & ","
                End If
            End If
            If (i Mod 100) = 0 Or i >= mshAutoCalc.Rows - 1 Then
                gstrSQL = "zl_自动计价项目_Modify('" & str病区ID & "','" & str细目ID & "','" & str计算标志 & "','" & str启用日期 & "' )"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                str病区ID = ""
                str细目ID = ""
                str计算标志 = ""
                str启用日期 = ""
            End If
        Next
        '按病区
        For i = 1 To Bill(bill_自动计算).Rows - 1
            lngTemp = Bill(bill_自动计算).RowData(i)
            If lngTemp <> 0 And Bill(bill_自动计算).TextMatrix(i, 1) <> "" Then
                If Bill(bill_自动计算).TextMatrix(i, 1) <> "" Then
                    str病区ID = str病区ID & lngTemp & ","
                    str细目ID = str细目ID & Bill(bill_自动计算).TextMatrix(i, 1) & ","
                    str计算标志 = str计算标志 & Switch(Left(Bill(bill_自动计算).TextMatrix(i, 3), 1) = "1", "6", Left(Bill(bill_自动计算).TextMatrix(i, 3), 1) = "2", "8", True, "7") & ","
                    str启用日期 = str启用日期 & Bill(bill_自动计算).TextMatrix(i, 4) & ","
                End If
            End If
            If (i Mod 100) = 0 Or i >= Bill(bill_自动计算).Rows - 1 Then
                gstrSQL = "zl_自动计价项目_Modify('" & str病区ID & "','" & str细目ID & "','" & str计算标志 & "','" & str启用日期 & "' )"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                str病区ID = ""
                str细目ID = ""
                str计算标志 = ""
                str启用日期 = ""
            End If
        Next
        gcnOracle.CommitTrans: blnTrans = False
        mshAutoCalc.Tag = ""
        Bill(bill_自动计算).Tag = ":"
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub


Private Sub Save单据操作()
    Dim lst As ListItem
    Dim i As Integer, blnTrans As Boolean
    
    '首先删除以前的所有单据操作
    On Error GoTo ErrHandle
    
    If lvw(lvw_单据).Tag = "已修改" Then
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "zl_单据操作控制_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        '再增加新的
        For Each lst In lvw(lvw_单据).ListItems
            gstrSQL = "zl_单据操作控制_Insert(" & lst.Tag & "," & lst.ListSubItems(1).Tag & _
                        "," & lst.SubItems(2) & "," & IIF(lst.SubItems(3) = "是", 1, 0) & "," & IIF(lst.SubItems(4) = "", "NULL", lst.SubItems(4)) & " )"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lvw(lvw_单据).Tag = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Save收费特定项目()
    Dim strTmp As String
    
    If fra特定收费项目.Tag = "已修改" Then
        '逐个对参数进行保存
        On Error GoTo ErrHandle
        If txtCmd(0).Text <> "" Then
            strTmp = "病历费," & txtCmd(0).Tag & ","
        End If
        If txtCmd(1).Text <> "" Then
            strTmp = strTmp & "工本费," & txtCmd(1).Tag & ","
        End If
        
        If txtCmd(3).Text <> "" Then
            strTmp = strTmp & "普通配置费," & txtCmd(3).Tag & ","
        End If
        
        If txtCmd(4).Text <> "" Then
            strTmp = strTmp & "肿瘤配置费," & txtCmd(4).Tag & ","
        End If
        
        If strTmp <> "" Then
            gstrSQL = "zl_收费特定项目_Modify('" & strTmp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        
        fra特定收费项目.Tag = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check记帐报警() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim arr类别() As String
        
    With Bill(bill_记帐报警)
        For lngRow = 1 To .Rows - 2
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) And .TextMatrix(lngTemp, 2) <> "" Then
                        MsgBox "病区“" & .TextMatrix(lngTemp, 0) & "”出现多次。", vbExclamation, gstrSysName
                        .Row = lngTemp: .Col = 0: .SetFocus: Exit Function
                    End If
                Next
                '刘兴洪 问题: 34770   日期:2010-12-21 10:54:02
                If Val(.TextMatrix(lngRow, 6)) > 999999999 Or Val(.TextMatrix(lngRow, 6)) < 0 Then
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”中的催款下限设置有误(应该在0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 6: .SetFocus: Exit Function
                End If
                If Val(.TextMatrix(lngRow, 7)) > 999999999 Or Val(.TextMatrix(lngRow, 7)) < 0 Then
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”中的催款标准有误(应该在0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 7: .SetFocus: Exit Function
                End If
                
            End If
        Next
        
        '检查同一病区不同报警方式的类别是否一个都没有设置或重复
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                If Trim(.TextMatrix(lngRow, 3)) = "" And Trim(.TextMatrix(lngRow, 4)) = "" And Trim(.TextMatrix(lngRow, 5)) = "" Then
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”未设置要报警的收费类别。", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If (.TextMatrix(lngRow, 3) = "所有类别" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 4) = "所有类别" And (Trim(.TextMatrix(lngRow, 3)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 5) = "所有类别" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 3)) <> "")) Then
                    
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”不同的报警方式包含相同的收费类别。", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If .TextMatrix(lngRow, 3) <> "所有类别" And Trim(.TextMatrix(lngRow, 4)) <> "所有类别" And Trim(.TextMatrix(lngRow, 5)) <> "所有类别" Then
                    For lngCol1 = 3 To 5
                        If Trim(.TextMatrix(lngRow, lngCol1)) <> "" Then
                            For lngCol2 = 3 To 5
                                If lngCol1 <> lngCol2 Then
                                    arr类别 = Split(.TextMatrix(lngRow, lngCol1), ",")
                                    For lngTemp = 0 To UBound(arr类别)
                                        If InStr("," & .TextMatrix(lngRow, lngCol2) & ",", "," & arr类别(lngTemp) & ",") > 0 Then
                                            MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”不同的报警方式包含相同的收费类别。", vbExclamation, gstrSysName
                                            .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With
    
    Check记帐报警 = True
End Function

Private Function ValidateData() As Boolean
    Dim lngRow As Long, lngTmp As Long
    Dim lngIndex As Long, strTmp As String
    Dim i As Integer
    
    
    '检查自动计算项目是否重复
    With Bill(bill_自动计算)
        For lngRow = 1 To .Rows - 2
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                For lngTmp = lngRow + 1 To .Rows - 1
                    If .RowData(lngRow) = .RowData(lngTmp) And .TextMatrix(lngRow, 1) = .TextMatrix(lngTmp, 1) Then
                        MsgBox "病区为“" & .TextMatrix(lngTmp, 0) & "”、收费细目为“" & _
                            .TextMatrix(lngTmp, 2) & "”" & vbCrLf & "这种组合出现多次。", vbExclamation, gstrSysName
                        .Row = lngTmp
                        .Col = 0
                        .SetFocus
                        Exit Function
                    End If
                Next
            End If
        Next
    End With
    
    '检查自动计算项目的启用日期
    With Bill(bill_自动计算)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                If Not IsDate(.TextMatrix(lngRow, 4)) Then
                    MsgBox "自动计算项目的启用日期未设置或日期格式不正确。", vbInformation, gstrSysName
                    .Row = lngRow
                    .Col = 4
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
   
    
    If CheckParChanged(txtUD, ud_费用金额保留位数, mrsPar) Then
        If MsgBox("你已调整了费用金额保留小数位，可能会引起小数计算误差！是否继续？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
         
            Exit Function
        End If
    End If
    
    If CheckParChanged(txtUD, ud_费用单价保留位数, mrsPar) Then
        If MsgBox("你已调整了费用单价保留小数位，可能会引起小数计算误差！是否继续？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then

            Exit Function
        End If
    End If
    If VsaliedData_收费 = False Then Exit Function
    
    ValidateData = True
End Function
Private Function VsaliedData_收费() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查收费相关参数设置的合法性
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-18 11:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    
    If cbo(cbo_收费_票据分配规则).ListIndex = 1 Then
        If chkBillRule(0).value = 0 And chkBillRule(1).value = 0 And chkBillRule(2).value = 0 And chkBillRule(3).value = 0 Then
            MsgBox "注意:" & vbCrLf & "    票据号分配规则按『" & cbo(cbo_收费_票据分配规则).Text & "』的必须设置一种分配规则,请检查!", vbInformation + vbOKOnly
           ' stab.Tab = 3
            If chkBillRule(0).Enabled And chkBillRule(0).Visible Then chkBillRule(0).SetFocus
            Exit Function
        End If
    End If
    VsaliedData_收费 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function VsaliedData_结帐() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结帐相关参数设置的合法性
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-18 11:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    With cbo(cbo_结帐_合约单位结帐打印)
        If .ListIndex < 0 Then
            MsgBox "注意:" & vbCrLf & _
                   "    你未选择合约单位结帐时所使用的何种票据!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    VsaliedData_结帐 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub tab报警_Click()
    Dim lngRow As Long
    
    mrsWarn.Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
    
    With Bill(bill_记帐报警)
        If mrsWarn.RecordCount = 0 Then
            .ClearBill
            mlngPreFind = 1
            .Rows = 2: .Row = 1: .Col = 1
            .RowData(1) = 0
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            .TextMatrix(1, 7) = ""
        Else
            .ClearBill
            mlngPreFind = 1
            .Rows = mrsWarn.RecordCount + 1: .Row = 1: .Col = 1
            lngRow = 1
            Do Until mrsWarn.EOF
                .RowData(lngRow) = NVL(mrsWarn!病区ID, 0)
                .TextMatrix(lngRow, 0) = IIF(IsNull(mrsWarn!病区ID), "*门诊*", mrsWarn!病区码 & "-" & mrsWarn!病区名)
                .TextMatrix(lngRow, 1) = IIF(mrsWarn!报警方法 = 1, "1-累计费用", "2-每日费用")
                .TextMatrix(lngRow, 2) = Format(mrsWarn!报警值, "##########0.00;-##########0.00;0.00;0.00")
                
                .TextMatrix(lngRow, 3) = Get类别名称串(NVL(mrsWarn!报警标志1), mrs类别)
                .TextMatrix(lngRow, 4) = Get类别名称串(NVL(mrsWarn!报警标志2), mrs类别)
                .TextMatrix(lngRow, 5) = Get类别名称串(NVL(mrsWarn!报警标志3), mrs类别)
                .TextMatrix(lngRow, 6) = Format(mrsWarn!催款下限, "###0.00;-###0.00;0.00;0.00")
                .TextMatrix(lngRow, 7) = Format(mrsWarn!催款标准, "###0.00;-###0.00;0.00;0.00")
                
                lngRow = lngRow + 1
                mrsWarn.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub pic提前颜色_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic提前颜色, 0, mrsPar)
End Sub

Private Function IsRecord(ByVal strFind As String) As Boolean
'功能:分析输入内容是否是有效的数据库中表的记录
'参数:strFind SQL语句的条件
'返回值:有效返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    
    rsTemp.CursorLocation = adUseClient
    IsRecord = False
    If InStr(strFind, "'") > 0 Then
        MsgBox "输入了非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    gstrSQL = "select distinct A.编码,A.名称,A.规格,A.计算单位 ,A.id from 收费细目 A,收费别名 B,收费类别 C " & _
         " where A.ID=B.收费细目ID and A.是否变价 <> 1 and A.末级=1 and  A.类别=C.编码 and  (A.编码 like [1] or B.名称 like [2] " & _
         " or  upper(B.简码) like [2]) and " & Where撤档时间("A")
          
    With Bill(bill_自动计算)
        If .TextMatrix(.Row, 3) <> "2-计算一次" Then
            gstrSQL = gstrSQL & " and C.编码 Not In('4','5','6','7') "
        End If
    End With
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind & "%", "%" & UCase(strFind) & "%")
    
    If rsTemp.RecordCount < 1 Then Exit Function
    If rsTemp.RecordCount > 1 Then
        gstrSQL = ""
        gstrSQL = frmSelCurr.ShowCurrSel(Me, rsTemp, "编码,1000,0,2;名称,1800,0,1;规格,2300,0,2;计算单位,1000,0,2;id,0,0,2", -1, "选择收费细目")
        If gstrSQL = "" Then
            Exit Function
        End If
        If Bill(bill_自动计算).TextMatrix(Bill(bill_自动计算).Row, 3) <> "2-计算一次" Then
            If Not IsRaiseByDate(Val(Split(gstrSQL, ";")(4))) Then
                MsgBox "项目[" & Split(gstrSQL, ";")(1) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_自动计算)
            .TextMatrix(.Row, 1) = Split(gstrSQL, ";")(4) ' rsTemp("ID")
            .TextMatrix(.Row, 2) = Split(gstrSQL, ";")(1) 'rsTemp("名称")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-按收治日"
            End If
        End With
    Else
        rsTemp.MoveFirst
        If Bill(bill_自动计算).TextMatrix(Bill(bill_自动计算).Row, 3) <> "2-计算一次" Then
            If Not IsRaiseByDate(Val(rsTemp!ID)) Then
                MsgBox "项目[" & rsTemp!名称 & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_自动计算)
            .TextMatrix(.Row, 1) = rsTemp("ID")
            .TextMatrix(.Row, 2) = rsTemp("名称")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-按收治日"
            End If
        End With
    End If
    IsRecord = True
End Function

Private Sub bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim lmX As Integer
    Dim lmY As Integer
    Dim strTmp As String
    
    With Bill(Index)
        If Index = bill_自动计算 Then
            If .Col = 4 And KeyCode = vbKeyReturn Then
                If .Text <> "" And Not IsDate(.Text) Then
                    If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                        .Text = ""
                        MsgBox "请输入正确的日期格式(yyyy-mm-dd或者yyyymmdd)。", vbInformation, gstrSysName
                    Else
                        .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                    End If
                    .TextMatrix(.Row, .Col) = .Text
                End If
            End If
                
            If .Col = 2 Then
                '收费细目列只处理回车键
                If KeyCode = vbKeyDelete Then .Tag = "已修改": Exit Sub   '118682
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '到下一个控件
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    '选择收费细目
                    If IsRecord(.Text) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = .TextMatrix(.Row, 2)
                    If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-按收治日"
                    
                End If
            End If
        End If
        
        If Index = bill_记帐报警 Then
            If .Col = 2 Then
                '报警值列只处理回车键
                If KeyCode = vbKeyDelete Then .Tag = "已修改": Exit Sub     '118682
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '到下一个控件
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    '判断输入的合法性
                    .Text = Format(.Text, "##########0.00;-##########0.00;0.00;0,00")
                    
                End If
            ElseIf .Col = 3 Then
                '禁止输入报警类别
                If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete Then KeyCode = 0: Cancel = True
            ElseIf .Col = 6 Or .Col = 6 Then
                .Text = Format(.Text, "###0.00;-###.00;0.00;0,00")
                
            End If
        End If
        
        .Tag = "已修改"
    End With
End Sub

Private Sub SetDrugStore()
    Dim lngType As Long, strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long, lngRow As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strTmp = "'西药房','中药房','成药房'"
    Set rsTmp = GetDepartments(strTmp, "1,2,3")
    
    With vsfDrugStore(0)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            lngRow = 1
            rsTmp.Filter = "工作性质='西药房'"
            If rsTmp.RecordCount > 0 Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(lngRow, 0) = 0
                    .TextMatrix(lngRow, 1) = rsTmp!名称
                    .TextMatrix(lngRow, 2) = "自动分配"
                    .RowData(lngRow) = Val(rsTmp!ID)
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Next
            End If
        End If
    End With
    
    With vsfDrugStore(1)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            lngRow = 1
            rsTmp.Filter = "工作性质='中药房'"
            If rsTmp.RecordCount > 0 Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(lngRow, 0) = 0
                    .TextMatrix(lngRow, 1) = rsTmp!名称
                    .TextMatrix(lngRow, 2) = "自动分配"
                    .RowData(lngRow) = Val(rsTmp!ID)
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Next
            End If
        End If
    End With
    
    With vsfDrugStore(2)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            lngRow = 1
            rsTmp.Filter = "工作性质='成药房'"
            If rsTmp.RecordCount > 0 Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(lngRow, 0) = 0
                    .TextMatrix(lngRow, 1) = rsTmp!名称
                    .TextMatrix(lngRow, 2) = "自动分配"
                    .RowData(lngRow) = Val(rsTmp!ID)
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Next
            End If
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_自动计算 Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            Else
                .TxtCheck = False
            End If
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '切换计算标志
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "0"
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "1-按床日"
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-计算一次"
                            Case Else
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                        MsgBox "药品类和卫材类的自动计算方式不能改变。", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "0-按收治日"
                        End Select
                        
                    Case vbKey0
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "药品类和卫材类的自动计算方式不能改变。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "0-按收治日"
                        
                    Case vbKey1
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "药品类和卫材类的自动计算类型不能改变。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "1-按床日"
                        
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-计算一次"
                        
                End Select
            End If
        
        ElseIf Index = bill_记帐报警 Then
            .TxtCheck = False
            If .Col = 1 Then
                
                '切换报警方法
                Select Case KeyAscii
                    Case Asc(" ")
                        '切换计算标志
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-每日费用"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-累计费用"
                        End Select
                        
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-累计费用"
                        
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-每日费用"
                        
                End Select
                If InStr(.TextMatrix(.Row, 1), "每日费用") > 0 Then
                    .TextMatrix(.Row, 4) = ""  '每日费用无报警方式2
                End If
            ElseIf InStr(1, "267", .Col) > 0 Then
                    .TxtCheck = True
                    .TextMask = "0123456789-"
                    .MaxLength = 10
            End If
        End If
        
        .Tag = "已修改"
    End With

End Sub

Private Sub Set类别选择(str类别 As String)
'功能：根据类似"检查,治疗..."的串设置列表的选择情况
    Dim i As Integer, j As Integer
    Dim arr类别() As String
    
    For i = 0 To lst类别.ListCount - 1
        lst类别.Selected(i) = False
    Next
    
    If Trim(str类别) = "" Then
        Exit Sub
    ElseIf str类别 = "所有类别" Then
        For i = 0 To lst类别.ListCount - 1
            lst类别.Selected(i) = (i = 0)
        Next
    Else
        lst类别.Selected(0) = False
        arr类别 = Split(str类别, ",")
        For i = 0 To UBound(arr类别)
            For j = 1 To lst类别.ListCount - 1
                If lst类别.List(j) = arr类别(i) Then
                    lst类别.Selected(j) = True: Exit For
                End If
            Next
        Next
    End If
    
    For i = 0 To lst类别.ListCount - 1
        If lst类别.Selected(i) Then
            lst类别.TopIndex = i: Exit For
        End If
    Next
End Sub

Private Sub bill_CommandClick(Index As Integer)
'通过按钮选择收费细目
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim rsTmp As New ADODB.Recordset
    
    If Index = bill_记帐报警 Then
        With Bill(Index)
            Call Set类别选择(.TextMatrix(.Row, .Col))
            
            lst类别.Left = .Left + .MsfObj.CellLeft
            If .Top + .MsfObj.CellTop + .MsfObj.CellHeight + lst类别.Height <= .Container.Height Then
                lst类别.Top = .Top + .MsfObj.CellTop + .MsfObj.CellHeight
            Else
                lst类别.Top = .Top + .MsfObj.CellTop - lst类别.Height - 30
            End If
            lst类别.Width = .MsfObj.CellWidth
            lst类别.ZOrder
            lst类别.Visible = True
            lst类别.SetFocus
        End With
    End If
    
    If Index = bill_自动计算 Then
        With Bill(bill_自动计算)
            If .TextMatrix(.Row, 3) <> "2-计算一次" Then
                blnRe = frmChargeListSel.ShowTree(strID, str名称, False)
            Else
                blnRe = frmChargeListSel.ShowTree(strID, str名称, True)
            End If
            If blnRe And strID <> "" Then
                If .TextMatrix(.Row, 3) <> "2-计算一次" Then
                    If Not IsRaiseByDate(strID) Then
                        MsgBox "项目[" & str名称 & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .SetFocus
                .TextMatrix(.Row, 1) = strID
                .TextMatrix(.Row, 2) = str名称
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-按收治日"
            End If
        End With
    End If
    Bill(Index).Tag = "已修改"
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            .RowData(.Row) = .ItemData(.ListIndex)

            If Index = bill_记帐报警 Then
                If .TextMatrix(.Row, 1) = "" Then .TextMatrix(.Row, 1) = "1-累计费用"
            End If
            
            Bill(Index).Tag = "已修改"
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'处理最后一列的变化
With Bill(Index)
    If .MouseRow = 0 Then Exit Sub
    
    If Index = bill_自动计算 Then
        If .MouseCol <> 3 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "0"
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或者其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "1-按床日"
            Case "1"
                .TextMatrix(.Row, .Col) = "2-计算一次"
            Case Else
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                        MsgBox "药品类和卫材类的自动计算方式不能改变。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或者其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "0-按收治日"
        End Select
    ElseIf Index = bill_记帐报警 Then
        If .MouseCol <> .Cols - 1 And .MouseCol <> 1 Then Exit Sub
        If .Col = 1 Then
            .TextMatrix(.Row, 1) = IIF(Left(.TextMatrix(.Row, 1), 1) = "1", "2-每日费用", "1-累计费用")
            If InStr(.TextMatrix(.Row, 1), "每日费用") > 0 Then
                .TextMatrix(.Row, 4) = ""  '每日费用无报警方式2
                
                '为“每日费用”时判断一下金额不能为负数
                If IsNumeric(.TextMatrix(.Row, 2)) Then
                    If Val(.TextMatrix(.Row, 2)) < 0 Then
                        .TextMatrix(.Row, 2) = "0.00"
                    End If
                Else
                    .TextMatrix(.Row, 2) = "0.00"
                End If
            End If
        End If
    End If
    .Tag = "已修改"
End With
    
End Sub


Private Sub lst类别_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 And lst类别.Selected(Item) Then
        For i = 1 To lst类别.ListCount - 1
            lst类别.Selected(i) = False
        Next
    ElseIf Item > 0 And lst类别.Selected(Item) Then
        lst类别.Selected(0) = False
    End If
End Sub

Private Sub lst类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lst类别_Validate(False)
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub lst类别_LostFocus()
    lst类别.Visible = False
End Sub

Private Sub lst类别_Validate(Cancel As Boolean)
    Dim objGrid As Object, i As Integer
    
    Set objGrid = Bill(bill_记帐报警)
    
    With objGrid
        .TextMatrix(.Row, .Col) = Get类别选择
        If .TextMatrix(.Row, .Col) = "所有类别" Then
            For i = 3 To 5
                If i <> .Col Then .TextMatrix(.Row, i) = " "
            Next
        End If
    End With
    
End Sub


Private Function Get类别选择() As String
'功能：根据类别选择框选择的情况返回类似"检查,治疗..."的串
    Dim i As Integer, strTmp As String
    
    If lst类别.Selected(0) Then
        Get类别选择 = "所有类别"
    Else
        For i = 1 To lst类别.ListCount - 1
            If lst类别.Selected(i) Then
                strTmp = strTmp & "," & lst类别.List(i)
            End If
        Next
        Get类别选择 = Mid(strTmp, 2)
        If Get类别选择 = "" Then Get类别选择 = " " '为了能回车新增行
    End If
End Function

Private Function Get类别名称串(str类别 As String, rs类别 As ADODB.Recordset) As String
'功能：将类似"CDEFG"的类别转换为类似"检查,检验..."串
    Dim i As Integer, strTmp As String
    
    If str类别 = "" Then
        Get类别名称串 = " " '为了能按回车新增行
        Exit Function
    End If
    
    If str类别 = "-" Then
        Get类别名称串 = "所有类别"
        Exit Function
    End If
    
    For i = 1 To Len(str类别)
        rs类别.Filter = "编码='" & Mid(str类别, i, 1) & "'"
        If Not rs类别.EOF Then strTmp = strTmp & "," & rs类别!类别
    Next
    Get类别名称串 = Mid(strTmp, 2)
End Function

Private Function Get类别编码串(str类别 As String) As String
'功能：根据类似"检查,治疗"的串返回类似"CDEFG"的串
    Dim i As Integer, j As Integer
    Dim arr类别() As String, strTmp As String
    
    If Trim(str类别) = "" Then Exit Function
    
    If str类别 = "所有类别" Then
        Get类别编码串 = "-"
    Else
        arr类别 = Split(str类别, ",")
        For i = 0 To UBound(arr类别)
            For j = 1 To lst类别.ListCount - 1
                If lst类别.List(j) = arr类别(i) Then
                    strTmp = strTmp & Chr(lst类别.ItemData(j))
                    Exit For
                End If
            Next
        Next
        Get类别编码串 = strTmp
    End If
End Function


Private Sub bill_AfterAddRow(Index As Integer, Row As Long)
    If Index = bill_记帐报警 Then
        With Bill(Index)
            .TextMatrix(Row, 3) = " "
            .TextMatrix(Row, 4) = " "
            .TextMatrix(Row, 5) = " "
            .TextMatrix(Row, 6) = ""
            .TextMatrix(Row, 7) = ""
        End With
    End If
    
    If Index = bill_自动计算 Then
        With Bill(Index)
            .TextMatrix(Row, 3) = "0-按收治日"
            .TextMatrix(Row, 4) = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd")
        End With
    End If
    
    Bill(Index).Tag = "已修改"
End Sub

Private Sub Bill_EditKeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_自动计算 Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            End If
        End If
    End With
End Sub

Private Sub bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    '禁止输入报警类别
    With Bill(Index)
        If Index = bill_记帐报警 And .Col >= 3 Then
            If .Col = 6 Or .Col = 7 Then
                .TxtEnable = True
            Else
                .TxtEnable = False
            End If
        Else
            .TxtEnable = True
        End If
        
        If Index = bill_记帐报警 And .Col = 4 Then  '报警方式2
            If InStr(.TextMatrix(.Row, 1), "每日费用") > 0 Then
                .ColData(4) = 5 '每日费用不能编辑报警方式2
            Else
                .ColData(4) = 1
            End If
        End If
        If Index = bill_记帐报警 Then
            Select Case .Col
            Case 6, 7
                .ColData(.Col) = 4
            Case Else
            End Select
        End If
    End With
    
End Sub

Private Sub bill_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Bill(Index)
        If Index = bill_记帐报警 And .MouseCol >= 3 And .MouseRow > 0 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub bill_Validate(Index As Integer, Cancel As Boolean)
    Dim lngRow As Long
    
    If Index = bill_记帐报警 Then
        
        If MouseInRect(cmdCancel.hwnd) Then Exit Sub
        
        '检查记帐报警设置
        If Not Check记帐报警 Then Cancel = True: Exit Sub
        
        '收集记帐报警数据
        With mrsWarn
            .Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
            Do While Not .EOF
                .Delete
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        
        With Bill(bill_记帐报警)
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                    mrsWarn.AddNew
                    mrsWarn!适用病人 = tab报警.SelectedItem.Caption
                    
                    If .RowData(lngRow) <> 0 Then
                        mrsWarn!病区ID = .RowData(lngRow)
                        mrsWarn!病区码 = Split(.TextMatrix(lngRow, 0), "-")(0)
                        mrsWarn!病区名 = Split(.TextMatrix(lngRow, 0), "-")(1)
                    End If
                    
                    mrsWarn!报警方法 = CInt(Left(.TextMatrix(lngRow, 1), 1))
                    mrsWarn!报警值 = CCur(.TextMatrix(lngRow, 2))
                    
                    mrsWarn!报警标志1 = Get类别编码串(.TextMatrix(lngRow, 3))
                    mrsWarn!报警标志2 = Get类别编码串(.TextMatrix(lngRow, 4))
                    mrsWarn!报警标志3 = Get类别编码串(.TextMatrix(lngRow, 5))
                    
                    mrsWarn!催款下限 = Round(Val(.TextMatrix(lngRow, 6)), 2)
                    mrsWarn!催款标准 = Round(Val(.TextMatrix(lngRow, 7)), 2)
                    
                    mrsWarn.Update
                End If
            Next
        End With
    End If
End Sub


Private Sub cmdOneCard_Click(Index As Integer)
    
    Select Case Index
        Case 0
            frmOneCard.mbytInFun = 0
            Call frmOneCard.ShowMe(Me)
            Call LoadOneCard
        Case 1
            If lvw(lvw_一卡通).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_一卡通).SelectedItem
                frmOneCard.mbytInFun = 1
                Call frmOneCard.ShowMe(Me, Mid(.Key, 2), .SubItems(1), .SubItems(2), .SubItems(3), IIF(.SubItems(4) = "启用:标准一卡通", 2, IIF(.SubItems(4) = "启用:仅涉及扣卡", 1, 0)))
                Call LoadOneCard
            End With
        Case 2
            If lvw(lvw_一卡通).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_一卡通).SelectedItem
                If MsgBox("你确实要删除“" & .SubItems(1) & "”吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    Call frmOneCard.DelOneCardRec(Val(Mid(.Key, 2)))
                    Call LoadOneCard
                End If
            End With
    End Select
End Sub


Private Sub cmdWarnDel_Click()
    If tab报警.SelectedItem.Caption = "普通病人" Then
        MsgBox """" & tab报警.SelectedItem.Caption & """报警方案不允许删除。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确实要删除""" & tab报警.SelectedItem.Caption & """报警方案吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    With mrsWarn
        .Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
        
        '记录删除的适用病人类型
        If InStr(1, mstrDel适用病人, tab报警.SelectedItem.Caption) = 0 Then
            mstrDel适用病人 = IIF(mstrDel适用病人 = "", "", mstrDel适用病人 & ";") & tab报警.SelectedItem.Caption
        End If
        
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    tab报警.Tabs.Remove tab报警.SelectedItem.Index
    tab报警.Tabs(1).Selected = True
    
    Bill(bill_记帐报警).Tag = "已修改"
End Sub

Private Sub cmdWarnNew_Click()
    Dim strName As String, strCopy As String
    Dim strSchemes As String, i As Integer
    Dim rsCopy As ADODB.Recordset
    
    For i = 1 To tab报警.Tabs.Count
        strSchemes = strSchemes & "," & tab报警.Tabs(i).Caption
    Next
    
    strName = frmWarnEdit.ShowMe(Me, Mid(strSchemes, 2), strCopy)
    If strName = "" Then Exit Sub
    
    '复制内容
    Set rsCopy = mrsWarn.Clone
    rsCopy.Filter = "适用病人='" & strCopy & "'"
    Do While Not rsCopy.EOF
        mrsWarn.AddNew
        mrsWarn!适用病人 = strName
        mrsWarn!病区ID = rsCopy!病区ID
        mrsWarn!病区码 = rsCopy!病区码
        mrsWarn!病区名 = rsCopy!病区名
        mrsWarn!报警方法 = rsCopy!报警方法
        mrsWarn!报警值 = rsCopy!报警值
        mrsWarn!报警标志1 = rsCopy!报警标志1
        mrsWarn!报警标志2 = rsCopy!报警标志2
        mrsWarn!报警标志3 = rsCopy!报警标志3
        mrsWarn!催款下限 = rsCopy!催款下限
        mrsWarn!催款标准 = rsCopy!催款标准
        mrsWarn.Update
        rsCopy.MoveNext
    Loop
    
    tab报警.Tabs.Add , , strName
    tab报警.Tabs(tab报警.Tabs.Count).Selected = True
    
    Bill(bill_记帐报警).Tag = "已修改"
End Sub



Private Function LoadOneCard() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    lvw(lvw_一卡通).ListItems.Clear
    
    strSQL = "Select 编号,名称,结算方式,医院编码,启用 From 一卡通目录 Order by 编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw(lvw_一卡通).ListItems.Add(, "_" & rsTmp!编号, rsTmp!编号)
        ObjItem.SubItems(1) = NVL(rsTmp!名称)
        ObjItem.SubItems(2) = NVL(rsTmp!结算方式)
        ObjItem.SubItems(3) = NVL(rsTmp!医院编码)
        ObjItem.SubItems(4) = IIF(NVL(rsTmp!启用, 0) = 2, "启用:标准一卡通", IIF(NVL(rsTmp!启用, 0) = 1, "启用:仅涉及扣卡", "停用"))
        rsTmp.MoveNext
    Loop
    
    If Not lvw(lvw_一卡通).SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw_一卡通, lvw(lvw_一卡通).SelectedItem)
    End If
    LoadOneCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsRaiseByDate(ByVal strID As String) As Boolean
    '判断该收费项目是否是按日调价
    '返回True-是按天条件
    '返回False-不是按天调价
    'strID='J' -床位项目
    'strID='H' -护理项目
    'strID=数字 -其他指定的项目
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    If strID = "J" Then
        strSQL = "Select ID" & _
              " From 收费价目 " & _
              " Where Nvl(终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And 执行日期 <> Trunc(执行日期, 'dd') And " & _
              " 收费细目id In " & _
              " (Select ID " & _
              " From 收费项目目录 " & _
              " Where 类别 = [1] " & _
              " Union All " & _
              " Select 从项id From 收费从属项目 Where 主项id In (Select ID From 收费项目目录 Where 类别 = [1])) "
    ElseIf strID = "H" Then
            strSQL = "Select ID" & _
              " From 收费价目 " & _
              " Where Nvl(终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And 执行日期 <> Trunc(执行日期, 'dd') And " & _
              " 收费细目id In " & _
              " (Select ID " & _
              " From 收费项目目录 " & _
              " Where 类别 = [1] " & _
              " Union All " & _
              " Select 从项id From 收费从属项目 Where 主项id In (Select ID From 收费项目目录 Where 类别 = [1])) "
    ElseIf Val(strID) <> 0 Then
        strSQL = "Select Id" & _
                " From 收费价目 " & _
                " Where Nvl(终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate " & _
                " And 执行日期<>trunc(执行日期,'dd') And (收费细目id = [2] or 收费细目id in (Select 从项id From 收费从属项目 Where 主项id = [2])) "
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID, Val(strID))
    
    IsRaiseByDate = Not (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    
    lvw(lvw_单据).Tag = "已修改"
    
    If Index = 0 Or Index = 1 Then
        If Index = 0 Then
            Set lst = lvw(lvw_单据).ListItems.Add(, , str姓名)
            lst.Selected = True
            lst.EnsureVisible
        Else
            Set lst = lvw(lvw_单据).SelectedItem
            lst.Text = str姓名
        End If
        lst.SubItems(1) = str单据
        lst.SubItems(2) = lng天数
        lst.SubItems(3) = IIF(bln修改他人 = True, "是", "否")
        lst.SubItems(4) = IIF(Val(dbl金额上限) = 0, "", Format(Val(dbl金额上限), "0.00"))
        lst.Tag = str人员ID
        lst.ListSubItems(1).Tag = lng单据
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
    If Me.Visible Then Call SetParChange(lvw, Index, mrsPar)
    
    Dim itemTemp As MSComctlLib.ListItem
    For Each itemTemp In lvw(Index).ListItems
        If Not itemTemp Is Item Then
            itemTemp.Checked = False
        End If
    Next
End Sub

Private Sub lvw_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim lng原值 As Long, blnValue As Boolean, strValue As String
    If Index = lvw_票据 Then
        lng原值 = Val(Item.SubItems(1))
        ud(ud_号码长度).Max = 20
        '设置最大值时，可能已经更改了列表中的值
        ud(ud_号码长度).value = IIF(lng原值 = 0, 7, lng原值)
        chk(chk_票号控制).value = IIF(Item.SubItems(2) = "√", 1, 0)
        Exit Sub
    ElseIf Index = lvw_一卡通 Then
        cmdOneCard(1).Enabled = Item.Text <> ""
        cmdOneCard(2).Enabled = cmdOneCard(1).Enabled
    End If
    
    If Not Me.Visible Then Exit Sub
    Call SetParChange(lvw, Index, mrsPar, blnValue, strValue)
    
    
End Sub

Private Sub lvw_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = lvw_票据 Then
        If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    ElseIf Index = lvw_单据 Then
        If KeyAscii = vbKeyReturn Then Call cmdOperate_Click(1)
    End If
End Sub

Private Sub Load单据操作()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, str单据 As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.人员ID,B.姓名,A.单据,A.时间限制,A.他人单据,A.金额上限 from 单据操作控制 A,人员表 B where A.人员ID=B.ID"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    lvw(lvw_单据).ListItems.Clear
    Do Until rsTemp.EOF
        Set lst = lvw(lvw_单据).ListItems.Add(, , rsTemp("姓名"))
        
        str单据 = Switch(rsTemp("单据") = 1, "挂号单据", rsTemp("单据") = 2, "收费单", rsTemp("单据") = 3, "划价单", rsTemp("单据") = 4, "门诊记帐", _
                       rsTemp("单据") = 5, "住院记帐", rsTemp("单据") = 6, "预交款", rsTemp("单据") = 7, "结帐单据", rsTemp("单据") = 8, "就诊卡", rsTemp("单据") = 9, "处方")
        lst.SubItems(1) = str单据
        lst.SubItems(2) = rsTemp("时间限制")
        lst.SubItems(3) = IIF(rsTemp("他人单据") = 1, "是", "否")
        lst.SubItems(4) = IIF(IsNull(rsTemp("金额上限")), "", Format(rsTemp("金额上限"), "0.00"))
        lst.Tag = rsTemp("人员ID")
        lst.ListSubItems(1).Tag = rsTemp("单据")
        
        rsTemp.MoveNext
    Loop
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Load病区()
    Dim rs病区 As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.ID,A.名称,A.编码 " & _
               " from  部门性质说明 b,部门表 a " & _
               " where B.服务对象 in(1,2,3) And B.工作性质='护理' and  b.部门ID=a.ID and " & _
               Where撤档时间("A") & " order by 编码"
    Call zlDatabase.OpenRecordset(rs病区, gstrSQL, Me.Caption)
    
    Bill(bill_自动计算).Clear
    Bill(bill_记帐报警).Clear
    
    If rs病区.RecordCount > 0 Then
        mshAutoCalc.Rows = rs病区.RecordCount + 1
        lngRow = 1
        Do Until rs病区.EOF
            Bill(bill_自动计算).AddItem rs病区("编码") & "-" & rs病区("名称")
            Bill(bill_自动计算).ItemData(Bill(bill_自动计算).NewIndex) = rs病区("ID")
            Bill(bill_记帐报警).AddItem rs病区("编码") & "-" & rs病区("名称")
            Bill(bill_记帐报警).ItemData(Bill(bill_自动计算).NewIndex) = rs病区("ID")
            mshAutoCalc.TextMatrix(lngRow, 0) = rs病区("编码") & "-" & rs病区("名称")
            mshAutoCalc.RowData(lngRow) = rs病区("ID")
            lngRow = lngRow + 1
            rs病区.MoveNext
        Loop
        Bill(bill_自动计算).ListIndex = 0
    End If
    Bill(bill_记帐报警).AddItem "*门诊*"
    Bill(bill_记帐报警).ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load费用类型()
'功能：初始化费用类型
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    gstrSQL = "Select 编码,名称 From 费用类型 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
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

Private Function GetDepartments(ByVal str性质 As String, _
    ByVal str服务对象 As String, _
    Optional ByVal bln仅操作员部门 As Boolean = False, _
    Optional ByVal blnCheck站点 As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定性质的部门列表
    '入参:str性质='临床','护理','中药房',...,允许为空
    '     str服务对象:以,分离:如1,3
    '     bln仅操作员部门-操作员的所属部门
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.工作性质||',')>0"
        Else
            strSQL = " And B.工作性质 = [1]"
        End If
    End If
    If bln仅操作员部门 Then strSQL = strSQL & "  And A.id=C.部门ID and C.人员id =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & IIF(bln仅操作员部门, ",部门人员 C", "") & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And Instr(',' || [2]|| ',',',' || B.服务对象 || ',')>0 " & strSQL & _
         IIF(blnCheck站点, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
        " Order by A.编码"
    Set GetDepartments = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str性质, str服务对象, glngUserId)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadOther()
'完成其余的初始化工作
    Dim rsTemp As New ADODB.Recordset
    Dim lngMaxRow As Long, lngRow As Long, lng单位 As Long
    Dim lngTmp As Long, i As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean, strCoding As String
    
    
    '收费特定项目
    On Error GoTo ErrHandle
    gstrSQL = "select a.特定项目 ,c.ID,c.名称  " & _
            " from 收费特定项目 a,收费细目 c " & _
            " where a.收费细目ID =c.id"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("特定项目")
            Case "病历费"
                txtCmd(0).Tag = rsTemp("ID")
                txtCmd(0).Text = rsTemp("名称")
            Case "工本费"
                txtCmd(1).Tag = rsTemp("ID")
                txtCmd(1).Text = rsTemp("名称")
            Case "普通配置费"
                txtCmd(3).Tag = rsTemp("ID")
                txtCmd(3).Text = rsTemp("名称")
            Case "肿瘤配置费"
                txtCmd(4).Tag = rsTemp("ID")
                txtCmd(4).Text = rsTemp("名称")
        End Select
        rsTemp.MoveNext
    Loop
    
    '病区自动计帐程序
    gstrSQL = "select A.病区ID,B.编码,b.名称 as 病区 ,a.收费细目ID,c.名称 as 收费细目 ,a.计算标志,a.启用日期 " & _
            " from 自动计价项目 A,部门表 B,收费细目 C " & _
            " where A.病区ID= B.id and A.收费细目ID =C.id(+) " & _
            " order by b.编码 "
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Bill(bill_自动计算)
        lngRow = 1
        Do Until rsTemp.EOF
            If IsNull(rsTemp("收费细目ID")) Then
                '床位费或护理费
                For lngTmp = 1 To mshAutoCalc.Rows - 1
                    If mshAutoCalc.RowData(lngTmp) = rsTemp("病区ID") Then
                        If rsTemp("计算标志") = 1 Then
                            '床位费
                            mshAutoCalc.TextMatrix(lngTmp, 1) = "√"
                            mshAutoCalc.TextMatrix(lngTmp, 2) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                            mshAutoCalc.TextMatrix(lngTmp, 5) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                        Else
                            '护理费
                            mshAutoCalc.TextMatrix(lngTmp, 3) = "√"
                            mshAutoCalc.TextMatrix(lngTmp, 4) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                            mshAutoCalc.TextMatrix(lngTmp, 6) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                        End If
                    End If
                Next
            Else
                '其它费用
                .Rows = lngRow + 1
                .RowData(lngRow) = rsTemp("病区ID")
                .TextMatrix(lngRow, 0) = rsTemp("编码") & "-" & rsTemp("病区")
                .TextMatrix(lngRow, 1) = rsTemp("收费细目ID")
                .TextMatrix(lngRow, 2) = rsTemp("收费细目")
                .TextMatrix(lngRow, 3) = Switch(rsTemp("计算标志") = 6, "1-按床日", rsTemp("计算标志") = 8, "2-计算一次", True, "0-按收治日")
                .TextMatrix(lngRow, 4) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                lngRow = lngRow + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    
    '记帐报警类别
    gstrSQL = "Select 编码,类别 From 收费类别 Order by 编码"
    Set mrs类别 = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrs类别, gstrSQL, Me.Caption)
    
    lst类别.Clear
    lst类别.AddItem "所有类别"
    Do While Not mrs类别.EOF
        lst类别.AddItem mrs类别!类别
        lst类别.ItemData(lst类别.NewIndex) = Asc(mrs类别!编码)
        mrs类别.MoveNext
    Loop
    
    '病区记帐报警线
    Set mrsWarn = New ADODB.Recordset
    mrsWarn.Fields.Append "病区ID", adBigInt, , adFldIsNullable
    mrsWarn.Fields.Append "病区码", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "病区名", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "适用病人", adVarChar, 100
    mrsWarn.Fields.Append "报警方法", adSmallInt
    mrsWarn.Fields.Append "报警值", adCurrency
    mrsWarn.Fields.Append "报警标志1", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "报警标志2", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "报警标志3", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "催款下限", adCurrency
    mrsWarn.Fields.Append "催款标准", adCurrency
    
    mrsWarn.CursorLocation = adUseClient
    mrsWarn.LockType = adLockOptimistic
    mrsWarn.CursorType = adOpenStatic
    mrsWarn.Open
    
    gstrSQL = "" & _
    "   Select a.病区ID,B.编码,b.名称 as 病区,a.适用病人,nvl(a.报警方法,1) as 报警方法, " & _
    "               a.报警值,a.报警标志1,a.报警标志2,a.报警标志3,A.催款下限,a.催款标准 " & _
    "   From 记帐报警线 a,部门表 b " & _
    "   Where a.病区ID= b.id(+)  " & _
    "   Order by Decode(a.适用病人,'普通病人',1,'医保病人',2,3),a.适用病人,B.编码 Desc"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    strCoding = ",普通病人" '至少有一个普通病人
    Do Until rsTemp.EOF
        mrsWarn.AddNew
        mrsWarn!病区ID = rsTemp!病区ID
        mrsWarn!病区码 = rsTemp!编码
        mrsWarn!病区名 = rsTemp!病区
        mrsWarn!适用病人 = rsTemp!适用病人
        mrsWarn!报警方法 = rsTemp!报警方法
        mrsWarn!报警值 = rsTemp!报警值
        mrsWarn!报警标志1 = rsTemp!报警标志1
        mrsWarn!报警标志2 = rsTemp!报警标志2
        mrsWarn!报警标志3 = rsTemp!报警标志3
        mrsWarn!催款下限 = Val(NVL(rsTemp!催款下限))
        mrsWarn!催款标准 = Val(NVL(rsTemp!催款标准))
        mrsWarn.Update
        
        If InStr(strCoding & ",", "," & rsTemp!适用病人 & ",") = 0 Then
            strCoding = strCoding & "," & rsTemp!适用病人
        End If
        rsTemp.MoveNext
    Loop
    strCoding = Mid(strCoding, 2)
    tab报警.Tabs.Clear
    For i = 0 To UBound(Split(strCoding, ","))
        tab报警.Tabs.Add , , Split(strCoding, ",")(i)
    Next
    tab报警.Tabs(1).Selected = True '之前不会激活Click事件,人为激活
   
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    '病历费和工本费限定为定价项目
    strSQL = "select id,编码,名称,计算单位,说明 from 收费项目目录 where 类别='Z' and nvl(是否变价,0)=0"

    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        If IsNumeric(txtCmd(Index).Tag) = False Then txtCmd(Index).Tag = 0
        strSQL = frmSelCurr.ShowCurrSel(Me, rsTmp, "id,0,0,2;编号,1000,0,2;名称,1800,0,1;单位,800,0,2;说明,2300,0,2", -1, "定价项目选择", , CStr(txtCmd(Index).Tag), 0, 3)
        If strSQL <> "" Then
            txtCmd(Index).Tag = CLng(Split(strSQL, ";")(0))
            txtCmd(Index).Text = Trim(Split(strSQL, ";")(2))
            txtCmd(Index).SetFocus
            
            fra特定收费项目.Tag = "已修改"
        End If
    Else
        MsgBox "无任何项目可用！", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
            If mshAutoCalc.Visible Or Bill(bill_自动计算).Visible Then
                If lblLocate(txt_Dept).Tag = "mshAutoCalc" Or lblLocate(txt_Dept).Tag = "" Then
                    Call LocateDept(strFind, mshAutoCalc)
                Else
                    Call LocateDept(strFind, Bill(bill_自动计算))
                End If
            ElseIf Bill(bill_记帐报警).Visible Then
                Call LocateDept(strFind, Bill(bill_记帐报警))
                
            End If
        End Select
    End If
End Sub

Private Sub LocateDept(ByVal strFind As String, ByRef objBill As Object)
'功能：查找科室
    Dim i As Long
    Dim strCode As String, strName As String
    
    With objBill
        For i = mlngPreFind To .Rows - 1
            '0列为病区列
            strCode = Split(.TextMatrix(i, 0), "-")(0)
            strName = Split(.TextMatrix(i, 0), "-")(1)
            
            If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                objBill.SetFocus
                .Row = i: .Col = 1
                .TopRow = i
                Exit For
            End If
        Next
        If i < .Rows Then
            mlngPreFind = i + 1
        Else
            If mlngPreFind = 1 Then
                MsgBox "没有找到匹配的科室，请检查输入的内容。", vbInformation, Me.Caption
                txtLocate(txt_Dept).SetFocus
            Else
                MsgBox "全部找完了，后面没有了。", vbInformation, Me.Caption
                mlngPreFind = 1
            End If
        End If
    End With
End Sub

Private Sub ud_Change(Index As Integer)
    Dim strValue As String
    If Not Me.Visible Then Exit Sub
    '动态改变票号长度
    If Index = ud_号码长度 Then
        lvw(lvw_票据).SelectedItem.SubItems(1) = ud(ud_号码长度).value
        strValue = GetBillLenSet
        Call SetParChange(lvw, lvw_票据, mrsPar, True, strValue)
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    If Not Me.Visible Then Exit Sub
    
    Select Case Index
    Case ud_挂号单, ud_急诊挂号单
        blnValue = True
        strValue = txtUD(ud_挂号单).Text & txtUD(ud_急诊挂号单).Text
    Case ud_号码长度
        strValue = GetBillLenSet
        Call SetParChange(lvw, lvw_票据, mrsPar, True, strValue)
        Exit Sub
    Case ud_票据张数
        strValue = IIF(chk(chk_票据剩余N张提醒操作员).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(Index).Text)
        Call SetParChange(chk, chk_票据剩余N张提醒操作员, mrsPar, True, strValue)
        txtUD(Index).ForeColor = chk(chk_票据剩余N张提醒操作员).ForeColor
        Exit Sub
    Case ud_收费_票据张数
        strValue = IIF(chk(chk_收费_票据剩余X张开始提醒).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(Index).Text)
        Call SetParChange(chk, chk_收费_票据剩余X张开始提醒, mrsPar, True, strValue)
        txtUD(Index).ForeColor = chk(chk_收费_票据剩余X张开始提醒).ForeColor
        Exit Sub
    Case ud_补结算_票据张数
        strValue = IIF(chk(chk_补结算_票据剩余X张开始提醒).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(Index).Text)
        Call SetParChange(chk, chk_补结算_票据剩余X张开始提醒, mrsPar, True, strValue)
        txtUD(Index).ForeColor = chk(chk_补结算_票据剩余X张开始提醒).ForeColor
        Exit Sub
    Case ud_补结算_票据张数
        strValue = Val(txtUD(Index).Text): blnValue = True
    End Select
    
    Call SetParChange(txtUD, Index, mrsPar, blnValue, strValue)
 
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case Index
        Case txt_挂号_预约失效次数
            With tbPage(Pg_挂号业务)
                .Item(3).Selected = True
            End With
        Case Else
            Call zlCommFun.PressKey(vbKeyTab)
        End Select
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim strValue As String
    If Val(txt(txt_单笔最大金额).Text) = 0 Then txt(txt_单笔最大金额).Text = ""
    If Index = txt_免密支付 Then
        If Val(txt(Index).Text) = 0 Then Exit Sub
        strValue = -1 * Val(txt(Index).Text)
        strValue = strValue & "|" & IIF(optBrushCard(11).value, 1, IIF(optBrushCard(12).value, 2, 0))
        Call SetParChange(optBrushCard, 3, mrsPar, True, strValue)
    End If
End Sub

Private Sub txtDateInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtDateInput
            If Not IsDate(.Text) Then
                If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                    MsgBox "请输入正确的日期格式(yyyy-mm-dd或者yyyymmdd)。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                End If
            End If
            mshAutoCalc.TextMatrix(mintCurRow, mintCurCol) = .Text
            mshAutoCalc.Tag = "已修改"
            .Visible = False
        End With
    End If
End Sub

Private Sub txtDateInput_LostFocus()
    txtDateInput.Text = ""
    txtDateInput.Visible = False
End Sub

Private Sub mshAutoCalc_Click()
    With Me.mshAutoCalc
        If .Row > 0 And (.Col = 2 Or .Col = 4) And .TextMatrix(.Row, IIF(.Col = 2, 1, 3)) <> "" Then
            mintCurRow = .Row
            mintCurCol = .Col
            txtDateInput.Move (.Left + .CellLeft - 10), (.Top + .CellTop - 10), .CellWidth, .CellHeight
            If .TextMatrix(.Row, .Col) <> "" Then
                txtDateInput.Text = .TextMatrix(.Row, .Col)
            End If
            txtDateInput.Visible = True
            txtDateInput.SetFocus
        End If
    End With
End Sub

Private Sub mshAutoCalc_DblClick()
    With mshAutoCalc
        If .MouseRow > 0 And .MouseCol > 0 And .RowData(.MouseRow) <> 0 Then
            If .Col = 1 Or .Col = 3 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "床位类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "护理类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "√", "")
                .TextMatrix(.Row, IIF(.Col = 1, 2, 4)) = IIF(.Text = "", "", IIF(.TextMatrix(.Row, IIF(.Col = 1, 5, 6)) = "", Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd"), .TextMatrix(.Row, IIF(.Col = 1, 5, 6))))
                
                .Tag = "已修改"
            End If
        End If
    End With
End Sub

Private Sub mshAutoCalc_KeyPress(KeyAscii As Integer)
    With mshAutoCalc
        If KeyAscii = vbKeyReturn Then
            If .Col = 1 Then
                .Col = 2
            ElseIf .Col = 4 Then
                If .Row = .Rows - 1 Then
                    Bill(bill_自动计算).SetFocus
                Else
                    .Row = .Row + 1
                    .Col = 1
                    If .Row - .TopRow > 8 Then .TopRow = .Row - 8
                End If
            End If
        ElseIf KeyAscii = Asc(" ") Then
            If .Row > 0 And (.Col = 1 Or .Col = 3) And .RowData(.Row) <> 0 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "床位类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "护理类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "√", "")
                .TextMatrix(.Row, IIF(.Col = 1, 2, 4)) = IIF(.Text = "", "", IIF(.TextMatrix(.Row, IIF(.Col = 1, 5, 6)) = "", Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd"), .TextMatrix(.Row, IIF(.Col = 1, 5, 6))))
                
            End If
        Else
            If .Row > 0 And (.Col = 2 Or .Col = 4) And .TextMatrix(.Row, IIF(.Col = 2, 1, 3)) <> "" Then
                mintCurRow = .Row
                mintCurCol = .Col
                txtDateInput.Move (.Left + .CellLeft - 10), (.Top + .CellTop - 10), .CellWidth, .CellHeight
                If .TextMatrix(.Row, .Col) <> "" Then
                    txtDateInput.Text = .TextMatrix(.Row, .Col)
                End If
                txtDateInput.Visible = True
                txtDateInput.SetFocus
            End If
        End If
        .Tag = "已修改"
    End With
End Sub

Private Sub txtCmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txtCmd(Index).Tag = ""
        txtCmd(Index).Text = ""
        
        fra特定收费项目.Tag = "已修改"
    End If
End Sub

Private Sub txtCmd_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = Asc("*") Then
        Call cmdSelect_Click(Index)
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case cbo_记帐操作_记帐后发药, cbo_一卡通_记帐票据格式
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case chk_存在预约挂号单禁止删除安排
        With tbPage(Pg_挂号业务)
            .Item(1).Selected = True
        End With
    Case chk_挂号_非严格控制为发卡
        With tbPage(Pg_挂号业务)
            .Item(2).Selected = True
        End With
    Case chk_借款_借出打印
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
        
End Sub


Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim i As Long
    If Not Me.Visible Then Exit Sub
    Select Case Index
    Case lst_医保病人, lst_公费病人
        blnValue = True
        strValue = Replace(Replace(GetTextFromList(lst(Index)), "'", ""), ",", "|")
    Case lst_补结算_结算方式
        blnValue = True
        strValue = Replace(Replace(GetTextFromList(lst(Index)), "'", ""), ",", "|")
    Case lst_刷卡密码
        blnValue = True
        With lst(lst_刷卡密码)
            For i = 0 To .ListCount - 1
                strValue = strValue & IIF(.Selected(i), 1, 0)
            Next
        End With
    Case lst_结帐_自费费用类别
        strValue = ""
        With lst(lst_结帐_自费费用类别)
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    strValue = strValue & "," & Chr(.ItemData(i))
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        blnValue = True
    Case lst_脱机医保结算方式
        strValue = ""
        With lst(lst_脱机医保结算方式)
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    strValue = strValue & "|" & zlCommFun.GetNeedName(.List(i), "-")
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        blnValue = True
    End Select
    Call SetParChange(lst, Index, mrsPar, blnValue, strValue)
End Sub

Private Sub Load脱机医保结算方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载自费费用类别
    '编制:刘兴洪
    '日期:2015-06-24 15:34:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    
    strSQL = "Select 编码,名称 From 结算方式 Where 性质 = 2 And Nvl(应收款,0)=0 And Nvl(应付款,0)=0 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With lst(lst_脱机医保结算方式)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!编码 & "-" & rsTemp!名称
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case lst_补结算_结算方式
         If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case lst_刷卡密码
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim intBillRull As Boolean
    If mblnNotChange Then Exit Sub
    
    If Not Me.Visible Then Exit Sub

    Select Case Index
    Case chk_首先输入收费类别
        If chk(Index).value = 1 Then
            If chk(chk_收费项目首位当类别简码).value = 1 Then chk(chk_收费项目首位当类别简码).value = 0
        End If
    Case chk_收费项目首位当类别简码
        If chk(Index).value = 1 Then
            If chk(chk_首先输入收费类别).value = 1 Then chk(chk_首先输入收费类别).value = 0
        End If
    Case chk_票号控制
        lvw(lvw_票据).SelectedItem.SubItems(2) = IIF(chk(Index).value = 1, "√", "")
        
        strValue = GetBillCtlSet: blnValue = True
        Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
        
    Case chk_病人姓名, chk_刷就诊卡, chk_挂号单号, chk_病人ID
        strValue = chk(chk_病人姓名).value & chk(chk_刷就诊卡).value & chk(chk_挂号单号).value & chk(chk_病人ID).value
        blnValue = True
        Call SetParChange(chk, chk_病人姓名, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_刷就诊卡, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_挂号单号, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_病人ID, mrsPar, blnValue, strValue)
    Case chk_姓名模糊查找
        If chk(chk_姓名模糊查找).value = 1 Then
            txt(txt_姓名查找天数).Enabled = True
        Else
            txt(txt_姓名查找天数).Enabled = False
        End If
    Case chk_姓名允许模糊查找
        txt(txt_姓名模糊查找天数).Enabled = chk(Index).value = 1
    
    Case chk_票据剩余N张提醒操作员
        strValue = IIF(chk(Index).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(ud_票据张数).Text)
        txtUD(ud_票据张数).Enabled = chk(Index).value = 1
        ud(ud_票据张数).Enabled = chk(Index).value = 1
        
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        txtUD(ud_票据张数).ForeColor = chk(Index).ForeColor
        Exit Sub
    Case chk_挂号_自动刷新挂号安排
        txt(txt_挂号_刷新时间).Enabled = chk(Index).value = 1
        If chk(Index).value <> 1 Then txt(txt_挂号_刷新时间).Text = 0
        strValue = Val(txt(txt_挂号_刷新时间).Text)
        Call SetParChange(txt, txt_挂号_刷新时间, mrsPar, True, strValue)
        chk(Index).ForeColor = txt(txt_挂号_刷新时间).ForeColor
        Exit Sub
    Case chk_挂号_姓名模糊查找
        txt(txt_挂号_姓名查找天数).Enabled = chk(Index).value = 1
    Case chk_挂号_病人挂号科室限制
        txt(txt_挂号_病人挂号科室限制).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_挂号_病人挂号科室限制).Text
        Else
            strValue = "0"
        End If
    Case chk_挂号_病人同科限约N个号
        txt(txt_挂号_病人同科限约N个号).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_挂号_病人同科限约N个号).Text
        Else
            strValue = "0"
        End If
    Case chk_专家号挂号限制
        txt(txt_专家号挂号限制).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_专家号挂号限制).Text
        Else
            txt(txt_专家号挂号限制).Text = ""
            strValue = "0"
        End If
    Case chk_专家号预约限制
        txt(txt_专家号预约限制).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_专家号预约限制).Text
        Else
            txt(txt_专家号预约限制).Text = ""
            strValue = "0"
        End If
    Case chk_挂号_病人预约科室数
        txt(txt_挂号_病人预约科室数).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_挂号_病人预约科室数).Text
        Else
            strValue = "0"
        End If
    Case chk_挂号_病人同科限挂N个号
        txt(txt_挂号_病人同科限挂N个号).Enabled = chk(Index).value = 1
        chk(chk_挂号_病人同科限挂N个号_急诊).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_挂号_病人同科限挂N个号).Text & "|" & IIF(chk(chk_挂号_病人同科限挂N个号_急诊).value = 1, "1", "0")
        Else
            strValue = "0|0"
        End If
    Case chk_挂号_病人同科限挂N个号_急诊
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_挂号_病人同科限挂N个号).Text & "|1"
        Else
            strValue = "0|0"
        End If
    Case chk_挂号包含科室安排, chk_预约包含科室安排
        blnValue = True
        strValue = chk(chk_挂号包含科室安排).value & "|" & chk(chk_预约包含科室安排).value
    Case chk_收费_未挂号自动加收挂号费
        blnValue = True
        strValue = ""
        
        If chk(Index).value = 1 Then
            If txt(txt_收费_自助加收挂号费).Text = "" And Me.Visible Then
                Call cmdAddedItem_Click: Exit Sub
            End If
            strValue = cmdAddedItem.Tag & ";" & txt(txt_收费_自助加收挂号费).Text
        Else
            mblnNotChange = True
            txt(txt_收费_自助加收挂号费).Text = "": cmdAddedItem.Tag = ""
            mblnNotChange = False
        End If
        Call SetParChange(txt, txt_收费_自助加收挂号费, mrsPar, blnValue, strValue)
        Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
        Exit Sub
    Case chk_收费_自动组合单据
        blnValue = True
        strValue = "0"
        If chk(Index).value = 1 Then
            strValue = cbo(cbo_收费_自动组合单据).ListIndex + 1
        End If
        cbo(cbo_收费_自动组合单据).Enabled = chk(Index).value = 1
        Call SetParChange(chk, chk_收费_自动组合单据, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_收费_自动组合单据, mrsPar, blnValue, strValue)
        Exit Sub
    Case chk_收费_票据剩余X张开始提醒
        strValue = IIF(chk(Index).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(ud_收费_票据张数).Text)
        txtUD(ud_收费_票据张数).Enabled = chk(Index).value = 1
        ud(ud_收费_票据张数).Enabled = chk(Index).value = 1
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        txtUD(ud_收费_票据张数).ForeColor = chk(Index).ForeColor
        Exit Sub
        
    Case chk_收费_票据生成方式
        intBillRull = IIF(cbo(cbo_收费_票据分配规则).ListIndex < 0, 0, cbo(cbo_收费_票据分配规则).ListIndex)
        If intBillRull <> 0 Then Exit Sub
        strValue = CStr(IIF(optBillMode(1).value, 1, 0) + Val(chk(chk_收费_票据生成方式).value) * 10)
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        Call SetParChange(optBillMode, 0, mrsPar, True, strValue)
        Exit Sub
    Case chk_收费_多张单据收费分别打印
        chk(chk_收费_体检病人按单据分别打印).Enabled = (chk(Index).value = vbChecked)
    Case chk_记帐_门诊姓名模糊查找
        chk(chk_记帐_只查找合约单位病人).Enabled = chk(chk_记帐_门诊姓名模糊查找).value = 1
        txt(txt_记帐_门诊姓名模糊查找天数).Enabled = chk(Index).value = 1
    Case chk_划价_门诊姓名模糊查找
        txt(txt_划价_门诊姓名模糊查找天数).Enabled = chk(Index).value = 1
    Case chk_收费_门诊姓名模糊查找
        txt(txt_收费_门诊姓名模糊查找天数).Enabled = chk(Index).value = 1
    Case chk_收费_搜寻划价单据
        txt(txt_收费_搜寻划价单据天数).Enabled = chk(Index).value = 1
    Case chk_补结算_姓名模糊查找
        txt(txt_补结算_姓名模糊查找天数).Enabled = chk(Index).value = 1
        strValue = IIF(chk(chk_补结算_姓名模糊查找).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txt(txt_补结算_姓名模糊查找天数).Text)
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        Call SetParChange(txt, txt_补结算_姓名模糊查找天数, mrsPar, True, strValue)
        Exit Sub
    Case chk_补结算_票据剩余X张开始提醒
        strValue = IIF(chk(Index).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(ud_补结算_票据张数).Text)
        txtUD(ud_补结算_票据张数).Enabled = chk(Index).value = 1
        ud(ud_补结算_票据张数).Enabled = chk(Index).value = 1
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        txtUD(ud_补结算_票据张数).ForeColor = chk(Index).ForeColor
        Exit Sub
    Case chk_记帐报警包含划价费用
        strValue = IIF(chk(Index).value = 1, 1, 0)
        If Val(strValue) = 1 Then
            lblInExseCharge.Enabled = True
            optInExseCharge(0).Enabled = True
            optInExseCharge(1).Enabled = True
        Else
            lblInExseCharge.Enabled = False
            optInExseCharge(0).Enabled = False
            optInExseCharge(1).Enabled = False
        End If
    Case chk_收费_按病人不打票据不分次数
        With vsBillFormat(vsGrid_收费票据格式)
            .ColHidden(.ColIndex("按病人补打票据格式")) = chk(Index).value <> 1
        End With
    Case chk_分诊_分诊台签到开始排队
        Call LoadTriageQueuingDep
        Call SetTriageQueuingEnalbe(chk(Index).value)
    Case chk_挂号_病人同一号源限挂N个号
        txt(txt_挂号_病人同一号源限挂N个号).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_挂号_病人同一号源限挂N个号).Text
        Else
            strValue = "0"
        End If
    End Select
    Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
End Sub

Public Function IsDrugOrStuff(ByVal strID As String) As Boolean
    '判断是否为药品类别
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select id From 收费细目 Where 类别 In('4','5','6','7') and id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
    
    IsDrugOrStuff = rs.RecordCount > 0
    rs.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmd公费费用类型_Click(Index As Integer)
    Dim i As Long
    
    With lst(lst_公费病人)
        For i = 0 To .ListCount - 1
            .Selected(i) = Index = 0    '将触发lst_ItemCheck事件
        Next
    End With
End Sub

Private Sub cmd医保费用类型_Click(Index As Integer)
    Dim i As Long
    
    With lst(lst_医保病人)
        For i = 0 To .ListCount - 1
            .Selected(i) = Index = 0    '将触发lst_ItemCheck事件
        Next
    End With
End Sub

Private Sub vsBillFormat_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = vsGrid_预交票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("预交打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
               
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    If Index = vsGrid_预交红票格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("退款打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
               
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    If Index = vsGrid_收费票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("收费打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("收费票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case .ColIndex("按病人补打票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 2), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_退费票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("收费打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("收费票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_补结算票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("收费打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_补结算退费票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("收费打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_结帐票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("结帐后打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_结帐红票格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("作废后打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
       
    If Index = vsGrid_发卡预交票据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("预交打印方式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
               
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_医疗卡收据格式 Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("票据格式")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
End Sub

Private Sub vsBillFormat_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
  If Index = vsGrid_预交票据格式 Then
    zl_vsGrid_Para_Save p预交款管理, vsBillFormat(Index), Me.Name, "预交发票打印方式", False, False
    Exit Sub
  End If
  If Index = vsGrid_预交红票格式 Then
    zl_vsGrid_Para_Save p预交款管理, vsBillFormat(Index), Me.Name, "预交退款打印方式", False, False
    Exit Sub
  End If
  If Index = vsGrid_收费票据格式 Then
    zl_vsGrid_Para_Save p门诊收费管理, vsBillFormat(Index), Me.Name, "收费票据格式", False, False
    Exit Sub
  End If
  If Index = vsGrid_退费票据格式 Then
    zl_vsGrid_Para_Save p门诊收费管理, vsBillFormat(Index), Me.Name, "退费票据格式", False, False
    Exit Sub
  End If
  If Index = vsGrid_补结算票据格式 Then
    zl_vsGrid_Para_Save p门诊补结算, vsBillFormat(Index), Me.Name, "补结算票据格式", False, False
    Exit Sub
  End If
  If Index = vsGrid_补结算退费票据格式 Then
    zl_vsGrid_Para_Save p门诊补结算, vsBillFormat(Index), Me.Name, "补结算退费票据格式", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_结帐票据格式 Then
    zl_vsGrid_Para_Save p病人结帐管理, vsBillFormat(Index), Me.Name, "结帐票据格式", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_结帐红票格式 Then
    zl_vsGrid_Para_Save p病人结帐管理, vsBillFormat(Index), Me.Name, "结帐红票格式", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_发卡预交票据格式 Then
    zl_vsGrid_Para_Save p医疗卡管理, vsBillFormat(Index), Me.Name, "发卡预交发票格式", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_医疗卡收据格式 Then
    zl_vsGrid_Para_Save p医疗卡管理, vsBillFormat(Index), Me.Name, "医疗卡收据格式", False, False
    Exit Sub
  End If
  
End Sub

Private Sub vsBillFormat_KeyPress(Index As Integer, KeyAscii As Integer)
    With vsBillFormat(Index)
        Select Case Index
        Case vsGrid_预交票据格式, vsGrid_收费票据格式, vsGrid_退费票据格式, vsGrid_补结算票据格式, vsGrid_补结算退费票据格式, _
            vsGrid_结帐票据格式, vsGrid_预交红票格式, vsGrid_结帐红票格式, vsGrid_发卡预交票据格式, vsGrid_医疗卡收据格式
           If KeyAscii <> vbKeyReturn Then Exit Sub
            KeyAscii = 0
            If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                If Index = vsGrid_收费票据格式 _
                    Or Index = vsGrid_补结算票据格式 _
                    Or Index = vsGrid_退费票据格式 _
                    Or vsGrid_补结算退费票据格式 Then
                    zlCommFun.PressKey vbKeyTab
                Else
                    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
                End If
               Exit Sub
            End If
            zlVsMoveGridCell vsBillFormat(Index), 1, .Cols - 1
        Case Else
        End Select
    End With
End Sub

Private Sub vsBillFormat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = vsGrid_预交票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("预交打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = vsGrid_预交红票格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("退款打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_收费票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("收费打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("收费票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("按病人补打票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = vsGrid_退费票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("收费打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("收费票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    

    If Index = vsGrid_补结算票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("收费打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_补结算退费票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("收费打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_结帐票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("结帐后打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_结帐红票格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("作废后打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_发卡预交票据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("预交打印方式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_医疗卡收据格式 Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("票据格式") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
End Sub

Private Sub Load结帐红票格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结帐红票格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_结帐红票格式
    
    rsPara.Filter = "模块=" & p病人结帐管理 & " And 参数名='作废发票打印方式'"
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "作废发票打印方式", p病人结帐管理, "", vsBillFormat(intIndex).ColIndex("作废后打印方式"))
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("作废后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("作废后打印方式")) = "0-不打印票据"
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(NVL(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("作废后打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next

            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p病人结帐管理, vsBillFormat(intIndex), Me.Name, "结帐红票格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load发卡预交票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载发卡预交票据格式
    '编制:李南春
    '日期:2016/9/23 15:35:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_发卡预交票据格式
    
    rsPara.Filter = "模块=" & p医疗卡管理 & " And 参数名='预交发票格式'"     '预交发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "预交发票格式", p医疗卡管理, "", vsBillFormat(intIndex).ColIndex("票据格式"))
    
    rsPara.Filter = "模块=" & p医疗卡管理 & " And 参数名='预交发票打印方式'"    '预交发票打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "预交发票打印方式", p医疗卡管理, "", vsBillFormat(intIndex).ColIndex("预交打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("预交打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
 
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
  '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat(intIndex)
        .Rows = 3
        .TextMatrix(1, 0) = "门诊预交"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "住院预交"
        .Cell(flexcpData, 2, 0) = 2
        
        .ColData(.ColIndex("票据格式")) = "0"
        .ColData(.ColIndex("预交打印方式")) = "0"
        
        .ColData(.ColIndex("票据格式")) = 1 ' IIF(intType = 5, 0, 1)
        .ColData(.ColIndex("预交打印方式")) = 1 'IIF(intType1 = 5, 0, 1)
        .Editable = flexEDKbdMouse
    End With
    
    
    With vsBillFormat(intIndex)
        .Clear 1: .Rows = 3
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, .ColIndex("预交打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("预交打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
        Next
    End With
    zl_vsGrid_Para_Restore p医疗卡管理, vsBillFormat(intIndex), Me.Name, "发卡预交发票格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub Load预交红票格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载预交红票格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_预交红票格式
    
    rsPara.Filter = "模块=" & p预交款管理 & " And 参数名='退款发票格式'"    '退款发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "退款发票格式", p预交款管理, "", vsBillFormat(intIndex).ColIndex("票据格式"))
    
    rsPara.Filter = "模块=" & p预交款管理 & " And 参数名='预交退款打印方式'"    '预交退款打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "预交退款打印方式", p预交款管理, "", vsBillFormat(intIndex).ColIndex("退款打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1103_1"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("退款打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
 
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat(intIndex)
        .TextMatrix(1, 0) = "门诊预交"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "住院预交"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("票据格式")) = "0"
        .ColData(.ColIndex("退款打印方式")) = "0"
        '.ForeColor = &H80000008:  .ForeColorFixed = &H80000008
'        Select Case intType
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("票据格式")) = 1 ' IIF(intType = 5, 0, 1)
'        End Select
'        Select Case intType1
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("退款打印方式")) = 1 'IIF(intType1 = 5, 0, 1)
'        End Select
'        If (Val(.ColData(.ColIndex("票据格式"))) = 1 Or _
'            Val(.ColData(.ColIndex("预交打印方式"))) = 1) Then
'            .Editable = flexEDKbdMouse
'        Else
            .Editable = flexEDKbdMouse
'        End If
    End With
    
    With vsBillFormat(intIndex)
        .Clear 1: .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("退款打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("退款打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
        Next
    End With
    zl_vsGrid_Para_Restore p预交款管理, vsBillFormat(intIndex), Me.Name, "预交退款打印方式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsDepositSort_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer, intIndex As Integer
    With vsDepositSort
        Select Case Col
            Case 2
                If Val(.TextMatrix(Row, Col)) = 0 Then
                    .TextMatrix(Row, Col) = -1
                Else
                    .TextMatrix(Row, 3) = 0
                    .TextMatrix(Row, 4) = 0
                End If
            Case 3
                If Val(.TextMatrix(Row, Col)) = 0 Then
                    .TextMatrix(Row, Col) = -1
                Else
                    .TextMatrix(Row, 2) = 0
                    .TextMatrix(Row, 4) = 0
                End If
            Case 4
                If Val(.TextMatrix(Row, Col)) = 0 Then
                    .TextMatrix(Row, Col) = -1
                Else
                    .TextMatrix(Row, 2) = 0
                    .TextMatrix(Row, 3) = 0
                End If
        End Select
        strValue = "1|"
        For i = 2 To 4
            If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
            If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
            If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
            If i <> 4 Then
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
            Else
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
            End If
        Next i
    End With
    Call SetParChange(optOrder, 1, mrsPar, True, strValue)
End Sub

Private Sub vsDepositSort_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or Col = 1 Then Cancel = True
End Sub

Private Sub vsfDelFeeDefaultType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer
    
    With vsfDelFeeDefaultType
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, .ColIndex("缺省退现")))) = 1 Then
                strValue = strValue & ";" & .TextMatrix(i, .ColIndex("收款方式"))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    
    Call SetParChange(vsfDelFeeDefaultType, 0, mrsPar, True, strValue)
End Sub

Private Sub vsfDelFeeDefaultType_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save p门诊收费管理, vsfDelFeeDefaultType, Me.Name, "退费缺省方式", False, False
End Sub

Private Sub vsfDelFeeDefaultType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    KeyAscii = 0
    With vsfDelFeeDefaultType
        If .Row = .Rows - 1 And .Col = .Cols - 1 Then
            zlCommFun.PressKey vbKeyTab
        Else
            zlVsMoveGridCell vsfDelFeeDefaultType, 1, .Cols - 1
        End If
    End With
End Sub

Private Sub vsfDelFeeDefaultType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsfDelFeeDefaultType, 0, mrsPar)
End Sub

Private Sub vsfDrugStore_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore(Index)
        Select Case Col
        Case .ColIndex("缺省"), .ColIndex("窗口")
            Cancel = Val(.Cell(flexcpData, Row, Col)) = 1
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick(Index As Integer)
    Dim strTmp As String, i As Long
    With vsfDrugStore(Index)
        If Not (.Row > 0 And .Col = 1) Then Exit Sub
        If .Cell(flexcpData, .Row, .ColIndex("缺省")) = 1 Then Exit Sub
'        .TextMatrix(.Row, .Col) = IIF(Val(.TextMatrix(.Row, .Col)) = 0, 1, 0)
        Call SetDrugStockDeFault(.Row, Index)
    End With
End Sub

Private Sub SetDrugStockDeFault(ByVal lngRow As Long, ByVal Index As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置药房的缺省值
    '入参:lngRow-指定行
    '编制:刘兴洪
    '日期:2009-09-02 14:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng缺省 As Long, strType As String
    With vsfDrugStore(Index)
        lng缺省 = Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省"))))
        If lng缺省 = 1 Then
            For i = 1 To .Rows - 1
                If i <> lngRow Then
                    .TextMatrix(i, .ColIndex("缺省")) = 0
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell(Index As Integer)
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore(Index)
        If .Row > 0 Then
            If .Col = .ColIndex("窗口") Then
                Set rsTmp = Read发药窗口(.RowData(.Row))
                strList = "自动分配|" & .BuildComboList(rsTmp, "名称")
                .ColComboList(.Col) = strList
            Else
                .ColComboList(.Col) = ""
            End If
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Function Read发药窗口(lngId As Long) As ADODB.Recordset
'功能：获取指定药房的发药窗口
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select 名称 From 发药窗口 Where 药房ID=[1] Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngId)
    Set Read发药窗口 = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfDrugStore_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        With vsfDrugStore(Index)
            If .MouseCol = .ColIndex("缺省") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("窗口") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = 1 Then
        With vsfDrugStore(Index)
            If .MouseCol = .ColIndex("缺省") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("窗口") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = 2 Then
        With vsfDrugStore(Index)
            If .MouseCol = .ColIndex("缺省") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("窗口") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
End Sub
 

Private Sub vsfDrugStore_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, strWindow As String
    Dim blnHave As Boolean
    If Index = 0 Then
        With vsfDrugStore(Index)
            Select Case Col
            Case .ColIndex("缺省")
                blnHave = False
                For i = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(i, .ColIndex("缺省")))) = 1 Then blnHave = True
                Next i
                If blnHave = True Then
                    If Abs(Val(.TextMatrix(Row, .ColIndex("缺省")))) = 1 Then
                        Call SetDrugStockDeFault(Row, Index)
                        Call SetParChange(vsfDrugStore, Index, mrsPar, True, .RowData(Row), CStr(Col))
                    End If
                Else
                    Call SetParChange(vsfDrugStore, Index, mrsPar, True, "", CStr(Col))
                End If
            Case .ColIndex("窗口")
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("窗口")) <> "自动分配" And .TextMatrix(i, .ColIndex("窗口")) <> "" Then
                        strWindow = strWindow & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                    End If
                Next i
                If strWindow <> "" Then strWindow = Mid(strWindow, 2)
                Call SetParChange(vsfDrugStore, Index, mrsPar, True, strWindow, CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = 1 Then
        With vsfDrugStore(Index)
            Select Case Col
            Case .ColIndex("缺省")
                If Abs(Val(.TextMatrix(Row, .ColIndex("缺省")))) = 1 Then
                    Call SetDrugStockDeFault(Row, Index)
                    Call SetParChange(vsfDrugStore, Index, mrsPar, True, .RowData(Row), CStr(Col))
                End If
            Case .ColIndex("窗口")
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("窗口")) <> "自动分配" And .TextMatrix(i, .ColIndex("窗口")) <> "" Then
                        strWindow = strWindow & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                    End If
                Next i
                If strWindow <> "" Then strWindow = Mid(strWindow, 2)
                Call SetParChange(vsfDrugStore, Index, mrsPar, True, strWindow, CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = 2 Then
        With vsfDrugStore(Index)
            Select Case Col
            Case .ColIndex("缺省")
                If Abs(Val(.TextMatrix(Row, .ColIndex("缺省")))) = 1 Then
                    Call SetDrugStockDeFault(Row, Index)
                    Call SetParChange(vsfDrugStore, Index, mrsPar, True, .RowData(Row), CStr(Col))
                End If
            Case .ColIndex("窗口")
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("窗口")) <> "自动分配" And .TextMatrix(i, .ColIndex("窗口")) <> "" Then
                        strWindow = strWindow & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                    End If
                Next i
                If strWindow <> "" Then strWindow = Mid(strWindow, 2)
                Call SetParChange(vsfDrugStore, Index, mrsPar, True, strWindow, CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
End Sub

Private Sub vsfTriageQueuingDep_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim l As Long
    Dim strTmp As String

    With vsfTriageQueuingDep
        If .TextMatrix(Row, .ColIndex("ID")) = "" Then Exit Sub
        If .RowData(Row) <> .TextMatrix(Row, .ColIndex("启用")) Then
            .Cell(flexcpForeColor, Row, .ColIndex("科室")) = vbRed
        Else
            .Cell(flexcpForeColor, Row, .ColIndex("科室")) = &H80000008
        End If
    End With
End Sub

Private Sub vsfTriageQueuingDep_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfTriageQueuingDep
        If .ColKey(Col) <> "启用" Then Cancel = 1
    End With
End Sub

Private Sub vsfTriageQueuingDep_DblClick()
    With vsfTriageQueuingDep
        If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        .TextMatrix(.RowSel, .ColIndex("启用")) = IIF(.TextMatrix(.RowSel, .ColIndex("启用")) <> "0", "0", "1")
        Call vsfTriageQueuingDep_AfterEdit(.RowSel, .ColSel)
    End With
End Sub

Private Sub vsInputItemSet_DblClick(Index As Integer)
    Call SetInputItemValue(Index)
End Sub
Private Sub vsInputItemSet_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsInputItemSet(Index)
            Select Case Index
            Case vsGrid_发卡输入项设置
                If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                   If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
                   Exit Sub
                End If
                
                zlVsMoveGridCell vsInputItemSet(Index), 1, .Cols - 1
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call SetInputItemValue(Index)
End Sub
Private Sub vsInputItemSet_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsInputItemSet, Index, mrsPar)
End Sub

Private Sub vsStationRegSort_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer
    With vsStationRegSort
        If Col <> .ColIndex("是否升序") Then Exit Sub
        For i = 1 To 5
            strValue = strValue & "|" & .TextMatrix(i, .ColIndex("排序字段")) & "," & IIF(.TextMatrix(i, .ColIndex("是否升序")) = -1, 1, 0)
        Next i
        strValue = Mid(strValue, 2)
    End With
    Call SetParChange(vsStationRegSort, 0, mrsPar, True, strValue)
End Sub

'----------------------------------------------------
Private Sub vs代收_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs代收
        Select Case Col
        Case .ColIndex("固定金额")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "###0.00;-###0.00;;")
            Call SetParChange(vs代收, 0, mrsPar, True, Get代收金额)
        Case .ColIndex("选择")
        Case Else
        End Select
    End With
End Sub
Private Sub vs代收_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs代收
        Select Case Col
        Case .ColIndex("固定金额")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vs代收_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs代收
        If .Col >= .ColIndex("固定金额") And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If .Row < .Rows - 1 Then
           .Row = .Row + 1
        End If
    End With
End Sub

Private Sub vs代收_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs代收
        Select Case Col
        Case .ColIndex("固定金额")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                Else
                    zlCommFun.PressKey vbKeyTab
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs代收_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs代收_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vs代收
        Select Case .Col
            Case .ColIndex("固定金额")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
        End Select
    End With
End Sub
Private Sub Load代收款(ByVal strValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载代收款
    '入参:strValue-代收款,格式:结算方式:金额|结算方式:金额....
    '编制:刘兴洪
    '日期:2011-07-19 15:13:59
    '问题:  34705
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, varData As Variant, varTemp As Variant, j As Long, strTmp As String
    
      
    On Error GoTo ErrHandle
    '结算方式
    strSQL = _
    " Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
    " From 结算方式应用 A,结算方式 B" & _
    " Where A.应用场合='预交款' And B.名称=A.结算方式 And Nvl(B.性质,1)=5" & _
    " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '结算方式:金额|结算方式:金额....
    varData = Split(strValue, "|")
    If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
    With vs代收
        .Tag = "1": .Editable = IIF(rsTmp.RecordCount = 0, flexEDNone, flexEDKbdMouse): i = 1
        .Rows = IIF(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        Do While rsTmp.EOF = False
            .TextMatrix(i, .ColIndex("代收款项")) = NVL(rsTmp!名称)
            For j = 0 To UBound(varData)
                varTemp = Split(varData(j) & ":", ":")
                If NVL(rsTmp!名称) = varTemp(0) Then
                    .TextMatrix(i, .ColIndex("固定金额")) = Format(Val(varTemp(1)), "###0.00;-###0.00;;")
                    Exit For
                End If
            Next
            i = i + 1
            rsTmp.MoveNext
        Loop
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function Get代收金额() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取代收结算金额
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-09 16:25:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTmp  As String
    On Error GoTo ErrHandle
    With vs代收
        strTmp = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("固定金额"))) <> 0 And Trim(.TextMatrix(i, .ColIndex("代收款项"))) <> "" Then
                strTmp = strTmp & "|" & Trim(.TextMatrix(i, .ColIndex("代收款项"))) & ":" & Val(.TextMatrix(i, .ColIndex("固定金额")))
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End With
    Get代收金额 = strTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub vs代收_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vs代收, 0, mrsPar)
End Sub
Private Sub SetParRelations(ByRef arrObj As Variant, ByRef rsPar As ADODB.Recordset, _
                        Optional ByVal varPar As Variant, Optional ByVal lngModule As Long, _
                        Optional ByVal strObjTag As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置关联参数值
    '入参:arrObj-关联对像集
    '     varPar-参数号(数字)或参数名(字符)
    '     lngModule-模块号
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-09 17:56:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngIndex As Long, blnNotClear As Boolean
    
    For i = 0 To UBound(arrObj)
        If i <> 0 Then
            Call zlDatabase.zlInsertCurrRowData(rsPar, mrsPar, "")
            varPar = 0: lngModule = 0
        End If
        lngIndex = 0: blnNotClear = False
        If GetControlIndex(arrObj(i)) >= 0 Then lngIndex = arrObj(i).Index: blnNotClear = True
        Call SetParRelation(arrObj(i), lngIndex, mrsPar, varPar, lngModule, , , blnNotClear)
    Next
End Sub

Private Function GetControlIndex(ByVal obj As Object) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取控件的索引值
    '返回:-1表示未获取到索引（即不是控件数组),否则为索引值
    '编制:刘兴洪
    '日期:2015-06-10 15:48:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    GetControlIndex = obj.Index
    Exit Function
ErrHand:
    GetControlIndex = -1
End Function

Private Sub optDepsoitDelSet_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDepsoitDelSet, Index, mrsPar)
End Sub

Private Sub optDepsoitDelSet_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDepsoitDelSet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDepsoitDelSet, Index, mrsPar)
End Sub

Private Sub Load预交票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载预交票据格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_预交票据格式
    
    rsPara.Filter = "模块=" & p预交款管理 & " And 参数名='预交发票格式'"    '预交发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "预交发票格式", p预交款管理, "", vsBillFormat(intIndex).ColIndex("票据格式"))
    
    rsPara.Filter = "模块=" & p预交款管理 & " And 参数名='预交发票打印方式'"    '预交发票打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "预交发票打印方式", p预交款管理, "", vsBillFormat(intIndex).ColIndex("预交打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("预交打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
 
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat(intIndex)
        .TextMatrix(1, 0) = "门诊预交"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "住院预交"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("票据格式")) = "0"
        .ColData(.ColIndex("预交打印方式")) = "0"
        '.ForeColor = &H80000008:  .ForeColorFixed = &H80000008
'        Select Case intType
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("票据格式")) = 1 ' IIF(intType = 5, 0, 1)
'        End Select
'        Select Case intType1
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("预交打印方式")) = 1 'IIF(intType1 = 5, 0, 1)
'        End Select
'        If (Val(.ColData(.ColIndex("票据格式"))) = 1 Or _
'            Val(.ColData(.ColIndex("预交打印方式"))) = 1) Then
'            .Editable = flexEDKbdMouse
'        Else
            .Editable = flexEDKbdMouse
'        End If
    End With
    
    With vsBillFormat(intIndex)
        .Clear 1: .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("预交打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("预交打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
        Next
    End With
    zl_vsGrid_Para_Restore p预交款管理, vsBillFormat(intIndex), Me.Name, "预交发票打印方式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDrugStockEdit(ByVal Index As Integer, ByVal strType As String, ByVal intType As Integer, ByVal lngEditCol As Long, Optional strMachValue As String = "", Optional strDefaultValue As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置药房的编辑属性
    '入参:strType-类别
    '     intType-返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    '     lngEditCol-控制的编辑列
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-02 14:53:10
    '问题:25132
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSetDefault As Boolean '设置了缺省值了,随后不能再设置缺省值
    Dim lngEditForColor As Long, blnAllowEdit As Boolean, bytLockEdit As Integer '1-锁定,0-不锁定
    
    '刘兴洪:由于可能参数权限发生变更,因此,不能统一进行设置,需要设置某一部分:
    With vsfDrugStore(Index)
        blnSetDefault = False: blnAllowEdit = True
        bytLockEdit = 0
        If InStr(1, ",1,3,15,", "," & intType & ",") > 0 Then
'            lngEditForColor = IIF(blnAllowEdit, vbBlue, &H8000000C)
            bytLockEdit = IIF(blnAllowEdit, 0, 1)
        ElseIf intType = 5 Then
'            lngEditForColor = vbBlue
        Else
'            lngEditForColor = &H80000008
        End If
        
        For i = 1 To .Rows - 1
            If lngEditCol = .ColIndex("缺省") Then
                '设置药房
                If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                    .TextMatrix(i, .ColIndex("缺省")) = IIF(Val(strMachValue) > 0, 1, 0)
                    blnSetDefault = True
                End If
'                 .Cell(flexcpForeColor, i, .ColIndex("缺省")) = lngEditForColor
'                 .Cell(flexcpForeColor, i, .ColIndex("药房")) = lngEditForColor:
            Else
                If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                    .TextMatrix(i, lngEditCol) = strDefaultValue
                End If
                '设置窗口
'                 .Cell(flexcpForeColor, i, .ColIndex("窗口")) = lngEditForColor
            End If
            .Cell(flexcpData, i, lngEditCol) = bytLockEdit
        Next
    End With
End Sub

Private Sub Load药房(ByVal rsPara As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset, k As Integer
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long, arrWindow As Variant
    Dim intType As Integer
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    intType = 1
    rsPara.Filter = "模块=" & p一卡通消费操作 & " And 参数名='缺省西药房'"    '缺省西药房
    Call SetParRelation(vsfDrugStore, 0, mrsPar, "缺省西药房", p一卡通消费操作, "", vsfDrugStore(0).ColIndex("缺省"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!参数值)
    If Val(strTmp) > 0 Then
        Call SetDrugStockEdit(0, "西药房", intType, vsfDrugStore(0).ColIndex("缺省"), Val(strTmp))
    Else
        Call SetDrugStockEdit(0, "西药房", intType, vsfDrugStore(0).ColIndex("缺省"), "")
    End If
    
    rsPara.Filter = "模块=" & p一卡通消费操作 & " And 参数名='缺省中药房'"    '缺省中药房
    Call SetParRelation(vsfDrugStore, 1, mrsPar, "缺省中药房", p一卡通消费操作, "", vsfDrugStore(1).ColIndex("缺省"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!参数值)
    If Val(strTmp) > 0 Then
        Call SetDrugStockEdit(1, "中药房", intType, vsfDrugStore(1).ColIndex("缺省"), Val(strTmp))
    Else
        Call SetDrugStockEdit(1, "中药房", intType, vsfDrugStore(1).ColIndex("缺省"), "")
    End If
    
    rsPara.Filter = "模块=" & p一卡通消费操作 & " And 参数名='缺省成药房'"    '缺省成药房
    Call SetParRelation(vsfDrugStore, 2, mrsPar, "缺省成药房", p一卡通消费操作, "", vsfDrugStore(2).ColIndex("缺省"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!参数值)
    If Val(strTmp) > 0 Then
        Call SetDrugStockEdit(2, "成药房", intType, vsfDrugStore(2).ColIndex("缺省"), Val(strTmp))
    Else
        Call SetDrugStockEdit(2, "成药房", intType, vsfDrugStore(2).ColIndex("缺省"), "")
    End If
    
    rsPara.Filter = "模块=" & p一卡通消费操作 & " And 参数名='西药房窗口'"    '西药房窗口
    Call SetParRelation(vsfDrugStore, 0, mrsPar, "西药房窗口", p一卡通消费操作, "", vsfDrugStore(0).ColIndex("窗口"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!参数值)
    If strTmp <> "" Then
        arrWindow = Split(strTmp, ",")
        For k = 0 To UBound(arrWindow)
            If arrWindow(k) <> "" Then
                Call SetDrugStockEdit(0, "西药房", intType, vsfDrugStore(0).ColIndex("窗口"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
            End If
        Next
    Else
        Call SetDrugStockEdit(0, "西药房", intType, vsfDrugStore(0).ColIndex("窗口"), "")
    End If
    
    rsPara.Filter = "模块=" & p一卡通消费操作 & " And 参数名='中药房窗口'"    '中药房窗口
    Call SetParRelation(vsfDrugStore, 1, mrsPar, "中药房窗口", p一卡通消费操作, "", vsfDrugStore(1).ColIndex("窗口"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!参数值)
    If strTmp <> "" Then
        arrWindow = Split(strTmp, ",")
        For k = 0 To UBound(arrWindow)
            If arrWindow(k) <> "" Then
                Call SetDrugStockEdit(1, "中药房", intType, vsfDrugStore(1).ColIndex("窗口"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
            End If
        Next
    Else
        Call SetDrugStockEdit(1, "中药房", intType, vsfDrugStore(1).ColIndex("窗口"), "")
    End If
    
    rsPara.Filter = "模块=" & p一卡通消费操作 & " And 参数名='成药房窗口'"    '成药房窗口
    Call SetParRelation(vsfDrugStore, 2, mrsPar, "成药房窗口", p一卡通消费操作, "", vsfDrugStore(2).ColIndex("窗口"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!参数值)
    If strTmp <> "" Then
        arrWindow = Split(strTmp, ",")
        For k = 0 To UBound(arrWindow)
            If arrWindow(k) <> "" Then
                Call SetDrugStockEdit(2, "成药房", intType, vsfDrugStore(2).ColIndex("窗口"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
            End If
        Next
    Else
        Call SetDrugStockEdit(2, "成药房", intType, vsfDrugStore(2).ColIndex("窗口"), "")
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load门诊费用转住院预交票格式()
    '功能:加载门诊费用转住院预交票格式
    Dim strReport As String, strBillFormat As String, strBillFormat1 As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    With cbo(cbo_门诊费用转住院预交发票格式)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!序号) & "-" & NVL(rsTemp!说明)
            .ItemData(.NewIndex) = Val(NVL(rsTemp!序号))
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load一卡通票据格式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载一卡通消费票据格式的
    '编制:刘兴洪
    '日期:2015-06-29 13:57:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strReport As String, strBillFormat As String, strBillFormat1 As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    strReport = "ZL" & glngSys \ 100 & "_BILL_1151"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    With cbo(cbo_一卡通_收票票据格式)
        .Clear: cbo(cbo_一卡通_记帐票据格式).Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!序号) & "-" & NVL(rsTemp!说明)
            .ItemData(.NewIndex) = Val(NVL(rsTemp!序号))
            cbo(cbo_一卡通_记帐票据格式).AddItem NVL(rsTemp!序号) & "-" & NVL(rsTemp!说明)
            cbo(cbo_一卡通_记帐票据格式).ItemData(cbo(cbo_一卡通_记帐票据格式).NewIndex) = Val(NVL(rsTemp!序号))
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 Then .ListIndex = 0
        If cbo(cbo_一卡通_记帐票据格式).ListCount <> 0 Then cbo(cbo_一卡通_记帐票据格式).ListIndex = 0
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

Private Sub Load收费票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费票据格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varPatiData As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    Dim strPatiBillFormat As String '按病人补打票据格式
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_收费票据格式
    
    rsPara.Filter = "模块=" & p门诊收费管理 & " And 参数名='收费发票格式'"    '收费发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "收费发票格式", p门诊收费管理, "", vsBillFormat(intIndex).ColIndex("收费票据格式"))
    
    rsPara.Filter = "模块=" & p门诊收费管理 & " And 参数名='按病人补打发票格式'"    '按病人补打发票格式
    If Not rsPara.EOF Then strPatiBillFormat = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "按病人补打发票格式", p门诊收费管理, "", vsBillFormat(intIndex).ColIndex("按病人补打票据格式"))
    
    
    rsPara.Filter = "模块=" & p门诊收费管理 & " And 参数名='收费发票打印方式'"    '收费发票打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "收费发票打印方式", p门诊收费管理, "", vsBillFormat(intIndex).ColIndex("收费打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    varPatiData = Split(strPatiBillFormat, "|")
    
    strReport = "ZL" & glngSys \ 100 & "_BILL_1121_1"
    Set rsTemp = zlGetBillFormatRec(strReport)
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("收费票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("按病人补打票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("收费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
 
 
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(vsGrid_收费票据格式)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("收费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("收费票据格式")) = "0"
            
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("收费票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varPatiData)
                varTemp = Split(varPatiData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("按病人补打票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("收费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p门诊收费管理, vsBillFormat(intIndex), Me.Name, "收费票据格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load退费票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载退费票据格式
    '编制:冉俊明
    '日期:2016-06-1
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varPatiData As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_退费票据格式
    
    rsPara.Filter = "模块=" & p门诊收费管理 & " And 参数名='退费发票格式'"    '退费发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "退费发票格式", p门诊收费管理, "", vsBillFormat(intIndex).ColIndex("收费票据格式"))
    
    rsPara.Filter = "模块=" & p门诊收费管理 & " And 参数名='退费发票打印方式'"    '退费发票打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "退费发票打印方式", p门诊收费管理, "", vsBillFormat(intIndex).ColIndex("收费打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    strReport = "ZL" & glngSys \ 100 & "_BILL_1121_7"
    Set rsTemp = zlGetBillFormatRec(strReport)
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("收费票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("收费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
 
 
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("收费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("收费票据格式")) = "0"
            
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("收费票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("收费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p门诊收费管理, vsBillFormat(intIndex), Me.Name, "退费票据格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub Load补结算票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载补结算票据格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_补结算票据格式
    
    rsPara.Filter = "模块=" & p门诊补结算 & " And 参数名='收费发票格式'"     '收费发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "收费发票格式", p门诊补结算, "", vsBillFormat(intIndex).ColIndex("票据格式"))
    rsPara.Filter = "模块=" & p门诊补结算 & " And 参数名='收费发票打印方式'"    '收费发票打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "收费发票打印方式", p门诊补结算, "", vsBillFormat(intIndex).ColIndex("收费打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1124"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("收费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    If zlStartFactUseType(1) Then
        rsTemp.Filter = "ID<>0"
    End If
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("收费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("收费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p门诊补结算, vsBillFormat(intIndex), Me.Name, "补结算票据格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load医疗卡票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载补结算票据格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_医疗卡收据格式
    
    rsPara.Filter = "模块=" & p医疗卡管理 & " And 参数名='医疗卡收据格式'"      '收费发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "医疗卡收据格式", p医疗卡管理, "", vsBillFormat(intIndex).ColIndex("票据格式"))
    
    varData = Split(strBillFormat, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1107"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = 3
        lngRow = 1
        Dim j As Integer
        For i = 1 To 2
            .TextMatrix(lngRow, .ColIndex("发卡类型")) = IIF(i = 1, "发卡", "绑定卡")
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = IIF(i = 1, Val(varData(0)), Val(varData(1)))
            lngRow = lngRow + 1
        Next
    End With
    zl_vsGrid_Para_Restore p医疗卡管理, vsBillFormat(intIndex), Me.Name, "医疗卡收据格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load补结算退费票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载补结算退费票据格式
    '编制:冉俊明
    '日期:2016-06-1
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_补结算退费票据格式
    
    rsPara.Filter = "模块=" & p门诊补结算 & " And 参数名='退费发票格式'"     '退费发票格式
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "退费发票格式", p门诊补结算, "", vsBillFormat(intIndex).ColIndex("票据格式"))
    
    rsPara.Filter = "模块=" & p门诊补结算 & " And 参数名='退费发票打印方式'"    '退费发票打印方式
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "退费发票打印方式", p门诊补结算, "", vsBillFormat(intIndex).ColIndex("收费打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1124_3"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("收费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    If zlStartFactUseType(1) Then
        rsTemp.Filter = "ID<>0"
    End If
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("收费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("收费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p门诊补结算, vsBillFormat(intIndex), Me.Name, "补结算退费票据格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load结帐票据格式(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结帐票据格式
    '编制:刘兴洪
    '日期:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_结帐票据格式
    
    rsPara.Filter = "模块=" & p病人结帐管理 & " And 参数名='病人结帐打印'"
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!参数值)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "病人结帐打印", p病人结帐管理, "", vsBillFormat(intIndex).ColIndex("结帐后打印方式"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("结帐后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("结帐后打印方式")) = "0-不打印票据"
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(NVL(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("结帐后打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next

            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p病人结帐管理, vsBillFormat(intIndex), Me.Name, "结帐票据格式", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlReadBillFormat(ByVal ReportCode As String) As ADODB.Recordset
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定报表的打印格式
    '入参:ReportCode-报表名称
    '返回:报表打印格式的记录集
    '编制:李南春
    '日期:2014-10-20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo ErrHandle
    
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='" & ReportCode & "'  " & _
    "   Order by  序号"
    Set zlReadBillFormat = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetBillFormat(ByVal intIndex As Integer, ByVal intType As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据参数设置
    '入参:intIndex-票据打印格式索引
    '     intType-0-获取票据打印方式;1-获取票据格式;2-获取按病人补打票据格式
    '返回:返回票据打印格式或打印方式
    '编制:刘兴洪
    '日期:2015-06-10 14:01:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrintMode As String, strPrintFormat As String, strPatiPrintFormat As String
    Dim i As Long
    
    On Error GoTo ErrHandle
    strPrintFormat = "": strPrintMode = ""
    Select Case intIndex
    Case vsGrid_预交票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("预交打印方式")), 1))
                End If
            Next
        End With
    Case vsGrid_预交红票格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("退款打印方式")), 1))
                End If
            Next
        End With
    Case vsGrid_收费票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                '80943,冉俊明,2014-12-18,票据未使用“收费类别”时，加入设置收费类别为空的打印方式和票据格式
                'If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("收费票据格式")))
                    strPatiPrintFormat = strPatiPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("按病人补打票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("收费打印方式")), 1))
                'End If
            Next
        End With
    Case vsGrid_补结算票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                '80943,冉俊明,2014-12-18,票据未使用“收费类别”时，加入设置收费类别为空的打印方式和票据格式
                'If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("收费打印方式")), 1))
                'End If
            Next
        End With
    Case vsGrid_退费票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("收费票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("收费打印方式")), 1))
            Next
        End With
    Case vsGrid_补结算退费票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("收费打印方式")), 1))
            Next
        End With
    
    Case vsGrid_结帐票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    'strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("结帐后打印方式")), 1))
                End If
            Next
        End With
        
    Case vsGrid_结帐红票格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    'strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("作废后打印方式")), 1))
                End If
            Next
        End With
    Case vsGrid_发卡预交票据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("预交打印方式")), 1))
                End If
            Next
        End With
    Case vsGrid_医疗卡收据格式
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("发卡类型"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Val(.TextMatrix(i, .ColIndex("票据格式")))
                End If
            Next
        End With
    End Select
    If strPrintFormat <> "" Then strPrintFormat = Mid(strPrintFormat, 2)
    If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
    If strPatiPrintFormat <> "" Then strPatiPrintFormat = Mid(strPatiPrintFormat, 2)
    '0-获取票据打印方式;1-获取票据格式;2-获取按病人补打票据格式
    If intType = 2 Then GetBillFormat = strPatiPrintFormat: Exit Function
    GetBillFormat = IIF(intType = 0, strPrintMode, strPrintFormat)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub SetInputItemValue(ByVal intIndex As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置当前项目的相关值
    '入参:intIndex-网格控件数组的索引值
    '编制:刘兴洪
    '日期:2015-06-11 17:58:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
       
    On Error GoTo ErrHandle
    With vsInputItemSet(intIndex)
        Select Case .Col
        Case .ColIndex("禁止录入")
            .TextMatrix(.Row, .ColIndex("禁止录入")) = IIF(.TextMatrix(.Row, .ColIndex("禁止录入")) = "", "√", "")
            If .TextMatrix(.Row, .ColIndex("禁止录入")) = "√" Then
                .TextMatrix(.Row, .ColIndex("光标进入")) = ""
                .TextMatrix(.Row, .ColIndex("必输项")) = ""
                .Cell(flexcpBackColor, .Row, .ColIndex("必输项"), .Row, .ColIndex("光标进入")) = &H8000000F
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("必输项"), .Row, .ColIndex("光标进入")) = &H8000000E
            End If
            .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入")) = &H8000000E
        Case .ColIndex("必输项")
        
            .TextMatrix(.Row, .ColIndex("必输项")) = IIF(.TextMatrix(.Row, .ColIndex("必输项")) = "", "√", "")
            If .TextMatrix(.Row, .ColIndex("必输项")) = "√" Then
                .TextMatrix(.Row, .ColIndex("禁止录入")) = ""
                .TextMatrix(.Row, .ColIndex("光标进入")) = "√"
                .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入")) = &H8000000F
                .Cell(flexcpBackColor, .Row, .ColIndex("光标进入")) = &H8000000E
            ElseIf .TextMatrix(.Row, .ColIndex("光标进入")) = "√" Then
                .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入")) = &H8000000F
                .Cell(flexcpBackColor, .Row, .ColIndex("光标进入")) = &H8000000E
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入"), .Row, .ColIndex("光标进入")) = &H8000000E
            End If
             .Cell(flexcpBackColor, .Row, .ColIndex("必输项")) = &H8000000E
        Case .ColIndex("光标进入")
            .TextMatrix(.Row, .ColIndex("光标进入")) = IIF(.TextMatrix(.Row, .ColIndex("光标进入")) = "", "√", "")
             .Cell(flexcpBackColor, .Row, .ColIndex("光标进入")) = &H8000000E
            If .TextMatrix(.Row, .ColIndex("光标进入")) = "√" Then
                .TextMatrix(.Row, .ColIndex("禁止录入")) = ""
                
                .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入")) = &H8000000F
            ElseIf .TextMatrix(.Row, .ColIndex("必输项")) = "√" Then
                .TextMatrix(.Row, .ColIndex("禁止录入")) = ""
                .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入")) = &H8000000F
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("禁止录入"), .Row, .ColIndex("光标进入")) = &H8000000E
            End If
        End Select
    End With
    Call SetParChange(vsInputItemSet, intIndex, mrsPar, True, GetInputItemSetValue(intIndex))
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function GetInputItemSetValue(ByVal intIndex As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取输入项的设置值
    '入参:intIndex-控件索引
    '返回:输入项设置的值,格式:输入项,是否禁用,光标是否跳过,是否必输项|....
    '编制:刘兴洪
    '日期:2015-06-11 18:10:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTmp As String
    On Error GoTo ErrHandle
        
    With vsInputItemSet(intIndex)
        For i = 1 To .Rows - 1
            strTmp = strTmp & "|" & .TextMatrix(i, .ColIndex("输入项目"))
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("禁止录入")) = "√", 1, 0)
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("必输项")) = "√", 1, 0)
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("光标进入")) = "√", 1, 0)
        Next
    End With
    GetInputItemSetValue = Mid(strTmp, 2)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitPage(ByVal intIndex As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '入参:intIndex-页面控件数组的索引
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-15 16:04:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo ErrHand:
     
    If intIndex = Pg_挂号业务 Then
        With tbPage(intIndex)
            picRegistPlan.BorderStyle = 0
            picRegist.BorderStyle = 0
            pic预约.BorderStyle = 0
            picOtherRegister.BorderStyle = 0
            
            Set ObjItem = .InsertItem(Pg_挂号_安排, "安排", picRegistPlan.hwnd, 0)
            ObjItem.Tag = Pg_挂号_安排
            Set ObjItem = .InsertItem(Pg_挂号_挂号, "挂号窗口", picRegist.hwnd, 0)
            ObjItem.Tag = Pg_挂号_挂号
            
            Set ObjItem = .InsertItem(Pg_挂号_预约, "预约窗口", pic预约.hwnd, 0)
            ObjItem.Tag = Pg_挂号_预约
            Set ObjItem = .InsertItem(Pg_挂号_其他, "其他(医生站/分诊台等)", picOtherRegister.hwnd, 0)
            ObjItem.Tag = Pg_挂号_其他
            
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        Exit Sub
    End If
    If intIndex = Pg_结帐业务 Then
        With tbPage(intIndex)
            picSettlePar(1).BorderStyle = 0
            picSettlePar(0).BorderStyle = 0
            picSettlePar(2).BorderStyle = 0
            picSettlePar(1).BackColor = &H8000000F
            picSettlePar(0).BackColor = &H8000000F
            picSettlePar(2).BackColor = &H8000000F
        
            Set ObjItem = .InsertItem(Pg_结帐_结帐参数, "结帐参数", picSettlePar(0).hwnd, 0)
            ObjItem.Tag = Pg_结帐_结帐参数
            Set ObjItem = .InsertItem(Pg_结帐_票据控制, "票据控制", picSettlePar(1).hwnd, 0)
            ObjItem.Tag = Pg_结帐_票据控制
            Set ObjItem = .InsertItem(Pg_结帐_界面风格, "界面风格", picSettlePar(2).hwnd, 0)
            ObjItem.Tag = Pg_结帐_界面风格
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        Exit Sub
    End If
    If intIndex = Pg_门诊收费 Then
      With tbPage(intIndex)
            picChargePg(1).BorderStyle = 0
            picChargePg(0).BorderStyle = 0
            picChargePg(1).BackColor = &H8000000F
            picChargePg(0).BackColor = &H8000000F
            
            Set ObjItem = .InsertItem(Pg_收费_单据控制, "单据控制", picChargePg(0).hwnd, 0)
            ObjItem.Tag = Pg_收费_单据控制
            Set ObjItem = .InsertItem(Pg_收费_票据控制, "票据控制", picChargePg(1).hwnd, 0)
            ObjItem.Tag = Pg_收费_票据控制
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        Exit Sub
    End If
    
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cmdAddedItem_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strValue As String
    
    mblnNotChange = True
    strSQL = "Select ID, 编码, 名称, 计算单位, 说明" & vbNewLine & _
            "From 收费项目目录" & vbNewLine & _
            "Where 类别 = 'Z' And Nvl(是否变价, 0) = 0 And 服务对象 In(1,3)" & vbNewLine & _
            "Order By 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "(定价)收费项目")
    If Not rsTmp Is Nothing Then
        txt(txt_收费_自助加收挂号费).Text = NVL(rsTmp!名称)
        cmdAddedItem.Tag = NVL(rsTmp!ID)
        If chk(chk_收费_未挂号自动加收挂号费).value = 0 Then chk(chk_收费_未挂号自动加收挂号费).value = 1
    End If
    strValue = cmdAddedItem.Tag & ";" & txt(txt_收费_自助加收挂号费).Text
    Call SetParChange(txt, txt_收费_自助加收挂号费, mrsPar, True, strValue)
    Call SetParChange(chk, chk_收费_未挂号自动加收挂号费, mrsPar, True, strValue)
    
    mblnNotChange = False
    Exit Sub
errH:
    mblnNotChange = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPrintListHaveData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据打印明细是否有数据
    '返回:有数据返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-05-17 14:24:40
    '说明:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHandle
    strSQL = "Select 1 From 票据打印明细 where Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    GetPrintListHaveData = rsTemp.RecordCount >= 1
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetBillRuleParaLocale()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置票据分配规则的位置
    '编制:刘兴洪
    '日期:2013-03-26 15:43:12
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intIndex As Integer
    On Error GoTo ErrHandle
    
    intIndex = cbo(cbo_收费_票据分配规则).ListIndex
    
    If intIndex < 0 Or intIndex > 2 Then intIndex = 0
    
    For i = 0 To 2
        picRuleBack(i).Visible = intIndex = i
    Next
    
    If intIndex = 0 Then
        '根据实际打印分配票号
        '设置容器
        Set chk(chk_收费_自动加收工本费).Container = picRuleBack(0)
        '工本费
        With chk(chk_收费_自动加收工本费)
            .Top = fraActuallyPrint.Top
            .Left = chk(chk_收费_每次只用一张票据).Left
            .TabIndex = chk(chk_收费_每次只用一张票据).TabIndex + 1
        End With
     End If

    If intIndex = 1 Then
        '根据预定规则分配票号
        '设置容器
        Set chk(chk_收费_自动加收工本费).Container = picRuleBack(1)
        '工本费
       With chk(chk_收费_自动加收工本费)
            .Top = fraRuleSystem.Top
            .Left = fraRuleSystem.Left + 100
            .TabIndex = chkBillRule(0).TabIndex - 1
       End With
    End If
    If intIndex = 2 Then
        '根据用户自定义规则处理
        '设置容器
        Set chk(chk_收费_自动加收工本费).Container = picRuleBack(2)
        '工本费
       With chk(chk_收费_自动加收工本费)
            .Top = lblCustomInfor.Top - .Height - 50
            .Left = lblCustomInfor.Left
            .TabIndex = cbo(intIndex).TabIndex + 1
        End With
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub chkBillRule_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    
    If Index <> 0 And chkBillRule(Index).value = 1 Then
        If Val(txtBillRuleNum(Index - 1).Text) = 0 Then
            updBillRuleNum(Index - 1).value = Val(txtBillRuleNum(Index - 1).Tag)    '恢复缺省值
        End If
    End If
    Call SetBillRuleEnable
    Call ShowRuleInfor
    If Not optRuleTotal(2).Visible Then
         If optRuleTotal(2).value Then optRuleTotal(0).value = True
    End If
    Call SaveBillRuleChange
End Sub


Private Sub ShowRuleInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示票号的分配规则
    '编制:刘兴洪
    '日期:2013-03-26 14:14:08
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, i As Integer
    Dim strName As String
    
    On Error GoTo ErrHandle
    strInfor = ""
    If chkBillRule(0).value = 1 Then
       strInfor = strInfor & "+ NO"
    End If
    For i = 1 To 3
        If chkBillRule(i).value = 1 Then
            strName = Switch(i = 1, "执行科室", i = 2, "收据费目", True, "收据细目")
            strInfor = strInfor & "+" & strName & "(" & txtBillRuleNum(i - 1).Text & ")"
        End If
    Next
    If strInfor <> "" Then strInfor = Mid(strInfor, 2)
    lblInfor.Caption = strInfor
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetBillRuleEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据票据分配规则,设置相应控件的Enabled属性
    '编制:刘兴洪
    '日期:2013-03-26 17:55:47
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, blnEnable As Boolean
    On Error GoTo ErrHandle
    '汇总条件(0-不汇总;1-首页汇总(按第1页汇总),2-分组汇总(选择明细时有效))
    '1.分组汇总:有费用明细时,同时存在勾选执行科室或者收据费目或按单据才会存在分组汇总项,才会存在汇总项
    blnEnable = chkBillRule(3).Enabled And chkBillRule(3).value = 1 And (chkBillRule(2).value = 1 Or chkBillRule(1).value = 1 Or chkBillRule(0).value = 1)
    optRuleTotal(2).Visible = blnEnable
    optRuleTotal(2).Enabled = blnEnable
    '2.首页汇总:都允许设置成汇总额
    optRuleTotal(1).Enabled = chkBillRule(3).Enabled
    optRuleTotal(0).Enabled = chkBillRule(3).Enabled
    
    '设置分组汇总相题
    If chkBillRule(0).value = 1 Then
        optRuleTotal(2).Caption = "按单据号分组汇总"
    ElseIf chkBillRule(1).value = 1 Then
        optRuleTotal(2).Caption = "按执行科室分组汇总"
    ElseIf chkBillRule(3).value = 1 Then
        optRuleTotal(2).Caption = "按收据费目分组汇总"
    ElseIf chkBillRule(3).value = 1 Then
        optRuleTotal(2).Caption = "按分组条件汇总"
    End If
    For intIndex = 1 To 3
        txtBillRuleNum(intIndex - 1).Enabled = chkBillRule(intIndex).value = 1 And chkBillRule(intIndex).Enabled
        updBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
        lblBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function Save票据分配规则() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存票据分配规则发生变化
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-18 12:13:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intList As Integer
    On Error GoTo ErrHandle
    intList = cbo(cbo_收费_票据分配规则).ListIndex
    
    If mint原票据分配规则 <> intList And intList <> 0 And mint原票据分配规则 <= 0 Then
       '如果当前切换成新模式,需要将票据打印格式记录下来,以便在重打或部分退费时按切换前的票据格式打印
       Call zlDatabase.ExecuteProcedure("Zl_Update_Bill_Printformat(" & glngSys & ")", Me.Caption)
    End If
    Save票据分配规则 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub updBillRuleNum_Change(Index As Integer)
    If Me.Visible = False Then Exit Sub
    
    If updBillRuleNum(Index).value = 0 Then
        chkBillRule(Index + 1).value = 0
    End If
     Call SaveBillRuleChange
End Sub
Private Sub InitBillRuleCtrl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化票据规则的相关控件
    '编制:刘兴洪
    '日期:2015-06-19 14:43:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    On Error GoTo ErrHandle
    For i = 0 To picRuleBack.UBound
        picRuleBack(i).BorderStyle = 0
    Next
    cbo(cbo_收费_票据分配规则).ZOrder 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function zlStartFactUseType(ByVal bytBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '入参:bytBillType-票种
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo ErrHandle
    strSQL = "Select  1 as 存在 From 票据领用记录 where 票种=[1] and nvl(使用类别,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查票据是否启用了使用类别的", bytBillType)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Load补充结算方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载补充结算方式
    '编制:刘兴洪
    '日期:2015-06-23 14:14:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    '排除消费卡对应的结算方式(性质8)，以及未启用一卡通的老一卡通结算方式(性质7)
    strSQL = _
        "Select Distinct b.编码, b.名称" & vbNewLine & _
        "From 结算方式应用 A, 结算方式 B" & vbNewLine & _
        "Where a.结算方式 = b.名称 And a.应用场合 In ('挂号', '收费')" & vbNewLine & _
        "     And Instr(',3,4,', ',' || b.性质 || ',') = 0" & vbNewLine & _
        "     And (b.性质 <> 7 Or b.性质 = 7 And Exists (Select 1 From 一卡通目录 Where 结算方式 = b.名称 And 启用 = 1))" & vbNewLine & _
        "     And (b.性质 <> 8 Or b.性质 = 8 And Not Exists (Select 1 From 消费卡类别目录 Where 结算方式 = b.名称))" & vbNewLine & _
        "Order By LPad(编码, 3, ' ')"
    strSQL = _
        "Select 编码, 名称" & vbNewLine & _
        "From (" & strSQL & ")" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select '00', '冲预存款' From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "保险补充结算")
    
    With lst(lst_补结算_结算方式)
        .Clear
        Do Until rsTemp.EOF
            .AddItem NVL(rsTemp!名称)
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetBillUseTypeRec(ByRef rsUseType As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据使用类别
    '出参:rsUseType-使用类别集
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-24 10:35:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    If Not mrsBillUseType Is Nothing Then
        If mrsBillUseType.State = 1 Then
            Set rsUseType = mrsBillUseType: GetBillUseTypeRec = True: Exit Function
        End If
    End If
    strSQL = "" & _
    "   Select rowNum as ID,编码, 名称 From 票据使用类别" & _
    "   Union All" & _
    "   Select 0 as ID, '', '' From Dual " & _
    "   Order By 编码"
    Set mrsBillUseType = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set rsUseType = mrsBillUseType
    GetBillUseTypeRec = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Load自费费用类别()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载自费费用类别
    '编制:刘兴洪
    '日期:2015-06-24 15:34:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    
    strSQL = "" & _
    "   Select 编码,名称 as 类别 " & _
    "   From 收费项目类别 " & _
    "   Where 编码 <> '1'" & _
    "   Order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With lst(lst_结帐_自费费用类别)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!类别
            .ItemData(.NewIndex) = Asc(rsTemp!编码)
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetBillLenSet() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据长度设置
    '返回:返回票据长度设置,格式:1-收费,2-预交,3-结帐,4-挂号
    '编制:刘兴洪
    '日期:2015-07-28 17:07:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    strTemp = strTemp & lvw(lvw_票据).ListItems("C1").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C2").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C3").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C4").SubItems(1)
    GetBillLenSet = strTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetBillCtlSet() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据控制设置
    '返回:返回票据控制设置:格式:1111 分别1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '编制:刘兴洪
    '日期:2015-07-28 17:07:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    On Error GoTo ErrHandle

    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C1").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C2").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C3").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C4").SubItems(2) = "√", "1", "0")
    GetBillCtlSet = strTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub optPrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintMode, Index, mrsPar)
End Sub

Private Sub optPrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintMode, Index, mrsPar)
End Sub


Private Sub optToExcelMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optToExcelMode, Index, mrsPar)
End Sub

Private Sub optToExcelMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optToExcelMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optToExcelMode, Index, mrsPar)
End Sub

Private Sub optVisitTablePrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optVisitTablePrintMode, Index, mrsPar)
End Sub

Private Sub optVisitTablePrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optVisitTablePrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optVisitTablePrintMode, Index, mrsPar)
End Sub

Private Sub LoadDelFeeDefaultType()
    '设置缺省退费方式
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long, strValue As String
    
    Call SetParRelation(vsfDelFeeDefaultType, 0, mrsPar, "非三方卡退费缺省方式", p门诊收费管理)
    mrsPar.Filter = "模块=" & p门诊收费管理 & " And 参数名='非三方卡退费缺省方式'"
    If Not mrsPar.EOF Then strValue = NVL(mrsPar!参数值)
    
    strSQL = _
        "Select 名称" & vbNewLine & _
        "From 结算方式 A, 结算方式应用 B" & vbNewLine & _
        "Where a.名称 = b.结算方式 And b.应用场合 = '收费' And a.性质 = 2 And Nvl(a.应付款, 0) = 0 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsfDelFeeDefaultType
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("收款方式")) = NVL(rsTemp!名称)
            If InStr(strValue, NVL(rsTemp!名称)) > 0 Then
                .TextMatrix(lngRow, .ColIndex("缺省退现")) = 1
            End If
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    zl_vsGrid_Para_Restore p门诊收费管理, vsfDelFeeDefaultType, Me.Name, "退费缺省方式", False, False
End Sub

Private Sub SaveTriageQueuingDep()
'功能:保存部门分诊台签到排队参数
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    With vsfTriageQueuingDep
        For i = 1 To .Rows - 1
            strSQL = "zl_Parameters_Update('分诊台签到排队','" & IIF(.TextMatrix(i, .ColIndex("启用")) = "0", "0", "1") & "'," & _
                                          glngSys & "," & p分诊管理 & ",1," & Val(.TextMatrix(i, .ColIndex("ID"))) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "保存参数【分诊台签到排队】")
            Call zlDatabase.ClearParaCache
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTriageQueuingDep()
'功能：加载分诊台签到排队参数适用科室
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHandle
    
    mrsPar.Filter = "模块=" & p分诊管理 & " And 参数名='分诊台签到排队'"
    If mrsPar.RecordCount <= 0 Then Exit Sub
        
    With vsfTriageQueuingDep
        If chk(chk_分诊_分诊台签到开始排队).value = 0 Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("启用")) = 0
            Next
            Exit Sub
        End If
        .Rows = 1
        strSQL = " Select a.Id, a.编码, a.名称 As 科室, Nvl(b.参数值, 1) As 启用" & vbNewLine & _
                 " From 部门表 A," & vbNewLine & _
                 "     (Select b.部门id, b.参数值" & vbNewLine & _
                 "       From zlParameters A, Zldeptparas B" & vbNewLine & _
                 "       Where a.Id = b.参数id And a.系统 = 100 And a.模块 = 1113 And a.参数名 = '分诊台签到排队') B, 部门性质说明 C" & vbNewLine & _
                 " Where a.Id = b.部门id(+) And a.Id = c.部门id And c.工作性质 = '临床'" & vbNewLine & _
                 " Order By a.名称"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取分诊台签到排队适用科室")
        
        Do While rsTemp.EOF = False
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(i, .ColIndex("编码")) = NVL(rsTemp!编码)
            .TextMatrix(i, .ColIndex("科室")) = NVL(rsTemp!科室)
            .TextMatrix(i, .ColIndex("启用")) = NVL(rsTemp!启用)
            .RowData(i) = NVL(rsTemp!启用)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End With
    
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetTriageQueuingEnalbe(Optional ByVal blnEnable As Boolean = False)
    '功能:设置 "分诊台签到开始排队"参数相关控件的Enable属性
    On Error GoTo ErrHandle
    With vsfTriageQueuingDep
        .Cell(flexcpForeColor, 1, .ColIndex("科室"), .Rows - 1, .ColIndex("科室")) = IIF(blnEnable = True, &H80000008, &H8000000C)
        .Enabled = blnEnable
    End With
    cmdDepSelectAll.Enabled = blnEnable
    cmdDepSelectAll.Visible = blnEnable
    cmdDepClearAll.Enabled = blnEnable
    cmdDepClearAll.Visible = blnEnable

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadStationRegOrder(ByVal strOrder As String)
    '功能:加载参数"医生站挂号排序控制"到vsStationRegSort列表中
    Dim varOrder As Variant, varData As Variant
    Dim i As Integer
    
    If strOrder = "" Then Exit Sub
    varOrder = Split(strOrder, "|")
    With vsStationRegSort
        For i = 0 To UBound(varOrder)
            varData = Split(varOrder(i), ",")
            If varData(0) <> "" Then
                .TextMatrix(i + 1, .ColIndex("排序字段")) = varData(0)
                .TextMatrix(i + 1, .ColIndex("是否升序")) = IIF(varData(1) = 0, 0, -1)
            End If
        Next
    End With
End Sub
