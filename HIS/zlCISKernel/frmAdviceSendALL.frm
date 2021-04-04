VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceSendALL 
   AutoRedraw      =   -1  'True
   Caption         =   "住院医嘱发送"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   270
   ClientWidth     =   11760
   Icon            =   "frmAdviceSendALL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAdviceSendALL.frx":058A
   ScaleHeight     =   7905
   ScaleWidth      =   11760
   Begin MSComCtl2.DTPicker dkpExecTime 
      Height          =   300
      Left            =   5400
      TabIndex        =   54
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   216727555
      CurrentDate     =   40976
   End
   Begin VB.PictureBox pic药房 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   5400
      ScaleHeight     =   1800
      ScaleWidth      =   3885
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3870
      Width           =   3885
      Begin VB.CheckBox chkDeptExchange 
         Caption         =   "本次发送启用药房置换"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDept 
         Height          =   1545
         Left            =   0
         TabIndex        =   25
         Top             =   255
         Width           =   3840
         _cx             =   6773
         _cy             =   2725
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         Rows            =   6
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAdviceSendALL.frx":0B14
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
   Begin VB.PictureBox pic基本 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   90
      ScaleHeight     =   6375
      ScaleWidth      =   3840
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   3840
      Begin VB.Frame fra医嘱条件 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   75
         TabIndex        =   48
         Top             =   1155
         Width           =   3460
         Begin VB.ComboBox cboDrugType 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1260
            Width           =   2715
         End
         Begin VB.CommandButton cmd执行科室 
            Height          =   240
            Left            =   3200
            Picture         =   "frmAdviceSendALL.frx":0B69
            Style           =   1  'Graphical
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "选择执行科室(F4)"
            Top             =   600
            Width           =   240
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            ItemData        =   "frmAdviceSendALL.frx":0C5F
            Left            =   750
            List            =   "frmAdviceSendALL.frx":0C61
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   0
            Width           =   900
         End
         Begin VB.ComboBox cbo药房 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   915
            Width           =   2715
         End
         Begin VB.TextBox txt执行科室 
            Height          =   300
            Left            =   750
            TabIndex        =   10
            Top             =   585
            Width           =   2715
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1645
            TabIndex        =   9
            Top             =   0
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   216727555
            CurrentDate     =   37953
         End
         Begin VB.Label lblDrugType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药嘱分类"
            Height          =   180
            Left            =   0
            TabIndex        =   57
            Top             =   1305
            Width           =   720
         End
         Begin VB.Label lblEatDrug 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "长嘱口服药结束时间：2015-11-25 14:59"
            Height          =   180
            Left            =   0
            TabIndex        =   56
            Top             =   345
            Width           =   3240
         End
         Begin VB.Label lbl执行科室 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行科室"
            Height          =   180
            Left            =   0
            TabIndex        =   51
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lbl结束时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束时间"
            Height          =   180
            Left            =   0
            TabIndex        =   50
            Top             =   60
            Width           =   720
         End
         Begin VB.Label lbl药房 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发药药房"
            Height          =   180
            Left            =   0
            TabIndex        =   49
            Top             =   960
            Width           =   720
         End
      End
      Begin VB.Frame fra药品 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   90
         TabIndex        =   47
         Top             =   0
         Width           =   3500
         Begin VB.CheckBox chkSendNow 
            Caption         =   "药品长嘱从当天开始发送"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            TabIndex        =   55
            Top             =   300
            Width           =   2655
         End
         Begin VB.CheckBox chk药品期效 
            Caption         =   "药品临嘱"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   1
            Left            =   2220
            TabIndex        =   1
            Top             =   0
            Width           =   1200
         End
         Begin VB.CheckBox chk药品期效 
            Caption         =   "药品长嘱"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   0
            Left            =   750
            TabIndex        =   0
            Top             =   0
            Width           =   1300
         End
         Begin VB.Image imgCheck 
            Height          =   360
            Left            =   140
            Picture         =   "frmAdviceSendALL.frx":0C63
            Top             =   -30
            Width           =   360
         End
      End
      Begin VB.Frame fra其他 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   90
         TabIndex        =   46
         Top             =   600
         Width           =   3500
         Begin VB.CheckBox chk其他期效 
            Caption         =   "其他临嘱"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   1
            Left            =   2220
            TabIndex        =   3
            Top             =   0
            Width           =   1200
         End
         Begin VB.CheckBox chk其他期效 
            Caption         =   "其他长嘱"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   0
            Left            =   750
            TabIndex        =   2
            Top             =   0
            Width           =   1300
         End
      End
      Begin VB.Frame fra状态 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   44
         Top             =   960
         Width           =   3490
         Begin VB.OptionButton opt状态 
            Caption         =   "已校对"
            Height          =   180
            Index           =   1
            Left            =   1660
            TabIndex        =   5
            Top             =   0
            Width           =   900
         End
         Begin VB.OptionButton opt状态 
            BackColor       =   &H00D0FFFF&
            Caption         =   "新开"
            Height          =   180
            Index           =   0
            Left            =   750
            TabIndex        =   4
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton opt状态 
            Caption         =   "全部"
            Height          =   180
            Index           =   2
            Left            =   2780
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.Label lbl状态 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "状态"
            Height          =   180
            Left            =   350
            TabIndex        =   45
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.Frame fra婴儿 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   5835
         Visible         =   0   'False
         Width           =   3210
         Begin VB.OptionButton opt婴儿 
            Caption         =   "婴儿医嘱"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   19
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt婴儿 
            Caption         =   "所有医嘱"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt婴儿 
            Caption         =   "病人医嘱"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   18
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk加班加价 
         Caption         =   "医嘱发送所产生的费用执行加班加价"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   6105
         Width           =   3180
      End
      Begin VB.Frame fraPati 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   3360
         Left            =   60
         TabIndex        =   42
         Top             =   2415
         Width           =   3495
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   765
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   0
            Width           =   2715
         End
         Begin VB.CommandButton cmdNoPati 
            Caption         =   "全清"
            Height          =   370
            Left            =   2820
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   2940
            Width           =   675
         End
         Begin VB.CommandButton cmdAllPati 
            Caption         =   "全选"
            Height          =   370
            Left            =   2115
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   2940
            Width           =   675
         End
         Begin VB.CommandButton cmdQuick 
            Caption         =   "排开欠费病人"
            Height          =   370
            Left            =   30
            TabIndex        =   14
            Top             =   2940
            Width           =   1380
         End
         Begin MSComctlLib.ListView lvwPati 
            Height          =   2490
            Left            =   15
            TabIndex        =   13
            Top             =   375
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   4392
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "姓名"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "住院号"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "床号"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "剩余款"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "住院医师"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "费别"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "护理等级"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "科室"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "入院日期"
               Object.Width           =   2857
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "病人类型"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "留观号"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lbl病区 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病区病人"
            Height          =   180
            Left            =   15
            TabIndex        =   43
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5000
         Y1              =   5790
         Y2              =   5790
      End
   End
   Begin VB.PictureBox pic类别 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   105
      ScaleHeight     =   915
      ScaleWidth      =   3630
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6510
      Width           =   3630
      Begin VB.ListBox lstClass 
         Columns         =   4
         Height          =   900
         ItemData        =   "frmAdviceSendALL.frx":1CE5
         Left            =   0
         List            =   "frmAdviceSendALL.frx":1CE7
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   0
         Width           =   3450
      End
   End
   Begin VB.PictureBox pic给药 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   5430
      ScaleHeight     =   2280
      ScaleWidth      =   3630
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1485
      Width           =   3630
      Begin VB.CommandButton cmdUnWay 
         Caption         =   "反选"
         Height          =   330
         Left            =   1440
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   1920
         Width           =   675
      End
      Begin VB.CommandButton cmdNoWay 
         Caption         =   "全清"
         Height          =   330
         Left            =   2805
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   1920
         Width           =   675
      End
      Begin VB.CommandButton cmdAllWay 
         Caption         =   "全选"
         Height          =   330
         Left            =   2130
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   1920
         Width           =   675
      End
      Begin MSComctlLib.ListView lvwWay 
         Height          =   1905
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3360
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "给药途径"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Left            =   4065
      MousePointer    =   9  'Size W E
      TabIndex        =   35
      Top             =   870
      Width           =   45
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7350
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   7665
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2175
      TabIndex        =   30
      Top             =   7620
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   7545
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceSendALL.frx":1CE9
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17489
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
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
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   4155
      MousePointer    =   7  'Size N S
      TabIndex        =   33
      Top             =   5910
      Width           =   7530
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1425
      Left            =   4155
      TabIndex        =   28
      Top             =   6030
      Width           =   7545
      _cx             =   13309
      _cy             =   2514
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceSendALL.frx":257D
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4215
      TabIndex        =   31
      Top             =   525
      Width           =   7425
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4980
      Left            =   4155
      TabIndex        =   27
      Top             =   825
      Width           =   7530
      _cx             =   13282
      _cy             =   8784
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
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceSendALL.frx":2618
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   7275
      Left            =   45
      TabIndex        =   36
      Top             =   180
      Width           =   3900
      _Version        =   589884
      _ExtentX        =   6879
      _ExtentY        =   12832
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceSendALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMainPrivs As String 'IN
Private mlng病区ID As Long 'IN:用于记录主界面的病区及上次发送病区，选择已转出病人时，是原病区ID
Private mlng病人ID As Long 'IN
Private mlng主页ID As Long 'IN,单病人调用时传入
Private mblnSend As Boolean 'OUT:是否成功发送过。
Private mblnRefresh As Boolean 'OUT：发送后是否要求刷新
Private mblnOnePati As Boolean     '单病人还是多病人模式

'----------------------------------------------
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库房的出库检查方式
Private mrs药房 As ADODB.Recordset

Private mrsBill As ADODB.Recordset
Private mrsWarn As ADODB.Recordset
Private mrsPrice As ADODB.Recordset '包含计价关系

'条件相关变量，根据条件读取医嘱后要使用
Private mstrEnd As String '本次发送的结束时间
Private mblnLimit As Boolean '本次发送给药途径计算是否以结束时间限制

Private mlngNOSequence As Long
Private mlng药品类别ID As Long '药品入出类别ID
Private mlng卫材类别ID As Long
Private mbln领药号 As Boolean
Private mstr领药号 As String
Private mstrAutoExe As String '本科执行自动完成
Private mbln医技后续 As Boolean
Private mint简码 As Integer
Private mstrLike As String
Private mstrRollNotify As String
Private mblnAutoVerify As Boolean   '发送之前自动校对（包括读取未校对的医嘱）
Private mblnChangeIF As Boolean     '是否改变了关键条件，须重新读取医嘱
Private mblnCheckAdvice As Boolean  '发送转科出院医嘱前检查未校对或未发送临嘱
Private mlngRefModld As Long        '0常规医嘱，1=备用医嘱
Private mobjCustom As CommandBarControlCustom
Private mobjlblMsg As CommandBarControl
Private mblnIsSendInterval As Boolean '判断护士是否有权限修改 发送时间间隔
Private mdatCurr As Date
Private mstrInfDepIDs As String  '在输液配置中心发药的病人病区
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mblnSendMsg As Boolean '消息平台是不是可用
Private mstr配液给药途径 As String '输液配置中心参数，配液的给药途径
Private mint输液配置期效 As Integer '启用输液配置中心的医嘱期效
Private mbln药房置换 As Boolean '是否进行了药房置换
Private mstrEndPoint As String '长嘱口服药发送结束时间
Private mbln检验单独产生单据 As Boolean  '检验医嘱单独产生单据
Private mint住院领药部门 As Integer   '0-给药执行科室；1-病人病区
Private mstrNoneIDs As String
Private mbln阳性用药 As Boolean  '皮试阳性用药 参数，当启用此参数后不用判断皮试结果，但需要填写皮试阳性用药说明
Private mstrAdDrugIDs As String '需进一步添加阳性说明的药品行医嘱ID串儿
Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级
Private mbln记帐提醒忽略 As Boolean
Private mintBnt As Integer  '临床批量记帐判断：-1表示未曾选择过.=vbCancel时表示中断，=vbIgnore时表示继续

Private Enum COND_GROUP_IDX
    GP_发送条件 = 0
    GP_给药途径 = 1
    GP_药房置换 = 2
    GP_诊疗类别 = 3
End Enum
'----------------------------------------------
Private Enum COL_ADVICE
    COL_选择 = 0
    COL_科室 = 1
    COL_姓名 = 2
    COL_住院号 = 3
    COL_床号 = 4
    COL_费别 = 5
    COL_婴儿 = 6
    COL_医嘱期效 = 7
    col_医嘱内容 = 8
    COL_规格 = 9
    COL_总量 = 10
    COL_总量单位 = 11
    COL_单量 = 12
    COL_单量单位 = 13
    COL_金额 = 14
    COL_频率 = 15
    COL_用法 = 16 '###
    COL_医生嘱托 = 17 'Data用于存放摘要(医保)
    COL_执行时间 = 18   '执行时间方案，Data中存长嘱的开始执行时间
    COL_首次时间 = 19
    COL_末次时间 = 20
    COL_执行科室 = 21
    COL_附加执行 = 22
    COL_执行性质 = 23
    COL_病人ID = 24 '隐藏列
    COL_主页ID = 25
    col_性别 = 26
    COL_年龄 = 27
    COL_险类 = 28
    COL_ID = 29
    COL_相关ID = 30
    COL_病人病区ID = 31
    COL_病人科室ID = 32
    COL_开嘱科室ID = 33
    COL_开嘱医生 = 34
    COL_诊疗类别 = 35
    COL_诊疗项目ID = 36
    COL_计价特性 = 37
    COL_执行性质ID = 38
    COL_执行科室ID = 39
    COL_执行标记 = 40
    COL_收费细目ID = 41
    COL_剂量系数 = 42
    COL_住院包装 = 43
    COL_住院单位 = 44
    COL_可否分零 = 45
    COL_药房分批 = 46 '###
    COL_是否变价 = 47
    COL_库存 = 48 '###
    COL_次数 = 49
    COL_分解时间 = 50
    COL_操作类型 = 51 '其他医嘱专用
    COL_试管编码 = 52
    COL_标本部位 = 53
    COL_检查方法 = 54
    COL_跟踪在用 = 55
    COL_紧急标志 = 56
    COL_医嘱状态 = 57
    COL_执行频率 = 58
    COL_新开操作时间 = 59
    COL_计算方式 = 60
    COL_执行安排 = 61
    COL_病人性质 = 62
    COL_开嘱时间 = 63
    COL_开始时间 = 64
    COL_申请序号 = 65
    COL_执行分类 = 66
    COL_毒理分类 = 67
    COL_用药理由 = 68
    COL_会诊医嘱ID = 69
End Enum
'-------------------------------------------------
Private Enum COL_PRICE
    COLP_行号 = 0
    COLP_收费细目ID = 1
    COLP_固定 = 2
    COLP_变价 = 3
    COLP_计价医嘱 = 4 '可见列
    COLP_类别 = 5
    COLP_收费项目 = 6
    COLP_计价数量 = 7
    COLP_付数 = 8
    COLP_数量 = 9
    COLP_单位 = 10
    COLP_单价 = 11
    COLP_应收金额 = 12
    COLP_实收金额 = 13
    COLP_执行科室 = 14
    COLP_费用类型 = 15
    COLP_从项 = 16
    COLP_收费方式 = 17
    COLP_收费类别 = 18 '隐藏列
    COLP_执行科室ID = 19
    COLP_跟踪在用 = 20
    COLP_费用性质 = 21
End Enum

Private Const BackColorNew = &HD0FFFF   '浅黄色

Public Function ShowMe(frmParent As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strMainPrivs As String, _
    blnRefresh As Boolean, blnOnePati As Boolean, Optional ByVal lng医护科室ID As Long, Optional ByVal lng婴儿病区ID As Long, Optional ByRef objMip As Object) As Boolean
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    If lng婴儿病区ID <> 0 Then
        If lng婴儿病区ID = lng医护科室ID Then
            mlng病区ID = lng婴儿病区ID
        End If
    End If
    mMainPrivs = strMainPrivs
    mblnOnePati = blnOnePati
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    If mclsMipModule Is Nothing Then
        mblnSendMsg = False
    ElseIf mclsMipModule.IsConnect Then
        mblnSendMsg = True
    Else
        mblnSendMsg = False
    End If
    
    On Error Resume Next
    Me.Show 1, frmParent
    
    blnRefresh = mblnRefresh
    ShowMe = mblnSend
End Function

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub cboDrugType_Click()
    mblnChangeIF = True
End Sub

Private Sub cboTime_Click()
    Dim strTmp As String, lngTmp As Long
    
    dtpEnd.Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    mdatCurr = zlDatabase.Currentdate
    Select Case cboTime.ListIndex
    Case 0 '今天
        dtpEnd.value = Format(mdatCurr, "yyyy-MM-dd 23:59:59")
    Case 1 '明天
        dtpEnd.value = Format(mdatCurr + 1, "yyyy-MM-dd 23:59:59")
    Case 2 '后天
        dtpEnd.value = Format(mdatCurr + 2, "yyyy-MM-dd 23:59:59")
    Case 3 '[指定..]
        strTmp = zlDatabase.GetPara("发送结束时点", glngSys, p住院医嘱发送, "23:59:59", Array(dtpEnd))
        lngTmp = Val(zlDatabase.GetPara("发送时间间隔", glngSys, p住院医嘱发送, "0", Array(dtpEnd)))
        mblnIsSendInterval = dtpEnd.Enabled
        dtpEnd.value = Format(mdatCurr + lngTmp, "yyyy-MM-dd " & strTmp)
        If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
    End Select
    If mstrEndPoint <> "" And lblEatDrug.Visible Then
        lblEatDrug.Caption = "长嘱口服药结束时间：" & Format(dtpEnd.value, "yyyy-MM-dd") & " " & Mid(mstrEndPoint, 1, 5)
    End If
    mblnChangeIF = True
End Sub

Private Sub chkSendNow_Click()
    mblnChangeIF = True
End Sub

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long
    Dim lngUnitID As Long
    Dim lngColor As Long
        
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    If DeptIsWoman(0, Get科室IDs(lngUnitID)) Then
        fra婴儿.Visible = True
        '医嘱处理范围
        opt婴儿(Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))).value = True
    Else
        fra婴儿.Visible = False
        opt婴儿(0).value = True
    End If
    strSQL = "Select 适用病人,报警方法,报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线 Where 病区ID=[1]"
    Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUnitID)
        
    If Not mblnOnePati Then
        str病人IDs = zlDatabase.GetPara("发送病人", glngSys, p住院医嘱发送)
       
        If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
            lng病区ID = Val(Split(str病人IDs, ":")(0))
            str病人IDs = Split(str病人IDs, ":")(1)
        End If
    End If
        
    If Me.Visible Then
        Set rsTmp = GetPatiRsByUnit(lngUnitID, 0, True, True, False)
    Else
        Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng病人ID, True, True, False)
    End If
    
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!审核标志 & "") < 1 Or gbyt病人审核方式 <> 1 Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!姓名)
            objItem.SubItems(1) = IIF(IsNull(rsTmp!住院号), "", rsTmp!住院号)
            objItem.SubItems(2) = IIF(IsNull(rsTmp!床号), "", rsTmp!床号)
            objItem.SubItems(3) = Format(NVL(rsTmp!剩余款, 0), "0.00")
            objItem.SubItems(4) = IIF(IsNull(rsTmp!住院医师), "", rsTmp!住院医师)
            objItem.SubItems(5) = IIF(IsNull(rsTmp!费别), "", rsTmp!费别)
            objItem.SubItems(6) = IIF(IsNull(rsTmp!护理等级), "", rsTmp!护理等级)
            objItem.SubItems(7) = IIF(IsNull(rsTmp!科室), "", rsTmp!科室)
            objItem.SubItems(8) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
            objItem.SubItems(9) = NVL(rsTmp!病人类型)
            objItem.SubItems(10) = NVL(rsTmp!留观号)
        
            '附加信息
            objItem.ListSubItems(1).Tag = NVL(rsTmp!适用病人)
            objItem.ListSubItems(2).Tag = NVL(rsTmp!担保额, 0)
            objItem.ListSubItems(3).Tag = NVL(rsTmp!病人状态, 0)
            objItem.ListSubItems(7).Tag = Val("" & rsTmp!科室ID)
            objItem.ListSubItems(9).Tag = Val("" & rsTmp!主页ID)
            
            '病人颜色
            lngColor = zlDatabase.GetPatiColor(NVL(rsTmp!病人类型))
            objItem.ListSubItems(1).ForeColor = lngColor
            objItem.ListSubItems(9).ForeColor = lngColor
            
            '上次是否选择
            If lngUnitID = lng病区ID And str病人IDs <> "" Then
                If str病人IDs = "ALL" _
                    Or Left(str病人IDs, 1) <> "-" And InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 _
                    Or Left(str病人IDs, 1) = "-" And InStr("," & Mid(str病人IDs, 2) & ",", "," & rsTmp!病人ID & ",") = 0 Then
                    objItem.Checked = True
                    If k = 0 Then '为了看到有选择的
                        objItem.EnsureVisible
                        objItem.Selected = True
                        k = 1
                    End If
                End If
            '出院病人和已转出病人通过医嘱提醒进入
            ElseIf rsTmp!病人ID = mlng病人ID Then
                objItem.Checked = True '缺省只选择当前病人
                objItem.EnsureVisible
                objItem.Selected = True
            End If
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecSend()
'功能：调用医嘱发送
'参数：lngModle=0 发送常规医嘱，=1发送备用医嘱
    Dim lng发送号 As Long, i As Long
    Dim objCbo As CommandBarComboBox
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "当前没有可以发送的医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    If mblnChangeIF Then
        MsgBox "医嘱发送的条件已改变，将自动重新读取数据，请检查后再发送。", vbInformation, gstrSysName
        Call RefreshData(mlngRefModld)
        Exit Sub
    End If
    
    '执行发送
    lng发送号 = SendAdvice
    If lng发送号 <> 0 Then
        mblnSend = True
        
        '发送了特殊医嘱时检查并提醒超期收回(自动)停止的医嘱
        If mstrRollNotify <> "" Then
            Call ShowRollNotify(mstrRollNotify)
        End If
        
        '使用了新领药号的处理
        If mstr领药号 <> "" Then
            Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
            i = objCbo.FindItem(mstr领药号)
            If i = 0 Then
                objCbo.AddItem mstr领药号, 2
                objCbo.ListIndex = 2
            End If
        End If
        
        '打印诊疗单据
        Call frmSendBillPrint.ShowMe(lng发送号, 2, Me)
    End If
End Sub

Private Sub FuncExecReStop()
'功能：调用确认停止
    Dim str病人IDs As String, i As Long
    
    If lvwPati.Visible Then
        mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        For i = 1 To lvwPati.ListItems.Count
            If lvwPati.ListItems(i).Checked Then
                str病人IDs = str病人IDs & ";" & Mid(lvwPati.ListItems(i).Key, 2) & "," & Val(lvwPati.ListItems(i).ListSubItems(9).Tag)
            End If
        Next
        If str病人IDs = "" Then
            MsgBox "请至少选择一个需要确认停止医嘱的病人。", vbInformation, gstrSysName
            If lvwPati.Visible And lvwPati.Enabled Then lvwPati.SetFocus: Exit Sub
        Else
            str病人IDs = Mid(str病人IDs, 2)
        End If
    End If
    
    Call frmAdviceOperate.ShowMe(Me, mMainPrivs, 2, mlng病人ID, mlng主页ID, mlng病区ID, 0, True, , , , , str病人IDs = "", str病人IDs, True, , , mlng病区ID, , mclsMipModule)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_View_Refresh  '读取常规医嘱
        If Me.ActiveControl Is dtpEnd Then cboTime.SetFocus
        Call RefreshData(0)
        mobjCustom.Visible = False
        mobjlblMsg.Visible = False
    Case conMenu_View_RefreshSpare  '读取备用医嘱
        If Me.ActiveControl Is dtpEnd Then cboTime.SetFocus
        Call RefreshData(1)
        mobjCustom.Visible = True
        mobjlblMsg.Visible = True
    Case conMenu_Edit_ReStop    '确认停止
        Call FuncExecReStop
    Case conMenu_Edit_Send      '发送
        Call FuncExecSend
        
    Case conMenu_View_Show
        tkpMain.Visible = True
        fraLR.Visible = True
        Call Form_Resize
    Case conMenu_View_Hide
        tkpMain.Visible = False
        fraLR.Visible = False
        Call Form_Resize
    Case conMenu_Edit_SelAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 Then
                    If Not (InStr(mstrNoneIDs, "," & .TextMatrix(i, COL_ID) & ",") > 0 And Not mbln阳性用药) Then
                        If Not (Val(.TextMatrix(i, COL_操作类型)) = 5 And .TextMatrix(i, COL_诊疗类别) = "Z") Then
                            Set .Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                        End If
                    End If
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Edit_ClsAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 Then
                    Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    Me.tkpMain.Left = lngLeft
    Me.tkpMain.Top = lngTop
    Me.tkpMain.Height = lngBottom - lngTop - stbThis.Height
    
    Me.fraLR.Left = lngLeft + tkpMain.Width
    Me.fraLR.Top = lngTop
    Me.fraLR.Height = lngBottom - lngTop - stbThis.Height
    
    If tkpMain.Visible Then
        lngLW = fraLR.Width + tkpMain.Width
    End If
    
    fraInfo.Top = lngTop
    fraInfo.Left = lngLeft + lngLW
    fraInfo.Width = lngRight - lngLeft - lngLW
    
    vsAdvice.Left = lngLeft + lngLW
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = lngRight - lngLeft - lngLW
    vsAdvice.Height = lngBottom - lngTop - fraInfo.Height - vsPrice.Height - fraUD.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = vsAdvice.Left
    fraUD.Width = vsAdvice.Width
    
    vsPrice.Left = vsAdvice.Left
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = vsAdvice.Width
    
    psb.Top = stbThis.Top + Screen.TwipsPerPixelY * 4
    psb.Left = stbThis.Panels(2).Left + Screen.TwipsPerPixelX * 2
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - Screen.TwipsPerPixelX * 7
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
       
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Show
        Control.Visible = Not tkpMain.Visible
    Case conMenu_View_Hide
        Control.Visible = tkpMain.Visible
    Case conMenu_View_Find
        Control.Visible = (chk药品期效(0).value = 1 Or chk药品期效(1).value = 1) And mbln领药号
    Case conMenu_Edit_ReStop
        If InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱确认停止;") = 0 Then Control.Visible = False
    End Select
End Sub

Private Sub chk其他期效_Click(Index As Integer)
    Call SetFaceBy医嘱期效
End Sub

Private Sub chk药品期效_Click(Index As Integer)
    Call SetFaceBy医嘱期效
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

'反选
Private Sub cmdUnWay_Click()
    Dim i As Long
    For i = 1 To lvwWay.ListItems.Count
        lvwWay.ListItems(i).Checked = Not lvwWay.ListItems(i).Checked
    Next
    lvwWay.SetFocus
End Sub

Private Sub cmdAllWay_Click()
    Call SelectLVW(lvwWay, True)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdNoWay_Click()
    Call SelectLVW(lvwWay, False)
    lvwWay.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdQuick_Click()
    Dim i As Long, blnDo As Boolean
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                '只根据累计报警方法进行处理
                mrsWarn.Filter = "报警方法=1 And 适用病人='" & .ListItems(i).ListSubItems(1).Tag & "'"
                If Not mrsWarn.EOF Then
                    blnDo = False
                    Select Case BeSureMode(NVL(mrsWarn!报警标志1), NVL(mrsWarn!报警标志2), NVL(mrsWarn!报警标志3))
                    Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) <= 0
                    Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) <= 0
                    Case 3 '低于报警值禁止记帐
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < NVL(mrsWarn!报警值, 0)
                    End Select
                    If blnDo Then
                        .ListItems(i).Checked = False
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub cmd执行科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    strSQL = _
        " Select 0 as ID,'-' as 编码,'所有科室' as 名称,NULL as 简码 From Dual" & _
        " Union ALL" & _
        " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " Order by 编码"
    vRect = zlControl.GetControlRect(txt执行科室.hwnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "执行科室", , , , , , True, vRect.Left, vRect.Top, txt执行科室.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有可用的科室，请先到部门管理中设置。", vbInformation, gstrSysName
        End If
        txt执行科室.Text = txt执行科室.Tag
        Call zlControl.TxtSelAll(txt执行科室)
    Else
        txt执行科室.Text = rsTmp!名称
        txt执行科室.Tag = rsTmp!名称
        cmd执行科室.Tag = rsTmp!ID
    End If
    txt执行科室.SetFocus
End Sub

Private Sub dkpExecTime_Change()
    Call RefreshData(1)
End Sub

Private Sub dtpEnd_Change()
    mblnChangeIF = True
End Sub

Private Sub dtpEnd_Validate(Cancel As Boolean)
    If mstrEndPoint <> "" And lblEatDrug.Visible Then
        lblEatDrug.Caption = "长嘱口服药结束时间：" & Format(dtpEnd.value, "yyyy-MM-dd") & " " & Mid(mstrEndPoint, 1, 5)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdAllPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdAllWay_Click
        ElseIf Me.ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = True
            Next
            lstClass.ListIndex = j
        Else
            cbsMain.FindControl(, conMenu_Edit_SelAll).Execute
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdNoPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdNoWay_Click
        ElseIf Me.ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = False
            Next
            lstClass.ListIndex = j
        Else
            cbsMain.FindControl(, conMenu_Edit_ClsAll).Execute
        End If
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If cmdQuick.Visible And cmdQuick.Enabled Then Call cmdQuick_Click
    ElseIf KeyCode = 13 Then
        If Not ActiveControl Is vsDept And Not ActiveControl Is vsAdvice _
            And Not ActiveControl Is vsPrice And Not Me.ActiveControl Is txt执行科室 Then
            Call zlcommfun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ActiveControl Is vsDept And Not ActiveControl Is vsAdvice _
            And Not ActiveControl Is vsPrice And Not Me.ActiveControl Is txt执行科室 Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub SetPic基本_BackColor()
    Dim i As Integer
    
    fra药品.BackColor = pic基本.BackColor
    fra其他.BackColor = pic基本.BackColor
    fra状态.BackColor = pic基本.BackColor
    fra婴儿.BackColor = pic基本.BackColor
    For i = 0 To 2
        If i <> 2 Then
            chk药品期效(i).BackColor = pic基本.BackColor
            chk其他期效(i).BackColor = pic基本.BackColor
        End If
        opt状态(i).BackColor = pic基本.BackColor
        opt婴儿(i).BackColor = pic基本.BackColor
    Next
    chkSendNow.BackColor = pic基本.BackColor
    fra医嘱条件.BackColor = pic基本.BackColor
    chk加班加价.BackColor = pic基本.BackColor
    
    fraPati.BackColor = pic基本.BackColor
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlcommfun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "隐藏")
        objControl.IconId = conMenu_View_Show
        objControl.ToolTipText = "隐藏发送条件区域"
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "显示")
        objControl.IconId = conMenu_View_Hide
        objControl.ToolTipText = "显示发送条件区域"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选")
        objControl.BeginGroup = True
        objControl.ToolTipText = "选中所有可以发送的医嘱(Ctrl+A)"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清")
        objControl.ToolTipText = "清除所有已选择发送医嘱的选择状态(Ctrl+R)"
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "读取常规医嘱"): objControl.BeginGroup = True
        objControl.ToolTipText = "根据当前条件读取常规发送医嘱"
        Set objControl = .Add(xtpControlButton, conMenu_View_RefreshSpare, "读取备用医嘱")
        objControl.ToolTipText = "根据当前条件读取备用发送医嘱"
        Set mobjlblMsg = .Add(xtpControlLabel, conMenu_View_RefreshSpare * 100# + 1, "【执行时间】:")
        mobjlblMsg.Visible = False
        Set mobjCustom = .Add(xtpControlCustom, conMenu_View_RefreshSpare * 100# + 2, "")
        mobjCustom.ToolTipText = "请输入该备用医嘱执行的时间。"
        dkpExecTime.value = zlDatabase.Currentdate
        mobjCustom.Handle = dkpExecTime.hwnd
        mobjCustom.Visible = False
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送医嘱"): objControl.BeginGroup = True
        objControl.ToolTipText = "发送所有已选择的医嘱"
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止")
        objControl.ToolTipText = "确认已停止的医嘱"
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyE, conMenu_Edit_Send
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With
    
    '主菜单右侧的领药号
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objCbo = .Add(xtpControlComboBox, conMenu_View_Find, "领药号")
            objCbo.BeginGroup = True
            objCbo.Flags = xtpFlagRightAlign
            objCbo.Style = xtpComboLabel
            objCbo.Width = 200
    End With
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem, blnDo As Boolean, i As Long
    Dim strPar As String
    
    If Not PatiFeeUsable(mlng病人ID, mlng主页ID) Then Unload Me: Exit Sub
    Call InitAdviceTable
    Call InitPriceTable
    fraLR.BackColor = Me.BackColor
    fraUD.BackColor = Me.BackColor
    
    mblnChangeIF = False
    mblnSend = False
    mblnRefresh = False
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    mstrAutoExe = zlDatabase.GetPara("本科执行自动完成", glngSys, p住院医嘱发送)
    mbln医技后续 = Val(zlDatabase.GetPara("医技医嘱后续处理", glngSys, p住院医嘱发送)) <> 0
    mbln领药号 = Val(zlDatabase.GetPara(27, glngSys)) <> 0
    mblnAutoVerify = Val(zlDatabase.GetPara("发送前自动校对", glngSys, p住院医嘱发送, 0)) = 1
    mblnCheckAdvice = Val(zlDatabase.GetPara("特殊医嘱发送前检查未生效医嘱", glngSys, p住院医嘱发送, 0)) = 1
    mblnLimit = Val(zlDatabase.GetPara("药嘱发送限制结束时间", glngSys, p住院医嘱发送, 0)) = 1
    mbln检验单独产生单据 = Val(zlDatabase.GetPara("检验医嘱单独产生单据", glngSys, p住院医嘱发送, "0")) = 1
    mint住院领药部门 = Val(zlDatabase.GetPara("住院领药部门", glngSys, p住院医嘱发送, "0"))
    mstrInfDepIDs = zlDatabase.GetPara("来源病区", glngSys, p输液配置中心, "")
    mstr配液给药途径 = zlDatabase.GetPara("输液给药途径", glngSys, p输液配置中心)
    mint输液配置期效 = Val(zlDatabase.GetPara("医嘱类型", glngSys, p输液配置中心, "1")) - 1
    
    mstrEndPoint = zlDatabase.GetPara("长嘱口服药发送结束时间", glngSys, p住院医嘱发送, "0")
    If Mid(mstrEndPoint, 1, 1) = 1 And InStr(mstrEndPoint, "|") > 0 Then
        mstrEndPoint = Split(mstrEndPoint, "|")(1)
    Else
        mstrEndPoint = ""
    End If
     
    mbln阳性用药 = Val(zlDatabase.GetPara("皮试阳性用药", glngSys, p住院医嘱下达)) <> 0
    
    If DeptIsWoman(0, Get科室IDs(mlng病区ID)) Then
        fra婴儿.Visible = True
        '医嘱处理范围
        opt婴儿(Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))).value = True
    End If
    
    '单病人模式
    '-----------------------------------------
    If mblnOnePati Then '不显示病人，必须在加分组之前改变pic基本的高度
        fraPati.Visible = False
        pic基本.Height = pic基本.Height - fraPati.Height + 60
    End If
    
    Call InitCommandBar
    
    '分组控件------------------------------------------
    'tkpMain.VisualTheme = xtpTaskPanelThemeListViewOffice2003
    
    Call tkpMain.SetMargins(0, 0, 0, 0, 0)
    Call tkpMain.SetItemInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetItemOuterMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupOuterMargins(3, 5, 3, 0)
        
    Set objGroup = tkpMain.Groups.Add(GP_发送条件, "发送条件")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic基本
    pic基本.BackColor = objItem.BackColor
    Call SetPic基本_BackColor
    
    Set objGroup = tkpMain.Groups.Add(GP_给药途径, "给药途径")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic给药
    pic给药.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(GP_药房置换, "药房置换")
    objGroup.Expanded = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic药房
    pic药房.BackColor = objItem.BackColor
    chkDeptExchange.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(GP_诊疗类别, "诊疗类别")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic类别
    pic类别.BackColor = objItem.BackColor
    
        
        
    '初始读取一些数据---------------------------------
    '各个库房药品出库检查方式,包括发料部门
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    Call Init药房置换
    
    '药品卫材类别
    mlng卫材类别ID = ExistIOClass(41) '不能确定是否使用了卫材收费,后面再判断
    mlng药品类别ID = ExistIOClass(9)
    If mlng药品类别ID = 0 Then
        MsgBox "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    '初始化参数数据-----------------------------------
    cboDrugType.AddItem "0-全部"
    cboDrugType.AddItem "1-毒品类"
    cboDrugType.AddItem "2-麻醉和精神I类"
    cboDrugType.AddItem "3-其它(非1和2类)"
    cboDrugType.ListIndex = 0
    
    '缺省结束时间
    cboTime.AddItem "1-今天"
    cboTime.AddItem "2-明天"
    cboTime.AddItem "3-后天"
    cboTime.AddItem "4-指定"
    
    '读取可用药房
    Call Load药房
    
    '缺省执行科室
    txt执行科室.Text = "所有科室"
    txt执行科室.Tag = txt执行科室.Text
    cmd执行科室.Tag = ""
        
    '病区/病人
    Call InitUnits
                        
    '给药途径
    Call Load给药途径
    
    '药房置换
    Call Show药房
        
    '诊疗类别
    Call Load诊疗类别
        
    '可以送医嘱类型--------------------------------
    Call Init医嘱主要条件
    
    
    '不校对即发送模式（双击医嘱提醒时调用）
    If mblnAutoVerify Then
        If mblnOnePati Then
            '如果病人存在以下特殊医嘱，则自动弹出校对窗体:持续护理等级,病重/危医嘱,术前术后医嘱不发送,记录入出量,转科，出院，转院，死亡
            blnDo = CheckSpecialAdvice(mlng病人ID, mlng主页ID)
            If blnDo Then
                mblnRefresh = True  '护理等级或病况校对后需要刷新病人清单
                If frmAdviceOperate.ShowMe(Me, mMainPrivs, 3, mlng病人ID, mlng主页ID, mlng病区ID, 0, True, False, 0, , , mblnOnePati, , , , , mlng病区ID, True, mclsMipModule) = False Then
                    '取消退出
                    Unload Me: Exit Sub
                End If
            End If
            
            Call RefreshData(0)
            
            If blnDo Then   '不存在特殊医嘱时，即使没有数据，也不退出，因为可能是之前保存的条件值引起的无数据
                If vsAdvice.Rows = vsAdvice.FixedRows + 1 Then
                    If Val("" & vsAdvice.TextMatrix(vsAdvice.Rows - 1, COL_ID)) = 0 Then
                        mblnRefresh = True
                        Unload Me: Exit Sub
                    End If
                End If
            End If
        Else
            For i = 1 To lvwPati.ListItems.Count
                If lvwPati.ListItems(i).Checked Then
                    Exit For
                End If
            Next
            If i <= lvwPati.ListItems.Count Then
                '自动读取新开医嘱
                Call RefreshData(0)
            End If
        End If
    End If

    '-------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
    
End Sub

Private Function Load诊疗类别() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str类别s As String
    
    On Error GoTo errH
    
    str类别s = zlDatabase.GetPara("其他发送诊疗类别", glngSys, p住院医嘱发送, "", Array(lstClass))
    
    strSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('5','6','7','8','9') Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstClass.AddItem rsTmp!名称
        lstClass.ItemData(lstClass.NewIndex) = Asc(rsTmp!编码)
        If str类别s <> "" Then
            If InStr(str类别s, "'" & rsTmp!编码 & "'") > 0 Then
                lstClass.Selected(lstClass.NewIndex) = True
            End If
        Else
            lstClass.Selected(lstClass.NewIndex) = True
        End If
        rsTmp.MoveNext
    Next
    If lstClass.ListCount > 0 Then lstClass.ListIndex = 0
    Load诊疗类别 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mMainPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病区ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Load给药途径() As Boolean
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str给药IDs As String
    Dim rs给药途径 As ADODB.Recordset
    
    On Error GoTo errH
    str给药IDs = zlDatabase.GetPara("药嘱发送给药途径", glngSys, p住院医嘱发送, "", Array(lvwWay))
    
    Set rs给药途径 = New ADODB.Recordset
    strSQL = "Select ID,编码,名称,执行分类 From 诊疗项目目录 Where 类别='E' And 操作类型='2' And 服务对象 in(2,3) And (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 编码"
    
    Call zlDatabase.OpenRecordset(rs给药途径, strSQL, Me.Caption)
        
    With rs给药途径
        lvwWay.ListItems.Clear
        For i = 1 To .RecordCount
            Set objItem = lvwWay.ListItems.Add(, "_" & !ID, !编码 & "-" & !名称)
            
            If str给药IDs <> "" Then
                If InStr("," & str给药IDs & ",", "," & !ID & ",") > 0 Then
                    objItem.Checked = True
                End If
            Else
                objItem.Checked = True
            End If
            .MoveNext
        Next
    End With
    
    Load给药途径 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Show药房()
    Dim strTmp As String, i As Long, j As Long
    Dim str置换 As String, arr置换 As Variant
    Dim lng原药房 As Long
    
    str置换 = zlDatabase.GetPara("药嘱发送药房置换", glngSys, p住院医嘱发送, "", Array(vsDept))
    arr置换 = Split(str置换, ",")
    
    If UBound(arr置换) >= 0 Then
        If arr置换(0) Like "启用-*" Then
            str置换 = Replace(str置换, CStr(arr置换(0)), "")
            If Mid(str置换, 1, 1) = "," Then str置换 = Mid(str置换, 2)
            arr置换 = Split(str置换, ",")
        End If
    End If
    
    mrs药房.Filter = 0
    If Not mrs药房.EOF Then
        vsDept.Rows = vsDept.FixedRows + mrs药房.RecordCount
        For i = 1 To mrs药房.RecordCount
            vsDept.Cell(flexcpData, i, 0) = CLng(mrs药房!ID)
            vsDept.TextMatrix(i, 0) = mrs药房!编码 & "-" & mrs药房!名称
            strTmp = strTmp & "|#" & mrs药房!ID & ";" & mrs药房!编码 & "-" & mrs药房!名称
            
            mrs药房.MoveNext
        Next
        
        
        For i = 1 To vsDept.Rows - 1
            lng原药房 = CLng(vsDept.Cell(flexcpData, i, 0))
            mrs药房.Filter = "ID=" & lng原药房
            For j = 0 To UBound(arr置换)
                If arr置换(j) Like mrs药房!ID & "-*" Then Exit For
            Next
            If j <= UBound(arr置换) Then
                mrs药房.Filter = "ID=" & Val(Split(arr置换(j), "-")(1))
                If Not mrs药房.EOF Then
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrs药房!ID)
                    mrs药房.Filter = "ID=" & lng原药房
                    mrs药房!现ID = CLng(vsDept.Cell(flexcpData, i, 1))
                    mrs药房.Update
                Else
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrs药房!现ID)
                End If
            Else
                vsDept.Cell(flexcpData, i, 1) = CLng(mrs药房!现ID)
            End If
            
            mrs药房.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 1))
            vsDept.TextMatrix(i, 1) = mrs药房!编码 & "-" & mrs药房!名称
            vsDept.Cell(flexcpBackColor, i, 1) = COLEditBackColor
        Next
        If strTmp <> "" Then vsDept.ColComboList(1) = Mid(strTmp, 2)
    Else
        vsDept.Rows = vsDept.FixedRows + 1
        vsDept.Editable = flexEDNone
    End If
    vsDept.Row = vsDept.FixedRows: vsDept.Col = 1
    Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
    chkDeptExchange.Enabled = InStr(GetInsidePrivs(p住院医嘱发送), ";允许置换药房;") > 0
    pic药房.Enabled = chkDeptExchange.Enabled
End Sub

Private Function Load药房() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    cbo药房.Clear
    cbo药房.AddItem "所有药房"
    cbo药房.ListIndex = 0
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cbo药房.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo药房.ItemData(cbo药房.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    Load药房 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function TheStockCheck(ByVal lng库房ID As Long, ByVal str类别 As String) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    If InStr(",5,6,7,", str类别) > 0 Then
        intStyle = mcolStock1("_" & lng库房ID)
    ElseIf str类别 = "4" Then
        intStyle = mcolStock2("_" & lng库房ID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Function Init药房置换() As Boolean
'功能：'初始读取一些数据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    '读取可用药房到集合中:用于药房置换
    Set mrs药房 = New ADODB.Recordset
    mrs药房.Fields.Append "ID", adBigInt
    mrs药房.Fields.Append "编码", adVarChar, 100
    mrs药房.Fields.Append "名称", adVarChar, 200
    mrs药房.Fields.Append "现ID", adBigInt
    mrs药房.CursorLocation = adUseClient
    mrs药房.LockType = adLockOptimistic
    mrs药房.CursorType = adOpenStatic
    mrs药房.Open
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房')" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        mrs药房.AddNew
        mrs药房!ID = rsTmp!ID
        mrs药房!编码 = rsTmp!编码
        mrs药房!名称 = rsTmp!名称
        mrs药房!现ID = rsTmp!ID
        mrs药房.Update
        rsTmp.MoveNext
    Next
    
    Init药房置换 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '释放私有及IN变量
    mMainPrivs = ""
    mlng病区ID = 0
    mlng病人ID = 0
    mstrInfDepIDs = ""
    mstrEnd = ""
    mblnLimit = False
    
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    Set mrs药房 = Nothing
    Set mrsBill = Nothing
    Set mrsWarn = Nothing
    Set mcolStock1 = Nothing
    Set mcolStock2 = Nothing
    Set mobjCustom = Nothing
    Set mobjlblMsg = Nothing
    Set mclsMipModule = Nothing
    gbln加班加价 = False
End Sub

Private Sub Refresh领药号()
    Dim objCbo As CommandBarComboBox
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strPre As String
    
    On Error GoTo errH
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    
    If objCbo.ListIndex > 0 Then strPre = objCbo.List(objCbo.ListIndex)
    
    objCbo.Clear
    objCbo.AddItem "<使用新的领药号>"
    objCbo.ListIndex = 1
    
    strSQL = "Select Distinct 领药号 From 未发药品记录 Where 填制日期>=Trunc(Sysdate) And 单据=9 And 对方部门ID=[1] And 领药号 is Not NULL Order by 领药号 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病区ID)
    Do While Not rsTmp.EOF
        objCbo.AddItem rsTmp!领药号
        If rsTmp!领药号 = strPre Then
            objCbo.ListIndex = objCbo.ListCount
        End If
        rsTmp.MoveNext
    Loop

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get领药号() As String
    Dim objCbo As CommandBarComboBox
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    If objCbo.ListIndex = 1 Then
        Get领药号 = zlDatabase.GetNextNo(122, mlng病区ID)
    ElseIf objCbo.ListIndex > 1 Then
        Get领药号 = objCbo.List(objCbo.ListIndex)
    End If
End Function

Private Sub RefreshData(ByVal lngRefModld As Long)
'功能：重置发送条件
'参数：lngModle=0 读取常规医嘱，=1读取备用医嘱
    Dim strSel As String, strUnSel As String
    Dim str病人IDs, str主页IDs As String, str给药IDs As String
    Dim str类别s As String, strTmp As String
    Dim i As Long
    Dim str病人科室IDs As String
    
    '检查和获取条件
    '--------------------------------------------------------------------------------
    mlngRefModld = lngRefModld
    If cboUnit.ListIndex = -1 Then
        MsgBox "请选择一个病区。", vbInformation, gstrSysName
        If cboUnit.Visible Then cboUnit.SetFocus: Exit Sub
    End If
    If chk药品期效(0).value = 0 And chk药品期效(1).value = 0 And chk其他期效(0).value = 0 And chk其他期效(1).value = 0 Then
        MsgBox "请选择一种医嘱期效。", vbInformation, gstrSysName
        Exit Sub
        '可能没有权限，某些期效不可见，或不可设置
    End If
    
    
    '期效与时间
    If chk药品期效(0).value = 1 Or chk其他期效(0).value = 1 Then
        mstrEnd = Format(dtpEnd.value, "yyyy-MM-dd HH:mm:ss")
    Else
        mstrEnd = ""
    End If
    
    '住院病人
    mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    str病人IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            If Val(lvwPati.ListItems(i).ListSubItems(3).Tag & "") = ps预出 Or Val(lvwPati.ListItems(i).ListSubItems(3).Tag & "") = ps出院 Then
                Call MsgBox("病人""" & lvwPati.ListItems(i) & """已" & IIF(Val(lvwPati.ListItems(i).ListSubItems(3).Tag & "") = ps预出, "预", "") & "出院，不允许进行医嘱发送！", vbInformation, gstrSysName)
                lvwPati.ListItems(i).Checked = False
                Exit Sub
            End If
            str病人IDs = str病人IDs & "," & Mid(lvwPati.ListItems(i).Key, 2)
            strSel = strSel & "," & Mid(lvwPati.ListItems(i).Key, 2)
            str主页IDs = str主页IDs & "," & lvwPati.ListItems(i).ListSubItems(9).Tag
            str病人科室IDs = str病人科室IDs & "," & lvwPati.ListItems(i).ListSubItems(7).Tag
        Else
            strUnSel = strUnSel & "," & Mid(lvwPati.ListItems(i).Key, 2)
        End If
    Next
    str病人IDs = Mid(str病人IDs, 2)
    str主页IDs = Mid(str主页IDs, 2)
    str病人科室IDs = Mid(str病人科室IDs, 2)
    If str病人IDs = "" Then
        MsgBox "请至少选择一个需要发送医嘱病人。", vbInformation, gstrSysName
        If lvwPati.Visible And lvwPati.Enabled Then lvwPati.SetFocus: Exit Sub
    End If
    
    '校对特殊医嘱
    If mblnAutoVerify Then
        strTmp = str病人IDs
        If CheckSpecialAdvice(0, 0, strTmp) Then
             '护理等级或病况校对后需要刷新病人清单
            mblnRefresh = True
            If frmAdviceOperate.ShowMe(Me, mMainPrivs, 3, mlng病人ID, mlng主页ID, mlng病区ID, _
                0, True, True, 0, , , mblnOnePati, strTmp, True, , , mlng病区ID, True, mclsMipModule) = False Then
                Exit Sub
            End If
            '弹出自动校对窗口后条件自动选择为:全部
            opt状态(2).value = True
        End If
        strTmp = ""
        '如果是备用医嘱，则默认选择全部
        If lngRefModld = 1 Then
            If Not opt状态(2).value Then opt状态(2).value = True
        End If
    End If
    
    strSel = Mid(strSel, 2)
    strUnSel = Mid(strUnSel, 2)
    If strSel = "" Or (UBound(Split(strSel, ",")) = 0 And Val(strSel) = mlng病人ID) Then
        strSel = ""
    Else
        If strUnSel = "" Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":ALL"
        ElseIf UBound(Split(strSel, ",")) > UBound(Split(strUnSel, ",")) Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":-" & strUnSel
        Else
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":" & strSel
        End If
    End If
    
    If chk药品期效(0).value = 1 Or chk药品期效(1).value = 1 Then
        '给药途径
        str给药IDs = ""
        For i = 1 To lvwWay.ListItems.Count
            If lvwWay.ListItems(i).Checked Then
                str给药IDs = str给药IDs & "," & Mid(lvwWay.ListItems(i).Key, 2)
            End If
        Next
        str给药IDs = Mid(str给药IDs, 2)
        If str给药IDs = "" Then
            MsgBox "请至少选择一种给药途径。", vbInformation, gstrSysName
            If lvwWay.Visible And lvwWay.Enabled Then lvwWay.SetFocus
            Exit Sub
        End If
        If UBound(Split(str给药IDs, ",")) + 1 = lvwWay.ListItems.Count Then
            str给药IDs = ""
        End If
    Else
        '诊疗类别
        str类别s = ""
        For i = 0 To lstClass.ListCount - 1
            If lstClass.Selected(i) Then
                str类别s = str类别s & ",'" & Chr(lstClass.ItemData(i)) & "'"
            End If
        Next
        str类别s = Mid(str类别s, 2)
        If str类别s = "" Then
            MsgBox "请至少选择一种诊疗类别。", vbInformation, gstrSysName
            If lstClass.Visible And lstClass.Enabled Then lstClass.SetFocus
            Exit Sub
        End If
        If UBound(Split(str类别s, ",")) + 1 = lstClass.ListCount Then
            str类别s = ""
        End If
    End If
    
    gbln加班加价 = chk加班加价.value = 1
    
    '读取数据
    '--------------------------------------------------------------------------------
    Call InitPriceRecordset '计价关系表
    
    Call LoadAdviceSend(str病人IDs, str主页IDs, mstrEnd, str给药IDs, str类别s, mlngRefModld, str病人科室IDs)
    mblnChangeIF = False
    
    
    '保存条件值
    '--------------------------------------------------------------------------------
    If Me.Visible Then  '自动读取(医嘱提醒)时不进行保存操作

        Call zlDatabase.SetPara("发送结束时间", cboTime.ListIndex, glngSys, p住院医嘱发送)
        If cboTime.ListIndex = cboTime.ListCount - 1 Then
            Call zlDatabase.SetPara("发送结束时点", Format(dtpEnd.value, "HH:mm:ss"), glngSys, p住院医嘱发送)
            Call zlDatabase.SetPara("发送时间间隔", Int(CDate(Format(dtpEnd.value, "yyyy-MM-dd")) - CDate(Format(mdatCurr, "yyyy-MM-dd"))), glngSys, p住院医嘱发送)
        End If
        
        '单病人模式不保存
        If Not mblnOnePati Then
            Call zlDatabase.SetPara("发送病人", strSel, glngSys, p住院医嘱发送)
        End If
        
        '通过医嘱提醒打开时，只显示新开的，不保存状态，和期效
        If Not (mblnAutoVerify And mblnOnePati) Then
            '医嘱期效
            If chk药品期效(0).value = 1 And chk药品期效(1).value = 1 Then
                strTmp = "2"
            ElseIf chk药品期效(0).value = 1 Then
                strTmp = "0"
            ElseIf chk药品期效(1).value = 1 Then
                strTmp = "1"
            Else
                strTmp = "-1"
            End If
            Call zlDatabase.SetPara("药嘱发送医嘱期效", strTmp, glngSys, p住院医嘱发送)
            
            If chk其他期效(0).value = 1 And chk其他期效(1).value = 1 Then
                strTmp = "2"
            ElseIf chk其他期效(0).value = 1 Then
                strTmp = "0"
            ElseIf chk其他期效(1).value = 1 Then
                strTmp = "1"
            Else
                strTmp = "-1"
            End If
            Call zlDatabase.SetPara("其他发送医嘱期效", strTmp, glngSys, p住院医嘱发送)
        End If
        
        
        If chk药品期效(0).Visible And chk药品期效(0).Enabled Or (chk药品期效(1).Visible And chk药品期效(1).Enabled) Then
            Call zlDatabase.SetPara("药嘱发送给药途径", str给药IDs, glngSys, p住院医嘱发送)
                        
            '药房置换
            mrs药房.Filter = 0
            strTmp = ""
            For i = 1 To mrs药房.RecordCount
                strTmp = strTmp & "," & mrs药房!ID & "-" & mrs药房!现ID
                mrs药房.MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            Call zlDatabase.SetPara("药嘱发送药房置换", strTmp, glngSys, p住院医嘱发送)
        End If
        
        If chk其他期效(0).Visible And chk其他期效(0).Enabled Or (chk其他期效(1).Visible And chk其他期效(1).Enabled) Then
            If chk其他期效(0).value = 1 Or chk其他期效(1).value = 1 Then
                Call zlDatabase.SetPara("其他发送诊疗类别", Replace(str类别s, "'", "''"), glngSys, p住院医嘱发送)
            End If
        End If
        
        If chkSendNow.Enabled Then Call zlDatabase.SetPara("药品长嘱从当天开始发送", chkSendNow.value, glngSys, p住院医嘱发送)
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tkpMain.Width + X < 3000 Or vsAdvice.Width - X < 3000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tkpMain.Width = tkpMain.Width + X
        
        fraInfo.Left = fraInfo.Left + X
        fraInfo.Width = fraInfo.Width - X
        
        vsAdvice.Left = vsAdvice.Left + X
        vsAdvice.Width = vsAdvice.Width - X
        
        vsPrice.Left = vsPrice.Left + X
        vsPrice.Width = vsPrice.Width - X
        
        fraUD.Left = fraUD.Left + X
        fraUD.Width = fraUD.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Top = vsPrice.Top + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, Optional rsUpload As ADODB.Recordset)
'功能：根据可见行的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_选择 Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            '一组申请序号的会诊医嘱
            If Val(.TextMatrix(lngRow, COL_申请序号)) <> 0 And .TextMatrix(lngRow, COL_诊疗类别) = "Z" Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_申请序号)) = Val(.TextMatrix(lngRow, COL_申请序号)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_申请序号)) = Val(.TextMatrix(lngRow, COL_申请序号)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
            End If
            
            '取消选择时
            If Not (.Cell(flexcpData, lngRow, COL_选择) = 0 And Not .Cell(flexcpPicture, lngRow, COL_选择) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_相关ID)))
                '1.清除对应的费用及发送记录填写
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "医嘱ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '因为要使用BookMark，因此恢复
                End If
                '2.清除对应的发送计价数量累计
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "医嘱ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.清除对应的医保上传单据号
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "医嘱ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'功能：根据指定医嘱行，返回该医嘱中可见的行
    Dim lng组ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        '一并给药的定位到第一药品行
        If blnFirst Then
            If .TextMatrix(lngRow, COL_诊疗类别) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 _
                And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_相关ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "C" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_诊疗类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, COL_诊疗类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "7" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'功能：根据当前医嘱行获取可选择的计价医嘱内容
'参数：lngRow=可见行(药疗或非药)
'说明：注意这里是根据具体医嘱在取
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If Val(.Cell(flexcpData, lngRow, COL_诊疗类别)) = 3 Then
            '中药用法：中药用法,中药煎法
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", Val(.Cell(flexcpData, i, COL_诊疗类别))) > 0 Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!固定, 0) = 0 Then
                                    If Val(.Cell(flexcpData, i, COL_诊疗类别)) = 2 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药煎法-" & .Cell(flexcpData, i, col_医嘱内容)
                                    ElseIf Val(.Cell(flexcpData, i, COL_诊疗类别)) = 3 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药用法-" & .Cell(flexcpData, i, col_医嘱内容)
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            If Val(.Cell(flexcpData, i, COL_诊疗类别)) = 2 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药煎法-" & .Cell(flexcpData, i, col_医嘱内容)
                            ElseIf Val(.Cell(flexcpData, i, COL_诊疗类别)) = 3 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药用法-" & .Cell(flexcpData, i, col_医嘱内容)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
            And .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
            '采集方法行
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!固定, 0) = 0 Then
                                If .TextMatrix(i, COL_诊疗类别) = "C" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";检验项目-" & .Cell(flexcpData, i, col_医嘱内容)
                                ElseIf .TextMatrix(i, COL_诊疗类别) = "E" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";采集方法-" & .Cell(flexcpData, i, col_医嘱内容)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        If .TextMatrix(i, COL_诊疗类别) = "C" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";检验项目-" & .Cell(flexcpData, i, col_医嘱内容)
                        ElseIf .TextMatrix(i, COL_诊疗类别) = "E" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";采集方法-" & .Cell(flexcpData, i, col_医嘱内容)
                        End If
                        If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                            strCombo = strCombo & "|#" & strTmp
                        End If
                    End If
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '首行成药：给药途径
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(.TextMatrix(lngRow, COL_相关ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_执行性质ID))) = 0 Then
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!固定, 0) = 0 Then
                                strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";给药途径-" & .Cell(flexcpData, lngTmp, col_医嘱内容)
                                Exit For
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";给药途径-" & .Cell(flexcpData, lngTmp, col_医嘱内容)
                    End If
                End If
            End If
        Else
            '一组手术或检查，或输血医嘱，或独立医嘱
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!固定, 0) = 0 Then
                                    If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";附加手术-" & .Cell(flexcpData, i, col_医嘱内容)
                                    ElseIf .TextMatrix(i, COL_诊疗类别) = "G" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";手术麻醉-" & .Cell(flexcpData, i, col_医嘱内容)
                                    ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";检查部位-" & .TextMatrix(i, COL_标本部位) & "(" & .TextMatrix(i, COL_检查方法) & ")"
                                    ElseIf .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";输血途径-" & .Cell(flexcpData, i, col_医嘱内容)
                                    Else
                                        If mrsPrice!费用性质 <> 0 Then
                                            '加收费用：目前包含检查的床旁和术中加收
                                            lngTmp = -1 * Val(mrsPrice!费用性质 & Val(.TextMatrix(i, COL_ID)))
                                            strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容) & _
                                                "(" & decode(Val(.TextMatrix(i, COL_执行标记)), 1, "床旁", 2, "术中", "") & "加收)"
                                        Else
                                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容)
                                        End If
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            '未设置计价的，可能选择添加计价项目
                            If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";附加手术-" & .Cell(flexcpData, i, col_医嘱内容)
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "G" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";手术麻醉-" & .Cell(flexcpData, i, col_医嘱内容)
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";检查部位-" & .TextMatrix(i, COL_标本部位) & "(" & .TextMatrix(i, COL_检查方法) & ")"
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";输血途径-" & .Cell(flexcpData, i, col_医嘱内容)
                            Else
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                            
                            '加收费用：目前包含检查的床旁或术中加收
                            If .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 _
                                And (Val(.TextMatrix(i, COL_执行标记)) = 1 Or Val(.TextMatrix(i, COL_执行标记)) = 2) Then
                                lngTmp = -1 * Val(1 & Val(.TextMatrix(i, COL_ID)))
                                strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容) & _
                                    "(" & decode(Val(.TextMatrix(i, COL_执行标记)), 1, "床旁", 2, "术中", "") & "加收)"
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'功能：根据医嘱计价关系，计算并显示指定医嘱的费用(整个医嘱，可能多行)
'参数：lngRow=可见行(药疗或非药)
    Dim rsTmp As New ADODB.Recordset
    Dim rsExeDays As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str计价医嘱 As String
    Dim str单位 As String, dbl数量 As Double, int付数 As Integer
    Dim bln附加手术 As Boolean, strCombo As String, str行号 As String, str分解时间 As String
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim dbl当前单价 As Double, dbl当前应收 As Double, cur当前应收 As Currency, cur当前实收 As Currency
    Dim lng行号 As Long, cur合计 As Currency, bln付数 As Boolean
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    Dim strPriceType As String
        
    On Error GoTo errH
    
    '用于汇总计算折扣的临时记录集
    rsMain.Fields.Append "医嘱行号", adBigInt
    rsMain.Fields.Append "费用性质", adInteger
    rsMain.Fields.Append "主项行号", adBigInt
    rsMain.Fields.Append "主收入ID", adBigInt
    rsMain.Fields.Append "医嘱合计", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnFirst = False '一并给药中是否第一药品行
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or 医嘱ID=" & Val(.TextMatrix(lngRow, COL_相关ID))
            Else
                mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or 相关ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '计价医嘱
            bln附加手术 = False
            lng行号 = .FindRow(CStr(mrsPrice!医嘱ID), , COL_ID)
            If .TextMatrix(lng行号, COL_诊疗类别) = "4" Then
                str计价医嘱 = "卫生材料-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf InStr(",5,6,7", .TextMatrix(lng行号, COL_诊疗类别)) > 0 Then
                str计价医嘱 = "药品医嘱-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf Val(.Cell(flexcpData, lng行号, COL_诊疗类别)) = 1 Then
                str计价医嘱 = "给药途径-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf Val(.Cell(flexcpData, lng行号, COL_诊疗类别)) = 2 Then
                str计价医嘱 = "中药煎法-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf Val(.Cell(flexcpData, lng行号, COL_诊疗类别)) = 3 Then
                str计价医嘱 = "中药用法-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "E" And Val(.TextMatrix(lng行号, COL_相关ID)) = 0 _
                And .TextMatrix(lng行号 - 1, COL_诊疗类别) = "C" And Val(.TextMatrix(lng行号 - 1, COL_相关ID)) = Val(.TextMatrix(lng行号, COL_ID)) Then
                str计价医嘱 = "采集方法-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "E" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 _
                And .TextMatrix(lng行号 - 1, COL_诊疗类别) = "K" And Val(.TextMatrix(lng行号 - 1, COL_ID)) = Val(.TextMatrix(lng行号, COL_相关ID)) Then
                str计价医嘱 = "输血途径-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "C" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "检验项目-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "F" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                bln附加手术 = True
                str计价医嘱 = "附加手术-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "G" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "手术麻醉-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "D" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "检查部位-" & .TextMatrix(lng行号, COL_标本部位) & "(" & .TextMatrix(lng行号, COL_检查方法) & ")"
            Else
                If NVL(mrsPrice!费用性质, 0) = 1 Then
                    '床旁或术中加收费用
                    str计价医嘱 = .Cell(flexcpData, lng行号, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, lng行号, col_医嘱内容) & _
                        "(" & decode(Val(.TextMatrix(lng行号, COL_执行标记)), 1, "床旁", 2, "术中", "") & "加收)"
                Else
                    str计价医嘱 = .Cell(flexcpData, lng行号, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, lng行号, col_医嘱内容)
                End If
            End If
            str计价医嘱 = Replace(str计价医嘱, "'", "''")
            
            '数量:药品按住院单位的数量,其它按零售数量
            int付数 = 1
            If InStr(",5,6,", .TextMatrix(lng行号, COL_诊疗类别)) > 0 Then
                dbl数量 = Val(.TextMatrix(lng行号, COL_总量))
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "7" Then
                '中药药房单位按不可分零处理:每付
                int付数 = Val(.TextMatrix(lng行号, COL_总量))
                If Val(.TextMatrix(lng行号, COL_可否分零)) = 0 Then
                    dbl数量 = Val(.TextMatrix(lng行号, COL_单量)) / Val(.TextMatrix(lng行号, COL_剂量系数)) / Val(.TextMatrix(lng行号, COL_住院包装))
                Else
                    dbl数量 = IntEx(Val(.TextMatrix(lng行号, COL_单量)) / Val(.TextMatrix(lng行号, COL_剂量系数)) / Val(.TextMatrix(lng行号, COL_住院包装)))
                End If
            Else
                If InStr(",3,4,5,6,", Val("" & mrsPrice!收费方式)) > 0 Then '一天只收一次的
                     '分解时间
                    If .TextMatrix(lng行号, COL_分解时间) <> "" Then
                        str分解时间 = .TextMatrix(lng行号, COL_分解时间)
                    Else
                        str分解时间 = .Cell(flexcpData, lng行号, COL_分解时间)    '开始执行时间
                    End If
                    
                    Set rsExeDays = GetExecDays(str分解时间)
                    dbl数量 = rsExeDays.RecordCount
                ElseIf InStr(",1,2,", Val("" & mrsPrice!收费方式)) > 0 Then '一次发送只收一次
                    dbl数量 = 1
                Else
                    dbl数量 = Val(.TextMatrix(lng行号, COL_总量))
                End If
            End If
            dbl数量 = Format(dbl数量 * NVL(mrsPrice!数量, 0), "0.00000")
                        
            '组合SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as 序号," & mrsPrice!医嘱ID & " as 医嘱ID,ID," & _
                NVL(mrsPrice!固定, 0) & " as 固定,'" & str计价医嘱 & "' as 计价医嘱,类别,名称,产地,规格," & _
                "计算单位 as 单位," & NVL(mrsPrice!数量, 0) & " as 计价数量," & int付数 & " as 付数," & dbl数量 & " as 数量," & _
                Format(NVL(mrsPrice!单价, 0), gstrDecPrice) & " as 单价,费用类型," & lng行号 & " as 行号," & _
                " 是否变价,加班加价," & IIF(bln附加手术, 1, 0) & " as 附加手术," & mrsPrice!从项 & " as 从项," & _
                NVL(mrsPrice!执行科室ID, 0) & " as 执行科室ID,屏蔽费别," & mrsPrice!费用性质 & " as 费用性质," & _
                mrsPrice!收费方式 & " as 收费方式 From 收费项目目录 Where ID=" & mrsPrice!收费细目ID
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '需要计价的医嘱选择
        '根据待发送医嘱取可计价医嘱(不能从mrsPrice取,因为可能无收费关系或已删除,而且也许现在计价已全部删除)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_计价医嘱) = strCombo
            .Editable = flexEDKbdMouse '可以选择则可以编辑
        Else
            .ColData(COLP_计价医嘱) = ""
        End If
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lngRow, COL_病人ID)), Val(vsAdvice.TextMatrix(lngRow, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        '显示已有的计价项目
        If strSQL <> "" Then
            strSQL = "Select A.行号,A.ID AS 收费细目ID,A.固定,A.从项,A.计价医嘱,A.类别,C.名称 as 类别名称,A.执行科室ID,G.名称 as 执行科室," & _
                " Nvl(E.名称,A.名称)||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 名称," & _
                " A.单位,A.计价数量,A.付数,A.数量,D.住院包装,D.住院单位,Decode(A.是否变价,1,A.单价,B.现价) as 单价,F.跟踪在用," & _
                " A.费用性质,A.收费方式,A.费用类型,A.屏蔽费别,A.是否变价,A.加班加价,B.加班加价率,B.原价,B.现价,A.附加手术,B.附术收费率,B.收入项目ID" & _
                " From (" & strSQL & ") A,收费价目 B,收费项目类别 C,药品规格 D,收费项目别名 E,材料特性 F,部门表 G" & _
                " Where A.ID=B.收费细目ID And A.类别=C.编码 And A.ID=D.药品ID(+)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "1", "2", "3") & _
                " And A.ID=F.材料ID(+) And A.执行科室ID=G.ID(+)" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIF(gbyt药品名称显示 = 0, 1, 3) & _
                " Order by A.序号"
                '因为输入后是调用本函数刷新,要保持动态记录集中记录顺序
                '要保证主项排在前面,LoadAdvicePrice时，主项是排在前面，而且编辑后只可能加了从项
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级) '没法
            
            If Not rsTmp.EOF And gbln从项汇总折扣 Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str行号 <> rsTmp!行号 & "_" & rsTmp!费用性质 & "_" & rsTmp!收费细目ID Then
                    If str行号 <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_变价)) = 1 And dbl单价 = 0) Then
                            .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, gstrDecPrice)
                            .Cell(flexcpData, .Rows - 1, COLP_单价) = .TextMatrix(.Rows - 1, COLP_单价) '记录用于恢复输入
                            .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                        End If
                        cur合计 = cur合计 + Format(cur实收, gstrDec)
                    End If
                    str行号 = rsTmp!行号 & "_" & rsTmp!费用性质 & "_" & rsTmp!收费细目ID
                    dbl单价 = 0: cur应收 = 0: cur实收 = 0
                    .Rows = .Rows + 1
                    
                    '标识固定对照为灰色
                    If rsTmp!固定 <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_行号) = rsTmp!行号
                    .TextMatrix(.Rows - 1, COLP_收费细目ID) = rsTmp!收费细目ID
                    .TextMatrix(.Rows - 1, COLP_固定) = rsTmp!固定
                    .TextMatrix(.Rows - 1, COLP_计价医嘱) = rsTmp!计价医嘱
                    .TextMatrix(.Rows - 1, COLP_费用性质) = rsTmp!费用性质
                    .TextMatrix(.Rows - 1, COLP_收费方式) = getChargeMode(Val(NVL(rsTmp!收费方式, 0)))
                        .Cell(flexcpData, .Rows - 1, COLP_收费方式) = Val(NVL(rsTmp!收费方式, 0))
                    .TextMatrix(.Rows - 1, COLP_类别) = rsTmp!类别名称
                    .TextMatrix(.Rows - 1, COLP_收费类别) = rsTmp!类别
                    .TextMatrix(.Rows - 1, COLP_收费项目) = rsTmp!名称
                    .TextMatrix(.Rows - 1, COLP_计价数量) = NVL(rsTmp!计价数量, 0) '相对数量
                    
                    int付数 = NVL(rsTmp!付数, 1)
                    
                    dbl数量 = NVL(rsTmp!数量, 0) '售价数量用于后面按成本打折计算
                    If InStr(",5,6,7,", rsTmp!类别) > 0 Then '住院包装
                        .TextMatrix(.Rows - 1, COLP_单位) = NVL(rsTmp!住院单位)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(NVL(rsTmp!数量, 0), 5)
                            dbl数量 = dbl数量 * NVL(rsTmp!住院包装, 1)
                        Else
                            '中药药房单位按不可分零处理:每付
                            '非药嘱药品计价:因为这里预定了售价数量,因此转换为药房单位显示时不作不分零处理
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(NVL(rsTmp!数量, 0) / NVL(rsTmp!住院包装, 1), 5)
                        End If
                        
                        If rsTmp!类别 = "7" Then
                            .TextMatrix(.Rows - 1, COLP_付数) = int付数
                            bln付数 = True
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_单位) = NVL(rsTmp!单位)
                        .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(NVL(rsTmp!数量, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_执行科室) = NVL(rsTmp!执行科室)
                    .TextMatrix(.Rows - 1, COLP_执行科室ID) = NVL(rsTmp!执行科室ID, 0)
                    
                    '显示医保费用类型
                    If Val(rsTmp!收费细目ID & "") <> 0 Then
                        strPriceType = GetPriceType(Val(vsAdvice.TextMatrix(lngRow, COL_病人ID)), Val(rsTmp!收费细目ID & ""), Val(vsAdvice.TextMatrix(lngRow, COL_险类)), False)
                    End If
                    '费用类型
                    If strPriceType = "" Then
                        .TextMatrix(.Rows - 1, COLP_费用类型) = NVL(rsTmp!费用类型)
                    Else
                        .TextMatrix(.Rows - 1, COLP_费用类型) = strPriceType
                    End If
                    
                    
                    .TextMatrix(.Rows - 1, COLP_从项) = IIF(NVL(rsTmp!从项, 0) = 0, "", "√")
                    .TextMatrix(.Rows - 1, COLP_跟踪在用) = NVL(rsTmp!跟踪在用, 0)
                    
                    '记录用于输入恢复
                    .Cell(flexcpData, .Rows - 1, COLP_计价医嘱) = .TextMatrix(.Rows - 1, COLP_计价医嘱)
                    .Cell(flexcpData, .Rows - 1, COLP_收费项目) = .TextMatrix(.Rows - 1, COLP_收费项目)
                    .Cell(flexcpData, .Rows - 1, COLP_计价数量) = .TextMatrix(.Rows - 1, COLP_计价数量)
                    .Cell(flexcpData, .Rows - 1, COLP_执行科室) = .TextMatrix(.Rows - 1, COLP_执行科室)
                    
                    '记录从属主项的信息，以便计算
                    If gbln从项汇总折扣 And rsTmp!从项 = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!行号 & "_" & rsTmp!费用性质 & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!行号 & "_" & rsTmp!费用性质 & ",") = 0 Then
                            rsClone.Filter = "行号=" & rsTmp!行号 & " And 费用性质=" & rsTmp!费用性质 & " And 从项=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!医嘱行号 = rsTmp!行号
                                rsMain!费用性质 = rsTmp!费用性质
                                rsMain!主项行号 = .Rows - 1
                                rsMain!主收入ID = rsTmp!收入项目ID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!行号 & "_" & rsTmp!费用性质
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!行号 & "_" & rsTmp!费用性质
                            End If
                        End If
                    End If
                    
                    '非药品、卫材医嘱的药品和跟踪卫材计价：即使固定也可以修改执行科室
                    If InStr(",5,6,7,", rsTmp!类别) > 0 _
                        Or rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '单价计算处理
                If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                    If NVL(rsTmp!是否变价, 0) = 0 Then
                        dbl当前单价 = NVL(rsTmp!单价, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), Format(int付数 * NVL(rsTmp!数量, 0) * NVL(rsTmp!住院包装, 1), gstrDecPrice), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        Else
                            dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), Format(int付数 * NVL(rsTmp!数量, 0), gstrDecPrice), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                        dbl当前单价 = dbl当前单价 * NVL(rsTmp!住院包装, 1)
                        dbl当前应收 = Format(int付数 * NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    Else
                        dbl当前应收 = Format(int付数 * NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                        dbl当前单价 = dbl当前单价 * NVL(rsTmp!住院包装, 1)
                    End If
                ElseIf rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 And NVL(rsTmp!是否变价, 0) = 1 Then
                    '跟踪在用的时价卫材和药品一样计算
                    dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), Format(NVL(rsTmp!数量, 0), "0.00000"), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    dbl当前应收 = Format(NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                Else
                    dbl当前单价 = NVL(rsTmp!单价, 0) '其它如果为变价则是用户输入的
                    dbl当前应收 = Format(NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    If NVL(rsTmp!是否变价, 0) = 1 Then '记录非药变价范围
                        .TextMatrix(.Rows - 1, COLP_变价) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_应收金额) = CCur(NVL(rsTmp!原价, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_实收金额) = CCur(NVL(rsTmp!现价, 0))
                        .Editable = flexEDKbdMouse '非药品变价,即使固定也可以定价
                        .Cell(flexcpBackColor, .Rows - 1, COLP_单价) = COLEditBackColor       '浅绿
                    End If
                End If
                '应收
                If rsTmp!附加手术 = 1 Then
                    dbl当前应收 = dbl当前应收 * NVL(rsTmp!附术收费率, 100) / 100
                End If
                '处理加班加价
                If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                    dbl当前应收 = dbl当前应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                End If
                cur当前应收 = Format(dbl当前应收, gstrDec)
                
                '实收
                If gbln从项汇总折扣 And (rsTmp!从项 = 1 Or InStr(strHaveSub & ",", "," & rsTmp!行号 & "_" & rsTmp!费用性质 & ",") > 0) Then
                    cur当前实收 = Format(cur当前应收, gstrDec)
                    '累计医嘱合计来计算折扣
                    rsMain.Filter = "医嘱行号=" & rsTmp!行号 & " And 费用性质=" & rsTmp!费用性质
                    rsMain!医嘱合计 = NVL(rsMain!医嘱合计, 0) + cur当前实收
                    rsMain.Update
                ElseIf NVL(rsTmp!屏蔽费别, 0) = 0 And vsAdvice.TextMatrix(lngRow, COL_费别) <> "" Then
                    cur当前实收 = Format(ActualMoney(vsAdvice.TextMatrix(lngRow, COL_费别), rsTmp!收入项目ID, cur当前应收, rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), _
                        int付数 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                Else
                    cur当前实收 = Format(cur当前应收, gstrDec)
                End If
                
                dbl单价 = dbl单价 + dbl当前单价
                cur应收 = cur应收 + cur当前应收
                cur实收 = cur实收 + cur当前实收
                
                rsTmp.MoveNext
            Next
            If str行号 <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_变价)) = 1 And dbl单价 = 0) Then
                    .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, gstrDecPrice)
                    .Cell(flexcpData, .Rows - 1, COLP_单价) = .TextMatrix(.Rows - 1, COLP_单价) '记录用于恢复输入
                    .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                End If
                cur合计 = cur合计 + Format(cur实收, gstrDec)
            End If
        End If
        
        '汇总计算折扣
        If gbln从项汇总折扣 And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur当前实收 = Format(ActualMoney(vsAdvice.TextMatrix(lngRow, COL_费别), rsMain!主收入ID, rsMain!医嘱合计), gstrDec)
                cur合计 = cur合计 - Val(.TextMatrix(rsMain!主项行号, COLP_实收金额))
                .TextMatrix(rsMain!主项行号, COLP_实收金额) = Format(Val(.TextMatrix(rsMain!主项行号, COLP_实收金额)) + (cur当前实收 - rsMain!医嘱合计), gstrDec)
                cur合计 = cur合计 + Val(.TextMatrix(rsMain!主项行号, COLP_实收金额))
                rsMain.MoveNext
            Loop
        End If
        
        '付数是否显示
        .ColHidden(COLP_付数) = Not bln付数
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '定位缺省单元
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_计价医嘱 And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_计价医嘱
        End If
        '定位表格输入位置
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_计价医嘱 And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '重新汇总显示可见行的发送医嘱金额
    vsAdvice.TextMatrix(lngRow, COL_金额) = Format(cur合计, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub Init医嘱主要条件()
    Dim lngTmp As Long, strTmp As String
    Dim strPrivs As String
    
     '可以进入这项表示两者至少有一个
    strPrivs = GetInsidePrivs(p住院医嘱发送)
    lblEatDrug.Visible = mstrEndPoint <> ""
    lblDrugType.Visible = gbln特殊药品分开发送: cboDrugType.Visible = gbln特殊药品分开发送
    If InStr(strPrivs, ";发送药疗临嘱;") = 0 And InStr(strPrivs, ";发送药疗长嘱;") = 0 Then
        fra药品.Visible = False
        chk药品期效(0).Enabled = False
        chk药品期效(1).Enabled = False
        lblEatDrug.Visible = False
        lblDrugType.Visible = False: cboDrugType.Visible = False
        cbo药房.Visible = False: lbl药房.Visible = False
        txt执行科室.Top = lbl结束时间.Top + 250
        lbl执行科室.Top = txt执行科室.Top + 50
        cmd执行科室.Top = txt执行科室.Top + 30
        
        fra医嘱条件.Height = txt执行科室.Top + txt执行科室.Height + 50
    ElseIf InStr(strPrivs, ";发送其他临嘱;") = 0 And InStr(strPrivs, ";发送其他长嘱;") = 0 Then
        fra其他.Visible = False
        chk其他期效(0).Enabled = False
        chk其他期效(1).Enabled = False
        cbo药房.Top = lblEatDrug.Top + IIF(lblEatDrug.Visible, 230, 0)
        If gbln特殊药品分开发送 Then
            lblDrugType.Top = cbo药房.Top + 50
            cboDrugType.Top = cbo药房.Top
            cbo药房.Top = cboDrugType.Top + cboDrugType.Height + 10
        End If
        lbl药房.Top = cbo药房.Top + 50
        txt执行科室.Visible = False: lbl执行科室.Visible = False: cmd执行科室.Visible = False
        fra医嘱条件.Height = cbo药房.Top + cbo药房.Height + 50
    Else
     
        txt执行科室.Top = IIF(lblEatDrug.Visible, lblEatDrug.Top, lbl结束时间.Top) + 250
        If gbln特殊药品分开发送 Then
            lblDrugType.Top = txt执行科室.Top + 50
            cboDrugType.Top = txt执行科室.Top
            txt执行科室.Top = cboDrugType.Top + cboDrugType.Height + 10
        End If
        lbl执行科室.Top = txt执行科室.Top + 50
        cmd执行科室.Top = txt执行科室.Top + 30

        cbo药房.Top = lbl执行科室.Top + 270
        lbl药房.Top = cbo药房.Top + 50
        fra医嘱条件.Height = cbo药房.Top + cbo药房.Height + 50
    End If
    '可用药品医嘱期效：可以选择这项表示两者至少有一个
    If InStr(strPrivs, "发送药疗长嘱") = 0 Then
        chk药品期效(0).value = 0
        chk药品期效(0).Enabled = False
    ElseIf InStr(strPrivs, "发送药疗临嘱") = 0 Then
        chk药品期效(1).value = 0
        chk药品期效(1).Enabled = False
    End If
    '可用其他医嘱期效：可以选择这项表示两者至少有一个
    If InStr(strPrivs, "发送其他长嘱") = 0 Then
        chk其他期效(0).value = 0
        chk其他期效(0).Enabled = False
    ElseIf InStr(strPrivs, "发送其他临嘱") = 0 Then
        chk其他期效(1).value = 0
        chk其他期效(1).Enabled = False
    End If
    
    '通过医嘱提醒自动读取之前，凡有权限的都自动勾上
    If mblnOnePati And mblnAutoVerify Then
        If chk药品期效(0).Enabled Then chk药品期效(0).value = 1
        If chk药品期效(1).Enabled Then chk药品期效(1).value = 1
        If chk其他期效(0).Enabled Then chk其他期效(0).value = 1
        If chk其他期效(1).Enabled Then chk其他期效(1).value = 1
    Else
        '缺省药品医嘱期效
        lngTmp = Val(zlDatabase.GetPara("药嘱发送医嘱期效", glngSys, p住院医嘱发送, "2", Array(chk药品期效(0), chk药品期效(1))))
        If lngTmp = 0 Or lngTmp = 2 Then
            If chk药品期效(0).Enabled Then chk药品期效(0).value = 1
        End If
        If lngTmp = 1 Or lngTmp = 2 Then
            If chk药品期效(1).Enabled Then chk药品期效(1).value = 1
        End If
        '缺省其他医嘱期效
        lngTmp = Val(zlDatabase.GetPara("其他发送医嘱期效", glngSys, p住院医嘱发送, "2", Array(chk其他期效(0), chk其他期效(1))))
        If lngTmp = 0 Or lngTmp = 2 Then
            If chk其他期效(0).Enabled Then chk其他期效(0).value = 1
        End If
        If lngTmp = 1 Or lngTmp = 2 Then
            If chk其他期效(1).Enabled Then chk其他期效(1).value = 1
        End If
        '异常，至少勾一个(即使没有权限),以便决定其他控件的状态
        If chk药品期效(0).value = 0 And chk药品期效(1).value = 0 And chk其他期效(0).value = 0 And chk其他期效(1).value = 0 Then
            If InStr(strPrivs, ";发送药疗临嘱;") = 0 And InStr(strPrivs, ";发送药疗长嘱;") = 0 Then
                If chk其他期效(0).Enabled Then
                    chk其他期效(0).value = 1
                Else
                    chk其他期效(1).value = 1
                End If
            Else
                If chk药品期效(0).Enabled Then
                    chk药品期效(0).value = 1
                Else
                    chk药品期效(1).value = 1
                End If
            End If
        End If
    End If
    
    lngTmp = Val(zlDatabase.GetPara("药品长嘱从当天开始发送", glngSys, p住院医嘱发送, "0", Array(chkSendNow), InStr(strPrivs, ";参数设置;") > 0))
    chkSendNow.Tag = IIF(chkSendNow.Enabled, "", "0") '是否初始可用
    chkSendNow.value = lngTmp
    
    Call SetFaceBy医嘱期效
    
    '医嘱状态
    If mblnAutoVerify Then
        '缺省读取新开的
        fra状态.Visible = True
        opt状态(0).value = True
    Else
        '只读取已校对的
        fra状态.Visible = False
        opt状态(1).value = True
    End If
    
        
    '发送结束时间
    strTmp = zlDatabase.GetPara("发送结束时间", glngSys, p住院医嘱发送, "0", Array(cboTime))
    cboTime.ListIndex = Val(strTmp)
    If cboTime.ListIndex = cboTime.ListCount - 1 Then
        mdatCurr = zlDatabase.Currentdate
        strTmp = zlDatabase.GetPara("发送结束时点", glngSys, p住院医嘱发送, "23:59:59", Array(dtpEnd))
        lngTmp = Val(zlDatabase.GetPara("发送时间间隔", glngSys, p住院医嘱发送, "0", Array(dtpEnd)))
        mblnIsSendInterval = dtpEnd.Enabled
        dtpEnd.value = Format(mdatCurr + lngTmp, "yyyy-MM-dd " & strTmp)
    End If
    Call SetFaceBy医嘱期效

End Sub

Private Sub SetFaceBy医嘱期效()
'功能：根据四种期效的选择，设置界面的可用条件
    Dim bln药品 As Boolean, bln其他 As Boolean
    
    bln药品 = chk药品期效(0).value = 1 Or chk药品期效(1).value = 1
    bln其他 = chk其他期效(0).value = 1 Or chk其他期效(1).value = 1
    cboDrugType.Enabled = bln药品
    cbo药房.Enabled = bln药品
    tkpMain.Groups.Find(GP_给药途径).Enabled = bln药品
    tkpMain.Groups.Find(GP_药房置换).Enabled = bln药品
    If Not bln药品 Then
        tkpMain.Groups.Find(GP_给药途径).Expanded = False
        tkpMain.Groups.Find(GP_药房置换).Expanded = False
    End If
    
    txt执行科室.Enabled = bln其他
    cmd执行科室.Enabled = bln其他
    tkpMain.Groups.Find(GP_诊疗类别).Enabled = bln其他
    If Not bln其他 Then tkpMain.Groups.Find(GP_诊疗类别).Expanded = False
       
    '长嘱的截止时间
    cboTime.Enabled = chk药品期效(0).value = 1 Or chk其他期效(0).value = 1
    cboTime.BackColor = IIF(Not cboTime.Enabled, Me.BackColor, &H80000005)
    If mblnIsSendInterval Then
        '可将结束时间间隔参数设置成模块参数，即可控制护士最多发送多少天的，小于3天则没有意义，因为选择后天即为3天
        dtpEnd.Enabled = cboTime.Enabled And cboTime.ListIndex = cboTime.ListCount - 1
    End If
    chkSendNow.Enabled = chkSendNow.Tag <> "0" And chk药品期效(0).value = 1 And chk药品期效(0).Enabled = True
    '如果受医嘱期效控制，则不可用时清空参数值
    If Not chkSendNow.Enabled And chkSendNow.Tag <> "0" Then chkSendNow.value = 0
End Sub

Private Sub imgCheck_Click()
    If chk药品期效(0).Visible And chk药品期效(0).Enabled Then
        chk药品期效(0).value = IIF(chk药品期效(0).value = 0, 1, 0)
    End If
    If chk药品期效(1).Visible And chk药品期效(1).Enabled Then
        chk药品期效(1).value = IIF(chk药品期效(1).value = 0, 1, 0)
    End If
    
    If chk其他期效(0).Visible And chk其他期效(0).Enabled Then
        chk其他期效(0).value = IIF(chk其他期效(0).value = 0, 1, 0)
    End If
    If chk其他期效(1).Visible And chk其他期效(1).Enabled Then
        chk其他期效(1).value = IIF(chk其他期效(1).value = 0, 1, 0)
    End If
End Sub

Private Sub pic给药_Resize()
    On Error Resume Next
    
    lvwWay.Left = 0
    lvwWay.Top = 0
    lvwWay.Width = pic给药.ScaleWidth
    lvwWay.Height = pic给药.ScaleHeight - cmdAllWay.Height - Screen.TwipsPerPixelY * 2
    lvwWay.ColumnHeaders(1).Width = lvwWay.Width - (GetSystemMetrics(SM_CXVSCROLL) + 4) * Screen.TwipsPerPixelX

    cmdNoWay.Left = pic给药.ScaleWidth - cmdNoWay.Width - Screen.TwipsPerPixelX * 4
    cmdAllWay.Left = cmdNoWay.Left - cmdAllWay.Width
    cmdUnWay.Left = cmdAllWay.Left - cmdUnWay.Width
End Sub

Private Sub pic基本_Resize()
    On Error Resume Next
    Dim lngTop As Long
    
    If Me.Visible = False Then Exit Sub
    
    If fra药品.Visible Then lngTop = fra药品.Height + 60
    If fra其他.Visible Then
        fra其他.Top = lngTop
        lngTop = lngTop + fra其他.Height + 60
    End If
     
    If fra状态.Visible Then
        fra状态.Top = lngTop
        lngTop = lngTop + fra状态.Height + 60
    End If
    fra医嘱条件.Top = lngTop
       
    
    fraPati.Top = fra医嘱条件.Top + fra医嘱条件.Height + 30
    Line1.Y1 = pic基本.ScaleHeight - fra婴儿.Height - chk加班加价.Height - 150
    Line1.Y2 = Line1.Y1
    fraPati.Height = Line1.Y1 - fraPati.Top - 60
    lvwPati.Height = fraPati.Height - lvwPati.Top - cmdAllPati.Height - 30
    cmdAllPati.Top = lvwPati.Top + lvwPati.Height + 30
    cmdNoPati.Top = cmdAllPati.Top
    cmdQuick.Top = cmdAllPati.Top
    
    
    fra婴儿.Top = Line1.Y1 + 60
    chk加班加价.Top = fra婴儿.Top + fra婴儿.Height + 60
                
    fra医嘱条件.Width = pic基本.ScaleWidth - fra医嘱条件.Left
    cbo药房.Width = fra医嘱条件.Width - cbo药房.Left - Screen.TwipsPerPixelX * 3
    cboDrugType.Left = cbo药房.Left
    cboDrugType.Width = cbo药房.Width
    txt执行科室.Width = fra医嘱条件.Width - txt执行科室.Left - Screen.TwipsPerPixelX * 3
    cmd执行科室.Left = txt执行科室.Left + txt执行科室.Width - cmd执行科室.Width - Screen.TwipsPerPixelX * 2
    
    
    fraPati.Width = pic基本.ScaleWidth - fraPati.Left
    cboUnit.Width = fraPati.Width - cboUnit.Left - Screen.TwipsPerPixelX * 3
    lvwPati.Width = fraPati.Width - lvwPati.Left - Screen.TwipsPerPixelX * 3
    cmdNoPati.Left = fraPati.Width - cmdNoPati.Width - Screen.TwipsPerPixelX * 3
    cmdAllPati.Left = cmdNoPati.Left - cmdAllPati.Width
    
    Line1.X2 = pic基本.ScaleWidth
    
End Sub

Private Sub pic类别_Resize()
    On Error Resume Next
    
    lstClass.Left = 0
    lstClass.Top = 0
    lstClass.Width = pic类别.ScaleWidth
    lstClass.Height = pic类别.ScaleHeight
End Sub

Private Sub pic药房_Resize()
    On Error Resume Next
    
    vsDept.Width = pic药房.ScaleWidth
    vsDept.Height = pic药房.ScaleHeight - chkDeptExchange.Height
End Sub

Private Sub txt执行科室_GotFocus()
    Call zlControl.TxtSelAll(txt执行科室)
End Sub

Private Sub txt执行科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call cmd执行科室_Click
End Sub

Private Sub txt执行科室_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt执行科室.Text = txt执行科室.Tag Then
            Call zlcommfun.PressKey(vbKeyTab)
        ElseIf txt执行科室.Text = "" Then
            txt执行科室.Text = "所有科室"
            txt执行科室.Tag = txt执行科室.Text
            cmd执行科室.Tag = ""
            Call zlcommfun.PressKey(vbKeyTab)
        Else
            strSQL = _
                " Select 0 as ID,'-' as 编码,'所有科室' as 名称,NULL as 简码 From Dual" & _
                " Union ALL" & _
                " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
            strSQL = "Select * From (" & strSQL & ")" & _
                " Where 编码 Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2]" & _
                " Order by 编码"
            vRect = zlControl.GetControlRect(txt执行科室.hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行科室", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txt执行科室.Height, blnCancel, False, True, _
                UCase(txt执行科室.Text) & "%", mstrLike & UCase(txt执行科室.Text) & "%")
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配的科室。", vbInformation, gstrSysName
                End If
                txt执行科室.Text = txt执行科室.Tag
                Call zlControl.TxtSelAll(txt执行科室)
                txt执行科室.SetFocus
            Else
                txt执行科室.Text = rsTmp!名称
                txt执行科室.Tag = rsTmp!名称
                cmd执行科室.Tag = rsTmp!ID
                Call zlcommfun.PressKey(vbKeyTab)
            End If
        End If
    End If
End Sub

Private Sub txt执行科室_Validate(Cancel As Boolean)
    If txt执行科室.Text = "" Then
        txt执行科室.Text = "所有科室"
        txt执行科室.Tag = txt执行科室.Text
        cmd执行科室.Tag = ""
    ElseIf txt执行科室.Text <> txt执行科室.Tag Then
        txt执行科室.Text = txt执行科室.Tag '恢复人为的清除
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：更改成药规格
    Dim rsDrug As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng次数 As Long, lng最小次数 As Long
    Dim dbl总量 As Double, str分解时间 As String
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim cur金额 As Currency
    Dim bln口服 As Boolean
    
    If Col = COL_执行科室 Or Col = COL_附加执行 Then
        With vsAdvice
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        End With
    ElseIf Col = COL_规格 Then
        With vsAdvice
            If Val(.TextMatrix(Row, COL_收费细目ID)) = .ComboData Then Exit Sub
            '药品相关信息
            .TextMatrix(Row, COL_收费细目ID) = .ComboData
            Set rsDrug = GetDrugInfo(Val(.TextMatrix(Row, COL_诊疗项目ID)), Val(.TextMatrix(Row, COL_收费细目ID)), Val(.TextMatrix(Row, COL_执行科室ID)))
            .TextMatrix(Row, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
            .TextMatrix(Row, COL_剂量系数) = rsDrug!剂量系数
            .TextMatrix(Row, COL_住院包装) = rsDrug!住院包装
            .TextMatrix(Row, COL_住院单位) = NVL(rsDrug!住院单位)
            .TextMatrix(Row, COL_是否变价) = rsDrug!是否变价
            .TextMatrix(Row, COL_药房分批) = rsDrug!药房分批
            .TextMatrix(Row, COL_可否分零) = NVL(rsDrug!可否分零, 0)
            .TextMatrix(Row, COL_库存) = Format(NVL(rsDrug!库存, 0), "0.00000")
   
            '医嘱相关信息
            strSQL = _
                " Select A.ID,a.相关id as 组ID,A.诊疗类别,A.开始执行时间,A.上次执行时间,A.执行终止时间,A.执行时间方案," & _
                " A.频率次数,A.频率间隔,A.间隔单位,A.单次用量,A.可否分零,B.入院日期,A.医嘱状态,A.首次用量,A.医嘱期效,A.紧急标志,A.审核状态" & _
                " From 病人医嘱记录 A,病案主页 B" & _
                " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.ID=[1]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(Row, COL_ID)))
            
            If mstrEndPoint <> "" Then
                If InStr(",5,6,7,", .TextMatrix(Row, COL_诊疗类别)) > 0 Then
                    For i = Row + 1 To .Rows - 1
                        If .TextMatrix(i, COL_诊疗类别) = "E" And (Val(.TextMatrix(i, COL_操作类型)) = 2 Or Val(.TextMatrix(i, COL_操作类型)) = 4) Then
                            If Val(.TextMatrix(i, COL_执行分类)) = 4 Then bln口服 = True
                            Exit For
                        End If
                    Next
                End If
            End If
            
            '重新计算总量,次数,分解时间
            Call Calc总量次数时间(dbl总量, lng次数, str分解时间, mstrEnd, rsTmp, rsDrug, bln口服)
            
            .TextMatrix(Row, COL_总量) = FormatEx(dbl总量, 5)
            .TextMatrix(Row, COL_总量单位) = NVL(rsDrug!住院单位)
            
            .TextMatrix(Row, COL_次数) = lng次数
            .TextMatrix(Row, COL_分解时间) = str分解时间
            If str分解时间 <> "" Then
                .TextMatrix(Row, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                .TextMatrix(Row, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
            End If
                        
            '同步更改给药途径的次数
            i = .FindRow(.TextMatrix(Row, COL_相关ID), , COL_ID)
            .TextMatrix(i, COL_次数) = .TextMatrix(Row, COL_次数)
            .TextMatrix(i, COL_总量) = .TextMatrix(Row, COL_次数) '相同
            .TextMatrix(i, COL_分解时间) = .TextMatrix(Row, COL_分解时间)
            .TextMatrix(i, COL_首次时间) = .TextMatrix(Row, COL_首次时间)
            .TextMatrix(i, COL_末次时间) = .TextMatrix(Row, COL_末次时间)
                                    
            '一并给药的按最小次数计算：其他药品的总量不变
            If RowIn一并给药(Row, lngBegin, lngEnd) Then
                For i = lngBegin To lngEnd
                    If Val(.TextMatrix(i, COL_次数)) < lng最小次数 Or lng最小次数 = 0 Then
                        lng最小次数 = Val(.TextMatrix(i, COL_次数))
                    End If
                Next
                For i = lngBegin To lngEnd + 1
                    If Val(.TextMatrix(i, COL_次数)) > lng最小次数 Then
                        .TextMatrix(i, COL_次数) = lng最小次数
                        .TextMatrix(i, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(i, COL_分解时间))
                        .TextMatrix(i, COL_首次时间) = Format(Split(.TextMatrix(i, COL_分解时间), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_末次时间) = Format(Split(.TextMatrix(i, COL_分解时间), ",")(lng最小次数 - 1), "yyyy-MM-dd HH:mm")
                    End If
                Next
            Else
                lngBegin = Row: lngEnd = Row
            End If
            
            '重新计算并显示金额当前药品及给药途径的金额和计价
            mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngBegin, COL_ID)) & " Or 医嘱ID=" & Val(.TextMatrix(lngEnd + 1, COL_ID))
            Do While Not mrsPrice.EOF
                mrsPrice.Delete
                mrsPrice.Update
                mrsPrice.MoveNext
            Loop
            
            '标记计价内容变化
            .Cell(flexcpData, lngBegin, COL_金额) = 1
            .Cell(flexcpData, lngEnd + 1, COL_金额) = 1
            
            cur金额 = 0
            Call LoadAdvicePrice(lngBegin, cur金额, rsDrug)
            .TextMatrix(lngBegin, COL_金额) = Format(cur金额, gstrDec)
            cur金额 = 0
            Call LoadAdvicePrice(lngEnd + 1, COL_金额)
            .TextMatrix(lngEnd + 1, COL_金额) = Format(cur金额, gstrDec)
            
            '一并给药的第一行(如果是)才显示包含给药途径的金额
            .TextMatrix(lngBegin, COL_金额) = Format(Val(.TextMatrix(lngBegin, COL_金额)) + Val(.TextMatrix(lngEnd + 1, COL_金额)), gstrDec)
            
            '根据库存决定选择状态
            If Val(.TextMatrix(Row, COL_总量)) > Val(.TextMatrix(Row, COL_库存)) Then
                If TheStockCheck(Val(.TextMatrix(Row, COL_执行科室ID)), .TextMatrix(Row, COL_诊疗类别)) = 2 _
                    Or Val(.TextMatrix(Row, COL_药房分批)) = 1 Or Val(.TextMatrix(Row, COL_是否变价)) = 1 Then
                    .Cell(flexcpData, Row, COL_选择) = 1
                    Set .Cell(flexcpPicture, Row, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                ElseIf TheStockCheck(Val(.TextMatrix(Row, COL_执行科室ID)), .TextMatrix(Row, COL_诊疗类别)) = 1 Then
                    .Cell(flexcpData, Row, COL_选择) = 0
                    Set .Cell(flexcpPicture, Row, COL_选择) = Nothing
                ElseIf TheStockCheck(Val(.TextMatrix(Row, COL_执行科室ID)), .TextMatrix(Row, COL_诊疗类别)) = 0 Then
                    .Cell(flexcpData, Row, COL_选择) = 0
                    Set .Cell(flexcpPicture, Row, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                End If
            ElseIf Val(.TextMatrix(Row, COL_总量)) <= Val(.TextMatrix(Row, COL_库存)) Then
                .Cell(flexcpData, Row, COL_选择) = 0
                Set .Cell(flexcpPicture, Row, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
            End If
            Call RowSelectSame(Row, COL_选择)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
            Call ShowSendTotal
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            '根据可否编辑设置编辑特性及光标特性
            If NewCol = COL_规格 Then
                .ComboList = .Cell(flexcpData, NewRow, NewCol)
                .FocusRect = flexFocusLight
            ElseIf CellEditable(NewRow, NewCol) Then
                .ComboList = "..."
                Set .CellButtonPicture = Me.Picture
                .FocusRect = flexFocusHeavy
            Else
                .ComboList = ""
                .FocusRect = flexFocusLight
            End If
            
            If OldRow <> NewRow Then
                If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                    Call ShowAdvicePrice(NewRow)
                End If
            End If
        End If
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, COL_频率)
    End With
End Sub

Private Function Should附加执行(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱行(可见行)是否可以设置附加的执行科室
    Dim lngRow2 As Long, i As Long
    
    If lngRow = 0 Then Exit Function
    
    lngRow2 = -1
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) = 0 Then Exit Function
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
            And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
            '中药用法
            lngRow2 = lngRow
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '给药途径
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1, COL_ID)
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "F" Then
            '手术麻醉
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If .TextMatrix(i, COL_诊疗类别) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" _
            And .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
            '采集方式
            lngRow2 = lngRow
        End If
        
        '叮嘱或院外执行
        If lngRow2 <> -1 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_执行性质ID))) = 0 Then
                Should附加执行 = True
            End If
        End If
    End With
End Function

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_选择 + 1 - .FixedCols Then
            .FrozenCols = COL_选择 + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_医嘱内容 Or Col = COL_规格 Then
            If Not .ColHidden(COL_规格) Then
                .AutoSize col_医嘱内容, COL_规格
            Else
                .AutoSize col_医嘱内容
            End If
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_选择 Then Cancel = True
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vPoint As PointAPI, blnCancel As Boolean
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " Order by A.编码"
    With vsAdvice
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "执行科室", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetDeptInput(Row, Col, rsTmp)
            Call vsAdvice_AfterRowColChange(-1, -1, Row, Col) '重新显示计价执行科室
        Else
            If Not blnCancel Then
                MsgBox "没有可用的科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_ChangeEdit()
    If vsAdvice.Col = COL_规格 Then
        Call vsAdvice_AfterEdit(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_频率: lngRight = COL_用法
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_科室: lngRight = COL_医嘱期效
            If Not Between(Col, lngLeft, lngRight) Then
                Exit Sub
            End If
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Val(.TextMatrix(Row, COL_医嘱状态)) = 1 Then
                SetBkColor hDC, OS.SysColor2RGB(BackColorNew)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then '解决直接输入汉字的问题
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        If i > .Rows - 1 Then .Row = .FixedRows
        If .RowHidden(.Row) Then .Row = lngRow
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub EnterNextCellPrice(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：定位到价表中下一个可以输入的单元格
    Dim i As Long, j As Long
    
    With vsPrice
        '当前单元格如果未输入完整,则退出
        If CellEditablePrice(lngRow, lngCol) Then
            If lngCol = COLP_单价 And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '从下一单元开始循环搜索
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_计价医嘱) To .Cols - 1
                If CellEditablePrice(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '当前表格内没有找到下一个可编辑单元,如果有需计价医嘱,则增加一新行
            If CStr(.ColData(COLP_计价医嘱)) <> "" Then
                '当前行未输入完整,则定位到不完整单元
                If .TextMatrix(lngRow, COLP_计价医嘱) = "" Then
                    .Col = COLP_计价医嘱
                ElseIf .TextMatrix(lngRow, COLP_计价数量) = "" Then
                    .Col = COLP_计价数量
                ElseIf .TextMatrix(lngRow, COLP_收费项目) = "" Then
                    .Col = COLP_收费项目
                ElseIf Val(.TextMatrix(lngRow, COLP_变价)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_单价)) = 0 _
                    And CellEditablePrice(lngRow, COLP_单价) Then
                    .Col = COLP_单价
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_计价医嘱
                    
                    '缺省选择计价医嘱(如果可能)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '不可编辑时随意定一个
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub ShowDefaultRow()
'功能：对于可以计价的医嘱,缺省增加一行并设置缺省计价医嘱
'说明：ComboList="#医嘱ID1;计价医嘱1|#医嘱ID2;计价医嘱2|..."
'      仅在第一次显示计价表和回车新增行时调用
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lng医嘱ID As Long, lng行号 As Long, str计价医嘱 As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_计价医嘱) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_计价医嘱), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_行号)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_收费细目ID)) <> 0 Then
                '第一次显示时缺省增加一行
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '不是第一次显示时缺省计价医嘱与上一行相同
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_固定)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_行号)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lng医嘱ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str计价医嘱 = Replace(arrCombo(i), "#" & lng医嘱ID & ";", "")
                lng行号 = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
                If blnHave Then
                    If lng行号 = Val(.TextMatrix(lngRow - 1, COLP_行号)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            '模拟选中这个计价医嘱
            .TextMatrix(lngRow, COLP_行号) = lng行号
            .TextMatrix(lngRow, COLP_计价医嘱) = str计价医嘱
            .Cell(flexcpData, lngRow, COLP_计价医嘱) = .TextMatrix(lngRow, COLP_计价医嘱)
            
            '只有一个计价医嘱时不必停留
            If UBound(arrCombo) = 0 Then
                .Col = COLP_收费项目
            Else
                .Col = COLP_计价医嘱
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_选择 Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_选择) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_选择) Is Nothing Then
                    If InStr(mstrNoneIDs, "," & .TextMatrix(.Row, COL_ID) & ",") > 0 And Not mbln阳性用药 Then
                        MsgBox "该医嘱无有效的阴性皮试结果，不允许发送！", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        Set .Cell(flexcpPicture, .Row, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                    End If
                Else
                    Set .Cell(flexcpPicture, .Row, COL_选择) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        Else
            If CellEditable(.Row, .Col) And .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsAdvice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As PointAPI, blnCancel As Boolean
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COL_规格 Then
                Call vsAdvice_KeyPress(13)
            ElseIf Col = COL_附加执行 And .EditText <> "" Then
                strInput = UCase(.EditText)
                strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
                    " Order by A.编码"
                With vsAdvice
                    vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行科室", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                    If Not rsTmp Is Nothing Then
                        Call SetDeptInput(Row, Col, rsTmp)
                        .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                        Call EnterNextCell(Row, Col)
                    Else
                        If Not blnCancel Then
                            MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                        End If
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    End If
                End With
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlcommfun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not CellEditable(Row, Col) Then Cancel = True
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：判断发送医嘱清单中单元格是否可以编辑
    Dim bln采集 As Boolean, blnDo As Boolean, i As Long
    Dim bln改科室 As Boolean
    
    If lngRow = 0 Then Exit Function
    
    bln改科室 = InStr(";" & GetInsidePrivs(p住院医嘱发送) & ";", ";修改非药医嘱的执行科室;") > 0
    
    With vsAdvice
        CellEditable = .Editable
        If lngCol = COL_规格 Then
            CellEditable = .ComboList <> ""
        ElseIf lngCol = COL_执行科室 Then
            If InStr("5,6,7", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                '药品执行科室不允许直接输入更改
                CellEditable = False
            ElseIf bln改科室 Then
                '检验组合中只有有一个可以设置，就允许选择
                If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 And .TextMatrix(lngRow, COL_诊疗类别) = "E" And .TextMatrix(lngRow - 1, COL_诊疗类别) <> "E" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then bln采集 = True
                
                If bln采集 Then
                    blnDo = False
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                                blnDo = True: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                Else
                    blnDo = InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0
                End If
                If Not blnDo Then CellEditable = False
            Else
                CellEditable = False
            End If
        ElseIf lngCol = COL_附加执行 And bln改科室 Then
            CellEditable = Should附加执行(lngRow)
        Else
            CellEditable = False
        End If
    End With
End Function

Private Function CellEditablePrice(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln非本科 As Boolean) As Boolean
'功能：判断价表中单元格是否可以编辑
    Dim lng行号 As Long
    
    With vsPrice
        bln非本科 = False
        CellEditablePrice = .Editable
        lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
        If lngCol = COLP_执行科室 Then
            '跟踪在用的卫材,非药嘱药品计价的执行科室可以修改
            If Not ((.TextMatrix(lngRow, COLP_收费类别) = "4" And Val(.TextMatrix(lngRow, COLP_跟踪在用)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_收费类别)) > 0) And InStr(",4,5,6,7,", vsAdvice.TextMatrix(lng行号, COL_诊疗类别)) = 0) Then
                CellEditablePrice = False
            End If
            If .TextMatrix(lngRow, COLP_收费项目) = "" Or .TextMatrix(lngRow, COLP_行号) = "" Then
                CellEditablePrice = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_固定)) <> 0 Then
            '固定对照行仅可以修改变价
            If Not (Val(.TextMatrix(lngRow, COLP_变价)) = 1 And lngCol = COLP_单价) Then
                CellEditablePrice = False
            End If
        Else
            If lngCol = COLP_单价 Then
                If Val(.TextMatrix(lngRow, COLP_变价)) <> 1 Then
                    CellEditablePrice = False
                Else
                    '非本科执行的变价项目不允许定价格
                    If lng行号 <> 0 Then
                        If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                            bln非本科 = True: CellEditablePrice = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_计价医嘱 And lngCol <> COLP_计价数量 And lngCol <> COLP_收费项目 Then
                CellEditablePrice = False
            End If
        End If
    End With
End Function

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsDept.Cell(flexcpData, Row, Col) = CLng(vsDept.ComboData)
    mrs药房.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, Row, 0))
    mrs药房!现ID = CLng(vsDept.Cell(flexcpData, Row, Col))
    mrs药房.Update
    mblnChangeIF = True
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsDept.Editable <> flexEDNone And NewCol = 1 Then
        vsDept.FocusRect = flexFocusSolid
    Else
        vsDept.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsDept_ChangeEdit()
    Call vsDept_AfterEdit(vsDept.Row, vsDept.Col)
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.Col = 1 Then
            If vsDept.Row + 1 <= vsDept.Rows - 1 Then
                vsDept.Row = vsDept.Row + 1
            Else
                Call zlcommfun.PressKey(vbKeyTab)
                vsDept.Row = vsDept.FixedRows + 1
            End If
        Else
            vsDept.Col = 1
        End If
        Call vsDept.ShowCell(vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub vsDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.ComboIndex <> -1 Then
            Call vsDept_KeyPress(13)
        End If
    End If
End Sub

Private Sub vsDept_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lng原嘱ID As Long, lng医嘱ID As Long
    Dim int费用性质 As Integer, int原费用性质 As Integer
    Dim lng收费细目ID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_计价医嘱 Then
            '如果绑定了ComboData,TextMatrix取值就为ComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lng医嘱ID = .ComboData
                If lng医嘱ID < 0 Then
                    int费用性质 = Val(Left(Abs(lng医嘱ID), 1))
                    lng医嘱ID = Val(Mid(Abs(lng医嘱ID), 2))
                End If
                lng原嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
                int原费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                                
                '检查该计价医嘱是否已有相同收费细目
                If lng收费细目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """已经设置了收费项目""" & .TextMatrix(Row, COLP_收费项目) & """。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '原来的医嘱如果有从项至少要保留一个(主项是固定不可动的)
                If lng原嘱ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 费用性质=" & int原费用性质 & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '标明输入了的计价医嘱部份
                i = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
                .TextMatrix(Row, COLP_行号) = i
                .TextMatrix(Row, COLP_费用性质) = int费用性质
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng收费细目ID <> 0 Then
                    '新选择的医嘱是否有从项决定修改后的项目是否从项
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 从项=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_从项) = IIF(blnHaveSub, "√", "")
                
                    '更新或增加记录集内容
                    If lng原嘱ID = 0 Then
                        mrsPrice.AddNew '加入
                    Else '更新
                        mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 费用性质=" & int原费用性质 & " And 收费细目ID=" & lng收费细目ID
                    End If
                    mrsPrice!医嘱ID = lng医嘱ID
                    If Val(vsAdvice.TextMatrix(i, COL_相关ID)) <> 0 Then
                        mrsPrice!相关ID = vsAdvice.TextMatrix(i, COL_相关ID)
                    Else
                        mrsPrice!相关ID = Null
                    End If
                    mrsPrice!费用性质 = int费用性质
                    mrsPrice!收费方式 = 0
                    If lng原嘱ID = 0 Then
                        mrsPrice!收费细目ID = lng收费细目ID
                        mrsPrice!数量 = Val(.TextMatrix(Row, COLP_计价数量))
                        mrsPrice!单价 = Val(.TextMatrix(Row, COLP_单价))
                        mrsPrice!在用 = Val(.TextMatrix(Row, COLP_跟踪在用))
                        mrsPrice!变价 = Val(.TextMatrix(Row, COLP_变价))
                        mrsPrice!固定 = 0
                    End If
                    mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    '标记计价内容变化
                    If lng原嘱ID <> 0 Then
                        vsAdvice.Cell(flexcpData, vsAdvice.FindRow(CStr(lng原嘱ID), , COL_ID), COL_金额) = 1
                    End If
                    vsAdvice.Cell(flexcpData, i, COL_金额) = 1
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
            End If
        ElseIf Col = COLP_收费项目 Or Col = COLP_执行科室 Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        ElseIf Col = COLP_计价数量 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
            int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                '标记计价内容变化
                vsAdvice.Cell(flexcpData, Val(.TextMatrix(Row, COLP_行号)), COL_金额) = 1
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        ElseIf Col = COLP_单价 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_应收金额), .Cell(flexcpData, Row, COLP_实收金额), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
            int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                '标记计价内容变化
                vsAdvice.Cell(flexcpData, Val(.TextMatrix(Row, COLP_行号)), COL_金额) = 1
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '根据可否编辑设置
    If Not CellEditablePrice(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_计价医嘱 Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_收费项目 Or NewCol = COLP_执行科室 Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_行号))
            If lngRow <> 0 And .Cell(flexcpData, NewRow, COLP_类别) <> "" Then
                If InStr(",5,6,7,", .Cell(flexcpData, NewRow, COLP_类别)) > 0 _
                    Or .Cell(flexcpData, NewRow, COLP_类别) = "4" And Val(.Cell(flexcpData, NewRow, COLP_费用类型)) = 1 Then
                    '显示药品及跟踪卫材的库存
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                        If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_规格) & "，" & vsAdvice.TextMatrix(lngRow, COL_执行科室) & IIF(Val(vsAdvice.TextMatrix(lngRow, COL_库存)) > 0, "有库存", "无库存")
                        Else
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_规格) & "，" & vsAdvice.TextMatrix(lngRow, COL_执行科室) & "可用库存：" & _
                                FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_库存)), 5) & vsAdvice.TextMatrix(lngRow, COL_住院单位)
                        End If
                    Else
                        '同一个函数取:药品按住院单位,卫材按售价单位
                        If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                            If GetStock(Val(.Cell(flexcpData, NewRow, COLP_收费项目)), Val(.Cell(flexcpData, NewRow, COLP_执行科室))) > 0 Then
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "，" & .TextMatrix(NewRow, COLP_执行科室) & "有库存"
                            Else
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "，" & .TextMatrix(NewRow, COLP_执行科室) & "无库存"
                            End If
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "，" & .TextMatrix(NewRow, COLP_执行科室) & "可用库存：" & _
                                FormatEx(GetStock(Val(.Cell(flexcpData, NewRow, COLP_收费项目)), Val(.Cell(flexcpData, NewRow, COLP_执行科室))), 5) & .TextMatrix(NewRow, COLP_单位)
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng行号 As Long, i As Long
    Dim str项目IDs As String, blnCancel As Boolean
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim int费用性质 As Integer, vPoint As PointAPI
    Dim strSQL2 As String
    
    With vsPrice
        lng行号 = Val(.TextMatrix(Row, COLP_行号))
        If Col = COLP_收费项目 Then
            '不能选择已有的项目
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_行号)) = lng行号 And lng行号 <> 0 And i <> Row Then
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                End If
            Next
            str项目IDs = Mid(str项目IDs, 2)
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lng行号, COL_病人ID)), Val(vsAdvice.TextMatrix(lng行号, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        
            
            strSQL = _
                " Select Distinct 0 as 末级,To_Number('999999999'||类型) as ID,-NULL as 上级ID," & _
                " CHR(13)||类型 as 编码,Decode(类型,1,'西成药',2,'中成药',3,'中草药',7,'卫生材料') as 名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 医保大类,NULL as 说明,NULL as 价格," & _
                " -NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,-ID as ID,Nvl(-上级ID,To_Number('999999999'||类型)) as 上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 医保大类,NULL as 说明,NULL as 价格," & _
                " -NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 医保大类," & _
                " NULL as 说明,NULL as 价格,-NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 收费分类目录 Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL2 = _
                " Select 末级,ID,上级ID,编码,名称,单位,规格,产地,类别,费用类型,医保大类,说明," & _
                " Decode(Nvl(是否变价,0),1,Decode(Instr('567',类别ID),0,Sum(Nvl(原价,0))||'-'||Sum(Nvl(现价,0)),'时价'),Sum(现价)) as 价格," & _
                " Sum(原价) as 原价ID,Sum(现价) as 现价ID,Sum(缺省价格) as 缺省价格ID,是否变价 as 是否变价ID,类别ID,跟踪在用ID" & _
                " From (" & _
                " Select Distinct 1 as 末级,A.ID,Decode(Instr('567',A.类别),0,A.分类ID,-E.分类ID) as 上级ID,A.编码,A.名称," & _
                " A.计算单位 as 单位,A.规格,A.产地,C.名称 as 类别,A.费用类型,N.名称 as 医保大类,A.说明,B.原价,B.现价,B.缺省价格,A.是否变价," & _
                " A.类别 as 类别ID,-Null as 跟踪在用ID" & _
                " From 收费项目目录 A,收费价目 B,收费项目类别 C,药品规格 D,诊疗项目目录 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.ID=B.收费细目ID [选择替换的过条件1]  And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.类别 Not IN('4','J','1') And A.类别=C.编码 And A.ID=D.药品ID(+) And D.药名ID=E.ID(+)" & _
                " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]" & _
                " And (Nvl(a.执行科室,0) <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = a.Id And (w.病人来源=2 or (w.病人来源 is Null And Nvl(w.开单科室id,[3]) = [3]))))" & _
                " And (a.类别 Not in ('5','6','7') Or Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[3])=[3]))"
            If DeptExist("发料部门", 2) Then
                strSQL2 = strSQL2 & " Union ALL " & _
                    " Select Distinct 1 as 末级,A.ID,-E.分类ID as 上级ID,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,C.名称 as 类别,A.费用类型,N.名称 as 医保大类,A.说明," & _
                    " B.原价,B.现价,B.缺省价格,A.是否变价,A.类别 as 类别ID,D.跟踪在用 as 跟踪在用ID" & _
                    " From 收费项目目录 A,收费价目 B,收费项目类别 C,材料特性 D,诊疗项目目录 E,保险支付项目 M,保险支付大类 N" & _
                    " Where A.ID=B.收费细目ID [选择替换的过条件2]  And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.类别='4' And A.类别=C.编码 And A.ID=D.材料ID And D.诊疗ID=E.ID And D.核算材料=0" & _
                    " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]" & _
                    " And Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[3])=[3])"
            End If
            strSQL2 = strSQL2 & " ) Group by 末级,ID,上级ID,类别,编码,名称,单位,规格,产地,费用类型,医保大类,说明,是否变价,类别ID,跟踪在用ID"
            '[选择替换的过条件1],[选择替换的过条件2],这两个串在选器中处理的
            '要确保 "占位参数" 在最后一位，该参数在选择器中拼接，要解决4000长度的问题
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "收费项目", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str项目IDs & ",", Val(vsAdvice.TextMatrix(lng行号, COL_险类)), Val(vsAdvice.TextMatrix(lng行号, COL_病人科室ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "占位参数")
            If Not rsTmp Is Nothing Then
                '非本科执行的医嘱不允许输入变价项目
                If lng行号 <> 0 Then
                    If NVL(rsTmp!是否变价ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!类别ID) > 0 Or rsTmp!类别ID = "4" And NVL(rsTmp!跟踪在用ID, 0) = 1) Then
                        If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                            MsgBox "该医嘱非本科执行，不允许对变价项目""" & rsTmp!名称 & """定价。该计价项目需要手工计价。", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                '医保对码检查
                If CheckItemInsure(rsTmp, lng行号) Then
                    .SetFocus: Exit Sub
                End If
                
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                Call SetItemInput(Row, rsTmp, lng医嘱ID, int费用性质, lng原项目ID)
                If lng行号 <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCellPrice(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有可用的收费项目，请先到收费项目管理中设置！", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_执行科室 Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_收费类别) = "4" Then
                '跟踪在用的卫材
                strSQL = _
                    " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                    " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                    " And A.收费细目ID=[1]" & _
                    " Order by B.服务对象,C.编码"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                '药品
                '药品从系统指定的储备药房中找
                If Not Check上班安排(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                    decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!名称
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!执行科室ID = rsTmp!ID
                    mrsPrice.Update
                    
                    '标记计价内容变化
                    vsAdvice.Cell(flexcpData, lng行号, COL_金额) = 1
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCellPrice(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset, ByVal lngRow As Long) As Boolean
'功能：检查输入(选择)计价项目是否医保对码
'返回：如果未对码，并且提示选择不继续，则返回真。
    
    If gint医保对码 = 0 Then Exit Function
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_险类)) <> 0 Then
            If Not ItemExistInsure(Val(.TextMatrix(lngRow, COL_病人ID)), rsInput!ID, Val(.TextMatrix(lngRow, COL_险类))) Then
                If gint医保对码 = 1 Then
                    If MsgBox("项目""" & rsInput!名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckItemInsure = True
                    End If
                ElseIf gint医保对码 = 2 Then
                    MsgBox "项目""" & rsInput!名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                    CheckItemInsure = True
                End If
            End If
        End If
    End With
End Function

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lng医嘱ID As Long, ByVal int费用性质 As Integer, ByVal lng原项目ID As Long)
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lng行号 As Long, dbl单价 As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '记录集内容
        '表格内容:仅临时显示标记输入了项目,也可以处理为未定计价医嘱不允许输入项目
        .TextMatrix(lngRow, COLP_类别) = rsInput!类别
        .TextMatrix(lngRow, COLP_收费类别) = rsInput!类别ID
        .TextMatrix(lngRow, COLP_收费细目ID) = rsInput!ID
        .TextMatrix(lngRow, COLP_收费项目) = rsInput!名称
        If Not IsNull(rsInput!产地) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & "(" & rsInput!产地 & ")"
        End If
        If Not IsNull(rsInput!规格) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & " " & rsInput!规格
        End If
        .TextMatrix(lngRow, COLP_单位) = NVL(rsInput!单位) '都按零售单位(包括非药嘱药品计价)
        .TextMatrix(lngRow, COLP_计价数量) = 1 '缺省相对计价1,药品为计1个零售单位
        
        '执行科室
        lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
        If lng行号 <> 0 Then
            lng执行科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lng行号, COL_病人ID)), Val(vsAdvice.TextMatrix(lng行号, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            '非药嘱药品和跟踪在用的卫材专门求执行科室
            If rsInput!类别ID = "4" And NVL(rsInput!跟踪在用ID, 0) = 1 Or InStr(",5,6,7,", rsInput!类别ID) > 0 Then
                lng病人科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人科室ID))
                lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsInput!类别ID, rsInput!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
            End If
        End If
        .TextMatrix(lngRow, COLP_执行科室) = sys.RowValue("部门表", lng执行科室ID, "名称")
        .TextMatrix(lngRow, COLP_执行科室ID) = lng执行科室ID
        
        '单价计算处理:药嘱的药品计价不可能在这里处理
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
            If NVL(rsInput!是否变价ID, 0) = 0 Then
                dbl单价 = NVL(rsInput!现价ID, 0)
            ElseIf lng行号 <> 0 Then
                '按每次缺省一个零售单位,当前发送数次计算
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, Val(vsAdvice.TextMatrix(lng行号, COL_总量)), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            End If
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, gstrDecPrice)
                        
            '时价药品不输入价格
            .TextMatrix(lngRow, COLP_变价) = 0
            .Cell(flexcpData, lngRow, COLP_应收金额) = 0
            .Cell(flexcpData, lngRow, COLP_实收金额) = 0
        ElseIf rsInput!类别ID = "4" And NVL(rsInput!跟踪在用ID, 0) = 1 And NVL(rsInput!是否变价ID, 0) = 1 Then
            '跟踪在用的时价卫材和药品一样计算
            dbl单价 = 0
            If lng行号 <> 0 Then
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, Val(vsAdvice.TextMatrix(lng行号, COL_总量)), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            End If
            .TextMatrix(lngRow, COLP_变价) = 0
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, gstrDecPrice)
            .Cell(flexcpData, lngRow, COLP_应收金额) = 0
            .Cell(flexcpData, lngRow, COLP_实收金额) = 0
        Else
            If NVL(rsInput!是否变价ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_变价) = 0
                .TextMatrix(lngRow, COLP_单价) = Format(NVL(rsInput!现价ID, 0), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_应收金额) = 0
                .Cell(flexcpData, lngRow, COLP_实收金额) = 0
            Else
                .TextMatrix(lngRow, COLP_变价) = 1
                .TextMatrix(lngRow, COLP_单价) = Format(NVL(rsInput!缺省价格ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_应收金额) = NVL(rsInput!原价ID, 0)
                .Cell(flexcpData, lngRow, COLP_实收金额) = NVL(rsInput!现价ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_费用类型) = NVL(rsInput!费用类型)
        .TextMatrix(lngRow, COLP_固定) = 0
        
        '用于输入恢复
        .Cell(flexcpData, lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目)
        .Cell(flexcpData, lngRow, COLP_计价数量) = .TextMatrix(lngRow, COLP_计价数量)
        .Cell(flexcpData, lngRow, COLP_单价) = .TextMatrix(lngRow, COLP_单价)
        .Cell(flexcpData, lngRow, COLP_执行科室) = .TextMatrix(lngRow, COLP_执行科室)
        
        '记录集内容
        If lng医嘱ID <> 0 Then
            If lng原项目ID = 0 Then
                '当前医嘱是否有从项决定新增的项目是否从项
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 从项=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_从项) = IIF(blnHaveSub, "√", "")
            
                mrsPrice.AddNew '加入
            Else '更新
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
            End If
            If lng原项目ID = 0 Then
                mrsPrice!医嘱ID = lng医嘱ID
                lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
                If Val(vsAdvice.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                    mrsPrice!相关ID = Val(vsAdvice.TextMatrix(lng行号, COL_相关ID))
                Else
                    mrsPrice!相关ID = Null
                End If
                mrsPrice!费用性质 = int费用性质
                mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!收费方式 = 0
            mrsPrice!收费类别 = rsInput!类别ID
            mrsPrice!收费细目ID = rsInput!ID
            If lng执行科室ID <> 0 Then
                mrsPrice!执行科室ID = lng执行科室ID
            Else
                mrsPrice!执行科室ID = Null
            End If
            mrsPrice!在用 = NVL(rsInput!跟踪在用ID, 0)
            mrsPrice!变价 = NVL(rsInput!是否变价ID, 0)
            mrsPrice!单价 = Val(.TextMatrix(lngRow, COLP_单价))
            mrsPrice!数量 = 1
            mrsPrice!固定 = 0
            mrsPrice.Update
            
            '标记计价内容变化
            vsAdvice.Cell(flexcpData, lng行号, COL_金额) = 1
        End If
    End With
End Sub

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditablePrice(.Row, .Col) And .Col = COLP_计价医嘱 Then
                Call zlcommfun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_固定)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_行号)) <> 0 And Val(.TextMatrix(.Row, COLP_收费细目ID)) <> 0 Then
                    '医嘱如果有从项至少要保留一个(主项是固定不可动的)
                    mrsPrice.Filter = "医嘱ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_行号)), COL_ID)) & _
                        " And 费用性质=" & Val(.TextMatrix(.Row, COLP_费用性质)) & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_计价医嘱) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("确实要删除当前计价行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "医嘱ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_行号)), COL_ID)) & _
                        " And 费用性质=" & Val(.TextMatrix(.Row, COLP_费用性质)) & " And 收费细目ID=" & Val(.TextMatrix(.Row, COLP_收费细目ID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_计价医嘱
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCellPrice(.Row, .Col)
        Else
            If CellEditablePrice(.Row, .Col) And (.Col = COLP_收费项目 Or .Col = COLP_执行科室) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng行号 As Long, i As Long
    Dim str项目IDs As String, int费用性质 As Integer
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI
    Dim lng病人科室ID As Long
    Dim lng西药房 As Long
    Dim lng成药房 As Long
    Dim lng中药房 As Long
    Dim lng发料部门 As Long
    Dim strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng行号 = Val(.TextMatrix(Row, COLP_行号))
            If Col = COLP_计价医嘱 Then
                '下拉时回车
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '不然EnterNextCellPrice函数要退出
                    Call EnterNextCellPrice(Row, Col)
                End If
            ElseIf Col = COLP_计价数量 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "计价数量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    
                    '标记计价内容变化
                    vsAdvice.Cell(flexcpData, lng行号, COL_金额) = 1

                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCellPrice(Row, Col)
            ElseIf Col = COLP_单价 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费单价输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '检查变价输入范围
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_应收金额), .Cell(flexcpData, Row, COLP_实收金额), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    
                    '标记计价内容变化
                    vsAdvice.Cell(flexcpData, lng行号, COL_金额) = 1
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCellPrice(Row, Col)
            ElseIf Col = COLP_收费项目 And .EditText <> "" Then
                '不能选择已有的项目
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_行号)), COL_ID)) = Val(vsAdvice.TextMatrix(lng行号, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng行号, COL_ID)) <> 0 And i <> Row Then
                        str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                    End If
                Next
                str项目IDs = Mid(str项目IDs, 2)
                
                lng病人科室ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID))
                lng中药房 = Val(zlDatabase.GetPara("住院缺省中药房", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
                lng西药房 = Val(zlDatabase.GetPara("住院缺省西药房", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
                lng成药房 = Val(zlDatabase.GetPara("住院缺省成药房", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
                lng发料部门 = Val(zlDatabase.GetPara("住院缺省发料部门", glngSys, p住院医嘱下达, , , , , lng病人科室ID))
                
                If lng西药房 <> 0 Or lng成药房 <> 0 Or lng中药房 <> 0 Or lng发料部门 <> 0 Then
                    strStock = _
                        "Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                        " From 药品库存 A,收费项目目录 B" & _
                        " Where A.性质 = 1 And (Nvl(A.批次,0)=0 Or A.效期 Is Null Or A.效期>Trunc(Sysdate))" & _
                        " And A.库房ID=Decode(B.类别,'5',[7],'6',[8],'7',[9],'4',[10],Null)" & _
                        " And A.药品ID=B.ID And B.类别 IN('4','5','6','7')" & _
                        " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
                Else
                    strStock = "Select Null as 药品ID,Null as 库存 From Dual"
                End If
                
                '不同的输入匹配方式
                strInput = UCase(.EditText)
                strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
                ElseIf zlcommfun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.简码 Like [2] And C.码类=[3]"
                ElseIf zlcommfun.IsCharChinese(strInput) Then
                    strMatch = " And C.名称 Like [2] And C.码类=[3]"
                End If
                If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lng行号, COL_病人ID)), Val(vsAdvice.TextMatrix(lng行号, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                
                strSQL = ""
                If Not DeptExist("发料部门", 2) Then strSQL = " And A.类别<>'4'"
                strSQL = "Select * From (" & _
                    " Select A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地," & _
                    " Decode(Instr('4567',A.类别ID),0,NULL,1," & _
                    "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存,'999990.0000'))||A.单位)," & _
                    "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存/Nvl(C.住院包装,1),'999990.0000'))||C.住院单位)) as 库存," & _
                    "   A.费用类型,N.名称 as 医保大类,A.说明," & _
                    " Decode(Nvl(A.是否变价,0),1,Decode(Instr('567',A.类别ID),0,Sum(Nvl(A.原价,0))||'-'||Sum(Nvl(A.现价,0)),'时价'),Sum(A.现价)) as 价格," & _
                    " Sum(A.原价) as 原价ID,Sum(A.现价) as 现价ID,Sum(A.缺省价格) as 缺省价格ID,A.是否变价 as 是否变价ID,A.类别ID,B.跟踪在用 as 跟踪在用ID,B.核算材料" & _
                    " From (" & _
                    " Select Distinct 1 as 末级,A.ID,a.执行科室,A.类别 as 类别ID,D.名称 as 类别,A.编码,A.名称,A.计算单位 as 单位," & _
                    " A.规格,A.产地,A.费用类型,A.说明,B.原价,B.现价,B.缺省价格,A.是否变价" & _
                    " From 收费项目目录 A,收费价目 B,收费项目别名 C,收费项目类别 D" & _
                    " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "11", "12", "13") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.收费细目ID And A.类别=D.编码 And A.类别 Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,材料特性 B,药品规格 C,保险支付项目 M,保险支付大类 N,(" & strStock & ") S" & _
                    " Where A.ID=B.材料ID(+) And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[5]  And A.ID=C.药品ID(+) And A.ID=S.药品ID(+)" & _
                    " And (Nvl(a.执行科室,0) <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = a.Id And (w.病人来源=2 or (w.病人来源 is Null And Nvl(w.开单科室id,[6]) = [6]))))" & _
                    " And (a.类别id not in ('4','5','6','7') Or Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[6])=[6]))" & _
                    " Group by A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地,A.费用类型,C.住院单位,C.住院包装,S.库存,N.名称,A.说明,A.是否变价,A.类别ID,B.跟踪在用,B.核算材料" & _
                    " ) Where Nvl(核算材料,0) = 0 Order by 类别, 编码"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "收费项目", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint简码 + 1, "," & str项目IDs & ",", Val(vsAdvice.TextMatrix(lng行号, COL_险类)), lng病人科室ID, lng西药房, lng成药房, lng中药房, lng发料部门, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                If Not rsTmp Is Nothing Then
                    '非本科执行的医嘱不允许输入变价项目
                    If lng行号 <> 0 Then
                        If NVL(rsTmp!是否变价ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!类别ID) > 0 Or rsTmp!类别ID = "4" And NVL(rsTmp!跟踪在用ID, 0) = 1) Then
                            If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                                MsgBox "该医嘱非本科执行，不允许对变价项目""" & rsTmp!名称 & """定价。该计价项目需要手工计价。", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    '医保对码检查
                    If CheckItemInsure(rsTmp, lng行号) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                        .SetFocus: Exit Sub
                    End If
                
                    lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                    int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    Call SetItemInput(Row, rsTmp, lng医嘱ID, int费用性质, lng原项目ID)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    If lng行号 <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                    End If
                    Call EnterNextCellPrice(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的收费项目！", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            ElseIf Col = COLP_执行科室 And .EditText <> "" Then '执行科室
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_收费类别) = "4" Then
                    '跟踪在用的卫材
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1] And (C.编码 Like [3] Or C.名称 Like [4] Or C.简码 Like [4])" & _
                        " Order by B.服务对象,C.编码"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                    '药品从系统指定的储备药房中找
                    If Not Check上班安排(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                            " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                        decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!名称
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    
                    '更新记录集
                    lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                    int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                        mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                        mrsPrice!执行科室ID = rsTmp!ID
                        mrsPrice.Update
                        
                        '标记计价内容变化
                        vsAdvice.Cell(flexcpData, lng行号, COL_金额) = 1
                        
                        Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                    End If
                    Call EnterNextCellPrice(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_计价数量 Or Col = COLP_单价 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明：返回的行号范围不包括给药途径的行号
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub InitAdviceTable()
'功能：初始化清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;科室,850,1;姓名,750,1;住院号,750,1;床号,500,4;费别,750,1;婴儿,550,1;期效,550,1;医嘱内容,2000,1;规格,2000,1;总量,600,7;单位,450,1;" & _
        "单量,600,7;单位,450,1;金额,850,7;频率,1000,1;用法,1000,1;医生嘱托,1500,1;执行时间,1000,1;首次时间,1530,1;末次时间,1530,1;执行科室,850,1;" & _
        "附加执行,850,1;执行性质,850,1;病人ID;主页ID;性别;年龄;险类;ID;相关ID;病人病区ID;病人科室ID;开嘱科室ID;开嘱医生;诊疗类别;诊疗项目ID;计价特性;" & _
        "执行性质ID;执行科室ID;执行标记;药品ID;剂量系数;住院包装;住院单位;可否分零;药房分批;是否变价;库存;次数;分解时间;操作类型;试管编码;标本部位;检查方法;跟踪在用;" & _
        "紧急标志;医嘱状态;执行频率;新开操作时间;计算方式;执行安排;病人性质;开嘱时间;开始时间;申请序号;执行分类;毒理分类;用药理由;会诊医嘱ID"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_选择 + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "行号;收费细目ID;固定;变价;计价医嘱,2000,1;类别,650,1;收费项目,2000,1;计价数量,900,7;" & _
        "付数,450,4;数量,800,7;单位,500,1;单价,1000,7;应收金额,1050,7;实收金额,1050,7;执行科室,1000,1;费用类型,850,1;" & _
        "从项,450,4;收费方式,1500,1;收费类别;执行科室ID;跟踪在用;费用性质"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Function Decode药房置换() As String
'功能：根据药房置换设置,产生字段Decode语句
'说明：SQL查询中,"病人医嘱记录"别名为"A"
    Dim i As Long, strTmp As String
    
    mbln药房置换 = False
    If chkDeptExchange.value = 1 And InStr(GetInsidePrivs(p住院医嘱发送), ";允许置换药房;") > 0 Then
        mrs药房.Filter = 0
        If Not mrs药房.EOF Then
            For i = 1 To mrs药房.RecordCount
                If mrs药房!ID <> mrs药房!现ID Then
                    strTmp = strTmp & "," & mrs药房!ID & "," & mrs药房!现ID
                End If
                mrs药房.MoveNext
            Next
        End If
    End If
    If strTmp <> "" Then
        Decode药房置换 = "Decode(A.执行科室ID" & strTmp & ",A.执行科室ID)"
        mbln药房置换 = True
    Else
        Decode药房置换 = "A.执行科室ID"
    End If
End Function

Private Sub DeleteCurRow(ByVal lngRow As Long, ByVal lng相关ID As Long)
'功能：在处理待发送清单的过程中删除最近加入的行
    Dim i As Long
    With vsAdvice
        '删除当前行
        .RemoveItem lngRow
        
        '删除配方或一并给药中已经加入的行
        If lng相关ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Function CheckWaitExecute(rsPati As ADODB.Recordset, ByVal lngRow As Long, ByVal byt项目检查方式 As Byte, ByVal byt药品检查方式 As Byte) As Boolean
'功能：按照指定的检查方式，对病人未执行的项圈或未发药品进行检查
'参数：byt检查方式=0-不检查,1-检查并提示,2-检查并禁止
'返回：是否继续
    Dim strTmp As String
        
    With vsAdvice
        If byt项目检查方式 <> 0 Then
            strTmp = ExistWaitExe(rsPati!病人ID, Val(.TextMatrix(lngRow, COL_主页ID)), -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_医嘱内容): .Refresh
                If byt项目检查方式 = 1 Then
                    If MsgBox("发现病人""" & rsPati!姓名 & """存在尚未执行完成的内容：" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(lngRow, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "发现病人""" & rsPati!姓名 & """存在尚未执行完成的内容：" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """将不被发送。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        If byt药品检查方式 <> 0 Then
            strTmp = ExistWaitDrug(rsPati!病人ID, Val(.TextMatrix(lngRow, COL_主页ID)), -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_医嘱内容): .Refresh
                If byt药品检查方式 = 1 Then
                    If MsgBox("发现病人""" & rsPati!姓名 & """" & _
                        strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(lngRow, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "发现病人""" & rsPati!姓名 & """" & _
                        strTmp & vbCrLf & vbCrLf & "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """将不被发送", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End With
    
    CheckWaitExecute = True
End Function

Private Function CheckStock(ByVal lngRow As Long, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean, Optional ByVal blnCurPati As Boolean) As Boolean
'功能：根据库存检查参数检查发送药品的库存
'参数：lngRow=医嘱行号
'      blnCurPati=是否只对当前病人进行汇总检查,用于发送过程中,因为是按病人提交,这时重新提取的库存是准确的
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = TheStockCheck(Val(.TextMatrix(lngRow, COL_执行科室ID)), .TextMatrix(lngRow, COL_诊疗类别))
        bln分批 = Val(.TextMatrix(lngRow, COL_药房分批)) = 1
        bln时价 = Val(.TextMatrix(lngRow, COL_是否变价)) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = .TextMatrix(lngRow, COL_住院单位) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品总量:住院包装
            If .TextMatrix(lngRow, COL_诊疗类别) = "7" Then
                '中药药房单位按不可分零处理:每付
                If Val(.TextMatrix(lngRow, COL_可否分零)) = 0 Then
                    dbl总量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量))
                    dbl总量 = dbl总量 / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装))
                Else
                    dbl总量 = IntEx(Val(.TextMatrix(lngRow, COL_单量)) / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装)))
                    dbl总量 = dbl总量 * Val(.TextMatrix(lngRow, COL_总量))
                End If
            Else
                dbl总量 = Val(.TextMatrix(lngRow, COL_总量))
            End If
            
            '当前可用库存:住院包装,减去前面相同药品要发送的库存
            For i = lngRow - 1 To .FixedRows Step -1
                If blnCurPati And Val(.TextMatrix(i, COL_病人ID)) = Val(.TextMatrix(lngRow, COL_病人ID)) Or Not blnCurPati Then
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_收费细目ID)) = Val(.TextMatrix(lngRow, COL_收费细目ID)) _
                            And Val(.TextMatrix(i, COL_执行科室ID)) = Val(.TextMatrix(lngRow, COL_执行科室ID))
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_诊疗类别) = "7" Then
                            '中药药房单位按不可分零处理:每付
                            If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                dbl已发库存 = dbl已发库存 + _
                                    Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) _
                                    / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                            Else
                                dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量)) _
                                    * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                            End If
                        Else
                            dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量))
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            dbl可用库存 = Val(.TextMatrix(lngRow, COL_库存))
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                            strTmp = "药房分批或时价药品""" & .TextMatrix(lngRow, COL_规格) & """：" & vbCrLf & vbCrLf & _
                                "在" & .TextMatrix(lngRow, COL_执行科室) & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "药房分批或时价药品""" & .TextMatrix(lngRow, COL_规格) & """库存不足：" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    Else
                        If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                            strTmp = "药品""" & .TextMatrix(lngRow, COL_规格) & """：" & vbCrLf & vbCrLf & _
                                "在" & .TextMatrix(lngRow, COL_执行科室) & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "药品""" & .TextMatrix(lngRow, COL_规格) & """库存不足：" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    End If
                    If .Cell(flexcpData, lngRow, COL_规格) <> "" Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "你可以在发送清单中选择该药品具有足够库存的其它规格。"
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该药品吗？"
                    End If
                    
                    strTmp = "病人" & .TextMatrix(lngRow, COL_姓名) & "：" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckStock = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function CheckDrugStorage(ByVal lngRow As Long, Optional bln存储库房提示 As Boolean) As Boolean
'功能：根据库存检查参数检查发送药品的存储库房
'参数：lngRow=医嘱行号
'      bln存储库房提示=是否继续提示
'返回：根据提示，是否对选择状态进行了处理
    Dim lng药品ID As Long, lng执行科室ID As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim strTmp As String
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        '如果本来就未勾选，则不检查
        If .Cell(flexcpData, lngRow, COL_选择) = 1 Then Exit Function
        '启用了置换药房的才检查
        If mbln药房置换 = False Then Exit Function
        '获取药品ID
        lng药品ID = Val(.TextMatrix(lngRow, COL_收费细目ID))
        If lng药品ID = 0 Then Exit Function
        lng执行科室ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        If lng执行科室ID = 0 Then Exit Function
        strSQL = "select 1 from 收费执行科室 where 收费细目ID = [1]  And Nvl(病人来源,2) = 2 And 执行科室id = [2] And Nvl(开单科室id, [3]) = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDrugStorage", lng药品ID, lng执行科室ID, Val(.TextMatrix(lngRow, COL_开嘱科室ID)))
        
        If rsTmp.RecordCount > 0 Then Exit Function
        strTmp = "库房""" & .TextMatrix(lngRow, COL_执行科室) & """中没有存储药品""" & .TextMatrix(lngRow, COL_规格) & """"
        strTmp = "病人" & .TextMatrix(lngRow, COL_姓名) & "：" & vbCrLf & vbCrLf & strTmp
        
        .Redraw = flexRDDirect:
        Call .ShowCell(lngRow, COL_选择)
        Screen.MousePointer = 0
        '勾了不再提示
        If bln存储库房提示 = True Then
            vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
            If vMsg = vbIgnore Then bln存储库房提示 = False
        End If
       
        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
        CheckDrugStorage = True
    
        Screen.MousePointer = 11
        .Refresh: .Redraw = flexRDNone
    End With
End Function

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    rsTotal As ADODB.Recordset, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean) As Boolean
'功能：发送过程中时，对非药嘱药品及跟踪在用的卫材计价进行库存检查(累计检查)
'参数：lngRow=医嘱行号
'      dbl数量=已计算好的计价数量(售价单位)
'      rsTotal=当前病人前面已累计发送的计价药品或卫材数量(售价单位)
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = TheStockCheck(lng库房ID, rsPrice!类别)
        bln分批 = NVL(rsPrice!分批, 0) = 1
        bln时价 = NVL(rsPrice!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = NVL(rsPrice!住院单位, NVL(rsPrice!计算单位)) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品或卫材总量:住院包装
            dbl总量 = Format(dbl数量 / NVL(rsPrice!住院包装, 1), "0.00000")
            
            '当前可用库存:住院包装,减去前面相同药品医嘱要发送的库存
            If InStr(",5,6,7,", rsPrice!类别) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_病人ID)) = Val(.TextMatrix(lngRow, COL_病人ID)) Then
                        blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                        If blnDo Then
                            blnDo = Val(.TextMatrix(i, COL_收费细目ID)) = rsPrice!ID And Val(.TextMatrix(i, COL_执行科室ID)) = lng库房ID
                        End If
                        If blnDo Then
                            blnDo = .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing
                        End If
                        If blnDo Then
                            If .TextMatrix(i, COL_诊疗类别) = "7" Then
                                '中药药房单位按不可分零处理:每付
                                If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                    dbl已发库存 = dbl已发库存 + _
                                        Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) _
                                        / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                                Else
                                    dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量)) _
                                        * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                                End If
                            Else
                                dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量))
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
            '计价部份要发送的累计数量
            rsTotal.Filter = "项目ID=" & rsPrice!ID & " And 库房ID=" & lng库房ID
            Do While Not rsTotal.EOF
                dbl已发库存 = dbl已发库存 + Format(rsTotal!数量 / NVL(rsPrice!住院包装, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl可用库存 = Format(GetStock(rsPrice!ID, lng库房ID, 2), "0.00000")
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的分批或时价计价项目：" & vbCrLf & vbCrLf & _
                                """" & rsPrice!名称 & """在" & sys.RowValue("部门表", lng库房ID, "名称") & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的分批或时价计价项目""" & rsPrice!名称 & """库存不足：" & _
                                vbCrLf & vbCrLf & sys.RowValue("部门表", lng库房ID, "名称") & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    Else
                        If InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0 Then
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的计价项目：" & vbCrLf & vbCrLf & _
                                """" & rsPrice!名称 & """在" & sys.RowValue("部门表", lng库房ID, "名称") & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的计价项目""" & rsPrice!名称 & """库存不足：" & _
                                vbCrLf & vbCrLf & sys.RowValue("部门表", lng库房ID, "名称") & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该医嘱吗？"
                    End If
                    strTmp = "病人" & .TextMatrix(lngRow, COL_姓名) & "：" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '如果未提示或要发送,加入累计发送数量
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, COL_相关ID))
            Else
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!项目ID = rsPrice!ID
            rsTotal!库房ID = lng库房ID
            rsTotal!数量 = dbl数量
            rsTotal.Update
        End If
    End With
End Function

Private Sub DeleteDrugRow(rsSend As ADODB.Recordset, ByVal lngRow As Long, lngDel相关ID As Long)
'功能：删除对应的药品行,用于药品停用或因其它原因找不到有效规格时
'返回：lngDel相关ID-需要同时删除的其它相关医嘱标识
    Dim strMsg As String
    
    With vsAdvice
        If rsSend!诊疗类别 = "7" Then
            strMsg = "该中草药对应的中药配方无法发送：" & vbCrLf & vbCrLf & "　　" & NVL(rsSend!医嘱内容)
        Else
            strMsg = "该药品(及一并给药的其他药品)无法发送：" & vbCrLf & vbCrLf & "　　" & NVL(rsSend!医嘱内容)
        End If
        strMsg = strMsg & vbCrLf & vbCrLf & "没有发现有效的药品规格信息，该药品可能已经被停用或不能用于住院病人。"
        strMsg = strMsg & vbCrLf & "请先到药品目录管理中处理，按[确定]继续处理其他医嘱。"
        .Redraw = flexRDDirect
        Call .ShowCell(lngRow, COL_选择)
        Screen.MousePointer = 0
        MsgBox strMsg, vbInformation, gstrSysName
        
        Screen.MousePointer = 11
        lngDel相关ID = NVL(rsSend!相关ID, 0)
        Call DeleteCurRow(lngRow, rsSend!相关ID)
        .Refresh: .Redraw = flexRDNone
    End With
End Sub

Private Sub SeekMatchDrug(rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset, ByVal dbl总量 As Double, vBookMark As Variant, Optional strList As String)
'功能：根据药品的多个规格定位缺省合适的规格,并设置相关药品信息到表格中
'参数：rsSend=要发送的医嘱信息
'      rsDrug=药品信息
'      dbl总量=要发送的药品总量,为0时表示还未计算出来
'      vBookMark=返回用于定位规格位置的书签
'      strList=返回有效可供选择的规格,用于设置下拉框数据
    Dim vPreBookMark As Variant
    Dim lng倍数 As Long
        
    vPreBookMark = 0
    If Not rsDrug.EOF And Not rsDrug.BOF Then
        vPreBookMark = rsDrug.Bookmark
    End If
    
    rsDrug.MoveFirst
    vBookMark = 0: strList = ""
    Do While Not rsDrug.EOF
        '排开停用的药品
        If NVL(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!服务对象, 0)) > 0 Then
            If CInt(NVL(rsSend!单次用量, 0)) <> 0 And (NVL(rsDrug!库存, 0) > dbl总量 Or NVL(rsDrug!库存, 0) = dbl总量 And dbl总量 <> 0) Then
                '寻找剂量单位为单量的最小倍数的规格
                If rsDrug!剂量系数 / rsSend!单次用量 = Int(rsDrug!剂量系数 / rsSend!单次用量) Then
                    If rsDrug!剂量系数 / rsSend!单次用量 < lng倍数 Or lng倍数 = 0 Then
                        vBookMark = rsDrug.Bookmark
                        lng倍数 = rsDrug!剂量系数 / rsSend!单次用量
                    End If
                End If
            End If
            strList = strList & "|#" & rsDrug!药品ID & ";" & rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "") & _
                vbTab & IIF(InStr(GetInsidePrivs(p住院医嘱发送), "显示药品库存") = 0, _
                    IIF(NVL(rsDrug!库存, 0) > 0, "有库存", "无库存"), "库存:" & NVL(rsDrug!库存, 0) & rsDrug!住院单位)
        End If
        rsDrug.MoveNext
    Loop
    If vBookMark = 0 Then
        rsDrug.MoveFirst
        Do While Not rsDrug.EOF
            If NVL(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!服务对象, 0)) > 0 Then
                If NVL(rsDrug!库存, 0) > dbl总量 Or NVL(rsDrug!库存, 0) = dbl总量 And dbl总量 <> 0 Then
                    vBookMark = rsDrug.Bookmark: Exit Do
                End If
                '确保能够选到一个未停用的规格；如果可用规格库存都为0，且rsDrug原有位置的记录是停用规格，这会导致界面加载停用规格，不能被发送
                vBookMark = rsDrug.Bookmark
            End If
            rsDrug.MoveNext
        Loop
    End If
    strList = Mid(strList, 2)
    
    If vBookMark = 0 And vPreBookMark <> 0 Then '没找到时恢复原有位置
        rsDrug.Bookmark = vPreBookMark
    End If
End Sub

Private Function Calc总量次数时间(dbl总量 As Double, lng次数 As Long, str分解时间 As String, ByVal strEnd As String, rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset, ByVal bln口服 As Boolean) As Boolean
'功能：对长期成药医嘱计算总量,执行次数,执行时间分解
'参数：rsDrug=包含药品规格的相关信息
'      rsSend=包含当前药品医嘱的相关信息
'      strEnd=本次发送的结束时间
'返回：dbl总量=住院包装
'      lng次数=执行次数(即为给药途径的执行次数)
'      str分解时间=具体的执行时间分解
    Dim datBegin As Date, datEnd As Date, strPause As String
    Dim datTmp As Date
    Dim intType As Integer
    Dim varArr As Variant
    Dim i As Long
    Dim strTmp As String
    Dim strToday As String
    Dim strTimRange As String
    
    Dim lng次数Tmp As Long 'bln口服 true 当程已经控制了分解时间时，在计算总量里不再重算 lng次数 str分解时间，(90797)
    Dim str分解时间Tmp As String
    
    On Error GoTo errH
    
    '当前医嘱的暂停时间段:"暂停时间,开始时间;...."
    If rsSend!医嘱状态 <> 1 Then
        strPause = GetAdvicePause(rsSend!ID, Val(rsSend!组ID & ""))
    End If
    
    '当前医嘱的发送计算时间段
    datBegin = rsSend!开始执行时间
    If Not IsNull(rsSend!上次执行时间) Then
        datBegin = Calc本周期开始时间(rsSend!开始执行时间, rsSend!上次执行时间, NVL(rsSend!频率间隔, 0), rsSend!间隔单位 & "")
        
        '本周期内已执行的时间不再计算,这里通过暂停方式来处理
        '备用医嘱不支持暂停
        If mlngRefModld = 0 Then
            strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!上次执行时间, "yyyy-MM-dd HH:mm:ss")
            If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
        End If
    End If
    
    If bln口服 And Val(rsDrug!是否摆药 & "") = 1 Then strEnd = Split(strEnd, " ")(0) & " " & mstrEndPoint
    
    datEnd = CDate(strEnd)
    If Not IsNull(rsSend!执行终止时间) Then
        If rsSend!执行终止时间 < CDate(strEnd) Then
            datEnd = rsSend!执行终止时间
        End If
    End If
    
    '药品长嘱从当天开始发送 只能准确控制 给药周期长度为 1 天 医嘱，一旦采取这种方式发送后，以前未发送的将不再被读取
    If mlngRefModld = 0 And chk药品期效(0).value = 1 And chkSendNow.value = 1 Then
        
        strToday = Format(mdatCurr, "yyyy-MM-dd 00:00:00")
        
        If rsSend!间隔单位 & "" = "天" And Val(NVL(rsSend!频率间隔, 0)) = 1 Or InStr(",小时,分钟,", "," & rsSend!间隔单位 & ",") > 0 Then
            datTmp = CDate(strToday)
            If datTmp > datBegin Then
                If rsSend!间隔单位 & "" = "小时" Then
                    datBegin = Calc每小时计算从当天发送开始时间(datBegin, datTmp, NVL(rsSend!频率间隔, 0))
                Else
                    datBegin = datTmp
                End If
            End If
        Else
            intType = 1
        End If
    End If
     
    '先按正常发送时间段计算分解时间及次数
    str分解时间 = Calc段内分解时间(datBegin, datEnd, strPause, NVL(rsSend!执行时间方案), NVL(rsSend!频率次数, 0), NVL(rsSend!频率间隔, 0), NVL(rsSend!间隔单位), rsSend!开始执行时间)
    
    If intType = 1 And str分解时间 <> "" Then
        If Format(datBegin, "yyyy-MM-dd HH:mm:ss") < strToday Then
            varArr = Split(str分解时间, ",")
            For i = UBound(varArr) To 0 Step -1
                If strToday <= CStr(varArr(i)) Then
                    strTmp = strTmp & "," & varArr(i)
                Else
                    Exit For
                End If
            Next
            str分解时间 = Mid(strTmp, 2)
        End If
    End If
    
    If Val(rsSend!医嘱期效 & "") = 0 And Val(rsSend!紧急标志 & "") = 1 And Val(rsSend!审核状态 & "") = 1 Then
        datBegin = rsSend!开始执行时间
        datEnd = DateAdd("d", 1, datBegin)
        strTimRange = Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(datEnd, "yyyy-MM-dd HH:mm:ss")
        str分解时间 = GetTimPointsInRange(strTimRange, str分解时间)
    End If
    
    If str分解时间 = "" Then
        dbl总量 = 0
        lng次数 = 0
        Calc总量次数时间 = True
        Exit Function
    End If

    lng次数 = UBound(Split(str分解时间, ",")) + 1
    
    If NVL(rsSend!诊疗类别) = "7" Then
        '中药配方长嘱
        dbl总量 = lng次数
    Else
        If bln口服 And Val(rsDrug!是否摆药 & "") = 1 Then
            lng次数Tmp = lng次数
            str分解时间Tmp = str分解时间
        End If
        
        '西药、中成药：再按药品分零特性计算总量(按住院单位),这时次数和分解时间可能增加
        dbl总量 = Calc发送药品总量( _
            rsSend!开始执行时间, lng次数, str分解时间, rsSend!单次用量, _
            rsDrug!剂量系数, rsDrug!住院包装, NVL(rsSend!可否分零, NVL(rsDrug!可否分零, 0)), _
            NVL(rsSend!执行终止时间, CDate("3000-01-01")), strPause, NVL(rsSend!执行时间方案), _
            rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位 & "", mblnLimit, NVL(rsSend!首次用量, 0), NVL(rsSend!上次执行时间, CDate(0)))
            
        If bln口服 And Val(rsDrug!是否摆药 & "") = 1 Then
            lng次数 = lng次数Tmp
            str分解时间 = str分解时间Tmp
        End If
    End If
    
    Calc总量次数时间 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetWhere(ByVal bytMode As Byte, ByRef bln会诊 As Boolean)
'功能：返回医嘱校对或发送的可操作医嘱条件（如果没有权限时，只能处理当前操作人员的所属病区的所有科室或者会诊科室下达的医嘱）
'参数：0-校对，1=发送
'       bln会诊 出参数，是否要读取会诊医嘱IDs
    Dim strTmp As String
    Dim blnDo As Boolean
    
    If bytMode = 0 Then
        blnDo = InStr(GetInsidePrivs(p住院医嘱发送), "全院医嘱校对") = 0
    Else
        blnDo = InStr(GetInsidePrivs(p住院医嘱发送), "全院医嘱发送") = 0
    End If
    
    If blnDo Then
        If gbln会诊科室下达医嘱处理 Then
            strTmp = " And (A.开嘱科室ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([4])) E) and nvl(a.会诊医嘱id,0)=0 or instr(','||[11]||',',','||nvl(a.会诊医嘱id,0)||',')>0)"
            bln会诊 = True
        Else
            strTmp = " And A.开嘱科室ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([4])) E) "
        End If
    End If
    
    GetWhere = strTmp
End Function

Private Function CheckSendPrivs(ByVal lng医嘱ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng会诊医嘱ID As Long) As Boolean
'功能：判断当前医嘱行的开嘱科室是否是本病区所属科室
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strDepts As String
    
    strDepts = GetUser科室IDs(True)   '当前操作人员的所属病区的所有科室'
    
    If gbln会诊科室下达医嘱处理 Then
        strSQL = " Select 1 From 病人医嘱记录 D Where D.ID = [1] And D.开嘱科室ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([2])) E) And nvl(D.会诊医嘱id,0)=0" & _
            " union all Select 1 From 病人医嘱记录 D Where D.ID = [3] And D.开嘱科室ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([2])) E)"
    Else
        strSQL = " Select 1 From 病人医嘱记录 D Where D.ID = [1] And D.开嘱科室ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([2])) E)"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, strDepts, lng会诊医嘱ID)
    CheckSendPrivs = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadAdviceSend(ByVal str病人IDs As String, ByVal str主页IDs As String, ByVal strEnd As String, ByVal str给药IDs As String, ByVal str类别s As String, ByVal lngModle As Long, ByVal str病人科室IDs As String)
'功能：按病人读取医嘱发送清单
    Dim rsSend药品 As ADODB.Recordset, rsSend其他 As ADODB.Recordset
    Dim arrPati As Variant, arrPatiPage As Variant, arrPatiDept As Variant
    
    Dim bln药品 As Boolean, bln其他 As Boolean, blnOnePati As Boolean, bln可接收病区 As Boolean
    Dim lng药房ID As Long, lng执行科室ID As Long, str执行科室 As String, str诊疗类别 As String
    
    Dim strSQL药品 As String, strSQL其他 As String, str主要条件 As String
    Dim str药房条件 As String, str给药途径 As String, str药房置换 As String
    Dim lng病人数 As Long, lng病人ID As Long, str科室 As String, bln品种药品 As Boolean, lng单量数 As Long
    Dim i As Long, k As Long, datEnd As Date
    Dim str临嘱 As String, str长嘱 As String, str医嘱期效 As String
    Dim bln时价提示 As Boolean, bln库存提示 As Boolean, bln默认发送 As Boolean, bln存储库房提示 As Boolean
    Dim strDepts As String, strTmp1 As String, strTmp2, strtmp3 As String
    Dim str备用医嘱 As String
    Dim str输液药品排除 As String '是否可以发送输液药品
    Dim strAdDrugIDs As String
    Dim str会诊医嘱IDs As String
    Dim bln会诊 As Boolean
    Dim str特殊医嘱排除 As String
    
    mstrNoneIDs = ","
    mstrAdDrugIDs = ""
    
    Screen.MousePointer = 11
    stbThis.Panels(3).Text = "" ': Call Form_Resize
    Call GetAdvicePause(0) '清除此方法中的缓存
    bln时价提示 = True: bln库存提示 = True: bln默认发送 = True: bln存储库房提示 = True
    bln药品 = chk药品期效(0).value = 1 Or chk药品期效(1).value = 1
    bln其他 = chk其他期效(0).value = 1 Or chk其他期效(1).value = 1
    
    If chk药品期效(0).value = 1 Or chk其他期效(0).value = 1 Then
        lblInfo.Caption = "本次发送：" & IIF(chk药品期效(1).value = 1 Or chk其他期效(1).value = 1, "临时医嘱，", "") & "长期医嘱，结束时间：" & strEnd
    Else
        lblInfo.Caption = "本次发送：临时医嘱"
    End If
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    
    With vsAdvice
        .Rows = .FixedRows '有删除行功能
        If mblnOnePati Then
            .ColHidden(COL_姓名) = True
            .ColHidden(COL_住院号) = True
            .ColHidden(COL_床号) = True
            .ColHidden(COL_费别) = True
        End If
        .ColHidden(COL_科室) = True
        .ColHidden(COL_婴儿) = True
        .ColHidden(COL_单量) = False
        .ColHidden(COL_单量单位) = False
        .ColHidden(COL_首次时间) = chk药品期效(1).value = 1 Or chk其他期效(1).value = 1
        .ColHidden(COL_末次时间) = chk药品期效(1).value = 1 Or chk其他期效(1).value = 1
        
        .ColHidden(COL_规格) = Not bln药品
        .ColHidden(COL_执行性质) = Not bln药品
        If chk药品期效(0).value = 0 And chk其他期效(0).value = 0 Or chk药品期效(1).value = 0 And chk其他期效(1).value = 0 Then
            .ColHidden(COL_医嘱期效) = True
        Else
            .ColHidden(COL_医嘱期效) = False
        End If
    End With
    Me.Refresh
    
    strDepts = GetUser科室IDs(True)    '当前操作人员的所属病区的所有科室
        
    str主要条件 = " And A.开始执行时间 is Not NULL And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3"
    '婴儿的处理
    If opt婴儿(1).value Or opt婴儿(2).value Then
        str主要条件 = str主要条件 & " And Nvl(A.婴儿,0)" & IIF(opt婴儿(1).value, "=0", ">0")
    End If
    str主要条件 = str主要条件 & IIF(Not mbln医技后续, " And A.前提ID is Null", "")
    
    If opt状态(1).value Then    '已校对
        '当前操作员病区包含的科室的所有医生
        str主要条件 = str主要条件 & GetWhere(1, bln会诊)
    Else
        If opt状态(0).value Then    '新开
            str主要条件 = str主要条件 & " And Exists(" & _
                "Select M.姓名 From 人员表 M,执业类别 N" & _
                " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1)," & _
            "2,Substr(A.开嘱医生,1,Decode(Instr(A.开嘱医生,'/'),0,length(A.开嘱医生),Instr(A.开嘱医生,'/')-1))," & _
            "Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
                " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')" & _
                " )"
                
            str主要条件 = str主要条件 & GetWhere(0, bln会诊)
        Else '两者
            str主要条件 = str主要条件 & " And (Nvl(A.医嘱状态,0)<>1 Or A.医嘱状态=1 And Exists(" & _
                "Select M.姓名 From 人员表 M,执业类别 N" & _
                " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1)," & _
            "2,Substr(A.开嘱医生,1,Decode(Instr(A.开嘱医生,'/'),0,length(A.开嘱医生),Instr(A.开嘱医生,'/')-1))," & _
            "Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
                " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')" & _
                " ))"
                
            strTmp1 = GetWhere(0, bln会诊)
            strTmp2 = GetWhere(1, bln会诊)
            If Not (strTmp1 = "" And strTmp2 = "") Then
                str主要条件 = str主要条件 & " And (Nvl(A.医嘱状态,0)<>1" & strTmp2 & " Or A.医嘱状态=1" & strTmp1 & ")"
            End If
        End If
    End If

    
    '置换的药房ID
    If bln药品 Then
        
        lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
        str药房置换 = Decode药房置换
        
        '只发送指定药房的药品:药房置换之后的为准
        If lng药房ID <> 0 Then
            str药房条件 = "Select ID From 病人医嘱记录 X" & _
                " Where 诊疗类别 IN('5','6','7') And (X.相关ID=A.相关ID Or X.相关ID=A.ID)" & _
                " And 执行科室ID+0=[3] And 病人ID=[2]"
            str药房条件 = " And Exists(" & str药房条件 & ")"
        End If
        
        '允许的给药途径部份(关联对应的成药)
        If str给药IDs <> "" Then
            str给药途径 = "Select ID From 病人医嘱记录 X" & _
                " Where 诊疗类别='E' And (X.ID=A.ID Or X.ID=A.相关ID)" & _
                " And 病人ID=[2] And 诊疗项目ID+0 IN(" & str给药IDs & ")"
            str给药途径 = " And Exists(" & str给药途径 & ")"
        End If
        
        str长嘱 = ""
        str临嘱 = ""
        str医嘱期效 = ""
        '不同期效的条件
        If chk药品期效(0).value = 1 Then    '长嘱
            strTmp1 = _
                "A.开始执行时间<=[1] And (A.上次执行时间<[1] Or A.上次执行时间 is NULL)" & _
                " And (A.执行终止时间>A.上次执行时间 Or A.执行终止时间 is NULL Or A.上次执行时间 Is NULL)" & _
                " And (A.执行终止时间>A.开始执行时间 Or A.执行终止时间 is NULL) And A.医嘱期效=0"
                
            If opt状态(1).value Then    '已校对
                str长嘱 = strTmp1 & " And Nvl(A.医嘱状态,0) Not IN(-1,1,2,4)"
            Else
                If opt状态(0).value Then    '新开(不管结束时间，发送时按开始执行时间大于指定的发送结束时间才发送)
                    str长嘱 = "A.医嘱状态=1 And A.医嘱期效=0"
                Else '两者
                    str长嘱 = "(A.医嘱状态=1 And A.医嘱期效=0 Or (" & strTmp1 & " And Nvl(A.医嘱状态,0) Not IN(-1,1,2,4)))"
                End If
            End If
        End If
        If chk药品期效(1).value = 1 Then    '临嘱
            If opt状态(1).value Then    '已校对
                str临嘱 = "Nvl(A.医嘱状态,0) Not IN(-1,1,2,4,8,9) And A.医嘱期效=1"
            Else
                If opt状态(0).value Then    '新开
                    str临嘱 = "A.医嘱状态=1 And A.医嘱期效=1"
                Else '两者
                    str临嘱 = "(A.医嘱状态=1 And A.医嘱期效=1 Or Nvl(A.医嘱状态,0) Not IN(-1,2,4,8,9) And A.医嘱期效=1)"
                End If
            End If
        End If
        If str长嘱 <> "" And str临嘱 <> "" Then '不可能同时为空
            strTmp1 = " And ((" & str长嘱 & ") Or (" & str临嘱 & "))"
            If strTmp1 = " And ((A.医嘱状态=1 And A.医嘱期效=0) Or (A.医嘱状态=1 And A.医嘱期效=1))" Then
                strTmp1 = " And A.医嘱状态=1 And A.医嘱期效 In(0,1)"
            End If
            str医嘱期效 = strTmp1
        ElseIf str长嘱 <> "" Then
            str医嘱期效 = " And " & str长嘱
        ElseIf str临嘱 <> "" Then
            str医嘱期效 = " And " & str临嘱
        End If
        
        If gblnKSSStrict Then
            If opt状态(0).value Or opt状态(2).value Then
                str医嘱期效 = str医嘱期效 & " And (A.医嘱状态<>1 Or A.医嘱状态=1 And  (Nvl(A.审核状态,0) Not in(1,3) or a.医嘱期效=0 and a.审核状态=1 and a.紧急标志=1 and (instr(',5,6,',A.诊疗类别)>0 or A.诊疗类别='E' and E.操作类型='2')))"
            End If
        End If
        
        '备用医嘱
        If lngModle = 1 Then
            str备用医嘱 = " And (NVL(a.执行频次,'无')='必要时' Or NVL(a.执行频次,'无')='需要时' And to_date([6],'yyyy-mm-dd hh24:mi') - a.开始执行时间<0.5) "
        Else
            str备用医嘱 = " And (NVL(a.执行频次,'无')<>'必要时' And NVL(a.执行频次,'无')<>'需要时') "
        End If
        
        '读取发送明细:(未排除正常的治疗医嘱等)
        '叮嘱不发送(给药途径,用法,煎法可能为),但这里先读取出来
        strSQL药品 = "Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " A.诊疗类别,A.诊疗项目ID,E.名称 as 诊疗项目,A.收费细目ID,A.婴儿,B.入院日期," & _
            " A.病人ID,A.主页ID,B.住院号,B.出院病床 as 床号,D.名称 as 科室,A.姓名,A.性别,A.年龄,B.费别,B.险类," & _
            IIF(mlngRefModld = 1, "Decode(Sign(A.开始执行时间-to_date([6],'yyyy-mm-dd hh24:mi')),-1,to_date([6],'yyyy-mm-dd hh24:mi'),A.开始执行时间) as 开始执行时间", "A.开始执行时间") & ",A.上次执行时间,A.医嘱内容,A.天数,A.总给予量,A.单次用量,E.计算单位,A.执行终止时间," & _
            " A.执行频次,Decode(A.执行频次,'必要时',1,A.频率次数) As 频率次数,Decode(A.执行频次,'必要时',1,A.频率间隔) As 频率间隔,Decode(A.执行频次,'必要时','天',A.间隔单位) as 间隔单位,A.医生嘱托," & _
            " Decode(A.执行频次,'必要时',[7],A.执行时间方案) As 执行时间方案,b.病人性质,a.开嘱时间,e.执行分类,e.操作类型," & _
            " [5] as 病人病区ID,A.病人科室ID,A.开嘱科室ID,A.开嘱医生," & IIF(mblnAutoVerify, "s.操作时间 as 新开操作时间,", "") & _
            " A.可否分零,A.计价特性,A.执行性质,A.执行标记," & str药房置换 & " as 执行科室ID,Nvl(F.名称,Decode(Nvl(A.执行性质,0),5,'-')) as 执行科室,A.摘要,A.医嘱状态,A.医嘱期效,A.首次用量,E.计算方式,e.执行安排,g.毒理分类,a.紧急标志,a.审核状态,a.用药理由,a.会诊医嘱ID" & _
            " From 病人医嘱记录 A,病案主页 B,病人信息 C,部门表 D,诊疗项目目录 E,部门表 F,药品特性 G,病人医嘱记录 X" & IIF(mblnAutoVerify, ",病人医嘱状态 S", "") & _
            " Where A.病人ID=[2] And A.病人ID=C.病人ID And B.出院科室ID=D.ID" & IIF(mblnAutoVerify, " And  s.医嘱ID=a.ID And s.操作类型=1 ", "") & _
            " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.主页ID = C.主页ID" & _
            " And A.相关ID=X.ID(+) And A.诊疗项目ID=E.ID And " & str药房置换 & "=F.ID(+) and e.id=g.药名id(+)" & _
            " And A.诊疗类别 IN('5','6','7','E')" & str主要条件 & str药房条件 & str给药途径 & str医嘱期效 & str备用医嘱 & _
            " And (B.婴儿科室ID is null or B.婴儿科室ID is not null and B.婴儿病区ID=[5] and NVL(A.婴儿,0)<>0 or B.婴儿科室ID is not null and B.婴儿病区ID<>[5] and NVL(A.婴儿,0)=0) "
        strtmp3 = strSQL药品
    End If
    
    If bln其他 Then
    
        lng执行科室ID = Val(cmd执行科室.Tag)
        
        '执行科室(以主要医嘱为准)
        If lng执行科室ID <> 0 Then
            '一般项目以及手术组合,检查组合;检验项目(组合)
            str执行科室 = _
                " And Exists(" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Null And (X.ID=A.ID Or X.ID=A.相关ID)" & _
                " And 病人ID=[2] And 执行科室ID+0=[3]" & _
                " Union ALL " & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Not Null And 诊疗类别='C' And (X.相关ID=A.相关ID Or X.相关ID=A.ID)" & _
                " And 病人ID=[2] And 执行科室ID+0=[3])"
        End If
        
        '允许的诊疗类别部份(以主要医嘱为准)
        If str类别s <> "" Then
            '一般项目以及手术组合,检查组合;检验项目(组合)
            str诊疗类别 = _
                " And Exists(" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Null And (X.ID=A.ID Or X.ID=A.相关ID)" & _
                " And 病人ID=[2] And 诊疗类别 IN(" & str类别s & ")" & _
                " Union ALL" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Not Null And 诊疗类别='C' And (X.相关ID=A.相关ID Or X.相关ID=A.ID)" & _
                " And 病人ID=[2] And 诊疗类别 IN(" & str类别s & "))"
        End If
        
        '排开给药途径，中药煎法、用法
        strSQL其他 = _
            " And Not(A.诊疗类别='E' And A.相关ID is Not NULL And E.操作类型='3')" & _
            " And Not Exists(Select ID From 病人医嘱记录 X" & _
            " Where 诊疗类别 IN('5','6','7') And X.相关ID=A.ID" & _
            " And 病人ID=[2])"
            
        '叮嘱不发送(手术,检查,检验不允许为叮嘱,采集方法、输血途径可能为叮嘱，这种要读出来,但最后不发送)
        '如果是不校对即发送模式，无需执行的不发送(手术,检查,检验不允许为叮嘱,采集方法、输血途径可能为叮嘱)
        If Not mblnAutoVerify Or opt状态(1).value Then
            strSQL其他 = strSQL其他 & " And A.诊疗类别 is Not Null And (Nvl(A.执行性质,0)<>0 Or A.诊疗类别='E' And E.操作类型 IN('6','8'))"
        ElseIf opt状态(2).value Then
            strSQL其他 = strSQL其他 & " And (A.医嘱状态=1 Or Nvl(A.医嘱状态,0)<>1 And A.诊疗类别 is Not Null" & _
                        " And (Nvl(A.执行性质,0)<>0 Or A.诊疗类别='E' And E.操作类型 IN('6','8')))"
        End If
                
                
        '不同期效的条件
        str长嘱 = ""
        str临嘱 = ""
        str医嘱期效 = ""
        If chk其他期效(0).value = 1 Then    '长嘱
            '术前术后医嘱不发送(先已检查已先校对)
            '新开的互斥长嘱：护理等级,病重/危医嘱,记录入出量医嘱不发送但要读出来校对(如果没有互斥医嘱，之前没有弹出先校对，所以这里要读出来校对)
            strTmp1 = _
                "A.开始执行时间<=[1] And (A.上次执行时间<[1] Or A.上次执行时间 is NULL)" & _
                " And (A.执行终止时间>A.上次执行时间 Or A.执行终止时间 is NULL Or A.上次执行时间 Is NULL)" & _
                " And (A.执行终止时间>A.开始执行时间 Or A.执行终止时间 is NULL) And A.医嘱期效=0" & _
                " And Not(Nvl(a.诊疗类别,'自由')='H' And E.操作类型='1' And E.执行频率=2)" & _
                " And Not(Nvl(a.诊疗类别,'自由')='Z' And E.操作类型 IN('4','14','9','10','12'))"
                
            If opt状态(1).value Then    '已校对
                str长嘱 = strTmp1 & " And Nvl(A.医嘱状态,0) Not IN(-1,1,2,4)"
            Else
                If opt状态(0).value Then    '新开(不管结束时间，发送时按开始执行时间大于指定的发送结束时间才发送)
                    str长嘱 = "A.医嘱状态=1 And A.医嘱期效=0"
                Else '两者
                    str长嘱 = "(A.医嘱状态=1 And A.医嘱期效=0 Or (" & strTmp1 & " And Nvl(A.医嘱状态,0) Not IN(-1,1,2,4)))"
                End If
            End If
        End If
        If chk其他期效(1).value = 1 Then    '临嘱
            If opt状态(1).value Then    '已校对
                str临嘱 = "Nvl(A.医嘱状态,0) Not IN(-1,1,2,4,8,9) And A.医嘱期效=1"
            Else
                If opt状态(0).value Then    '新开
                    str临嘱 = "A.医嘱状态=1 And A.医嘱期效=1"
                Else '两者
                    str临嘱 = "(A.医嘱状态=1 And A.医嘱期效=1 Or Nvl(A.医嘱状态,0) Not IN(-1,2,4,8,9) And A.医嘱期效=1)"
                End If
            End If
        End If
        If str长嘱 <> "" And str临嘱 <> "" Then '不可能同时为空
            strTmp1 = " And ((" & str长嘱 & ") Or (" & str临嘱 & "))"
            If strTmp1 = " And ((A.医嘱状态=1 And A.医嘱期效=0) Or (A.医嘱状态=1 And A.医嘱期效=1))" Then
                strTmp1 = " And A.医嘱状态=1 And A.医嘱期效 In(0,1)"
            End If
            str医嘱期效 = strTmp1
        ElseIf str长嘱 <> "" Then
            str医嘱期效 = " And " & str长嘱
        ElseIf str临嘱 <> "" Then
            str医嘱期效 = " And " & str临嘱
        End If
        
         '备用医嘱
        If lngModle = 1 Then
            str备用医嘱 = " And (NVL(a.执行频次,'无')='必要时' Or NVL(a.执行频次,'无')='需要时'  And to_date([6],'yyyy-mm-dd hh24:mi') - a.开始执行时间<0.5) "
        Else
            str备用医嘱 = " And (NVL(a.执行频次,'无')<>'必要时' And NVL(a.执行频次,'无')<>'需要时') "
        End If
        
        '手术分级管理/输血分级管理
        If gbln手术分级管理 Or gbln输血分级管理 Or gbln血库系统 Then
            If opt状态(0).value Or opt状态(2).value Then
                str医嘱期效 = str医嘱期效 & " And (A.医嘱状态<>1 Or A.医嘱状态=1 And Nvl(A.审核状态,0) Not in " & IIF(gbln血库系统 = True, "(1,3,7)", "(1,3,4,5,7)") & ")"
            End If
        End If
        
        '读取发送明细:
        strSQL其他 = "Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " A.诊疗类别,G.名称 as 类别名称,A.诊疗项目ID,E.名称 as 诊疗项目,A.收费细目ID,B.入院日期," & _
            " A.婴儿,A.病人ID,A.主页ID,B.住院号,B.出院病床 as 床号,D.名称 as 科室,A.姓名,A.性别,A.年龄,B.费别,B.险类," & _
            IIF(mlngRefModld = 1, "Decode(Sign(A.开始执行时间-to_date([6],'yyyy-mm-dd hh24:mi')),-1,to_date([6],'yyyy-mm-dd hh24:mi'),A.开始执行时间) as 开始执行时间", "A.开始执行时间") & ",A.上次执行时间,A.医嘱内容,A.总给予量,A.单次用量,E.计算规则," & _
            " Decode(A.诊疗类别,'4',H.计算单位,E.计算单位) as 计算单位,A.执行终止时间,b.病人性质,a.开嘱时间,a.申请序号,e.执行分类," & _
            " A.执行频次,Decode(A.执行频次,'必要时',1,A.频率次数) As 频率次数,Decode(A.执行频次,'必要时',1,A.频率间隔) As 频率间隔,Decode(A.执行频次,'必要时','天',A.间隔单位) as 间隔单位,A.医生嘱托," & _
            " Decode(A.执行频次,'必要时',[7],A.执行时间方案) As 执行时间方案," & _
            " [5] as 病人病区ID,A.病人科室ID,A.开嘱科室ID,A.开嘱医生,A.摘要," & _
            " A.标本部位,A.检查方法,A.执行标记,A.计价特性,E.操作类型,E.试管编码,E.执行频率," & IIF(mblnAutoVerify, "s.操作时间 as 新开操作时间,", "") & _
            " A.执行性质,A.执行科室ID,Nvl(F.名称,Decode(Nvl(A.执行性质,0),5,'-')) as 执行科室,H.是否变价,I.跟踪在用,A.紧急标志,A.医嘱状态,A.医嘱期效,h.撤档时间,A.首次用量,E.计算方式,e.执行安排,a.用药理由,a.会诊医嘱ID" & _
            " From 病人医嘱记录 A,病案主页 B,病人信息 C,部门表 D,诊疗项目目录 E," & _
            " 部门表 F,诊疗项目类别 G,收费项目目录 H,材料特性 I,病人医嘱记录 X" & IIF(mblnAutoVerify, ",病人医嘱状态 S", "") & _
            " Where A.病人ID=[2] And A.病人ID=C.病人ID And B.出院科室ID=D.ID And A.收费细目ID=H.ID(+) And A.收费细目ID=I.材料ID(+)" & IIF(mblnAutoVerify, " And  s.医嘱ID=a.ID And s.操作类型=1 ", "") & _
            " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.主页ID = C.主页ID And A.相关ID=X.ID(+)" & _
            " And A.诊疗项目ID=E.ID(+) And E.类别=G.编码(+) And A.执行科室ID=F.ID(+)" & strSQL其他 & _
            " And Nvl(a.诊疗类别,'自由') Not IN('5','6','7')" & str主要条件 & str执行科室 & str诊疗类别 & str医嘱期效 & str备用医嘱 & _
            " And (B.婴儿科室ID is null or B.婴儿科室ID is not null and B.婴儿病区ID=[5] and NVL(A.婴儿,0)<>0 or B.婴儿科室ID is not null and B.婴儿病区ID<>[5] and NVL(A.婴儿,0)=0)" & _
            " Order by A.医嘱期效,A.婴儿,组号,组ID,A.序号"
    End If
    
    On Error GoTo errH
    arrPati = Split(str病人IDs, ",")
    arrPatiPage = Split(str主页IDs, ",")
    arrPatiDept = Split(str病人科室IDs, ",")
    blnOnePati = UBound(arrPati) = 0
    datEnd = CDate(IIF(strEnd = "", "1990-01-01", strEnd))
    
    If cboUnit.ListIndex <> -1 Then
        bln可接收病区 = (mstrInfDepIDs = "" Or InStr("," & mstrInfDepIDs & ",", "," & Val(cboUnit.ItemData(cboUnit.ListIndex)) & ",") > 0)
    End If
    For k = 0 To UBound(arrPati)
        If bln会诊 Then str会诊医嘱IDs = Get会诊医嘱IDs(Val(arrPati(k)), arrPatiPage(k), strDepts)
        If bln药品 Then
            str输液药品排除 = ""
            '排除发送到配置中心的输液类医嘱,如果其中有一个药不发送到输液配置中心，则允许在这里发送(自备药、自取药、离院带药、不取药等特殊性质的药品的允许在这里发送、配置中心不接收的给药途径药品可以在这里发送、期效不满足的也可以在这里发送)
            If lngModle = 0 Then
                If gstr输液配置中心 <> "" Then
                    If bln可接收病区 Then
                        str特殊医嘱排除 = Get输液类医嘱(Val(arrPati(k)), arrPatiPage(k), 0)
                        str输液药品排除 = " and instr(','||[10]|| ',',','||Nvl(A.相关ID,A.ID)||',')=0"
                    Else
                        '如果科室启用了配置中心，则排除所有的静配流程的药品，如果科室没有启用，但全院启用了，则只排除静脉营养医嘱
                        str输液药品排除 = " And NVL(E.执行标记,0)<>2 And (Not Exists(Select 1 From 诊疗项目目录 Y Where X.诊疗项目id = y.Id And NVL(Y.执行标记,0)=2) OR x.诊疗项目id is null)"
                    End If
                End If
            End If
            strAdDrugIDs = ""
            mstrNoneIDs = mstrNoneIDs & GetNoneSendID(Val(arrPati(k)), arrPatiPage(k), 2, , , strAdDrugIDs) & ","
            If strAdDrugIDs <> "" Then
                mstrAdDrugIDs = IIF(mstrAdDrugIDs = "", "", mstrAdDrugIDs & ",") & strAdDrugIDs
            End If
            strSQL药品 = strtmp3
            strSQL药品 = strSQL药品 & str输液药品排除 & " Order by A.医嘱期效,A.婴儿,组号,组ID,A.序号"
            Set rsSend药品 = zlDatabase.OpenSQLRecord(strSQL药品, Me.Caption, datEnd, Val(arrPati(k)), lng药房ID, strDepts, mlng病区ID, Format(dkpExecTime.value, "YYYY-MM-DD HH:mm"), Format(dkpExecTime.value, "HH:mm"), _
                gstr输液配置中心, mstr配液给药途径, str特殊医嘱排除, str会诊医嘱IDs)
        End If
        If bln其他 Then
            Set rsSend其他 = zlDatabase.OpenSQLRecord(strSQL其他, Me.Caption, datEnd, Val(arrPati(k)), lng执行科室ID, strDepts, mlng病区ID, Format(dkpExecTime.value, "YYYY-MM-DD HH:mm"), Format(dkpExecTime.value, "HH:mm"), , , , str会诊医嘱IDs)
        End If
        
        '先显示新开的
        If mblnAutoVerify Then
            If bln药品 Then
                rsSend药品.Filter = "医嘱状态=1"
                If rsSend药品.RecordCount > 0 Then
                    Call LoadAdviceSendDrug(blnOnePati, strEnd, rsSend药品, lng病人数, str科室, bln品种药品, bln时价提示, bln库存提示, bln默认发送, lng病人ID, bln存储库房提示)
                End If
            End If
            If bln其他 Then
                rsSend其他.Filter = "医嘱状态=1"
                If rsSend其他.RecordCount > 0 Then
                    Call LoadAdviceSendOther(blnOnePati, strEnd, rsSend其他, lng病人数, str科室, lng单量数, lng病人ID)
                End If
            End If
        End If
        
        If bln药品 Then
            If mblnAutoVerify Then rsSend药品.Filter = "医嘱状态<>1"
            If rsSend药品.RecordCount > 0 Then
                Call LoadAdviceSendDrug(blnOnePati, strEnd, rsSend药品, lng病人数, str科室, bln品种药品, bln时价提示, bln库存提示, bln默认发送, lng病人ID, bln存储库房提示)
            End If
        End If
        If bln其他 Then
            If mblnAutoVerify Then rsSend其他.Filter = "医嘱状态<>1"
            If rsSend其他.RecordCount > 0 Then
                Call LoadAdviceSendOther(blnOnePati, strEnd, rsSend其他, lng病人数, str科室, lng单量数, lng病人ID)
            End If
        End If
        If Not blnOnePati Then
            Progress = k / (UBound(arrPati) + 1) * 100
        End If
    Next
    If Not blnOnePati Then Progress = 0

    If bln药品 Then
        If mbln领药号 Then Call Refresh领药号
    End If
    
    
    With vsAdvice
        If mblnOnePati Then
            If .Rows - 1 > .FixedRows Then
                lblInfo.Caption = "姓名：" & .TextMatrix(.Rows - 1, COL_姓名) & ",住院号：" & .TextMatrix(.Rows - 1, COL_住院号) & "。床号：" & .TextMatrix(.Rows - 1, COL_床号) & "," & lblInfo.Caption & IIF(str科室 = "", " ", "(" & Mid(str科室, 2) & ") ")
            Else
                lblInfo.Caption = "没有读取任何医嘱。"
            End If
        Else
            lblInfo.Caption = lblInfo.Caption & "，共有" & IIF(str科室 = "", " ", "(" & Mid(str科室, 2) & ") ") & lng病人数 & " 个病人的医嘱"
        End If
        
        .Redraw = flexRDNone
        If bln药品 Then
            .ColHidden(COL_规格) = Not bln品种药品 Or chk药品期效(0).value <> 1
        End If
        
        If bln其他 And bln药品 = False Then
            .ColHidden(COL_单量) = lng单量数 = 0
        Else
            .ColHidden(COL_单量) = False
        End If
        .ColHidden(COL_单量单位) = .ColHidden(COL_单量)
        
        If Not .ColHidden(COL_规格) Then
            .AutoSize col_医嘱内容, COL_规格
        Else
            .AutoSize col_医嘱内容
        End If
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
               
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    
    If VsfOnlyOneRow(vsAdvice) Then
        '只有一行
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 1 Then
            vsAdvice.BackColorSel = BackColorNew
        Else
            vsAdvice.BackColorSel = vbWhite
        End If
    Else
        vsAdvice.BackColorSel = COLSelBackColor
    End If
    
    If vsAdvice.Visible Then vsAdvice.SetFocus
    Call ShowSendTotal
    Screen.MousePointer = 0
        
    Exit Sub
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadAdviceSendDrug(ByVal blnOnePati As Boolean, ByVal strEnd As String, ByVal rsSend As ADODB.Recordset, ByRef lng病人数 As Long, ByRef str科室 As String, _
    ByRef bln品种药品 As Boolean, ByRef bln时价提示 As Boolean, ByRef bln库存提示 As Boolean, ByRef bln默认发送 As Boolean, ByRef lng病人ID As Long, ByRef bln存储库房提示 As Boolean) As Boolean
'功能：显示要发送的药品医嘱清单
'参数：strEnd=发送到的结束时间(yyyy-MM-dd HH:mm:ss),临嘱没有
'返回：lng病人数=有待发送医嘱的病人数
'      str科室=所有病人当前科室串
'      bln品种药品=是否存在未确定规格的品种药品
'说明：注意CellData中存放得有附加数据
'   RowData：0-未发送的,-1-已成功发送的
'   COL_选择：0-可自由选择的,1-禁止改变选择状态的
'   COL_婴儿：存放婴儿编号
'   COL_诊疗类别：1-给药途径，2-中药煎法，3-中药用法；只在本函数内使用
'   COL_医嘱内容：存放诊疗项目名称,用于显示计价医嘱
'   COL_分解时间:临嘱无分解时间时,存放费用发生时间
'   COL_规格：存放成药可选择的规格下拉数据(ComboList)
'   COL_金额：存放计价内容是否变更过
    
    Dim rsDrug As New ADODB.Recordset
    Dim i As Long, j As Long, k As Long, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDel相关ID As Long, vBookMark As Variant
    Dim lng次数 As Long, lng最小次数 As Long, str用法 As String
    Dim str分解时间 As String, dbl总量 As Double, cur金额 As Currency
    Dim blnReCalc As Boolean
    Dim rsTmp As Recordset, strSQL As String, strIDs As String
    Dim bln口服 As Boolean
    Dim lng组医嘱ID As Long
    Dim lng组医嘱IDTmp As Long
    Dim varBookmark As Variant
     
    On Error GoTo errH
    '计算并显示发送清单
    '----------------------------------------------------------------------------------------------------------
    With vsAdvice
        .Redraw = flexRDNone
        For i = 1 To rsSend.RecordCount
            If rsSend!诊疗类别 = "E" And IsNull(rsSend!相关ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_相关ID)) Then
                GoTo NextLoop '跳过非药正常的治疗医嘱或检验采集方法
            ElseIf rsSend!诊疗类别 = "E" And Not IsNull(rsSend!相关ID) And NVL(rsSend!相关ID, 0) <> Val(.TextMatrix(.Rows - 1, COL_相关ID)) Then
                GoTo NextLoop '跳过输血途径
            ElseIf (rsSend!ID = lngDel相关ID Or NVL(rsSend!相关ID, 0) = lngDel相关ID) And lngDel相关ID <> 0 Then
                GoTo NextLoop '一并给药或配方中的一个可能已经不能发送,则整组不能发送
            Else
                lngDel相关ID = 0
            End If
            
            If mstrEndPoint <> "" Then
                lng组医嘱IDTmp = IIF(IsNull(rsSend!相关ID), rsSend!ID, rsSend!相关ID)
                If lng组医嘱ID <> lng组医嘱IDTmp Then
                    lng组医嘱ID = lng组医嘱IDTmp
                    bln口服 = False
                    If InStr(",5,6,7,", "," & rsSend!诊疗类别 & ",") > 0 Then
                        varBookmark = rsSend.Bookmark
                        rsSend.MoveNext
                        Do While Not rsSend.EOF
                            If rsSend!诊疗类别 = "E" And (Val(rsSend!操作类型 & "") = 2 Or Val(rsSend!操作类型 & "") = 4) Then
                                If Val(rsSend!执行分类 & "") = 4 Then
                                    bln口服 = True
                                End If
                                Exit Do
                            End If
                            rsSend.MoveNext
                        Loop
                        rsSend.Bookmark = varBookmark
                    End If
                End If
            End If
            
            '加入当前行
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .Cell(flexcpPictureAlignment, lngRow, COL_选择) = 4
            
            If InStr(mstrNoneIDs, "," & CStr(rsSend!ID) & ",") > 0 And Not mbln阳性用药 Then
                Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing
            Else
                Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
            End If
            
            
            If mlngRefModld = 1 Then
                If Format(rsSend!上次执行时间, "YYYY-MM-DD") >= Format(mdatCurr, "YYYY-MM-DD") Then '当天已发送的(长嘱可能发送到将来)
                    .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HA08000           '海蓝
                End If
            End If
            
            '隐藏附加行
            If rsSend!诊疗类别 = "7" Then
                .RowHidden(lngRow) = True '中草药
            ElseIf rsSend!诊疗类别 = "E" Then
                If Not IsNull(rsSend!相关ID) Then
                    .RowHidden(lngRow) = True
                    .Cell(flexcpData, lngRow, COL_诊疗类别) = 2 '中药煎法
                ElseIf Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID _
                    And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                    .RowHidden(lngRow) = True
                    .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 '给药途径
                Else
                    .Cell(flexcpData, lngRow, COL_诊疗类别) = 3 '中药用法
                End If
            End If
            
            '一般列赋值
            '---------------------------------------------------------------
            .Cell(flexcpData, lngRow, COL_婴儿) = CLng(NVL(rsSend!婴儿, 0))
            If NVL(rsSend!婴儿, 0) = 0 Then
                .TextMatrix(lngRow, COL_婴儿) = "病人"
            Else
                .TextMatrix(lngRow, COL_婴儿) = "婴儿" & rsSend!婴儿
                .ColHidden(COL_婴儿) = False '有婴儿医嘱时才显示
            End If
            .TextMatrix(lngRow, COL_科室) = rsSend!科室
            If InStr(str科室 & ",", "," & rsSend!科室 & ",") = 0 Then
                If str科室 <> "" Then .ColHidden(COL_科室) = False
                str科室 = str科室 & "," & rsSend!科室
            End If
            
            .TextMatrix(lngRow, COL_病人ID) = rsSend!病人ID
            .TextMatrix(lngRow, COL_主页ID) = rsSend!主页ID
            .TextMatrix(lngRow, COL_姓名) = rsSend!姓名
            .TextMatrix(lngRow, col_性别) = NVL(rsSend!性别)
            .TextMatrix(lngRow, COL_年龄) = NVL(rsSend!年龄)
            .TextMatrix(lngRow, COL_险类) = NVL(rsSend!险类)
            .TextMatrix(lngRow, COL_住院号) = NVL(rsSend!住院号)
            .TextMatrix(lngRow, COL_床号) = NVL(rsSend!床号)
            .TextMatrix(lngRow, COL_费别) = NVL(rsSend!费别)
            
            .TextMatrix(lngRow, COL_ID) = rsSend!ID
            .TextMatrix(lngRow, COL_相关ID) = NVL(rsSend!相关ID)
            .TextMatrix(lngRow, COL_诊疗类别) = rsSend!诊疗类别
            .TextMatrix(lngRow, COL_诊疗项目ID) = rsSend!诊疗项目ID
            .TextMatrix(lngRow, COL_医嘱期效) = IIF(rsSend!医嘱期效 = 0, "长嘱", "临嘱")
            .Cell(flexcpData, lngRow, COL_医嘱期效) = Val(rsSend!医嘱期效)
            
            .TextMatrix(lngRow, col_医嘱内容) = NVL(rsSend!医嘱内容)
            .Cell(flexcpData, lngRow, col_医嘱内容) = CStr(NVL(rsSend!诊疗项目)) '用于显示计价医嘱
            
            .TextMatrix(lngRow, COL_医生嘱托) = NVL(rsSend!医生嘱托)
            .Cell(flexcpData, lngRow, COL_医生嘱托) = CStr(NVL(rsSend!摘要))
            
            .TextMatrix(lngRow, COL_执行时间) = NVL(rsSend!执行时间方案)
            If Not IsNull(rsSend!开始执行时间) Then
                .Cell(flexcpData, lngRow, COL_执行时间) = CStr(Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss"))
            End If
            
            .TextMatrix(lngRow, COL_频率) = NVL(rsSend!执行频次)
            .TextMatrix(lngRow, COL_病人病区ID) = NVL(rsSend!病人病区ID)
            .TextMatrix(lngRow, COL_病人科室ID) = NVL(rsSend!病人科室id)
            .TextMatrix(lngRow, COL_开嘱科室ID) = NVL(rsSend!开嘱科室id)
            .TextMatrix(lngRow, COL_开嘱医生) = NVL(rsSend!开嘱医生)
            
            .TextMatrix(lngRow, COL_计价特性) = NVL(rsSend!计价特性, 0)
            .TextMatrix(lngRow, COL_执行性质ID) = NVL(rsSend!执行性质, 0)
            .TextMatrix(lngRow, COL_执行标记) = NVL(rsSend!执行标记, 0)
            .TextMatrix(lngRow, COL_执行分类) = NVL(rsSend!执行分类, 0)
            .TextMatrix(lngRow, COL_操作类型) = NVL(rsSend!操作类型, 0)
            .TextMatrix(lngRow, COL_用药理由) = NVL(rsSend!用药理由)
            .TextMatrix(lngRow, COL_会诊医嘱ID) = NVL(rsSend!会诊医嘱ID, 0)
            '医嘱状态用于发送前对未校对的先进行自动校对
            .TextMatrix(lngRow, COL_医嘱状态) = rsSend!医嘱状态
            If rsSend!医嘱状态 = 1 Then
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = BackColorNew '浅黄色
            End If
                                                
            '显示主要执行科室
            .TextMatrix(lngRow, COL_执行科室) = NVL(rsSend!执行科室)
            
            '显示附加执行科室
            If rsSend!诊疗类别 = "E" And IsNull(rsSend!相关ID) Then
                If InStr(",7,E,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                    '中药用法
                    .TextMatrix(lngRow, COL_附加执行) = NVL(rsSend!执行科室)
                    .Cell(flexcpData, lngRow, COL_附加执行) = CStr(NVL(rsSend!执行科室))
                ElseIf InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                    '给药途径
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                            .TextMatrix(j, COL_附加执行) = NVL(rsSend!执行科室)
                            .Cell(flexcpData, j, COL_附加执行) = CStr(NVL(rsSend!执行科室))
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
            
            .TextMatrix(lngRow, COL_执行科室ID) = NVL(rsSend!执行科室ID)
            If mblnAutoVerify Then .TextMatrix(lngRow, COL_新开操作时间) = Format(rsSend!新开操作时间, "yyyy-MM-dd HH:mm:ss")
            If InStr(",4,5,6,7,", "," & rsSend!诊疗类别 & ",") = 0 Then .TextMatrix(lngRow, COL_计算方式) = NVL(rsSend!计算方式, 0)
            .TextMatrix(lngRow, COL_执行安排) = NVL(rsSend!执行安排, 0)
            .TextMatrix(lngRow, COL_病人性质) = NVL(rsSend!病人性质, 0)
            .TextMatrix(lngRow, COL_开嘱时间) = Format(rsSend!开嘱时间, "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, COL_开始时间) = Format(NVL(rsSend!开始执行时间), "yyyy-MM-dd HH:mm:ss")
                                                
            '读取药品相关信息
            '---------------------------------------------------------------
            If InStr(",5,6,7", rsSend!诊疗类别) > 0 Then
                '毒理分类判断
                If gbln特殊药品分开发送 Then
                    strTmp = ""
                    Select Case cboDrugType.ListIndex
                    Case 1
                        If rsSend!毒理分类 & "" <> "毒性药" Then strTmp = "1"
                    Case 2
                        If InStr(",麻醉药,精神I类,", "," & rsSend!毒理分类 & ",") = 0 Then strTmp = "1"
                    Case 3
                        If InStr(",毒性药,麻醉药,精神I类,", "," & rsSend!毒理分类 & ",") > 0 Then strTmp = "1"
                    End Select
                    
                    If strTmp <> "" Then
                        lngDel相关ID = NVL(rsSend!相关ID, 0)
                        Call DeleteCurRow(lngRow, rsSend!相关ID)
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                    .TextMatrix(lngRow, COL_毒理分类) = NVL(rsSend!毒理分类, "空")
                End If
                
                Set rsDrug = New ADODB.Recordset
                '先包括停用药品,待确认要发送的医嘱再检查停用
                Set rsDrug = GetDrugInfo(rsSend!诊疗项目ID, NVL(rsSend!收费细目ID, 0), NVL(rsSend!执行科室ID, 0), 2, False)
                '排除当前执行科室下没有存储的规格
                If NVL(rsSend!执行科室ID, 0) <> 0 And rsDrug.RecordCount > 1 And InStr("," & gstr输液配置中心 & ",", "," & NVL(rsSend!执行科室ID, 0) & ",") > 0 Then
                    strIDs = ""
                    Do While Not rsDrug.EOF
                        strIDs = strIDs & "," & rsDrug!药品ID
                        rsDrug.MoveNext
                    Loop
                    strSQL = "Select /*+ rule*/" & vbNewLine & _
                            "Distinct 收费细目id" & vbNewLine & _
                            "From 收费执行科室" & vbNewLine & _
                            "Where (开单科室id = [2] Or 开单科室id Is Null) And 执行科室ID = [3] And" & vbNewLine & _
                            "      收费细目id In (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)))"
    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2), Val(rsSend!开嘱科室id & ""), Val(rsSend!执行科室ID & ""))
                    If rsDrug.RecordCount > 0 Then rsDrug.MoveFirst
                    strIDs = ""
                    Do While Not rsDrug.EOF
                        rsTmp.Filter = "收费细目ID=" & rsDrug!药品ID
                        If rsTmp.RecordCount = 0 Then
                           strIDs = strIDs & " or 药品ID<>" & rsDrug!药品ID
                        End If
                        rsDrug.MoveNext
                    Loop
                    strIDs = Mid(strIDs, 4)
                    If strIDs <> "" Then rsDrug.Filter = strIDs
                    If rsDrug.RecordCount > 0 Then rsDrug.MoveFirst
                End If
                If rsDrug.EOF Then
                    '药品没有对应的规格信息
                    '删除当前行(及相关行),及处理下一医嘱
                    Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                    lng最小次数 = 0: GoTo NextLoop
                ElseIf rsDrug.RecordCount > 1 Then
                    '寻找合适的规格
                    Call SeekMatchDrug(rsSend, rsDrug, 0, vBookMark, strTmp)
                    If vBookMark <> 0 Then
                        rsDrug.Bookmark = vBookMark
                    Else
                        rsDrug.MoveFirst
                    End If
                    .Cell(flexcpData, lngRow, COL_规格) = strTmp '可选择的规格
                    '检查全部(指定)规格都停用的药品
                    If .Cell(flexcpData, lngRow, COL_规格) = "" Then
                        Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                Else
                    '检查全部(指定)规格都停用的药品：长期药品医嘱在确认要发送时再删除和提示
                    If Not (rsSend!医嘱期效 = 0 And InStr(",5,6,", rsSend!诊疗类别) > 0) _
                        And Not (NVL(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!服务对象, 0)) > 0) Then
                        Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                        lng最小次数 = 0: GoTo NextLoop
                    ElseIf Val(rsSend!医嘱期效 & "") = 0 And InStr(",5,6,", rsSend!诊疗类别) > 0 And Val(rsSend!执行科室ID & "") <> 0 And Val(rsSend!收费细目ID & "") <> 0 Then '检查收费执行科室是否改变
                        strSQL = "Select 1 From 收费执行科室 Where 收费细目id = [1] And Nvl(病人来源, 2) = 2 And Nvl(开单科室ID, [2]) = [2] And 执行科室ID = [3]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!收费细目ID & ""), Val(rsSend!开嘱科室ID & ""), Val(rsSend!执行科室ID & ""))
                        If rsTmp.EOF Then
                            Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                            lng最小次数 = 0: GoTo NextLoop
                        End If
                    End If
                End If
                .TextMatrix(lngRow, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
                .TextMatrix(lngRow, COL_收费细目ID) = rsDrug!药品ID
                .TextMatrix(lngRow, COL_库存) = Format(NVL(rsDrug!库存, 0), "0.00000") '按住院包装
                .TextMatrix(lngRow, COL_剂量系数) = NVL(rsDrug!剂量系数, 1)
                .TextMatrix(lngRow, COL_住院包装) = NVL(rsDrug!住院包装, 1)
                .TextMatrix(lngRow, COL_住院单位) = NVL(rsDrug!住院单位)
                .TextMatrix(lngRow, COL_可否分零) = NVL(rsSend!可否分零, NVL(rsDrug!可否分零, 0))
                .TextMatrix(lngRow, COL_药房分批) = NVL(rsDrug!药房分批, 0)
                .TextMatrix(lngRow, COL_是否变价) = NVL(rsDrug!是否变价, 0)
                
                '是否存在未确定规格的品种药品
                If .Cell(flexcpData, lngRow, COL_规格) <> "" Then
                    .Cell(flexcpForeColor, lngRow, COL_规格) = vbBlue '突出显示
                    bln品种药品 = True
                End If
            End If
                                                                    
            '计算发送次数，执行的分解时间，总量
            '---------------------------------------------------------------
            If rsSend!医嘱期效 = 0 Then
                '长嘱---------------------------------------------
                If rsSend!诊疗类别 = "7" Then
                    '当前医嘱的发送计算时间段
                    Call Calc总量次数时间(dbl总量, lng次数, str分解时间, strEnd, rsSend, rsDrug, bln口服)
                    If str分解时间 = "" Then
                        If rsSend!医嘱状态 = 1 Then '仅校对
                            lng最小次数 = 0
                        Else
                            '无法分解时间(如被暂停的)
                            lngDel相关ID = rsSend!相关ID
                            Call DeleteCurRow(lngRow, rsSend!相关ID)
                            lng最小次数 = 0: GoTo NextLoop
                        End If
                    ElseIf Not (NVL(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!服务对象, 0)) > 0) Then
                        '确认要继续发送，但已被撤档或不服务于病人的药品
                        Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    If Len(str分解时间) > 4000 Then
                        .TextMatrix(lngRow, COL_分解时间) = Mid(str分解时间, 1, InStr(Mid(str分解时间, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_分解时间) = str分解时间
                    End If
                    If str分解时间 <> "" Then
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = NVL(rsSend!单次用量)
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = lng次数 * NVL(rsSend!总给予量, 1) '总付数
                    
                    .TextMatrix(lngRow, COL_总量单位) = "付"
                ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                    blnReCalc = False
ReCalc:
                    '当前医嘱的发送计算时间段
                    Call Calc总量次数时间(dbl总量, lng次数, str分解时间, strEnd, rsSend, rsDrug, bln口服)
                    If str分解时间 = "" Then
                        If rsSend!医嘱状态 = 1 Then '仅校对
                            lng最小次数 = 0
                        Else
                            '无法分解时间(如被暂停的)
                            lngDel相关ID = rsSend!相关ID
                            Call DeleteCurRow(lngRow, rsSend!相关ID)
                            lng最小次数 = 0: GoTo NextLoop
                        End If
                    ElseIf Not (NVL(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!服务对象, 0)) > 0) Then
                        '确认要继续发送，但已被撤档或不服务于病人的药品
                        Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    If Len(str分解时间) > 4000 Then
                        .TextMatrix(lngRow, COL_分解时间) = Mid(str分解时间, 1, InStr(Mid(str分解时间, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_分解时间) = str分解时间
                    End If
                    If str分解时间 <> "" Then
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsDrug!住院单位)
                    If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
                    
                    '当有多个规格可选择时，根据库存是否足够再次定位规格
                    If Not blnReCalc And .Cell(flexcpData, lngRow, COL_规格) <> "" _
                        And Val(.TextMatrix(lngRow, COL_总量)) > Val(.TextMatrix(lngRow, COL_库存)) Then
                        Call SeekMatchDrug(rsSend, rsDrug, Val(.TextMatrix(lngRow, COL_总量)), vBookMark)
                        If vBookMark <> 0 Then
                            rsDrug.Bookmark = vBookMark
                            .TextMatrix(lngRow, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
                            .TextMatrix(lngRow, COL_收费细目ID) = rsDrug!药品ID
                            .TextMatrix(lngRow, COL_库存) = Format(NVL(rsDrug!库存, 0), "0.00000") '按住院包装
                            .TextMatrix(lngRow, COL_剂量系数) = NVL(rsDrug!剂量系数, 1)
                            .TextMatrix(lngRow, COL_住院包装) = NVL(rsDrug!住院包装, 1)
                            .TextMatrix(lngRow, COL_住院单位) = NVL(rsDrug!住院单位)
                            .TextMatrix(lngRow, COL_药房分批) = NVL(rsDrug!药房分批, 0)
                            .TextMatrix(lngRow, COL_是否变价) = NVL(rsDrug!是否变价, 0)
                            blnReCalc = True: GoTo ReCalc
                        End If
                    End If
                Else
                    '一并给药的按最小次数发送(影响给药途径计费及上次执行时间)(不分零的可能浪废)
                    If .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 Then '给药途径
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_次数)) > lng最小次数 Then
                                    .TextMatrix(j, COL_次数) = lng最小次数
                                    .TextMatrix(j, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(j, COL_分解时间))
                                    .TextMatrix(j, COL_首次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(0), "yyyy-MM-dd HH:mm")
                                    .TextMatrix(j, COL_末次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(lng最小次数 - 1), "yyyy-MM-dd HH:mm")
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng最小次数 = 0
                    End If
                    
                    If InStr(",2,3,", .Cell(flexcpData, lngRow, COL_诊疗类别)) > 0 Then
                        '中药煎法、用法为付数
                        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                    Else
                        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数)
                    End If
                    .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                    If .Cell(flexcpData, lngRow, COL_诊疗类别) = 3 Then '中药用法
                        .TextMatrix(lngRow, COL_总量单位) = "付"
                    End If
                    
                    .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                    .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                    .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                End If
                If mlngRefModld = 1 Then
                    If (dkpExecTime.value >= rsSend!执行终止时间 Or dkpExecTime.value <= rsSend!上次执行时间) Then
                        '备用医嘱删除范围外的备用医嘱
                        lngDel相关ID = rsSend!相关ID
                        Call DeleteCurRow(lngRow, rsSend!相关ID)
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                End If
            Else
                '临嘱---------------------------------------------
                If rsSend!诊疗类别 = "7" Then
                    .TextMatrix(lngRow, COL_次数) = NVL(rsSend!总给予量, 0)
                    If Not IsNull(rsSend!执行时间方案) Or NVL(rsSend!间隔单位) = "分钟" Then
                        .TextMatrix(lngRow, COL_分解时间) = Calc次数分解时间(rsSend!总给予量, rsSend!开始执行时间, CDate("3000-01-01"), "", NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(rsSend!总给予量 - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '无分解时间(临嘱可能未输入执行时间而无法分解)
                        '记录费用发生时间(以医嘱开始执行时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_单量) = NVL(rsSend!单次用量) '单量
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = NVL(rsSend!总给予量, 0) '付数
                    .TextMatrix(lngRow, COL_总量单位) = "付"
                ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                    '计算临嘱用药次数
                    If NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Then
                        lng次数 = 1 '设置为一次性的临嘱药品
                    ElseIf NVL(rsSend!天数, 0) <> 0 And Not IsNull(rsSend!执行频次) Then
                        '一个频率周期的次数
                        If rsSend!间隔单位 = "周" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / 7))
                        ElseIf rsSend!间隔单位 = "天" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔))
                        ElseIf rsSend!间隔单位 = "小时" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔) * 24)
                        ElseIf rsSend!间隔单位 = "分钟" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔) * (24 * 60))
                        End If
                    Else
                        '可分零药品时,按总量对单量的倍数计算给药途径的次数,不可分零与一次性使用药品时，按总量对（单量与剂量系数比值取整）的倍数计算给药途径的次数，
                        '否则按一个频率周期的次数计算
                        If NVL(rsSend!可否分零, NVL(rsDrug!可否分零, 0)) = 0 And NVL(rsSend!单次用量, 0) <> 0 Then
                            lng次数 = IntEx(rsSend!总给予量 * rsDrug!剂量系数 / rsSend!单次用量)
                        ElseIf (NVL(rsSend!可否分零, NVL(rsDrug!可否分零, 0)) = 1 Or NVL(rsSend!可否分零, NVL(rsDrug!可否分零, 0)) = 2) And NVL(rsSend!单次用量, 0) <> 0 Then
                            lng次数 = IntEx(rsSend!总给予量 / IntEx(rsSend!单次用量 / rsDrug!剂量系数))
                        Else
                            lng次数 = NVL(rsSend!频率次数, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!执行时间方案) Or NVL(rsSend!间隔单位) = "分钟" Then
                        str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        If str分解时间 <> "" Then
                            If Len(str分解时间) > 4000 Then
                                .TextMatrix(lngRow, COL_分解时间) = Mid(str分解时间, 1, InStr(Mid(str分解时间, 4001), ",") + 3999)
                            Else
                                .TextMatrix(lngRow, COL_分解时间) = str分解时间
                            End If
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        '无分解时间(一次性临嘱未输入执行时间而无法分解)
                        '记录费用发生时间(以医嘱开始执行时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = FormatEx(rsSend!总给予量 / rsDrug!住院包装, 5) '以住院单位显示
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsDrug!住院单位)
                    
                    If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
                Else
                    '临嘱：一并给药的按最小次数发送(影响给药途径计费)
                    If .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 Then '给药途径
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_次数)) > lng最小次数 Then
                                    .TextMatrix(j, COL_次数) = lng最小次数
                                    If .TextMatrix(j, COL_分解时间) <> "" Then
                                        .TextMatrix(j, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(j, COL_分解时间))
                                        .TextMatrix(j, COL_首次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(0), "yyyy-MM-dd HH:mm")
                                        .TextMatrix(j, COL_末次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(lng最小次数 - 1), "yyyy-MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng最小次数 = 0
                    End If
                    
                    .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数) '付数或次数
                    .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                    If .Cell(flexcpData, lngRow, COL_诊疗类别) = 3 Then '中药用法
                        .TextMatrix(lngRow, COL_总量单位) = "付"
                    End If
                    
                    .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                    .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                    .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                    .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                End If
            End If
            
            '计算项目的金额:用于查看及记帐报警
            '---------------------------------------------------------------
            cur金额 = 0
            Call LoadAdvicePrice(lngRow, cur金额, rsDrug)
            .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
            
            '相关行时的一些处理：累计显示金额,给药途径,用法,执行科室,执行性质
            '---------------------------------------------------------------
            If InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_诊疗类别))) > 0 Then '给药途径或中药用法
                cur金额 = 0
                lngTmp = .FindRow(CStr(rsSend!ID), , COL_相关ID)
                
                If .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 Then '给药途径
                    '一并给药时,给药途径的金额累加显示在第一个成药中
                    .TextMatrix(lngTmp, COL_金额) = Format(Val(.TextMatrix(lngTmp, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                    
                    '显示给药途径,执行性质
                    For j = lngTmp To lngRow - 1
                        strTmp = ""
                        If Val(.TextMatrix(j, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                            If Val(.TextMatrix(j, COL_执行标记)) = 2 Then
                                strTmp = "不取药"
                            Else
                                strTmp = "自备药"
                            End If
                        ElseIf Val(.TextMatrix(j, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                            strTmp = "离院带药"
                        Else
                            strTmp = IIF(Val(.TextMatrix(j, COL_执行标记)) = 1, "自取药", "")
                        End If
                        .TextMatrix(j, COL_执行性质) = strTmp
                        .TextMatrix(j, COL_用法) = rsSend!诊疗项目
                    Next
                Else
                    '药品的执行性质
                    strTmp = ""
                    If Val(.TextMatrix(lngTmp, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                        If Val(.TextMatrix(lngTmp, COL_执行标记)) = 2 Then
                            strTmp = "不取药"
                        Else
                            strTmp = "自备药"
                        End If
                    ElseIf Val(.TextMatrix(lngTmp, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                        strTmp = "离院带药"
                    Else
                        strTmp = IIF(Val(.TextMatrix(lngTmp, COL_执行标记)) = 1, "自取药", "")
                    End If
                    
                    '中药用法,煎法
                    str用法 = rsSend!诊疗项目
                    If Val(.Cell(flexcpData, lngRow - 1, COL_诊疗类别)) = 2 Then
                        str用法 = str用法 & "|" & sys.RowValue("诊疗项目目录", Val(.TextMatrix(lngRow - 1, COL_诊疗项目ID)), "名称")
                    End If
                    For j = lngTmp To lngRow
                        .TextMatrix(j, COL_用法) = str用法 '用于填写收发记录
                        cur金额 = cur金额 + Val(.TextMatrix(j, COL_金额))
                    Next
                    .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
                    '显示执行性质
                    .TextMatrix(lngRow, COL_执行性质) = strTmp
                    '显示配方执行科室
                    .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngTmp, COL_执行科室)
                End If
                
                '使相关医嘱选择状态相同(固为库存的原因)
                For j = lngTmp To lngRow
                    If .Cell(flexcpData, j, COL_选择) <> 0 Then
                        Call RowSelectSame(j, COL_选择)
                        Exit For '一个禁止,全部禁止
                    End If
                Next
                If j > lngRow Then
                    For j = lngRow To lngTmp Step -1
                        If InStr(",5,6,7,", .TextMatrix(j, COL_诊疗类别)) > 0 Then
                            If .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                                Call RowSelectSame(j, COL_选择)
                                Exit For '最后不选,全部不选
                            End If
                        End If
                    Next
                End If
            End If
            
            '药品库存检查:自备药不检查
            '---------------------------------------------------------------
            If InStr(",5,6,7,", rsSend!诊疗类别) > 0 And NVL(rsSend!执行性质, 0) <> 5 Then
                Call CheckStock(lngRow, bln库存提示, bln时价提示, bln默认发送)
                Call CheckDrugStorage(lngRow, bln存储库房提示)
            End If
            
            '其它处理
            '---------------------------------------------------------------
            '病人计数及分隔
            If rsSend!病人ID <> lng病人ID Then
                lng病人数 = lng病人数 + 1
                If lng病人ID <> 0 Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Not .RowHidden(j) Then
                            .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                            Exit For
                        End If
                    Next
                End If
            End If
            lng病人ID = rsSend!病人ID

NextLoop:           '---------------------------------------------------------------
            If blnOnePati Then Progress = i / rsSend.RecordCount * 100
            rsSend.MoveNext
        Next
        .Redraw = flexRDDirect
    End With
    
    If blnOnePati Then Progress = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadAdviceSendOther(ByVal blnOnePati As Boolean, ByVal strEnd As String, ByVal rsSend As ADODB.Recordset, _
    ByRef lng病人数 As Long, ByRef str科室 As String, ByRef lng单量数 As Long, ByRef lng病人ID As Long) As Boolean
'功能：根据条件读取并显示要发送的医嘱清单
'参数：strEnd=发送到的结束时间(yyyy-MM-dd HH:mm:ss),临嘱没有
'返回：lng病人数=有待发送医嘱的病人数
'      str科室=所有病人当前科室串
'      lng单量数=决定是否显示单量列
'说明：注意CellData中存放得有附加数据
'   RowData：0-未发送的,-1-已成功发送的
'   COL_选择：0-可自由选择的,1-禁止改变选择状态的
'   COL_婴儿：存放婴儿编号
'   COL_诊疗类别：存放诊疗类别名称，在显示计价医嘱时使用
'   COL_医嘱内容：存放诊疗项目名称或标本部位，在显示计价医嘱时使用
'   COL_分解时间:临嘱无分解时间时,存放费用发生时间
'   COL_首次时间,COL_末次时间：存放持续性长嘱的首末次执行时间
'   COL_频率：1-"一次性"临嘱，2-"持续性"长嘱
'   COL_执行科室：存放原执行科室名称
'   COL_执行科室ID：是否更改了执行科室
'   COL_金额：存放计价内容是否变更过
   
    Dim strTmp As String, i As Long, j As Long, k As Long
    Dim datBegin As Date, datEnd As Date, strPause As String
    Dim lng次数 As Long, dbl总量 As Double, bln采集方法 As Boolean
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim lngDel医嘱ID As Long, lngRow As Long
    Dim cur金额 As Currency
    Dim str停用 As String
             
        
    '计算并显示发送清单
    '----------------------------------------------------------------------------------------------------------
    With vsAdvice
        .Redraw = flexRDNone
        For i = 1 To rsSend.RecordCount
            If NVL(rsSend!相关ID, 0) = lngDel医嘱ID And lngDel医嘱ID <> 0 Then
                GoTo NextLoop '检查组合或手术组合,输血项目中的一个可能已经不能发送,则整组不能发送
            Else
                lngDel医嘱ID = 0
            End If
            
            bln采集方法 = False
            
            '加入当前行
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .Cell(flexcpPictureAlignment, lngRow, COL_选择) = 4
            If Not ("Z" = "" & rsSend!诊疗类别 And Val("" & rsSend!操作类型) = 5) Then
                Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
            End If
            
            If mlngRefModld = 1 Then
                If Format(rsSend!上次执行时间, "YYYY-MM-DD") >= Format(mdatCurr, "YYYY-MM-DD") Then '当天已发送的(长嘱可能发送到将来)
                    .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HA08000           '海蓝
                End If
            End If
            
            '隐藏:附加手术,手术麻醉,检查部位方法,检验项目,输血途径
            .RowHidden(lngRow) = Not IsNull(rsSend!相关ID)
            
            '一般列赋值
            '---------------------------------------------------------------
            .Cell(flexcpData, lngRow, COL_婴儿) = CLng(NVL(rsSend!婴儿, 0))
            If NVL(rsSend!婴儿, 0) = 0 Then
                .TextMatrix(lngRow, COL_婴儿) = "病人"
            Else
                .TextMatrix(lngRow, COL_婴儿) = "婴儿" & rsSend!婴儿
                .ColHidden(COL_婴儿) = False '有婴儿医嘱时才显示
            End If
            .TextMatrix(lngRow, COL_科室) = rsSend!科室
            If InStr(str科室 & ",", "," & rsSend!科室 & ",") = 0 Then
                If str科室 <> "" Then .ColHidden(COL_科室) = False
                str科室 = str科室 & "," & rsSend!科室
            End If
            
            .TextMatrix(lngRow, COL_病人ID) = rsSend!病人ID
            .TextMatrix(lngRow, COL_主页ID) = rsSend!主页ID
            .TextMatrix(lngRow, COL_姓名) = rsSend!姓名
            .TextMatrix(lngRow, col_性别) = NVL(rsSend!性别)
            .TextMatrix(lngRow, COL_年龄) = NVL(rsSend!年龄)
            .TextMatrix(lngRow, COL_险类) = NVL(rsSend!险类)
            .TextMatrix(lngRow, COL_住院号) = NVL(rsSend!住院号)
            .TextMatrix(lngRow, COL_床号) = NVL(rsSend!床号)
            .TextMatrix(lngRow, COL_费别) = NVL(rsSend!费别)
            
            .TextMatrix(lngRow, COL_ID) = rsSend!ID
            .TextMatrix(lngRow, COL_相关ID) = NVL(rsSend!相关ID)
            .TextMatrix(lngRow, COL_诊疗类别) = "" & rsSend!诊疗类别
            .TextMatrix(lngRow, COL_诊疗项目ID) = Val("" & rsSend!诊疗项目ID)
            .TextMatrix(lngRow, COL_医嘱期效) = IIF(rsSend!医嘱期效 = 0, "长嘱", "临嘱")
            .Cell(flexcpData, lngRow, COL_医嘱期效) = Val(rsSend!医嘱期效)
                        
            .TextMatrix(lngRow, COL_标本部位) = NVL(rsSend!标本部位)
            .TextMatrix(lngRow, COL_检查方法) = NVL(rsSend!检查方法)
            .TextMatrix(lngRow, COL_执行标记) = NVL(rsSend!执行标记, 0)
            
            '用于显示计价医嘱
            .Cell(flexcpData, lngRow, COL_诊疗类别) = CStr(NVL(rsSend!类别名称))
            .TextMatrix(lngRow, col_医嘱内容) = NVL(rsSend!医嘱内容)
            .Cell(flexcpData, lngRow, col_医嘱内容) = CStr(NVL(rsSend!诊疗项目))
            
            .TextMatrix(lngRow, COL_医生嘱托) = NVL(rsSend!医生嘱托)
            .Cell(flexcpData, lngRow, COL_医生嘱托) = CStr(NVL(rsSend!摘要))
            
            .TextMatrix(lngRow, COL_执行时间) = NVL(rsSend!执行时间方案)
            If Not IsNull(rsSend!开始执行时间) Then
                .Cell(flexcpData, lngRow, COL_执行时间) = CStr(Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss"))
            End If
            
            .TextMatrix(lngRow, COL_频率) = NVL(rsSend!执行频次)
            
            .TextMatrix(lngRow, COL_病人病区ID) = NVL(rsSend!病人病区ID)
            .TextMatrix(lngRow, COL_病人科室ID) = NVL(rsSend!病人科室id)
            .TextMatrix(lngRow, COL_开嘱科室ID) = NVL(rsSend!开嘱科室id)
            .TextMatrix(lngRow, COL_开嘱医生) = NVL(rsSend!开嘱医生)
            
            .TextMatrix(lngRow, COL_计价特性) = NVL(rsSend!计价特性, 0)
            .TextMatrix(lngRow, COL_操作类型) = NVL(rsSend!操作类型)
            .TextMatrix(lngRow, COL_执行频率) = NVL(rsSend!执行频率)
            .TextMatrix(lngRow, COL_执行性质ID) = NVL(rsSend!执行性质, 0)
            .TextMatrix(lngRow, COL_紧急标志) = NVL(rsSend!紧急标志, 0)
            If mblnAutoVerify Then .TextMatrix(lngRow, COL_新开操作时间) = Format(rsSend!新开操作时间, "yyyy-MM-dd HH:mm:ss")
            If InStr(",4,5,6,7,", "," & rsSend!诊疗类别 & ",") = 0 Then .TextMatrix(lngRow, COL_计算方式) = NVL(rsSend!计算方式, 0)
            .TextMatrix(lngRow, COL_执行安排) = NVL(rsSend!执行安排, 0)
            .TextMatrix(lngRow, COL_病人性质) = NVL(rsSend!病人性质, 0)
            .TextMatrix(lngRow, COL_开嘱时间) = Format(rsSend!开嘱时间, "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, COL_开始时间) = Format(NVL(rsSend!开始执行时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, COL_申请序号) = NVL(rsSend!申请序号, 0)
            .TextMatrix(lngRow, COL_执行分类) = NVL(rsSend!执行分类, 0)
            .TextMatrix(lngRow, COL_用药理由) = NVL(rsSend!用药理由)
            .TextMatrix(lngRow, COL_会诊医嘱ID) = NVL(rsSend!会诊医嘱ID, 0)
            '医嘱状态用于发送前对未校对的先进行自动校对
            .TextMatrix(lngRow, COL_医嘱状态) = rsSend!医嘱状态
            If rsSend!医嘱状态 = 1 Then
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = BackColorNew '浅黄色
            End If
            
            '采集方式的管码与一并的第一个检验相同
            If NVL(rsSend!诊疗类别) = "E" And IsNull(rsSend!相关ID) _
                And .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                j = .FindRow(CStr(rsSend!ID), .FixedRows, COL_相关ID)
                If j <> -1 Then
                    .TextMatrix(lngRow, COL_试管编码) = .TextMatrix(j, COL_试管编码)
                End If
            Else
                .TextMatrix(lngRow, COL_试管编码) = NVL(rsSend!试管编码)
            End If
            
            '用于卫材医嘱
            .TextMatrix(lngRow, COL_收费细目ID) = NVL(rsSend!收费细目ID, 0)
            .TextMatrix(lngRow, COL_是否变价) = NVL(rsSend!是否变价, 0)
            .TextMatrix(lngRow, COL_跟踪在用) = NVL(rsSend!跟踪在用, 0)
            
            '主项目执行科室显示
            If IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "E" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                '采集方法显示为检验项目的执行科室
                bln采集方法 = True
                .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngRow - 1, COL_执行科室)
                .Cell(flexcpData, lngRow, COL_执行科室) = .Cell(flexcpData, lngRow - 1, COL_执行科室)
            Else
                .TextMatrix(lngRow, COL_执行科室) = NVL(rsSend!执行科室)
                .Cell(flexcpData, lngRow, COL_执行科室) = CStr(NVL(rsSend!执行科室))
            End If
            
            '附加项目执行科室显示
            If NVL(rsSend!诊疗类别) = "E" And IsNull(rsSend!相关ID) _
                And .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                '采集方式在当前行显示附加执行科室
                .TextMatrix(lngRow, COL_附加执行) = NVL(rsSend!执行科室)
                .Cell(flexcpData, lngRow, COL_附加执行) = CStr(NVL(rsSend!执行科室))
            ElseIf NVL(rsSend!诊疗类别) = "G" And Not IsNull(rsSend!相关ID) Then
                '手术麻醉主手术行显示附加执行科室
                j = .FindRow(CStr(rsSend!相关ID), .FixedRows, COL_ID)
                If j <> -1 Then
                    .TextMatrix(j, COL_附加执行) = NVL(rsSend!执行科室)
                    .Cell(flexcpData, j, COL_附加执行) = CStr(NVL(rsSend!执行科室))
                End If
            End If
            
            .TextMatrix(lngRow, COL_执行科室ID) = NVL(rsSend!执行科室ID)
                            
            '计算发送次数，执行的分解时间，总量
            '---------------------------------------------------------------
            If rsSend!医嘱期效 = 0 Then
                '长嘱---------------------------------------------
                If (IsNull(rsSend!相关ID) And Not bln采集方法) _
                    Or (Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "C") Then '主要医嘱或一并采集的检验项目
                
                    '当前医嘱的暂停时间段:"暂停时间,开始时间;...."
                    If rsSend!医嘱状态 <> 1 Then
                        strPause = GetAdvicePause(rsSend!ID)
                    End If
                    
                    '当前医嘱的发送计算时间段
                    datBegin = rsSend!开始执行时间
                    If Not IsNull(rsSend!上次执行时间) Then
                        If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                            datBegin = DateAdd("s", 1, rsSend!上次执行时间) '"持续性"的项目
                        Else
                            datBegin = Calc本周期开始时间(rsSend!开始执行时间, rsSend!上次执行时间, rsSend!频率间隔, rsSend!间隔单位)
                            
                            '本周期内已执行的时间不再计算,这里通过暂停方式来处理
                            If mlngRefModld = 0 Then
                                strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!上次执行时间, "yyyy-MM-dd HH:mm:ss")
                                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                            End If
                        End If
                    End If
                    datEnd = CDate(strEnd)
                    If Not IsNull(rsSend!执行终止时间) Then
                        If rsSend!执行终止时间 < CDate(strEnd) Then
                            datEnd = rsSend!执行终止时间
                        End If
                    End If
                    
                    '计算分解时间及次数
                    If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                        '执行频率为"持续性"的项目,每天发送一次(00:00)
                        lng次数 = Calc持续性长嘱次数(datBegin, datEnd, _
                            Format(NVL(rsSend!上次执行时间), "yyyy-MM-dd HH:mm:ss"), _
                            Format(NVL(rsSend!执行终止时间), "yyyy-MM-dd HH:mm:ss"), _
                            strPause, str首次时间, str末次时间, str分解时间)
                        If lng次数 = 0 Then '本次无需发送
                            If rsSend!医嘱状态 = 1 Then '不发送仅校对
                                str分解时间 = ""
                                str首次时间 = ""
                                str末次时间 = ""
                            Else
                                lngDel医嘱ID = NVL(rsSend!ID, 0)
                                .RemoveItem lngRow
                                GoTo NextLoop
                            End If
                        End If
                        
                        '标记为"持续性"长嘱
                        .Cell(flexcpData, lngRow, COL_频率) = 2
                    Else
                        '执行频率为"可选频率"的项目
                        str分解时间 = Calc段内分解时间(datBegin, datEnd, strPause, NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位, rsSend!开始执行时间)
                        If str分解时间 = "" Then '无法分解时间(如被暂停的)
                            If rsSend!医嘱状态 = 1 Then '不发送仅校对
                                lng次数 = 0
                            Else
                                lngDel医嘱ID = NVL(rsSend!ID, 0)
                                .RemoveItem lngRow
                                GoTo NextLoop
                            End If
                        Else
                            lng次数 = UBound(Split(str分解时间, ",")) + 1
                        End If
                        
                    End If
                    If NVL(rsSend!计算规则, 0) = 1 Then
                        '取整计算,持续性长嘱无单量
                        dbl总量 = IntEx(NVL(rsSend!单次用量, 1)) * lng次数
                    Else
                        dbl总量 = NVL(rsSend!单次用量, 1) * lng次数
                    End If

                    .TextMatrix(lngRow, COL_次数) = lng次数
                    If Len(str分解时间) > 4000 Then
                        .TextMatrix(lngRow, COL_分解时间) = Mid(str分解时间, 1, InStr(Mid(str分解时间, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_分解时间) = str分解时间
                    End If
                    If str分解时间 <> "" Then   '要发送的长嘱不为空，为空的是仅校对的
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    If Not IsNull(rsSend!单次用量) Then
                        .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    End If
                    .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                ElseIf rsSend!诊疗类别 = "E" And Not IsNull(rsSend!相关ID) _
                    And .TextMatrix(lngRow - 1, COL_诊疗类别) = "K" And NVL(rsSend!相关ID, 0) = Val(.TextMatrix(lngRow - 1, COL_ID)) Then
                    '输血途径的执行次数
                    .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                    .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数)
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                    
                    .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                    .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                    .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                    .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                ElseIf Not IsNull(rsSend!相关ID) Or bln采集方法 Then '附加医嘱或标本采集方法
                    '检查组合和手术组合不可能为长嘱,所以此段不会执行
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    If Not IsNull(rsSend!单次用量) Then
                        .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    End If
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                    
                    If lngRow > .FixedRows Then
                        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                        .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                        .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                        .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                        .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                    End If
                End If
                If mlngRefModld = 1 Then
                    If (dkpExecTime.value >= rsSend!执行终止时间 Or dkpExecTime.value <= rsSend!上次执行时间) Then
                        '备用医嘱删除范围外的备用医嘱
                        lngDel医嘱ID = NVL(rsSend!ID, 0)
                        .RemoveItem lngRow
                        GoTo NextLoop
                    End If
                End If
            Else
                '临嘱---------------------------------------------
                If (IsNull(rsSend!相关ID) And Not bln采集方法) _
                    Or (Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "C") Then '主要医嘱或一并采集的检验项目
                    
                    If rsSend!诊疗类别 = "K" Then
                        dbl总量 = NVL(rsSend!总给予量, 0)
                        
                        '输血途径的执行次数
                        If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                            lng次数 = 1 '执行频率为"一次性"的项目
                        Else
                            lng次数 = NVL(rsSend!频率次数, 1) '执行频率为"可选频率"的项目
                        End If
                    Else
                        dbl总量 = NVL(rsSend!总给予量, 1)
                        lng次数 = IntEx(dbl总量 / NVL(rsSend!单次用量, 1))
                    End If
                    
                    If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                        '执行频率为"一次性"的项目
                        str分解时间 = "" '不需要
                        .Cell(flexcpData, lngRow, COL_频率) = 1
                    Else
                        '执行频率为"可选频率"的项目
                        If Not IsNull(rsSend!执行时间方案) Or NVL(rsSend!间隔单位) = "分钟" Then
                            str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        Else
                            str分解时间 = "" '临嘱也许未输入执行时间,无法分解
                        End If
                    End If
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    If Len(str分解时间) > 4000 Then
                        .TextMatrix(lngRow, COL_分解时间) = Mid(str分解时间, 1, InStr(Mid(str分解时间, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_分解时间) = str分解时间
                    End If
                    If str分解时间 <> "" Then
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '记录费用发生时间(当无分解时间时),以医嘱的开始执行时间
                        .Cell(flexcpData, lngRow, COL_分解时间) = CStr(Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss"))
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    If Not IsNull(rsSend!单次用量) Then
                        .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    End If
                    .TextMatrix(lngRow, COL_总量) = IIF(dbl总量 = 0, "", FormatEx(dbl总量, 5))
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                ElseIf rsSend!诊疗类别 = "E" And Not IsNull(rsSend!相关ID) _
                    And .TextMatrix(lngRow - 1, COL_诊疗类别) = "K" And NVL(rsSend!相关ID, 0) = Val(.TextMatrix(lngRow - 1, COL_ID)) Then
                    '输血途径的执行次数
                    .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                    .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数)
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                    
                    .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                    .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                    .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                    .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                ElseIf Not IsNull(rsSend!相关ID) Or bln采集方法 Then
                    '附加医嘱或标本采集方法
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    If Not IsNull(rsSend!单次用量) Then
                        .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    End If
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                    
                    If lngRow > .FixedRows Then
                        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                        .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                        .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                        If .TextMatrix(lngRow, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "6" And .Cell(flexcpData, lngRow - 1, COL_分解时间) <> .TextMatrix(lngRow, COL_开始时间) Then
                            .Cell(flexcpData, lngRow, COL_分解时间) = .TextMatrix(lngRow, COL_开始时间)
                        End If
                        .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                        .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                    End If
                End If
            End If
            
            If Not IsNull(rsSend!单次用量) Then
                lng单量数 = lng单量数 + 1 '决定是否显示单量列
            End If
            
            '如果是停用的，则提示不能发送
            If Format(NVL(rsSend!撤档时间, "3000-1-1"), "YYYY-MM-DD") <> Format("3000-1-1", "YYYY-MM-DD") Then
                .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                If InStr(str停用 & ",", "," & rsSend!医嘱内容 & ",") = 0 Then str停用 = str停用 & "," & rsSend!医嘱内容
            End If
            
            '计算项目的金额:用于查看及记帐报警
            '---------------------------------------------------------------
            cur金额 = 0
            If Not IsNull(rsSend!诊疗项目ID) Then
                Call LoadAdvicePrice(lngRow, cur金额)
            End If
            .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
            
            '相关行时的一些处理：累计显示一组医嘱的金额
            '---------------------------------------------------------------
            If Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 <> "C" Then
                '其它附加医嘱
                For j = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_ID)) = rsSend!相关ID Then
                        .TextMatrix(j, COL_金额) = Format(Val(.TextMatrix(j, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                        Exit For
                    End If
                Next
            ElseIf bln采集方法 Then
                '检验标本采集方法为显示行
                For j = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                        .TextMatrix(lngRow, COL_金额) = Format(Val(.TextMatrix(lngRow, COL_金额)) + Val(.TextMatrix(j, COL_金额)), gstrDec)
                    Else
                        Exit For
                    End If
                Next
            End If
            
            '其它处理
            '---------------------------------------------------------------
            '病人计数及分隔
            If rsSend!病人ID <> lng病人ID Then
                lng病人数 = lng病人数 + 1
                If lng病人ID <> 0 Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Not .RowHidden(j) Then
                            .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                            Exit For
                        End If
                    Next
                End If
            End If
            lng病人ID = rsSend!病人ID
            
NextLoop:           '---------------------------------------------------------------
            If blnOnePati Then Progress = i / rsSend.RecordCount * 100
            rsSend.MoveNext
        Next
        .Redraw = flexRDDirect
        
         '如果有停用的项目，则提示
        If str停用 <> "" Then
            Call MsgBox("诊疗项目：" & Mid(str停用, 2) & " 已经停用，不能发送。", vbInformation, Me.Caption)
        End If
    End With
    
    If blnOnePati Then Progress = 0
End Function

Private Function LoadAdvicePrice(ByVal lngRow As Long, cur合计 As Currency, Optional ByVal rsDrug As ADODB.Recordset) As Boolean
'功能：读取指定医嘱(仅当前行)的计价关系到临时记录集,并计算缺省发送金额(按费别打折)
'参数：rsDrug=包含待发送药品信息的记录集，发送药品医嘱时传入。因为可能按规格下达，医嘱中不一定有明确的药品ID。
'返回：cur合计=计算出的医嘱发送金额(非药变价未算,需要输入价格后才行)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, strPrice As String
    Dim str费用性质 As String, arr费用性质 As Variant
    Dim blnDo As Boolean, i As Long, k As Long
    Dim dbl数量 As Double, dbl单价 As Double, dbl应收 As Double
    Dim cur应收 As Currency, cur实收 As Currency
    Dim bln附加手术 As Boolean, lng项目ID As Long
    Dim lng主收入ID As Long, blnHaveSub As Boolean
    Dim lng执行科室ID As Long, cur金额 As Currency
    Dim lng材料ID As Long
    
    On Error GoTo errH
    
    cur金额 = 0
    With vsAdvice
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lngRow, COL_病人ID)), Val(vsAdvice.TextMatrix(lngRow, COL_主页ID)), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        
        If InStr(",5,6,7,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '不为院外执行(自备药),药品不可能为叮嘱,且固定正常计价
            If Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
                If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                    mrsPrice!相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
                End If
                mrsPrice!费用性质 = 0
                mrsPrice!收费方式 = 0
                mrsPrice!收费类别 = .TextMatrix(lngRow, COL_诊疗类别)
                mrsPrice!收费细目ID = rsDrug!药品ID
                mrsPrice!执行科室ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
                mrsPrice!数量 = 1
                mrsPrice!在用 = 1
                mrsPrice!变价 = NVL(rsDrug!是否变价, 0)
                mrsPrice!固定 = 1
                mrsPrice!从项 = 0
                                
                '发送的零售数量
                If .TextMatrix(lngRow, COL_诊疗类别) = "7" Then
                    '中药药房单位按不可分零处理:每付
                    If Val(.TextMatrix(lngRow, COL_可否分零)) = 0 Then
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量)) / NVL(rsDrug!剂量系数, 1)
                    Else
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_单量)) / NVL(rsDrug!剂量系数, 1) / NVL(rsDrug!住院包装, 1)) * NVL(rsDrug!住院包装, 1)
                    End If
                Else
                    dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * NVL(rsDrug!住院包装, 1)
                End If
                dbl数量 = Format(dbl数量, "0.00000")
                                
                '记录售价单价
                If NVL(rsDrug!是否变价, 0) = 0 Then
                    mrsPrice!单价 = Format(CalcPrice(rsDrug!药品ID, , , True, , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                Else '以售价计算药品时价,自备药时无对应药房
                    mrsPrice!单价 = Format(CalcDrugPrice(rsDrug!药品ID, Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                End If
                mrsPrice.Update
                                
                '计算医嘱发送金额(按费别打折的实收金额)
                If .TextMatrix(lngRow, COL_费别) <> "" Then
                    If NVL(rsDrug!是否变价, 0) = 0 Then
                        cur金额 = Format(CalcPrice(rsDrug!药品ID, .TextMatrix(lngRow, COL_费别), dbl数量, , Val(.TextMatrix(lngRow, COL_执行科室ID)), , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDec)
                    Else
                        cur金额 = Format(CalcDrugPrice(rsDrug!药品ID, Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, .TextMatrix(lngRow, COL_费别), , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), "0.00000")
                    End If
                Else
                    If gbln加班加价 Then
                        '处理加班加价
                        If NVL(rsDrug!是否变价, 0) = 0 Then
                            dbl单价 = Format(CalcPrice(rsDrug!药品ID, , , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        Else '以售价计算药品时价,自备药时无对应药房
                            dbl单价 = Format(CalcDrugPrice(rsDrug!药品ID, Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        End If
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * dbl单价, gstrDec)
                    Else
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * mrsPrice!单价, gstrDec)
                    End If
                End If
            End If
            
            cur合计 = cur金额
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "4" Then
            '不为院外执行(自备药),药品不可能为叮嘱,且固定正常计价
            If Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
                If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                    mrsPrice!相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
                End If
                mrsPrice!费用性质 = 0
                mrsPrice!收费方式 = 0
                mrsPrice!收费类别 = .TextMatrix(lngRow, COL_诊疗类别)
                mrsPrice!收费细目ID = Val(.TextMatrix(lngRow, COL_收费细目ID))
                mrsPrice!执行科室ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
                mrsPrice!数量 = 1
                mrsPrice!在用 = Val(.TextMatrix(lngRow, COL_跟踪在用))
                mrsPrice!变价 = Val(.TextMatrix(lngRow, COL_是否变价))
                mrsPrice!固定 = 1
                mrsPrice!从项 = 0
                                
                '发送的零售数量
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                                
                '记录售价单价
                If Val(.TextMatrix(lngRow, COL_是否变价)) = 0 Then
                    '定价卫材
                    mrsPrice!单价 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), , , True, , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                ElseIf Val(.TextMatrix(lngRow, COL_跟踪在用)) = 0 Then
                    '非跟踪在用的时价卫材，价格可能已保存在病人医嘱计价中
                    mrsPrice!单价 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), , , True, , Val(.TextMatrix(lngRow, COL_ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                Else
                    '计算跟踪在用卫材时价
                    mrsPrice!单价 = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                End If
                mrsPrice.Update
                                
                '计算医嘱发送金额(按费别打折的实收金额)
                If .TextMatrix(lngRow, COL_费别) <> "" Then
                    If Val(.TextMatrix(lngRow, COL_是否变价)) = 0 Then
                        cur金额 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), .TextMatrix(lngRow, COL_费别), dbl数量, , Val(.TextMatrix(lngRow, COL_执行科室ID)), , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDec)
                    ElseIf Val(.TextMatrix(lngRow, COL_跟踪在用)) = 0 Then
                        cur金额 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), .TextMatrix(lngRow, COL_费别), dbl数量, , Val(.TextMatrix(lngRow, COL_执行科室ID)), Val(.TextMatrix(lngRow, COL_ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDec)
                    Else
                        cur金额 = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, .TextMatrix(lngRow, COL_费别), , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), "0.00000")
                    End If
                Else
                    If gbln加班加价 Then
                        '处理加班加价
                        If Val(.TextMatrix(lngRow, COL_是否变价)) = 0 Then
                            dbl单价 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), , , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        ElseIf Val(.TextMatrix(lngRow, COL_跟踪在用)) = 0 Then
                            dbl单价 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), , , , , Val(.TextMatrix(lngRow, COL_ID)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        Else '以售价计算药品时价,自备药时无对应药房
                            dbl单价 = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        End If
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * dbl单价, gstrDec)
                    Else
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * mrsPrice!单价, gstrDec)
                    End If
                End If
            End If
            
            cur合计 = cur金额
        Else
            '取诊疗收费 关系中的对照(发送时才定计价):正常计价,不为叮嘱、院外执行
            If Val(.TextMatrix(lngRow, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                bln附加手术 = (.TextMatrix(lngRow, COL_诊疗类别) = "F" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0)
                
                '几种对应的计价情况
                If .TextMatrix(lngRow, COL_标本部位) <> "" And .TextMatrix(lngRow, COL_检查方法) <> "" Then
                    strPrice = " And 检查部位=[4] And 检查方法=[5] And Nvl(费用性质,0)=0"
                ElseIf Val(.TextMatrix(lngRow, COL_执行标记)) = 0 Then
                    strPrice = " And 检查部位 Is Null And 检查方法 is Null And Nvl(费用性质,0)=0"
                Else '目前包含床旁或术中加收的情况
                    strPrice = " And 检查部位 Is Null And 检查方法 is Null And Nvl(费用性质,0) IN(0,1)"
                End If
                
                strPrice = "Select 收费项目ID,固有对照 From (" & _
                    " Select c.收费项目ID, c.固有对照, c.适用科室id" & _
                    "   ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                    " From 诊疗收费关系 C Where C.诊疗项目ID=[2]" & strPrice & _
                    "       And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = Decode([3],0,[6],[3]) And C.病人来源 = 2)" & _
                    " ) Where Nvl(适用科室id, 0) = Top"
                
                '先读取已有的计价
                strSQL = _
                    " Select C.类别,A.收费细目ID as 收费项目ID,A.数量 as 收费数量,Nvl(E.固有对照,0) as 固有对照," & _
                    " B.收入项目ID,C.加班加价,B.加班加价率,Decode(C.是否变价,1,A.单价,B.现价)" & IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "") & " as 单价," & _
                    " C.是否变价,Nvl(A.从项,0) as 从项,D.跟踪在用,Nvl(A.执行科室ID,[3]) as 执行科室ID,C.屏蔽费别," & _
                    " Nvl(A.费用性质,0) as 费用性质,Nvl(A.收费方式,0) as 收费方式" & _
                    " From 病人医嘱计价 A,收费价目 B,收费项目目录 C,材料特性 D,(" & strPrice & ") E" & _
                    " Where A.医嘱ID=[1] And A.收费细目ID=0+E.收费项目ID(+)" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "B", "7", "8", "9") & _
                    " And A.收费细目ID=B.收费细目ID And A.收费细目ID=C.ID And A.收费细目ID=D.材料ID(+)" & _
                    " And C.服务对象 IN(2,3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " Order by 费用性质,从项,A.收费细目ID"
                
                
                '不读取默认的计价(不管是否有)
                '需要校对的模式下：这里发送的都是经过校对的，以实际已确定的计价内容为准，不能再读缺省计价，因为有可能校对或计价调整时已删除某些项目
                '不校对即发送的模式：只有在新开状态下才读取，因为发送后同上。
                If mblnAutoVerify And Val(.TextMatrix(lngRow, COL_医嘱状态)) = 1 Then
                    lng材料ID = 0 '检验试管费用,只收取试管对应的卫材费用
                    If .TextMatrix(lngRow, COL_试管编码) <> "" Then
                        lng材料ID = GetTubeMaterial(.TextMatrix(lngRow, COL_试管编码))
                    End If
                
                    '几种对应的计价情况
                    If .TextMatrix(lngRow, COL_标本部位) <> "" And .TextMatrix(lngRow, COL_检查方法) <> "" Then
                        strPrice = " And c.检查部位=[3] And c.检查方法=[4] And Nvl(c.费用性质,0)=0"
                    ElseIf Val(.TextMatrix(lngRow, COL_执行标记)) = 0 Then
                        strPrice = " And c.检查部位 Is Null And c.检查方法 is Null And Nvl(c.费用性质,0)=0"
                    Else '目前包含床旁或术中加收的情况
                        strPrice = " And c.检查部位 Is Null And c.检查方法 is Null And Nvl(c.费用性质,0) IN(0,1)"
                    End If
                    
                    strPrice = "Select * From (" & _
                        "Select C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,c.适用科室id" & _
                        " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                        " From 诊疗收费关系 C Where C.诊疗项目ID=[1]" & strPrice & _
                        "      And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = Decode([2],0,[6],[2]) And C.病人来源 = 2)" & _
                        " ) Where Nvl(适用科室id, 0) = Top"
                    
                    strSQL = _
                        " Select C.类别,A.收费项目ID,A.收费数量,A.固有对照,B.收入项目ID," & _
                        " C.加班加价,B.加班加价率,Decode(C.是否变价,1,B.缺省价格,B.现价)" & IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "") & " as 单价," & _
                        " C.是否变价,Nvl(A.从属项目,0) as 从项,D.跟踪在用,[2] as 执行科室ID,C.屏蔽费别," & _
                        " Nvl(A.费用性质,0) as 费用性质,Nvl(A.收费方式,0) as 收费方式" & _
                        " From (" & strPrice & ") A,收费价目 B,收费项目目录 C,材料特性 D" & _
                        " Where A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And A.收费项目ID=D.材料ID(+)" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "B", "7", "8", "9") & _
                        " And C.服务对象 IN(2,3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " And (Nvl(A.收费方式,0)=1 And C.类别='4' And A.收费项目ID=[5] Or Not(Nvl(A.收费方式,0)=1 And C.类别='4' And [5]<>0))" & _
                        " Order by 费用性质,从项,A.收费项目ID"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)), _
                        Val(.TextMatrix(lngRow, COL_执行科室ID)), .TextMatrix(lngRow, COL_标本部位), .TextMatrix(lngRow, COL_检查方法), lng材料ID, mlng病区ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                Else
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), _
                        Val(.TextMatrix(lngRow, COL_执行科室ID)), .TextMatrix(lngRow, COL_标本部位), .TextMatrix(lngRow, COL_检查方法), mlng病区ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                End If
                
                '确定计价之中是否包含从项以及主项收入ID
                arr费用性质 = Array()
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        If InStr(str费用性质 & ",", "," & rsTmp!费用性质 & ",") = 0 Then
                            str费用性质 = str费用性质 & "," & rsTmp!费用性质
                        End If
                        rsTmp.MoveNext
                    Loop
                    arr费用性质 = Split(Mid(str费用性质, 2), ",")
                End If
                                
                For k = 0 To UBound(arr费用性质)
                    rsTmp.Filter = "费用性质=" & arr费用性质(k)
                    
                    lng项目ID = 0: cur金额 = 0
                    lng主收入ID = 0: blnHaveSub = False
                    If Not rsTmp.EOF And gbln从项汇总折扣 Then
                        Do While Not rsTmp.EOF
                            If NVL(rsTmp!从项, 0) = 0 Then
                                'SQL中主项排在前面,只取主项目的第一个收入
                                If lng主收入ID = 0 Then lng主收入ID = rsTmp!收入项目ID
                            ElseIf NVL(rsTmp!从项, 0) = 1 Then
                                blnHaveSub = True: Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsTmp.MoveFirst
                    End If
                    
                    Do While True
                        blnDo = False
                        If rsTmp.EOF Then
                            If lng项目ID <> 0 Then blnDo = True
                        Else
                            If rsTmp!收费项目ID <> lng项目ID And lng项目ID <> 0 Then blnDo = True
                        End If
                        If blnDo Then
                            If Not IsNull(mrsPrice!单价) Then
                                mrsPrice!单价 = Format(mrsPrice!单价, gstrDecPrice)
                            End If
                            mrsPrice.Update
                            
                            '医嘱发送金额
                            cur金额 = cur金额 + Format(cur实收, gstrDec)
                        End If
                        If rsTmp.EOF Then Exit Do
                        
                        '------------------------------------
                        If rsTmp!收费项目ID <> lng项目ID Then
                            cur实收 = 0
                            mrsPrice.AddNew
                            mrsPrice!医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
                            If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                                mrsPrice!相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
                            End If
                            mrsPrice!费用性质 = NVL(rsTmp!费用性质, 0)
                            mrsPrice!收费方式 = NVL(rsTmp!收费方式, 0)
                            mrsPrice!收费类别 = rsTmp!类别
                            mrsPrice!收费细目ID = rsTmp!收费项目ID
                            mrsPrice!数量 = NVL(rsTmp!收费数量, 0)
                            mrsPrice!在用 = NVL(rsTmp!跟踪在用, 0)
                            mrsPrice!变价 = NVL(rsTmp!是否变价, 0)
                            mrsPrice!固定 = NVL(rsTmp!固有对照, 0)
                            mrsPrice!从项 = NVL(rsTmp!从项, 0)
                            
                            If .TextMatrix(lngRow, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5" And InStr(",5,6,", rsTmp!类别) > 0 Then
                                '原液皮试问题。绑定的药品费用如果没有指定科室则按原来逻辑
                                If Val(.TextMatrix(lngRow, COL_用药理由)) <> 0 Then
                                    lng执行科室ID = Val(.TextMatrix(lngRow, COL_用药理由))
                                Else
                                    lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
                                End If
                                lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!类别, rsTmp!收费项目ID, 4, Val(.TextMatrix(lngRow, COL_病人科室ID)), 0, 2, lng执行科室ID, , , 2)
                            Else
                                '执行科室:非药嘱药品及跟踪卫材的专门取
                                lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
                                If rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!类别) > 0 Then
                                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!类别, rsTmp!收费项目ID, 4, Val(.TextMatrix(lngRow, COL_病人科室ID)), 0, 2, lng执行科室ID, , , 2)
                                End If
                            End If
                            
                            If lng执行科室ID <> 0 Then
                                mrsPrice!执行科室ID = lng执行科室ID
                            Else
                                mrsPrice!执行科室ID = Null
                            End If
                        End If
                        lng项目ID = rsTmp!收费项目ID
                        
                        '计算单价和实收
                        If NVL(rsTmp!是否变价, 0) = 1 And InStr(",5,6,7,", rsTmp!类别) > 0 Then
                            '非药嘱药品计价按时价计算(仅一个收入),其它变价需要由医生输入
                            mrsPrice!单价 = CalcDrugPrice(rsTmp!收费项目ID, NVL(mrsPrice!执行科室ID, 0), dbl数量 * NVL(rsTmp!收费数量, 0), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            
                            dbl应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(mrsPrice!单价, gstrDecPrice)
                            
                            '处理加班加价
                            If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                                dbl应收 = dbl应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                            End If
    
                            cur应收 = Format(dbl应收, gstrDec)
                            
                            If .TextMatrix(lngRow, COL_费别) <> "" And Not (gbln从项汇总折扣 And blnHaveSub) And NVL(rsTmp!屏蔽费别, 0) = 0 Then
                                cur实收 = cur实收 + Format(ActualMoney(.TextMatrix(lngRow, COL_费别), rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                    mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur实收 + cur应收
                            End If
                        ElseIf NVL(rsTmp!是否变价, 0) = 1 And rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 Then
                            '跟踪在用的时价卫材和药品一样计算
                            mrsPrice!单价 = CalcDrugPrice(rsTmp!收费项目ID, NVL(mrsPrice!执行科室ID, 0), dbl数量 * NVL(rsTmp!收费数量, 0), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            
                            dbl应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(mrsPrice!单价, gstrDecPrice)
                            
                            '处理加班加价
                            If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                                dbl应收 = dbl应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                            End If
    
                            cur应收 = Format(dbl应收, gstrDec)
                            
                            If .TextMatrix(lngRow, COL_费别) <> "" And Not (gbln从项汇总折扣 And blnHaveSub) And NVL(rsTmp!屏蔽费别, 0) = 0 Then
                                cur实收 = cur实收 + Format(ActualMoney(.TextMatrix(lngRow, COL_费别), rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                    mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur实收 + cur应收
                            End If
                        Else '固定价格或普通变价(只有一个收入项目)
                            mrsPrice!单价 = NVL(mrsPrice!单价, 0) + NVL(rsTmp!单价, 0)
                            
                            dbl应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(NVL(rsTmp!单价, 0), gstrDecPrice)
                            
                            '处理加班加价
                            If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                                dbl应收 = dbl应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                            End If
                            
                            cur应收 = Format(dbl应收, gstrDec)
                            
                            If .TextMatrix(lngRow, COL_费别) <> "" And Not (gbln从项汇总折扣 And blnHaveSub) And NVL(rsTmp!屏蔽费别, 0) = 0 Then
                                cur实收 = cur实收 + Format(ActualMoney(.TextMatrix(lngRow, COL_费别), rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                    mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur实收 + cur应收
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                    
                    '从属项目汇总计算折扣
                    If gbln从项汇总折扣 And blnHaveSub And lng主收入ID <> 0 Then
                        cur金额 = Format(ActualMoney(.TextMatrix(lngRow, COL_费别), lng主收入ID, cur金额), gstrDec)
                    End If
                    
                    cur合计 = cur合计 + cur金额
                Next
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitBillSet()
'功能：初始化医嘱记帐单据生成记录集
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 30
    mrsBill.Fields.Append "费用序号", adBigInt
    mrsBill.Fields.Append "发送序号", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub InitSeekSet(rsSeek As ADODB.Recordset)
'功能：初始化用于汇总计算折扣的临时记录集
    Set rsSeek = New ADODB.Recordset
    rsSeek.Fields.Append "费用性质", adInteger
    rsSeek.Fields.Append "主项标签", adVariant
    rsSeek.Fields.Append "主收入ID", adBigInt
    rsSeek.Fields.Append "合计", adCurrency, , adFldIsNullable
    rsSeek.CursorLocation = adUseClient
    rsSeek.LockType = adLockOptimistic
    rsSeek.CursorType = adOpenStatic
    rsSeek.Open
End Sub

Private Sub InitPriceRecordset()
'功能：初始化医嘱计价记录集
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "费用性质", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "收费方式", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "收费类别", adVarChar, 1
    mrsPrice.Fields.Append "收费细目ID", adBigInt
    mrsPrice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble
    mrsPrice.Fields.Append "单价", adDouble, , adFldIsNullable '变价价格
    mrsPrice.Fields.Append "在用", adInteger '卫材是否跟踪在用
    mrsPrice.Fields.Append "变价", adInteger
    mrsPrice.Fields.Append "从项", adInteger
    mrsPrice.Fields.Append "固定", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset, rsSendDays As ADODB.Recordset)
'初始化记录集
    'SQL记录集
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-费用记录,2-医嘱记录,3-发送记录,4-发料记录
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '用于NO替换处理时排序
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '计价数量累计记录集
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsTotal.Fields.Append "项目ID", adBigInt
    rsTotal.Fields.Append "库房ID", adBigInt
    rsTotal.Fields.Append "数量", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    '医保上传记帐单
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsUpload.Fields.Append "NO", adVarChar, 30
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
    
    '计录试管编码
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "管码", adVarChar, 18
    rsNumber.Fields.Append "相关ID", adBigInt
    rsNumber.Fields.Append "样本条码", adVarChar, 18
    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
    rsNumber.Fields.Append "婴儿", adBigInt
    rsNumber.Fields.Append "紧急标志", adBigInt
    rsNumber.Fields.Append "标本", adVarChar, 18
    rsNumber.Fields.Append "采集科室ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '当前病人本次要发送的费用
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsMoneyNow.Fields.Append "诊疗项目ID", adBigInt
    rsMoneyNow.Fields.Append "收费项目ID", adBigInt
    rsMoneyNow.Fields.Append "试管编码", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "样本条码", adVarChar, 50, adFldIsNullable
    rsMoneyNow.Fields.Append "收费方式", adInteger
    rsMoneyNow.Fields.Append "收费时间", adVarChar, 10
    rsMoneyNow.Fields.Append "执行部门ID", adBigInt
    rsMoneyNow.Fields.Append "子医嘱ID", adBigInt '相关ID不为空的医嘱行的医嘱ID
    rsMoneyNow.Fields.Append "检查部位", adVarChar, 100
    rsMoneyNow.Fields.Append "检查方法", adVarChar, 100
    rsMoneyNow.Fields.Append "数量", adDouble '收费数量
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '当前病人本次发送的费用项目汇总
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "病人ID", adBigInt
    rsItems.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "医嘱ID", adBigInt
    rsItems.Fields.Append "收费类别", adVarChar, 1
    rsItems.Fields.Append "收费细目ID", adBigInt
    rsItems.Fields.Append "数量", adDouble
    rsItems.Fields.Append "单价", adDouble
    rsItems.Fields.Append "实收金额", adDouble
    rsItems.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
    '长嘱非药按天拆分发送记录
    Set rsSendDays = New ADODB.Recordset
    rsSendDays.Fields.Append "医嘱ID", adBigInt
    rsSendDays.Fields.Append "发送号", adBigInt
    rsSendDays.Fields.Append "计算时间", adVarChar, 100
    rsSendDays.Fields.Append "首次时间", adVarChar, 100
    rsSendDays.Fields.Append "末次时间", adVarChar, 100
    rsSendDays.Fields.Append "分解时间", adVarChar, 4000
    rsSendDays.Fields.Append "NO", adVarChar, 30
    rsSendDays.Fields.Append "发送数次", adDouble
    rsSendDays.CursorLocation = adUseClient
    rsSendDays.LockType = adLockOptimistic
    rsSendDays.CursorType = adOpenStatic
    rsSendDays.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng费用序号 As Long, lng发送序号 As Long)
'功能：获取当前记帐单据的NO及序号
'参数：lng费用序号=费用记录中的序号,为-1时表示不取费用序号
'      lng发送序号=发送记录中的序号,为-1时表示不取发送序号
'说明：strKey=根据记帐单据生成规则定的唯一关键字
'1.中西成药按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
'2.一个配方中的所有草药分配一个独立单据号
'3.材料医嘱与成药分号规则相同。
'4.其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
'5.检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
'6.一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        
        '取单据号
        'mrsBill!NO = zlDatabase.GetNextNo(14)
        mlngNOSequence = mlngNOSequence + 1
        mrsBill!NO = "TemporaryNO=" & Format(mlngNOSequence, "00000")
        
        mrsBill!费用序号 = IIF(lng费用序号 = -1, 0, 1)
        mrsBill!发送序号 = IIF(lng发送序号 = -1, 0, 1)
        mrsBill.Update
    Else
        If lng费用序号 <> -1 Then
            mrsBill!费用序号 = mrsBill!费用序号 + 1
        End If
        If lng发送序号 <> -1 Then
            mrsBill!发送序号 = mrsBill!发送序号 + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng费用序号 <> -1 Then lng费用序号 = mrsBill!费用序号
    If lng发送序号 <> -1 Then lng发送序号 = mrsBill!发送序号
End Sub

Private Sub ReplaceTrueNO(rsSQL As ADODB.Recordset, rsUpload As ADODB.Recordset)
'功能：将临时产生的NO替换成最终保存的真实NO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = zlDatabase.GetNextNo(14)
                            
                'rsUpload中一个NO只有一条记录
                rsUpload.Filter = "NO='" & rsSQL!NO & "'"
                If Not rsUpload.EOF Then
                    rsUpload!NO = strNO
                    rsUpload.Update
                End If
            End If
            
            rsSQL!Sql = Replace(rsSQL!Sql, rsSQL!NO, strNO)
            'rsSQL!NO = strNO '这个不更新，避免导致Sort后顺序紊乱
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Private Function CompletePatiSend(rsPati As ADODB.Recordset, rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, rsItems As ADODB.Recordset, ByVal cur合计 As Currency, ByVal cur记帐合计 As Currency, ByVal str类别 As String, _
    ByVal bln划价 As Boolean, blnTran As Boolean, ByVal lng发送号 As Long) As Boolean
'功能：提交一个病人的医嘱发送数据,在这之前处理记帐报警
'参数：rsPati=包含病人信息的记录集,用于记帐报警
'      rsSQL=包含所有要执行的SQL
'      rsUpload=用于医保上传的记帐单据号
'      rsItems=用于医保管控检查的项目汇总记录集
'      cur合计=病人本次要发送医嘱的记帐金额合计,用于记帐报警
'      cur记帐合计=病人本次要发送医嘱的记帐金额合计，包括本科执行后自动审核的划价费用，不含其它划价费用
'      bln划价=是否全部费用都是划价模式，用于报警的特殊处理
'      str类别=病人本次发送记帐费用的收费类别,用于记帐报警
'      lng发送号=本次发送的主关键字
'说明：如果出错,则在调用函数中处理,blnTran返回是否启用了事务
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intR As Integer, lng组ID As Long, str医嘱IDs As String, lngS As Long
    Dim i As Long, j As Long
    Dim strMsg As String, strAllmsg As String, strDiag As String, strAdviceInfo As String
    Dim strErr As String
    Dim str当前病情 As String
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim blnClearPatiCache As Boolean
    Dim blnPlugIn As Boolean
    Dim rsAdviceRis As ADODB.Recordset
    Dim strAdvices输血 As String
    Dim var输血 As Variant
    Dim bln记帐提醒忽略 As Boolean
    Dim intBnt As Integer
    
    lng病人ID = Val(rsPati!病人ID & "")
    lng主页ID = Val(rsPati!主页ID & "")
    blnClearPatiCache = True
    
    '调用外挂接口发送前检查医嘱费用
    If CreatePlugInOK(p住院医嘱发送, 1) Then
        blnPlugIn = True
        On Error Resume Next
        blnPlugIn = gobjPlugIn.AdviceCheckSendFee(glngSys, p住院医嘱发送, lng病人ID, lng主页ID, cur合计, 1)
        If Not blnPlugIn And err.Number <> 0 Then blnPlugIn = True
        Call zlPlugInErrH(err, "AdviceCheckSendFee")
        err.Clear: On Error GoTo 0
        If Not blnPlugIn Then
            Exit Function
        End If
    End If
     
    '病人费用报警
    If cur合计 > 0 Then
        If InitObjPublicExpense Then
            For i = 1 To Len(str类别)
                intBnt = mintBnt
                Call gobjPublicExpense.zlBillingWarn.zlBillingWarnCheck(Me, 1, IIF(bln划价, 1, 0), lng病人ID, lng主页ID, mlng病区ID, Mid(str类别, i, 1), IIF(gbln报警包含划价费用, cur合计, cur记帐合计), InStr(";" & GetInsidePrivs(p住院医嘱发送) & ";", ";欠费强制记帐;") > 0, False, blnClearPatiCache, intR, , , , True, True, bln记帐提醒忽略, intBnt)
                blnClearPatiCache = False
                If bln记帐提醒忽略 And Not mbln记帐提醒忽略 Then
                    mbln记帐提醒忽略 = True
                    mintBnt = IIF(InStr(",2,3,", intR) > 0, vbCancel, vbIgnore)
                End If
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        '医保管控实时监测
        If Not IsNull(rsPati!险类) Then
            If gclsInsure.GetCapability(support实时监控, rsPati!病人ID, rsPati!险类) Then
                rsItems.Filter = 0
                If Not rsItems.EOF Then
                    If Not gclsInsure.CheckItem(rsPati!险类, 1, 2, rsItems) Then
                        CompletePatiSend = True: Exit Function '可以继续下一个病人
                    End If
                End If
            End If
        End If
        Call InitObjLis(p住院护士站)
        'LIS申请接口
        If Not gobjLIS Is Nothing Then
            strAdviceInfo = Get检验医嘱信息(lng病人ID, lng主页ID)
            If strAdviceInfo <> "" Then
                Set rsTmp = Get病人诊断记录(lng病人ID, lng主页ID, "2")
                If rsTmp.RecordCount > 0 Then strDiag = rsTmp!诊断描述
            End If
        End If
            
        If gbln血库系统 Then
            If InitObjBlood(True) Then
                strAdvices输血 = Get输血医嘱信息(lng病人ID, lng主页ID)
                If strAdvices输血 <> "" Then
                    var输血 = Split(strAdvices输血, ",")
                End If
            End If
        End If
        
        Call ReplaceTrueNO(rsSQL, rsUpload)
        
        '执行顺序:1-费用,2-医嘱执行科室,3-发送,4-自动发料
        '1.先填写费用,因为发送时可能处理费用
        '2.对费用记录按收费细目ID排序插入
        rsSQL.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        rsSQL.Sort = "类型,项目ID,序号"
        rsUpload.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        
        If mblnSendMsg Then
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        If .TextMatrix(i, COL_诊疗类别) = "Z" And (Val(.TextMatrix(i, COL_操作类型)) = 9 Or Val(.TextMatrix(i, COL_操作类型)) = 10) Then
                            str当前病情 = Get病人当前病情(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)))
                            Exit For
                        End If
                    End If
                Next
            End With
        End If
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!Sql, Me.Caption)
            rsSQL.MoveNext
        Loop
                
        '调用LIS申请接口
        If strAdviceInfo <> "" Then
            If gobjLIS.SendLisApplicationForm(strAdviceInfo, strDiag) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                Call Del检验申请
                MsgBox "检验接口调用失败，不能发送检验医嘱。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        '医保数据上传
        strAllmsg = ""
        If Not IsNull(rsPati!险类) Then
            If gclsInsure.GetCapability(support医嘱上传, rsPati!病人ID, rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, rsPati!病人ID, rsPati!险类) Then
                Do While Not rsUpload.EOF
                    strMsg = "" '因为现在一张NO内肯定为一个病人的,所以最后病人参数可以不传
                    'strAdvance中传入“总单据数|当前单据数”以便医保接口处理
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!险类, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                        '未提交前上传失败则回滚并中止发送
                        gcnOracle.RollbackTrans: blnTran = False
                        Screen.MousePointer = 0
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName '每张提示
                        Else
                            MsgBox rsPati!姓名 & "的费用上传失败，发送操作将被中止。", vbExclamation, gstrSysName
                        End If
                        Exit Function
                    Else
                        If strMsg <> "" Then strAllmsg = strAllmsg & rsUpload!NO & ":" & strMsg & vbCrLf
                    End If
                    rsUpload.MoveNext
                Loop
            End If
            
            '医保档案上传接口(事务内以限制)
            If gclsInsure.GetCapability(support上传住院档案, rsPati!病人ID, rsPati!险类) Then
                If Not gclsInsure.TranElecDossier(2, rsPati!病人ID, rsPati!主页ID, rsPati!险类) Then Exit Function
            End If
        End If
        If strAdvices输血 <> "" Then
            strErr = ""
            For i = 0 To UBound(var输血)
                If gobjPublicBlood.AdviceOperation(p住院医嘱发送, Val(var输血(i)), 5, False, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        End If
        gcnOracle.CommitTrans: blnTran = False
        If strAllmsg <> "" Then
            Screen.MousePointer = 0
            MsgBox strAllmsg, vbInformation, gstrSysName
        End If
        
        '医保数据上传
        If Not IsNull(rsPati!险类) Then
            If gclsInsure.GetCapability(support医嘱上传, rsPati!病人ID, rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, rsPati!病人ID, rsPati!险类) Then
                Do While Not rsUpload.EOF
                    strMsg = ""
                    Screen.MousePointer = 0
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!险类, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                        '提交后上传失败,仅提示
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox rsPati!姓名 & "的记帐单""" & rsUpload!NO & """上传失败，HIS端数据已提交，按确定继续发送。", vbExclamation, gstrSysName
                        End If
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                    End If
                    Screen.MousePointer = 11
                    rsUpload.MoveNext
                Loop
            End If
        End If
        
        '消息发送
        If mblnSendMsg Then
            Call SendMsg发送(lng发送号, IIF(bln划价, 1, 2), str当前病情)
        End If
        
        'RIS接口
        If HaveRIS Then
            If GetAdviceRis(rsAdviceRis, lng病人ID, lng主页ID) Then
                On Error Resume Next
                If gobjRis.HISSendAdvice(rsAdviceRis, 2, lng病人ID, lng主页ID, "", lng发送号) <> 1 Then
                    MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISSendAdvice)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo 0
            End If
        ElseIf gbln启用影像信息系统接口 = True Then
            MsgBox "当前启用了影像信息系统接口， 但于由RIS接口创建失败未调用(HISSendAdvice)接口，请与系统管理员联系。", vbInformation, gstrSysName
        End If
        
        '提交成功,将病人医嘱行标记为可删除
        With vsAdvice
            lngS = .FindRow(CStr(rsPati!病人ID), , COL_病人ID)
            For i = lngS To .Rows - 1
                If Val(.TextMatrix(i, COL_病人ID)) = rsPati!病人ID Then
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        .RowData(i) = -1
                    End If
                Else
                    Exit For
                End If
            Next
            '调用外挂接口
            If CreatePlugInOK(p住院医嘱发送) Then
                On Error Resume Next
                Call gobjPlugIn.AdviceSend(glngSys, p住院医嘱发送, lng病人ID, lng主页ID, lng发送号)
                Call zlPlugInErrH(err, "AdviceSend")
                err.Clear: On Error GoTo 0
            End If
            If gobjExchange Is Nothing Then
                On Error Resume Next
                Set gobjExchange = CreateObject("zlExchange.clsExchange")
                If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
                err.Clear: On Error GoTo 0
            End If
            '调用数据交换平台，向LIS,PACS传递申请单
            If Not gobjExchange Is Nothing Then
                For i = lngS To .Rows - 1
                    If Val(.TextMatrix(i, COL_病人ID)) = rsPati!病人ID And Val(.TextMatrix(i, COL_医嘱状态)) <> 1 Then
                        If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                            'c-检验,d-检查
                            If .TextMatrix(i, COL_诊疗类别) = "C" Or .TextMatrix(i, COL_诊疗类别) = "D" Then
                                If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                    lng组ID = Val(.TextMatrix(i, COL_相关ID))
                                Else
                                    lng组ID = Val(.TextMatrix(i, COL_ID))
                                End If
                                If InStr(str医嘱IDs & ",", "," & lng组ID & ",") = 0 Then
                                    str医嘱IDs = str医嘱IDs & "," & lng组ID
                                    Call gobjExchange.SendMsg(IIF(.TextMatrix(i, COL_诊疗类别) = "C", 1, 2), "病人ID::" & rsPati!病人ID & "||主页ID::" & rsPati!病人ID & "||医嘱ID::" & lng组ID & "||操作类型::1")
                                End If
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
        End With
    End If
    CompletePatiSend = True
End Function

Private Sub SendMsg发送(ByVal lng发送号 As Long, ByVal int单据性质 As Integer, ByVal str当前病情 As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strIDs As String
    Dim lngTmp As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    Dim str开嘱科室 As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errH
    strSQL = "select 名称 from 部门表 where id=[1]"
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If Val(.TextMatrix(i, COL_执行安排)) = 1 Then '申请安排
                    Call ZLHIS_CIS_004(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_姓名), .TextMatrix(i, COL_住院号), , IIF(Val(.TextMatrix(i, COL_病人性质)) = 1, 1, 2), _
                        .TextMatrix(i, COL_主页ID), .TextMatrix(i, COL_病人科室ID), "", , .TextMatrix(i, COL_床号), Val(.TextMatrix(i, COL_ID)), IIF("长嘱" = .TextMatrix(i, COL_医嘱期效), 0, 1), .TextMatrix(i, COL_诊疗类别), .TextMatrix(i, COL_操作类型), _
                        lng发送号, .TextMatrix(i, COL_执行科室ID))
                End If
                '检验医嘱
                If .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 6 Then
                    strIDs = ""
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_相关ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_诊疗类别) = "C" Then
                                strIDs = strIDs & "," & Val(.TextMatrix(j, COL_ID))
                                lngTmp = Val(.TextMatrix(j, COL_执行科室ID))
                            End If
                        End If
                    Next
                    strIDs = Mid(strIDs, 2)
                    If strIDs <> "" Then
                        Call ZLHIS_CIS_016(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_姓名), .TextMatrix(i, COL_住院号), , IIF(Val(.TextMatrix(i, COL_病人性质)) = 1, 1, 2), _
                        .TextMatrix(i, COL_主页ID), Val(.TextMatrix(i, COL_病人科室ID)), , Val(.TextMatrix(i, COL_ID)), _
                            .TextMatrix(i, COL_标本部位), .TextMatrix(i, COL_诊疗项目ID), , .TextMatrix(i, COL_执行科室ID), , strIDs, , lngTmp, , lng发送号, "", _
                            int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), , "")
                    End If
                '检查申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    strTmp1 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_相关ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_诊疗类别) = "D" Then
                                strTmp1 = strTmp1 & "," & .TextMatrix(j, COL_标本部位)
                            End If
                        End If
                    Next
                    strTmp1 = Mid(strTmp1, 2)
                    Call ZLHIS_CIS_017(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_姓名), .TextMatrix(i, COL_住院号), , IIF(Val(.TextMatrix(i, COL_病人性质)) = 1, 1, 2), _
                        .TextMatrix(i, COL_主页ID), Val(.TextMatrix(i, COL_病人科室ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_诊疗项目ID), .TextMatrix(i, col_医嘱内容), strTmp1, .TextMatrix(i, COL_执行科室ID), , lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), , "")
                 
                '手术申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    strTmp1 = Get医嘱附项内容(Val(.TextMatrix(i, COL_ID)), "主刀医生")
                    strTmp2 = Get医嘱附项内容(Val(.TextMatrix(i, COL_ID)), "助手医生")
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_相关ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_诊疗类别) = "F" Then
                                strIDs = strIDs & "," & .TextMatrix(j, COL_ID)
                            ElseIf .TextMatrix(j, COL_诊疗类别) = "G" Then
                                lngTmp = Val(.TextMatrix(j, COL_ID))
                            End If
                        End If
                    Next
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_开嘱科室ID)))
                    If Not rsTmp.EOF Then str开嘱科室 = rsTmp!名称 & ""
                    strIDs = Mid(strIDs, 2)
                    Call ZLHIS_CIS_018(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_姓名), .TextMatrix(i, COL_住院号), , IIF(Val(.TextMatrix(i, COL_病人性质)) = 1, 1, 2), _
                        .TextMatrix(i, COL_主页ID), Val(.TextMatrix(i, COL_病人科室ID)), "", Val(.TextMatrix(i, COL_ID)), strIDs, , lngTmp, , strTmp1, strTmp2, .TextMatrix(i, COL_执行科室ID), .TextMatrix(i, COL_执行科室), lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), str开嘱科室)
                
                '输血申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "K" Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_开嘱科室ID)))
                    If Not rsTmp.EOF Then str开嘱科室 = rsTmp!名称 & ""
                    Call ZLHIS_CIS_019(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_姓名), .TextMatrix(i, COL_住院号), , IIF(Val(.TextMatrix(i, COL_病人性质)) = 1, 1, 2), _
                        .TextMatrix(i, COL_主页ID), Val(.TextMatrix(i, COL_病人科室ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_执行科室ID), .TextMatrix(i, COL_执行科室), lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), str开嘱科室)
                 
                '会诊申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",7,8,11,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0 _
                    Or .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 5 Then
                    If .TextMatrix(i, COL_操作类型) = "7" Then
                        strTmp1 = "ZLHIS_CIS_020"  '会诊申请
                    ElseIf .TextMatrix(i, COL_操作类型) = "8" Then
                        strTmp1 = "ZLHIS_CIS_021"  '患者抢救医嘱
                    ElseIf .TextMatrix(i, COL_操作类型) = "11" Then
                        strTmp1 = "ZLHIS_CIS_022"  '患者死亡医嘱
                    ElseIf .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 5 Then
                        strTmp1 = "ZLHIS_CIS_023"  '患者特殊治疗医嘱
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_开嘱科室ID)))
                    If Not rsTmp.EOF Then str开嘱科室 = rsTmp!名称 & ""
                    Call SendMsg(strTmp1, mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_姓名), .TextMatrix(i, COL_住院号), , IIF(Val(.TextMatrix(i, COL_病人性质)) = 1, 1, 2), _
                        .TextMatrix(i, COL_主页ID), Val(.TextMatrix(i, COL_病人科室ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_执行科室ID), .TextMatrix(i, COL_执行科室), lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), str开嘱科室)
                 
                '住院患者预出院
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 5 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 10, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_009(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_主页ID), .TextMatrix(i, COL_姓名), .TextMatrix(i, col_性别), .TextMatrix(i, COL_住院号), _
                        lngTmp, .TextMatrix(i, COL_开嘱时间), Val(.TextMatrix(i, COL_病人病区ID)), , Val(.TextMatrix(i, COL_病人科室ID)), "", , .TextMatrix(i, COL_床号), Val(.TextMatrix(i, COL_ID)))
                
                '住院患者病情变更
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And (Val(.TextMatrix(i, COL_操作类型)) = 9 Or Val(.TextMatrix(i, COL_操作类型)) = 10) Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 13, lngTmp, strTmp1)
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_病人科室ID)))
                    strTmp2 = "": If Not rsTmp.EOF Then strTmp2 = rsTmp!名称 & ""
                    Call ZLHIS_PATIENT_005(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_主页ID), .TextMatrix(i, COL_姓名), .TextMatrix(i, col_性别), .TextMatrix(i, COL_住院号), _
                        Val(.TextMatrix(i, COL_病人病区ID)), , Val(.TextMatrix(i, COL_病人科室ID)), strTmp2, str当前病情, lngTmp, .TextMatrix(i, COL_开嘱时间), strTmp1, .TextMatrix(i, COL_开嘱医生), Val(.TextMatrix(i, COL_ID)))
                 
                '住院患者转出科室
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 3, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_003(mclsMipModule, Val(.TextMatrix(i, COL_病人ID)), .TextMatrix(i, COL_主页ID), .TextMatrix(i, COL_姓名), .TextMatrix(i, col_性别), .TextMatrix(i, COL_住院号), _
                        Val(.TextMatrix(i, COL_病人病区ID)), , Val(.TextMatrix(i, COL_病人科室ID)), "", , .TextMatrix(i, COL_住院号), _
                        lngTmp, .TextMatrix(i, COL_开嘱时间), , , Val(.TextMatrix(i, COL_执行科室ID)), , Val(.TextMatrix(i, COL_ID)))
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub DeleteSendRow()
'功能：将待发送医嘱清单中已发送成功的的行删除
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_选择
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Get实收金额(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Get实收金额 = CCur(strSQL)
End Function

Private Function Set实收金额(ByVal strSQL As String, ByVal cur金额 As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Set实收金额 = strLeft & cur金额 & strRight
End Function

Private Function GetMergeDrugStore(ByVal lngRow As Long) As Long
'功能：获取一并给药的基准药房，用于生成发送NO的Key值
'说明：一并给药的药品发送到一起，包括自备药和不同药房的情况
    Dim lng药房ID As Long, lngBegin As Long, i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_相关ID)) <> Val(.TextMatrix(lngRow - 1, COL_相关ID)) And Val(.TextMatrix(lngRow, COL_执行科室ID)) <> 0 Then
            lng药房ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        Else
            lngBegin = .FindRow(.TextMatrix(lngRow, COL_相关ID), , COL_相关ID)
            For i = lngBegin To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    If Val(.TextMatrix(i, COL_执行科室ID)) <> 0 Then
                        lng药房ID = Val(.TextMatrix(i, COL_执行科室ID)): Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetMergeDrugStore = lng药房ID
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng项目ID As Long, ByVal int费用性质 As Integer, ByVal lngCol As Long)
'功能：定位到并显示指定医嘱的指定计价行
'参数：lngRow=医嘱行号
'      lng项目ID=计价项目ID
'      lngCol=计价表格显示列
    Dim k As Long
    
    With vsAdvice
        .Col = col_医嘱内容 '进入行自动ShowPrice,mrsPrice发生变化
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                '附加手术,手术麻醉,检查部位,检验组合项目
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '给药途径
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '中药煎法
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_行号)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_费用性质)) = int费用性质 _
                And Val(vsPrice.TextMatrix(k, COLP_收费细目ID)) = lng项目ID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Public Function SendAdvice() As Long
'功能：处理医嘱发送(这个过程中记帐报警)
'说明：逐个病人发送提交
'返回：如果成功，则返回发送号
'rsSQL!类型=1-校对(如果不需要先校对),2-医嘱计价,3-住院记帐,4-执行科室替换，5-医嘱发送，6-自动发料,7-输液配药
    Dim rsPati As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    
    Dim rsSQL As ADODB.Recordset '用于组织SQL语句的动态记录集
    Dim rsTotal As ADODB.Recordset '用于库存汇总检查的动态记录集
    Dim rsUpload As ADODB.Recordset '用于收集医保上传单据号的动态记录集
    Dim rsNumber As ADODB.Recordset '用于生成条码的动态记录集
    Dim rsItems As ADODB.Recordset '用于医保管控的费用记录集,动态记录集
    Dim rsMoneyNow As ADODB.Recordset '当前病人本次要发送的费用,动态记录集
    Dim rsMoneyDay As ADODB.Recordset '当前病人当天已发送的费用,静态记录集
    Dim rsAudit As ADODB.Recordset     '医保审批记录集
    Dim rsExec As ADODB.Recordset  '医嘱执行计价
    Dim rsClone As ADODB.Recordset, rsSeek As ADODB.Recordset '用汇总打折计算的动态记录集
    Dim rsDays As ADODB.Recordset  '长嘱非药按天拆分费用
    Dim rsSendDays As ADODB.Recordset '长嘱非药按天拆分发送记录
    Dim rsDayNo As ADODB.Recordset '发送号，如果拆分。则一天对应一个发送号
    Dim rsTmp As ADODB.Recordset
    
    Dim i As Long, j As Long
    Dim strSQL As String, strTmp As String
    Dim blnTran As Boolean, strCurDate As String, strCurDateTmp As String
    Dim str类别 As String, str毒理 As String
    
    Dim lng病人ID As Long, lng主页ID As Long, lng病人性质 As Long
    Dim lng发送号 As Long, int计费状态 As Integer, bln划价 As Boolean, int划价 As Integer, strNO As String
    Dim str收费项目 As String, lng费用序号 As Long, lng费用父号 As Long, lng发送序号 As Long, lng组ID As Long, lngOld组ID As Long
    Dim int付数 As Integer, dbl数量 As Double, cur合计 As Currency, cur记帐合计 As Currency
    Dim dbl单价 As Double, dbl应收 As Double, cur应收 As Currency, cur实收 As Currency
    Dim bln保险项目否 As Boolean, lng保险大类ID As Long, cur统筹金额 As Currency, str保险编码 As String, str费用类型 As String
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim int配方数 As Integer, strNOKey As String, str自动发料 As String
    Dim str发生时间 As String, str登记时间 As String
    Dim dbl发送数次 As Double, blnFirst As Boolean '配方数及分号关键字
    Dim lng病人科室ID As Long, lng执行科室ID As Long, int执行状态 As Integer
    Dim bln附加手术 As Boolean, int药品性质 As Integer, blnBool As Boolean
    
    Dim strHaveSub As String, strNoneSub As String
    Dim int父序号 As Integer, lng父项目ID As Long, str实收 As String
    
    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean
    Dim bln药品零差价提示 As Boolean
    Dim str领药号 As String, strCuvetteNumber As String '生成条码
    
    Dim strAudit As String
    Dim bln实时监控 As Boolean, blnSend As Boolean, blnOldSend As Boolean, blnSendPrivs As Boolean
    Dim lng费用次数 As Long '一天只收一次时，本次发送应收取的费用次数
    Dim strMsg As String, blnExists As Boolean, lngSpecialAdviceID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim str特殊 As String, str特殊医嘱 As String
    Dim lng主医嘱行 As Long
    Dim lng采集科室ID As Long
    Dim lng发送号Tmp As Long
    Dim strPlugIn发送号 As String '外挂接口用
    Dim lngLastPatiID As Long, str给药IDs As String, lngLastPageID As Long, lngLastPatiDeptID As Long
    Dim str部位方法 As String '检查项目的部位方法，固定格式，检查部位<sTab>检查方法，如："头部<sTab>平扫"
    Dim dblOther数量 As Double '费用项目收费次数
    Dim str关联药行  As String '关联的药品行医嘱 ,"皮试医嘱ID,药品行医嘱ID"
    Dim rs皮试 As ADODB.Recordset
    Dim strMinDate As String
    
    On Error GoTo errH
    
    '先检查并提示特殊医嘱:3-转科,5-出院,6-转院,11-死亡
    With vsAdvice
        If InitObjRecipeAudit(p住院医嘱下达) Then
            '处方审查系统产生待审数据
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    If (.TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6") And Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                        If lngLastPatiID <> Val(.TextMatrix(i, COL_病人ID)) Then
                            If Mid(str给药IDs, 2) <> "" Then
                                Call gobjRecipeAudit.BuildData(Mid(str给药IDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
                                str给药IDs = ""
                            End If
                        End If
                        lngLastPatiID = Val(.TextMatrix(i, COL_病人ID))
                        lngLastPageID = Val(.TextMatrix(i, COL_主页ID))
                        lngLastPatiDeptID = Val(.TextMatrix(i, COL_病人科室ID))
                        If InStr("," & str给药IDs & ",", "," & .TextMatrix(i, COL_相关ID) & ",") = 0 Then str给药IDs = str给药IDs & "," & .TextMatrix(i, COL_相关ID)
                    End If
                End If
            Next
            If Mid(str给药IDs, 2) <> "" Then
                Call gobjRecipeAudit.BuildData(Mid(str给药IDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
            End If
        End If
        
        strTmp = ""
        strMinDate = "3000-01-01 00:00"
        blnExists = False
        Call InitExecRecordset(rsExec)   '医嘱执行计价
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                    strTmp = strTmp & vbCrLf & .TextMatrix(i, COL_姓名) & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容)
                    
                    If mblnCheckAdvice Then
                        lngSpecialAdviceID = Val(.TextMatrix(i, COL_ID))
                        strMsg = CheckUnExecutedAdvice(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)), lngSpecialAdviceID, Val(.Cell(flexcpData, i, COL_婴儿)))
                        If strMsg <> "" Then
                            If blnExists = False Then blnExists = True
                            
                            Call GetPatiRowScope(i, lngBegin, lngEnd)
                            '取消勾选
                            For j = lngBegin To lngEnd
                                If .Cell(flexcpData, j, COL_选择) = 0 Then
                                    Set .Cell(flexcpPicture, j, COL_选择) = Nothing
                                End If
                            Next
                        End If
                    End If

                    '转科医嘱处理时判断长嘱可以继续发送
                    If Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                        If CheckCanSendAdvice(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)), lngSpecialAdviceID, Val(.Cell(flexcpData, i, COL_婴儿))) Then
                            Call MsgBox("发现转科医嘱：" & vbCrLf & .TextMatrix(i, COL_姓名) & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & "必须将可以发送的长期医嘱处理后才能发送。", vbInformation, gstrSysName)
                            Exit Function
                        End If
                    End If
                    
                    '转科时未审核销帐单据检查
                    If Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                        If CheckWaitQuittance(Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID))) Then Exit Function
                    End If
                    
                End If
                
                 '3-转科;4-术后;5-出院;6-转院,11-死亡,14-术前
                If Mid(gstrESign, 2, 1) = "1" Then  '住院医生站启用了电子签名才检查
                    If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,4,5,6,11,14,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                        If InStr(str特殊 & ",", "," & .TextMatrix(i, COL_病人ID) & ":" & .TextMatrix(i, COL_主页ID) & ",") = 0 Then
                            str特殊 = str特殊 & "," & .TextMatrix(i, COL_病人ID) & ":" & .TextMatrix(i, COL_主页ID)
                        End If
                        
                        str特殊医嘱 = str特殊医嘱 & vbCrLf & .TextMatrix(i, COL_姓名) & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容)
                    End If
                End If
                
                '如果是无需校对模式，则检查新开医嘱是否并发修改了(为提高性能，只检查一组中的主记录，因为一组医嘱的修改时间是相同的)
                If Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                    If mblnAutoVerify And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                        If CheckAdviceUpdate(Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_新开操作时间)) Then
                            MsgBox "医嘱：" & .TextMatrix(i, col_医嘱内容) & vbCrLf & "已经被修改，请重新读取医嘱后再发送。", vbInformation, "病人医嘱发送"
                            Exit Function
                        End If
                    End If
                End If
                
                '毒理分类判断信息收集
                If gbln特殊药品分开发送 Then
                    If cboDrugType.ListIndex = 0 Then
                        If InStr("," & str毒理 & ",", "," & .TextMatrix(i, COL_毒理分类) & ",") = 0 Then
                            str毒理 = str毒理 & "," & .TextMatrix(i, COL_毒理分类)
                        End If
                    ElseIf cboDrugType.ListIndex = 3 Then
                        str毒理 = ""
                    Else
                        str毒理 = ",毒性药"
                    End If
                End If
                If .TextMatrix(i, COL_首次时间) < strMinDate Then
                    strMinDate = .TextMatrix(i, COL_首次时间)
                End If
            End If
        Next
        If strMinDate = "3000-01-01 00:00" Then strMinDate = ""
        
        If str毒理 <> "" And cboDrugType.ListIndex = 0 Then
            If Not (str毒理 = ",毒性药" Or str毒理 = ",精神I类" Or str毒理 = ",麻醉药" Or str毒理 = ",麻醉药,精神I类" Or str毒理 = ",精神I类,麻醉药") Then
                If Not (InStr(str毒理 & ",", ",毒性药,") = 0 And InStr(str毒理 & ",", ",麻醉药,") = 0 And InStr(str毒理 & ",", ",精神I类,") = 0) Then
                    MsgBox "本次发送的医嘱中可能包含毒麻精类药品，需分别发送，请修改过滤条件重新读取医嘱后再发送。", vbInformation, gstrSysName
                    Exit Function
                Else
                    str毒理 = ""
                End If
            End If
        End If
        
        If strTmp <> "" Then
            If mblnCheckAdvice And blnExists Then
                Call MsgBox("发现以下特殊医嘱：" & vbCrLf & strTmp & vbCrLf & vbCrLf & "必须将" & strMsg & "处理后才能发送。", vbInformation, gstrSysName)
                
            ElseIf MsgBox("要发送的医嘱中包含下列特殊医嘱：" & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送当前选择的医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If MsgBox("确实要发送当前选择的医嘱吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        '如果启用了电子签名，检查存在"已停止但未确认停止"的医嘱，提示护士先进行确认停止
        '因为特殊医嘱校对时会将"已停止但未确认停止"的医嘱的"执行终止时间"调整为特殊医嘱的开始执行时间，医嘱停止的签名源文包含了"执行终止时间"，这会导致签名验证无法通过
        If str特殊 <> "" Then
            str特殊 = Mid(str特殊, 2)
            strTmp = ""
            If CheckStopedUnAffirm(str特殊, strTmp) Then
                MsgBox "要发送的医嘱中包含下列特殊医嘱：" & vbCrLf & str特殊医嘱 & _
                    vbCrLf & vbCrLf & "发送后会将未确认停止的医嘱重新停止，为了不影响签名验证，请先对以下病人进行确认停止操作：" & strTmp, vbInformation, gstrSysName
                Exit Function
            End If
            strTmp = ""
        End If
        
        '阳性用药
        If mbln阳性用药 Then
            blnBool = Set阳性用药()
            If Not blnBool Then
                GoTo FuncEnd
            End If
        End If
        
        If Not zlPluginAdviceBeforeSend Then
            Exit Function
        End If
    End With
    
    'RIS预约检查判断提示
    Call CheckRISScheduling
    
    Screen.MousePointer = 11
    
    blnSendPrivs = InStr(GetInsidePrivs(p住院医嘱发送), "全院医嘱发送") > 0
    mstrRollNotify = "": mstr领药号 = ""
    bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
    bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True
    bln药品零差价提示 = True
    mbln记帐提醒忽略 = False
    mintBnt = -1
    
    Call InitBillSet
    lng发送号 = zlDatabase.GetNextNo(10)        '如果全部是新开长嘱，且指定结束时间内无发送（次数为零），不执行发送时会浪费一个号
    mlngNOSequence = 0 '单据号序列重新初始
    mdatCurr = zlDatabase.Currentdate
    strCurDateTmp = Format(mdatCurr, "yyyy-MM-dd HH:mm:ss")
    strCurDate = "To_Date('" & strCurDateTmp & "','YYYY-MM-DD HH24:MI:SS')"
    int配方数 = 1 '表示发送的第几付配方,用于分单据号
    '一天一个发送号
    Set rsDayNo = New ADODB.Recordset
    rsDayNo.Fields.Append "发送号", adBigInt
    rsDayNo.Fields.Append "计算时间", adVarChar, 100
    rsDayNo.CursorLocation = adUseClient
    rsDayNo.LockType = adLockOptimistic
    rsDayNo.CursorType = adOpenStatic
    rsDayNo.Open
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                   
                '新开的长嘱，读取医嘱时按界面指定的结束时间计算后不需发送的（次数为零）
                '自由录入的临嘱和长嘱不发送
                '特殊长嘱只校对不发送:护理等级,病重/危医嘱,记录入出量医嘱不发送(如果没有互斥医嘱，之前没有弹出要求先校对)
                blnSend = True
                If Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then   '新开医嘱
                    If lng组ID = lngOld组ID Then
                        blnSend = blnOldSend
                    Else
                        If Val(.Cell(flexcpData, i, COL_医嘱期效)) = 0 And Val(.TextMatrix(i, COL_次数)) = 0 Or _
                            .TextMatrix(i, COL_诊疗类别) = "" And Val("" & .TextMatrix(i, COL_诊疗项目ID)) = 0 Then
                            blnSend = False
                        ElseIf Val(.Cell(flexcpData, i, COL_医嘱期效)) = 0 Then
                            If .TextMatrix(i, COL_诊疗类别) = "H" And .TextMatrix(i, COL_操作类型) = "1" And .TextMatrix(i, COL_执行频率) = "2" Then
                                blnSend = False
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",9,10,12,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0 Then
                                blnSend = False
                            End If
                        End If
                        If Not blnSendPrivs And blnSend Then
                            If Not CheckSendPrivs(Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_主页ID)), Val(.TextMatrix(i, COL_会诊医嘱ID))) Then
                                blnSend = False
                            End If
                        End If
                    End If
                End If
                blnOldSend = blnSend
                
                '提交当前病人的数据
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_病人ID)) <> lng病人ID Then
                    '提交当前病人数据
                    If lng病人ID <> 0 Then
                        If strAudit <> "" Then
                            MsgBox "病人""" & rsPati!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, gstrSysName
                            GoTo errH
                        End If
                                    
                        
                         '医嘱执行计价
                        If rsExec.RecordCount > 0 Then
                            rsExec.MoveFirst
                            Do While Not rsExec.EOF
                                rsSQL.AddNew
                                rsSQL!类型 = 8
                                rsSQL!项目ID = 0
                                rsSQL!序号 = 0
                                rsSQL!医嘱ID = lng组ID
                                rsSQL!Sql = "Zl_医嘱执行计价_Insert(" & rsExec!医嘱ID & "," & rsExec!发送号 & ",To_date('" & _
                                rsExec!要求时间 & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!收费细目ID & "")) & "," & rsExec!数量 & "," & rsExec!费用性质 & ")"
                                rsExec.MoveNext
                            Loop
                        End If
                    
                        If Not CompletePatiSend(rsPati, rsSQL, rsUpload, rsItems, cur合计, cur记帐合计, str类别, bln划价, blnTran, lng发送号) Then GoTo errH
                        SendAdvice = lng发送号 '只要提交成功则标注
                        Call InitExecRecordset(rsExec)   '医嘱执行计价
                    End If
                    
                    '重置病人相关变量
                    str自动发料 = ""
                    lng病人ID = Val(.TextMatrix(i, COL_病人ID))
                    lng主页ID = Val(.TextMatrix(i, COL_主页ID))
                    lng病人性质 = Val(.TextMatrix(i, COL_病人性质))
                    Set rs皮试 = Nothing
                    Call InitRecordSet(rsSQL, rsTotal, rsUpload, rsNumber, rsMoneyNow, rsItems, rsSendDays) '重置SQL数组
                    cur合计 = 0:  str类别 = "":  cur记帐合计 = 0 '重置报警变量
                    Set rsMoneyDay = Nothing
                    
                    '获取当前病人信息
                    strSQL = _
                        " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 类型 = 2 And 病人ID=[1]" & _
                        " Union ALL" & _
                        " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
                    strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
                    
                    '状态:0-正常；1-尚未入科；2-正在转科；3-已预出院
                    strSQL = "Select A.病人ID,B.主页ID,NVL(B.姓名,A.姓名) 姓名,B.险类,B.状态,Nvl(B.性别,A.性别) as 性别,Nvl(B.年龄,A.年龄) as 年龄,B.出院病床 as 床号,a.门诊号," & _
                        " zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人,C.剩余款,a.家庭电话 as PhoneNO,a.家庭地址 as Address,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI:SS') as Birthdate," & _
                        " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额,b.当前病区ID as 病区ID,b.出院科室ID as 科室ID,B.住院号" & _
                        " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
                        " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+) And A.病人ID=[1] And B.主页ID=[2]"
                    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
                    
                    If blnSend Then
                        '提取当前病人的审批项目清单
                        strAudit = ""
                        If Not IsNull(rsPati!险类) Then
                            If Val(zlDatabase.GetPara("检查医保审批", glngSys, p住院医嘱发送, "1")) = 1 Then
                                Set rsAudit = GetAuditRecord(lng病人ID, lng主页ID)
                            Else
                                Set rsAudit = Nothing
                            End If
                            bln实时监控 = gclsInsure.GetCapability(support实时监控, rsPati!病人ID, rsPati!险类)
                        Else
                            Set rsAudit = Nothing '以Nothing为标志该病人不需要判断
                            bln实时监控 = False
                        End If
                        
                        '检查更新并检查当前病人医嘱的药品库存,自备药不检查
                        '虽然提取时已汇总检查，但按品种下时如果改了规格可能发生变化
                        For j = i To .Rows - 1
                            If Val(.TextMatrix(j, COL_病人ID)) = lng病人ID Then
                                '可能根据前面库存检查提示的结果现在已不可用
                                If .Cell(flexcpData, j, COL_选择) = 0 And Not .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                                    If InStr(",5,6,7,", .TextMatrix(j, COL_诊疗类别)) > 0 And Val(.TextMatrix(j, COL_执行性质ID)) <> 5 Then
                                        '在不足禁止的情况下,包括分批或时价药品
                                        If TheStockCheck(Val(.TextMatrix(j, COL_执行科室ID)), .TextMatrix(j, COL_诊疗类别)) = 2 _
                                            Or Val(.TextMatrix(j, COL_药房分批)) = 1 Or Val(.TextMatrix(j, COL_是否变价)) = 1 Then
                                            .TextMatrix(j, COL_库存) = Format(GetStock(Val(.TextMatrix(j, COL_收费细目ID)), Val(.TextMatrix(j, COL_执行科室ID)), 2), "0.00000")
                                            If CheckStock(j, bln药品库存提示, bln药品时价提示, bln药品默认发送, True) Then
                                                Call RowSelectSame(j, COL_选择)
                                            End If
                                        End If
                                        If CheckDrug零差价(j, bln药品零差价提示) Then
                                            Call RowSelectSame(j, COL_选择)
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                                    
                '可能根据前面库存检查提示的结果现在已不可用
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                                         
                    '特殊医嘱：3-转科;5-出院;6-转院,11-死亡
                    If .TextMatrix(i, COL_诊疗类别) = "Z" Then
                        '转科,出院,转院,死亡医嘱发送时，病人要处于正常状态
                        If blnSend And .Cell(flexcpData, i, COL_婴儿) = 0 Then
                            If InStr(",3,5,6,11,", .TextMatrix(i, COL_操作类型)) > 0 And NVL(rsPati!状态, 0) <> 0 Then
                                MsgBox "病人""" & rsPati!姓名 & """当前处于""" & decode(NVL(rsPati!状态, 0), 1, "等待入科", 2, "正在转科", 3, "已预出院") & """状态，" & _
                                    "不能发送""" & .TextMatrix(i, col_医嘱内容) & """医嘱。", vbInformation, gstrSysName
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                GoTo NextAdvice
                            End If
                        End If
                        
                        '如果是转科、出院、转院医嘱,检查病人是否有未执行的医技项目及未发药品
                        If InStr(",3,", .TextMatrix(i, COL_操作类型)) > 0 Then
                            If Not CheckWaitExecute(rsPati, i, gbyt转科检查未执行, gbyt转科检查未发药) Then
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                GoTo NextAdvice
                            End If
                        End If
                        If InStr(",5,6,", .TextMatrix(i, COL_操作类型)) > 0 Then
                            If Not CheckWaitExecute(rsPati, i, gbyt出院检查未执行, gbyt出院检查未发药) Then
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                GoTo NextAdvice
                            End If
                        End If
                        
                        '因为自动停止医嘱，需要进行超期收回提醒
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_操作类型)) > 0 Then
                            If InStr(mstrRollNotify, lng病人ID & ":" & lng主页ID) = 0 Then
                                mstrRollNotify = mstrRollNotify & "," & lng病人ID & ":" & lng主页ID
                            End If
                        End If
                    End If
                    
                    '更改医嘱的执行科室
                    If .Cell(flexcpData, i, COL_执行科室ID) = 1 Then
                        rsSQL.AddNew
                        rsSQL!类型 = 4
                        rsSQL!医嘱ID = lng组ID
                        rsSQL!项目ID = 0
                        rsSQL!序号 = i
                        rsSQL!Sql = "ZL_医嘱执行科室_Update(" & Val(.TextMatrix(i, COL_ID)) & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & ",1)"
                        rsSQL.Update
                    End If
                    
                    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng病人ID, lng主页ID, "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    
                    '产生医嘱记帐费用:以最新价格计算
                    '-----------------------------------------------------------------------------------------
                    strSQL = "": str收费项目 = ""
                    If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
                        If Val(.TextMatrix(i, COL_执行性质ID)) <> 5 Then
                            strSQL = _
                                " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,B.收入项目ID," & _
                                " C.收据费目,Y.住院单位,Y.住院包装,Y.剂量系数,1 as 数量,B.现价 as 单价," & _
                                " A.加班加价,B.加班加价率,B.附术收费率,A.是否变价,Y.药房分批 as 分批,0 as 跟踪在用," & _
                                " 0 as 从项,[2] as 执行科室ID,A.屏蔽费别,A.费用确认,0 as 费用性质,0 as 收费方式,I.要求审批" & _
                                " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,药品规格 Y,保险支付项目 I" & _
                                " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码" & _
                                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                                " And A.ID=Y.药品ID(+) And A.ID=[1] And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                                " Order by A.编码"
                        End If
                    Else
                        '不计价,手工计价；叮嘱,院外执行的医嘱不读取
                        If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                            '先删除原非药医嘱的计价(发送时才校对的模式，没有必要先删除，因为之前没有产生计价)
                            If Val(.Cell(flexcpData, i, COL_金额)) = 1 And Val(.TextMatrix(i, COL_医嘱状态)) <> 1 Then
                                rsSQL.AddNew
                                rsSQL!类型 = 2: rsSQL!项目ID = 0: rsSQL!序号 = i
                                rsSQL!医嘱ID = lng组ID
                                rsSQL!Sql = "ZL_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                                rsSQL.Update
                            End If
                        
                            mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not mrsPrice.EOF Then
                                For j = 1 To mrsPrice.RecordCount
                                    If NVL(mrsPrice!收费细目ID, 0) <> 0 And NVL(mrsPrice!数量, 0) <> 0 Then '对照数量为0的自动过滤掉
                                        '普通项目的变价单价要求输入，包括非跟踪在用的时价卫材医嘱
                                        If NVL(mrsPrice!单价, 0) = 0 And NVL(mrsPrice!变价, 0) = 1 _
                                            And Not (InStr(",5,6,7,", mrsPrice!收费类别) > 0 Or mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1) Then
                                            Call SeekPriceRow(i, mrsPrice!收费细目ID, mrsPrice!费用性质, COLP_单价)
                                            Screen.MousePointer = 0
                                            MsgBox "必须为变价的收费项目确定一个收费价格。", vbInformation, gstrSysName
                                            vsPrice.SetFocus: GoTo FuncEnd
                                        End If
                                        
                                        '计价执行科室:只保存非药品及卫材医嘱的，药品和卫材计价的执行科室
                                        If InStr(",4,5,6,7,", .TextMatrix(i, COL_诊疗类别)) = 0 _
                                            And (InStr(",5,6,7,", mrsPrice!收费类别) > 0 Or mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1) Then
                                            lng执行科室ID = NVL(mrsPrice!执行科室ID, 0)
                                            
                                            '卫材必须设置执行科室
                                            If lng执行科室ID = 0 And mrsPrice!收费类别 = "4" Then
                                                Call SeekPriceRow(i, mrsPrice!收费细目ID, mrsPrice!费用性质, COLP_执行科室)
                                                Screen.MousePointer = 0
                                                MsgBox "卫材""" & vsPrice.TextMatrix(vsPrice.Row, COLP_收费项目) & """没有确定执行科室，请手工输入正确的执行科室。" & vbCrLf & _
                                                    "如果不能确定正确的执行科室，请到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                                vsPrice.SetFocus: GoTo FuncEnd
                                            End If
                                        Else
                                            lng执行科室ID = 0
                                        End If
                                        
                                        '药品、卫材医嘱的计价固定对应不保存；非跟踪在用的时价卫材的变价需要输入，因此要保存到计价表中
                                        If Val(.Cell(flexcpData, i, COL_金额)) = 1 Or Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                                            If InStr(",4,5,6,7,", .TextMatrix(i, COL_诊疗类别)) = 0 _
                                                Or .TextMatrix(i, COL_诊疗类别) = "4" And NVL(mrsPrice!在用, 0) = 0 And NVL(mrsPrice!变价, 0) = 1 Then
                                                rsSQL.AddNew
                                                rsSQL!类型 = 2: rsSQL!项目ID = mrsPrice!收费细目ID: rsSQL!序号 = i
                                                rsSQL!医嘱ID = lng组ID
                                                rsSQL!Sql = "ZL_病人医嘱计价_INSERT(" & _
                                                    mrsPrice!医嘱ID & "," & mrsPrice!收费细目ID & "," & _
                                                    NVL(mrsPrice!数量, 0) & "," & NVL(mrsPrice!单价, 0) & "," & _
                                                    NVL(mrsPrice!从项, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                                    NVL(mrsPrice!费用性质, 0) & "," & NVL(mrsPrice!收费方式, 0) & ")"
                                                rsSQL.Update
                                            End If
                                        End If
                                        
                                        '临时病人医嘱计价表
                                        If Val(.TextMatrix(i, COL_总量)) <> 0 Then '输血可能没有总量
                                            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                                "Select " & mrsPrice!收费细目ID & " as 收费细目ID," & _
                                                NVL(mrsPrice!执行科室ID, 0) & " as 执行科室ID," & _
                                                NVL(mrsPrice!数量, 0) & " as 数量," & Format(NVL(mrsPrice!单价, 0), gstrDecPrice) & " as 单价," & _
                                                NVL(mrsPrice!从项, 0) & " as 从项," & NVL(mrsPrice!费用性质, 0) & " as 费用性质," & _
                                                NVL(mrsPrice!收费方式, 0) & " as 收费方式 From Dual"
                                        End If
                                    End If
                                    
                                    mrsPrice.MoveNext
                                Next
                            End If
                        End If
                        
                        If strSQL <> "" Then
                            strSQL = _
                                " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,A.是否变价," & _
                                " A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,B.附术收费率,Y.住院单位,Y.住院包装,Y.剂量系数," & _
                                " Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批,E.跟踪在用,B.收入项目ID," & _
                                " C.收据费目,X.数量,Decode(A.是否变价,1,X.单价,B.现价) as 单价,X.执行科室ID," & _
                                " X.从项,X.费用性质,X.收费方式,I.要求审批" & _
                                " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,(" & strSQL & ") X,药品规格 Y,保险支付项目 I" & _
                                " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.ID=E.材料ID(+)" & _
                                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                                " And A.类别=D.编码 And X.收费细目ID=A.ID And A.ID=Y.药品ID(+)" & _
                                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                                " And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                                " Order by X.费用性质,X.从项,X.收费方式 Desc,A.ID"
                                '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                        End If
                    End If
                    
                    '医嘱校对,发送前自动校对(一组医嘱调用一次，所有叮嘱都要校对)
                    If mblnAutoVerify Then
                        If Val(.TextMatrix(i, COL_医嘱状态)) = 1 And lng组ID <> lngOld组ID Then
                            rsSQL.AddNew
                            rsSQL!类型 = 1
                            rsSQL!医嘱ID = lng组ID
                            rsSQL!项目ID = 0
                            rsSQL!序号 = i
                            rsSQL!Sql = "ZL_病人医嘱记录_校对(" & lng组ID & ",3," & strCurDate & ",Null,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                        End If
                    End If
                    
                    
                    
                    '执行发送和记帐费用
                    '-----------------------------------------
                    If blnSend Then
                        '汇总折扣变量初始
                        strHaveSub = "": strNoneSub = ""
                        int父序号 = 0: lng父项目ID = 0
                        Call InitSeekSet(rsSeek)
                        
                        '提前生成样本条码(参数"医嘱发送生成条形码"没有启用时也产生一个虚拟的条码，用于判断是否收采血管费用)
                        strCuvetteNumber = ""
                        If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then
                            j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                            If j > 0 Then lng采集科室ID = Val(.TextMatrix(j, COL_执行科室ID))
                            strCuvetteNumber = GetCuvetteNumber(rsNumber, .TextMatrix(i, COL_试管编码), _
                                Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)), .TextMatrix(i, COL_诊疗类别), Val(.TextMatrix(i, COL_操作类型)), _
                                Val(.TextMatrix(i, COL_执行科室ID)), Val(.TextMatrix(i, COL_婴儿)), Val(.TextMatrix(i, COL_诊疗项目ID)), _
                                Val(.TextMatrix(i, COL_紧急标志)), .TextMatrix(i, COL_标本部位), lng采集科室ID)
                        End If
                        
                        int计费状态 = IIF(Val(.TextMatrix(i, COL_计价特性)) = 1, -1, 0) '无需计费或未计费
                    
                
                        '产生单据号分配关键字
                        '-----------------------------------------------------------------------------------------
                        If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            '启用参数：特殊药品分开发送 时，特殊药品医嘱的药品行单独生成单据号，一组医嘱分配一个号
                            If str毒理 <> "" Then
                                strNOKey = "中西成药_" & .TextMatrix(i, COL_相关ID)
                            Else
                                '中西成药按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                                '一并给药的，发送到一起：包括自备药和不同药房的情况
                                strNOKey = "中西成药_" & lng病人ID & "_" & lng主页ID & "_" & .TextMatrix(i, COL_医嘱期效) & "_" & _
                                    Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                                    .TextMatrix(i, COL_开嘱医生) & "_" & GetMergeDrugStore(i)
                                '再按要打印的诊疗单据分号
                                strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), 2)
                                '给药执行科室不相同，则分配不同的NO号
                                j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                                If j > 0 Then strNOKey = strNOKey & "_" & Val(.TextMatrix(j, COL_执行科室ID))
                            End If
                        ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                            '一个配方中的所有草药分配一个独立单据号
                            strNOKey = "中药配方_" & lng病人ID & "_" & lng主页ID & "_" & .TextMatrix(i, COL_医嘱期效) & "_" & int配方数
                        ElseIf InStr(",4,M,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            '材料按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                            strNOKey = "材料医嘱_" & lng病人ID & "_" & lng主页ID & "_" & .TextMatrix(i, COL_医嘱期效) & "_" & _
                                Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                                .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                            '再按要打印的诊疗单据分号
                            strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), 2)
                            
                        ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "C" Then
                            '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
                            '同一个类检验型，同一个检验执行科室，同一采集管，同一个采集方式，同一个采集执行科室的检验分配相同的单据号
                            If mbln检验单独产生单据 Then
                                strNOKey = "一并采集_" & Val(.TextMatrix(i, COL_相关ID))
                            Else
                                lng主医嘱行 = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                                strNOKey = "一并采集_" & lng病人ID & "_" & lng主页ID & "_" & .TextMatrix(i, COL_标本部位) & "_" & _
                                    .TextMatrix(i, COL_执行科室ID) & "_" & .TextMatrix(i, COL_操作类型) & "_" & .TextMatrix(i, COL_试管编码) & "_" & _
                                    .TextMatrix(lng主医嘱行, COL_诊疗项目ID) & "_" & .TextMatrix(lng主医嘱行, COL_执行科室ID)
                            End If
                        ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
                            strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_相关ID))
                        Else
                            '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法，采集方式，麻醉方式，输血医嘱/输血途径)
                            strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_ID))
                        End If
                        
                                
                         '分解时间
                        If .TextMatrix(i, COL_分解时间) <> "" Then
                            str分解时间 = .TextMatrix(i, COL_分解时间)
                        Else
                            str分解时间 = .Cell(flexcpData, i, COL_分解时间)    '开始执行时间
                        End If
                        If str分解时间 = "" Then str分解时间 = .TextMatrix(i, COL_开始时间)
                        If Len(str分解时间) > 4000 Then
                            Screen.MousePointer = 0
                            MsgBox "当前发送的医嘱时间范围太长,共需执行" & CStr(UBound(Split(str分解时间, ",")) + 1) & "次。" & vbCrLf & _
                                "超过了支持的最大次数" & CStr(UBound(Split(Mid(str分解时间, 1, 4000), ",")) + 1) & "次,请调整结束时间后重新发送！", vbInformation, gstrSysName
                            Call DeleteSendRow: Call ShowSendTotal
                            Progress = 0: Exit Function
                        End If
                        
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "D" Then
                            str部位方法 = .TextMatrix(i, COL_标本部位) & "<sTab>" & .TextMatrix(i, COL_检查方法)
                        Else
                            str部位方法 = ""
                        End If
                        
                        '产生记帐费用
                        '------------------------------------------------------
                        If strSQL <> "" Then
                            '是否离院带药
                            int药品性质 = 0
                            If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                int药品性质 = decode(.TextMatrix(i, COL_执行性质), "离院带药", 3, "自取药", 4, int药品性质)
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                                j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                                If j <> -1 Then
                                    int药品性质 = decode(.TextMatrix(j, COL_执行性质), "离院带药", 3, "自取药", 4, int药品性质)
                                End If
                            End If
                        
                            Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_收费细目ID)), Val(.TextMatrix(i, COL_执行科室ID)), Val(NVL(rsPati!险类, 0)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            If Not rsPrice.EOF Then
                                int计费状态 = 1 '已计费
                                Set rsClone = rsPrice.Clone
                            End If
    
                            '处理收入项目级的费用明细
                            bln附加手术 = .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0
                            Do While Not rsPrice.EOF
MoneyItemBegin:
                                '执行科室ID
                                lng执行科室ID = NVL(rsPrice!执行科室ID, 0)
                                '在原值基础上取有效的非药嘱药品及跟踪卫材的执行科室
                                If rsPrice!类别 = "4" And NVL(rsPrice!跟踪在用, 0) = 1 _
                                    Or InStr(",5,6,7", rsPrice!类别) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                                    lng病人科室ID = Val(.TextMatrix(i, COL_病人科室ID))
                                    lng执行科室ID = Get收费执行科室ID(rsPati!病人ID, rsPati!主页ID, rsPrice!类别, rsPrice!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
                                    
                                    '卫材必须设置执行科室
                                    If lng执行科室ID = 0 And rsPrice!类别 = "4" Then
                                        .Row = GetVisibleRow(i, True)
                                        Call .ShowCell(.Row, .Col)
                                        Screen.MousePointer = 0
                                        MsgBox "系统不能为计价卫材""" & rsPrice!名称 & """确定合适的执行科室。" & vbCrLf & _
                                            "请使用计价调整功能人为确定，或到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                        Call DeleteSendRow: Call ShowSendTotal
                                        Progress = 0: Exit Function
                                    End If
                                End If
                                
                                '----------------------------------------
                                '根据收费方式，确定当前收费项目是否应收费
                                If rsPrice!费用性质 & "_" & rsPrice!ID <> str收费项目 Then
                                    If Not AdviceMoneyMake(lng病人ID, lng主页ID, rsMoneyNow, rsMoneyDay, _
                                        lng组ID, Val(.TextMatrix(i, COL_诊疗项目ID)), rsPrice!ID, lng执行科室ID, .TextMatrix(i, COL_试管编码), _
                                        rsPrice!类别, NVL(rsPrice!收费方式, 0), str分解时间, 2, lng费用次数, Val(.TextMatrix(i, COL_总量)), _
                                        Val(.TextMatrix(i, COL_ID)), lng发送号, Val(rsPrice!数量 & ""), rsExec, Val(.TextMatrix(i, COL_计算方式)), _
                                        .TextMatrix(i, COL_频率), Val(.TextMatrix(i, COL_单量)), .Cell(flexcpData, i, COL_医嘱期效), rsPrice!费用性质, .TextMatrix(i, COL_诊疗类别), strCuvetteNumber, str部位方法, dblOther数量, strMinDate) Then
                                        '跳过当前收费项目(多个收入项目)
                                        str收费项目 = rsPrice!费用性质 & "_" & rsPrice!ID
                                        Do While rsPrice!费用性质 & "_" & rsPrice!ID = str收费项目
                                            rsPrice.MoveNext
                                            If rsPrice.EOF Then Exit Do
                                        Loop
                                        If rsPrice.EOF Then Exit Do
                                        GoTo MoneyItemBegin
                                    End If
                                End If
                                '----------------------------------------
                                
                                '检查是否需要和已经审批
                                If NVL(rsPrice!要求审批, 0) = 1 And Not rsAudit Is Nothing Then
                                    rsAudit.Filter = "项目ID=" & rsPrice!ID
                                    If rsAudit.EOF Then
                                        If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                            If InStr(strAudit, "●" & rsPrice!名称) = 0 Then
                                                strAudit = strAudit & vbCrLf & "●" & rsPrice!名称
                                            End If
                                        ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                            strAudit = strAudit & vbCrLf & "… …"
                                        End If
                                    End If
                                End If
                                
                                If InStr(",5,6,7", rsPrice!类别) > 0 Then
                                    If InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                        If .TextMatrix(i, COL_诊疗类别) = "7" Then
                                            int付数 = Val(.TextMatrix(i, COL_总量))
                                            '中药药房单位按不可分零处理:每付
                                            If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                                dbl数量 = Val(.TextMatrix(i, COL_单量)) / NVL(rsPrice!剂量系数, 1)
                                            Else
                                                dbl数量 = IntEx(Val(.TextMatrix(i, COL_单量)) / NVL(rsPrice!剂量系数, 1) / NVL(rsPrice!住院包装, 1)) * NVL(rsPrice!住院包装, 1)
                                            End If
                                        Else
                                            int付数 = 1
                                            dbl数量 = Val(.TextMatrix(i, COL_总量)) * NVL(rsPrice!住院包装, 1)
                                            If rs皮试 Is Nothing Then
                                                Set rs皮试 = Get原液皮试(lng病人ID, lng主页ID, "")
                                            End If
                                            rs皮试.Filter = "药品ID=" & Val(rsPrice!ID & "")
                                            If Not rs皮试.EOF Then
                                                If Val(rs皮试!标号 & "") = 0 Then
                                                    '进行减总量计算
                                                    dbl数量 = (Val(.TextMatrix(i, COL_总量)) - 1) * NVL(rsPrice!住院包装, 1)
                                                    rs皮试!标号 = Val(.TextMatrix(i, COL_ID))
                                                    
                                                    str关联药行 = "'" & rs皮试!皮试医嘱ID & "," & rs皮试!标号 & "'"
                                                    rs皮试.Update
                                                    If dbl数量 <= 0 Then
                                                        rsPrice.MoveNext
                                                        If rsPrice.EOF Then Exit Do
                                                        GoTo MoneyItemBegin
                                                    End If
                                                End If
                                            End If
                                            
                                        End If
                                    Else
                                        int付数 = 1
                                        '中药药房单位按不可分零处理:每付
                                        '非药嘱药品计价:因为这里预定了售价数量,因此不作不分零处理
                                        '对于收费对照中的药品，且为当天只收取一次，数量为费用次数*对照数量
                                        If InStr(",2,3,4,5,6,7,9,", Val("" & rsPrice!收费方式)) > 0 Then
                                            If dblOther数量 > 0 Then
                                                dbl数量 = Format(dblOther数量, "0.00000")
                                            Else
                                                dbl数量 = Format(lng费用次数 * NVL(rsPrice!数量, 0), "0.00000")
                                            End If
                                        Else
                                            dbl数量 = Val(.TextMatrix(i, COL_总量)) * NVL(rsPrice!数量, 0)
                                        End If
                                    End If
                                    dbl数量 = Format(dbl数量, "0.00000")
                                    
                                    If NVL(rsPrice!是否变价, 0) = 1 Then
                                        dbl单价 = Format(CalcDrugPrice(rsPrice!ID, lng执行科室ID, int付数 * dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                                    Else
                                        dbl单价 = Format(NVL(rsPrice!单价, 0), gstrDecPrice)
                                    End If
                                ElseIf rsPrice!类别 = "4" And NVL(rsPrice!跟踪在用, 0) = 1 Then
                                    '检查卫生材料入出类别
                                    If mlng卫材类别ID = 0 Then
                                        Screen.MousePointer = 0
                                        MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                        Call DeleteSendRow: Call ShowSendTotal
                                        Progress = 0: Exit Function
                                    End If
                                    
                                    int付数 = 1
                                    If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsPrice!收费方式)) > 0 Then
                                        If dblOther数量 > 0 Then
                                            dbl数量 = Format(dblOther数量, "0.00000")
                                        Else
                                            dbl数量 = Format(lng费用次数 * NVL(rsPrice!数量, 0), "0.00000")
                                        End If
                                    Else
                                        dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * NVL(rsPrice!数量, 0), "0.00000")
                                    End If
                                    
                                    '计算时价卫材单价
                                    If NVL(rsPrice!是否变价, 0) = 1 Then
                                        dbl单价 = Format(CalcDrugPrice(rsPrice!ID, lng执行科室ID, dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                                    Else
                                        dbl单价 = Format(NVL(rsPrice!单价, 0), gstrDecPrice)
                                    End If
                                Else
                                    '总量等于单次用量乘数次。一天只收一次时，有多少天要执行，就收多少次，不管单次用量（例如：每天两次）,但要管收费对照的次数
                                    int付数 = 1
                                    If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsPrice!收费方式)) > 0 Then
                                        If dblOther数量 > 0 Then
                                            dbl数量 = Format(dblOther数量, "0.00000")
                                        Else
                                            dbl数量 = Format(lng费用次数 * NVL(rsPrice!数量, 0), "0.00000")
                                        End If
                                    Else
                                        dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * NVL(rsPrice!数量, 0), "0.00000")
                                    End If
                                    dbl单价 = Format(NVL(rsPrice!单价, 0), gstrDecPrice)
                                End If
                                
                                '非药嘱药品及跟踪卫材的库存检查
                                If rsPrice!类别 = "4" And NVL(rsPrice!跟踪在用, 0) = 1 _
                                    Or InStr(",5,6,7", rsPrice!类别) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                                    If TheStockCheck(lng执行科室ID, rsPrice!类别) <> 0 Or NVL(rsPrice!是否变价, 0) = 1 Or NVL(rsPrice!分批, 0) = 1 Then
                                        If rsPrice!类别 = "4" Then
                                            blnBool = CheckPriceStock(i, rsPrice, lng执行科室ID, int付数 * dbl数量, rsTotal, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                                        Else
                                            blnBool = CheckPriceStock(i, rsPrice, lng执行科室ID, int付数 * dbl数量, rsTotal, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                                        End If
                                        If blnBool Then
                                            Call RowSelectSame(i, COL_选择, rsSQL, rsTotal, rsUpload)
                                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                            GoTo NextAdvice
                                        End If
                                    End If
                                End If
                                
                                '发送金额
                                dbl应收 = int付数 * dbl数量 * dbl单价
                                If bln附加手术 Then
                                    dbl应收 = dbl应收 * NVL(rsPrice!附术收费率, 100) / 100
                                End If
                                
                                '处理加班加价
                                If gbln加班加价 And NVL(rsPrice!加班加价, 0) = 1 Then
                                    dbl应收 = dbl应收 * (1 + NVL(rsPrice!加班加价率, 0) / 100)
                                End If
                                
                                cur应收 = Format(dbl应收, gstrDec)
                                                            
                                'NO,序号---------------------------------------------------------------------
                                Call GetCurBillSet(strNOKey, strNO, lng费用序号, -1)
                                rsSQL.AddNew: blnBool = False
                                If rsPrice!费用性质 & "_" & rsPrice!ID <> str收费项目 Then
                                    lng费用父号 = lng费用序号
                                    If rsPrice!从项 = 0 Then
                                        '记录主项信息，主项肯定在从项前
                                        '即使不汇总折扣，也要记录主从项关系
                                        If InStr(strHaveSub & ",", "," & rsPrice!费用性质 & ",") = 0 _
                                            And InStr(strNoneSub & ",", "," & rsPrice!费用性质 & ",") = 0 Then
                                            rsClone.Filter = "费用性质=" & rsPrice!费用性质 & " And 从项=1"
                                            If Not rsClone.EOF Then
                                                int父序号 = lng费用序号
                                                lng父项目ID = rsPrice!ID
                                                
                                                rsSeek.AddNew
                                                rsSeek!费用性质 = rsPrice!费用性质
                                                rsSeek!主项标签 = rsSQL.Bookmark 'Variant(Double)
                                                rsSeek!主收入ID = rsPrice!收入项目ID
                                                rsSeek.Update
                                                strHaveSub = strHaveSub & "," & rsPrice!费用性质
                                                
                                                blnBool = True
                                            Else
                                                strNoneSub = strNoneSub & "," & rsPrice!费用性质
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '计算汇总折扣合计
                                If gbln从项汇总折扣 And (rsPrice!从项 = 1 Or InStr(strHaveSub & ",", "," & rsPrice!费用性质 & ",") > 0) Then
                                    cur实收 = cur应收
                                    
                                    '累计医嘱合计来计算折扣
                                    rsSeek.Filter = "费用性质=" & rsPrice!费用性质
                                    rsSeek!合计 = NVL(rsSeek!合计, 0) + cur实收
                                    rsSeek.Update
                                ElseIf NVL(rsPrice!屏蔽费别, 0) = 0 Then
                                    cur实收 = Format(ActualMoney(.TextMatrix(i, COL_费别), rsPrice!收入项目ID, cur应收, rsPrice!ID, lng执行科室ID, _
                                        int付数 * dbl数量, IIF(gbln加班加价 And NVL(rsPrice!加班加价, 0) = 1, NVL(rsPrice!加班加价率, 0) / 100, 0)), gstrDec)
                                Else
                                    cur实收 = cur应收
                                End If
                                If gbln从项汇总折扣 And blnBool Then
                                    '汇总折扣时，对主项的实收金额作特殊处理
                                    str实收 = Chr(0) & Chr(1) & "Begin" & cur实收 & "End" & Chr(0) & Chr(1)
                                Else
                                    str实收 = cur实收
                                End If
                                '----------------------------------------------------------------------------
                                
                                '医保相关字段
                                bln保险项目否 = False: lng保险大类ID = 0: cur统筹金额 = 0: str保险编码 = "": str费用类型 = ""
                                If Not IsNull(rsPati!险类) Then
                                    strTmp = gclsInsure.GetItemInsure(lng病人ID, rsPrice!ID, cur实收, False, rsPati!险类, .Cell(flexcpData, i, COL_医生嘱托) & "||" & int付数 * dbl数量)
                                    If strTmp <> "" Then
                                        bln保险项目否 = Val(Split(strTmp, ";")(0)) <> 0
                                        lng保险大类ID = Val(Split(strTmp, ";")(1))
                                        cur统筹金额 = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                        str保险编码 = CStr(Split(strTmp, ";")(3))
                                        If UBound(Split(strTmp, ";")) >= 5 Then
                                            If Split(strTmp, ";")(5) <> "" Then
                                                str费用类型 = Split(strTmp, ";")(5)
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '收集记帐报警类别
                                cur合计 = cur合计 + cur实收
                                If InStr(str类别, rsPrice!类别) = 0 Then
                                    str类别 = str类别 & rsPrice!类别
                                End If
                                                            
                                '是否划价
                                strTmp = mlng病区ID
                                If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                    int划价 = IIF(InStr(gstr住院发送划价单, "5") > 0, 1, 0)
                                    '问题号72503
                                    If mint住院领药部门 = 0 Then
                                        j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                                        If Val(.TextMatrix(j, COL_执行科室ID)) <> 0 Then strTmp = Val(.TextMatrix(j, COL_执行科室ID))
                                    End If
                                Else
                                    int划价 = IIF(InStr(gstr住院发送划价单, .TextMatrix(i, COL_诊疗类别)) > 0, 1, 0)
                                End If
                                If int划价 = 0 Then int划价 = IIF(NVL(rsPrice!费用确认, 0) = 1, 1, 0)
                                
                                If int划价 = 0 Or int执行状态 = 1 Then
                                    bln划价 = False
                                    cur记帐合计 = cur记帐合计 + cur实收
                                End If
                                '发生时间
                                If int药品性质 = 3 Then
                                    str发生时间 = strCurDate
                                ElseIf .TextMatrix(i, COL_分解时间) <> "" Then
                                    str发生时间 = "To_Date('" & Split(.TextMatrix(i, COL_分解时间), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                                Else
                                    str发生时间 = "To_Date('" & .Cell(flexcpData, i, COL_分解时间) & "','YYYY-MM-DD HH24:MI:SS')"
                                End If
                                
                                '登记时间
                                If int划价 = 1 Then '与非划价的时间上区分开
                                    str登记时间 = "To_Date('" & Format(DateAdd("s", 1, mdatCurr), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                Else
                                    str登记时间 = strCurDate
                                End If
                                
                                '长嘱非药按天拆分费用
                                Set rsDays = New ADODB.Recordset
                                rsDays.Fields.Append "收费时间", adVarChar, 10
                                rsDays.Fields.Append "数量", adDouble
                                rsDays.Fields.Append "实收金额", adDouble
                                rsDays.Fields.Append "应收金额", adDouble
                                rsDays.Fields.Append "发生时间", adVarChar, 100
                                rsDays.Fields.Append "NO", adVarChar, 30
                                rsDays.CursorLocation = adUseClient
                                rsDays.LockType = adLockOptimistic
                                rsDays.CursorType = adOpenStatic
                                rsDays.Open
                                '长嘱膳食、输氧、治疗、特殊治疗可拆分
                                If .TextMatrix(i, COL_医嘱期效) = "长嘱" And (InStr(",I,L,", .TextMatrix(i, COL_诊疗类别)) > 0 Or .TextMatrix(i, COL_诊疗类别) = "E" And InStr(",0,5,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0) Then
                                    rsExec.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID)) & " And 发送号=" & lng发送号 & " And 收费细目ID=" & rsPrice!ID
                                    If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                                    Do While Not rsExec.EOF
                                        If Val(rsExec!数量 & "") <> 0 Then
                                            rsDays.Filter = "收费时间='" & Format(rsExec!要求时间 & "", "YYYY-MM-DD") & "'"
                                            If rsDays.RecordCount = 0 Then
                                                rsSendDays.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID)) & " And 计算时间='" & Format(rsExec!要求时间 & "", "yyyy-MM-dd") & "'"
                                                If rsSendDays.RecordCount > 0 Then rsSendDays.MoveFirst
                                                rsDayNo.Filter = 0
                                                If rsDayNo.RecordCount = 0 Then
                                                    lng发送号Tmp = lng发送号
                                                    rsDayNo.AddNew
                                                    rsDayNo!发送号 = lng发送号Tmp
                                                    rsDayNo!计算时间 = Format(rsExec!要求时间 & "", "yyyy-MM-dd")
                                                    rsDayNo.Update
                                                Else
                                                    rsDayNo.Filter = "计算时间='" & Format(rsExec!要求时间 & "", "yyyy-MM-dd") & "'"
                                                    If rsDayNo.RecordCount = 0 Then
                                                        lng发送号Tmp = zlDatabase.GetNextNo(10)
                                                        rsDayNo.AddNew
                                                        rsDayNo!发送号 = lng发送号Tmp
                                                        rsDayNo!计算时间 = Format(rsExec!要求时间 & "", "yyyy-MM-dd")
                                                        rsDayNo.Update
                                                    Else
                                                        lng发送号Tmp = rsDayNo!发送号
                                                    End If
                                                End If
                                                Call GetCurBillSet(strNOKey & "_" & Format(rsExec!要求时间 & "", "yyyy-MM-dd"), strNO, -1, -1)
                                                rsDays.AddNew
                                                rsDays!收费时间 = Format(rsExec!要求时间 & "", "YYYY-MM-DD")
                                                rsDays!数量 = rsExec!数量
                                                rsDays!实收金额 = Format(rsExec!数量 / dbl数量 * str实收, gstrDec)
                                                rsDays!应收金额 = Format(rsExec!数量 / dbl数量 * cur应收, gstrDec)
                                                rsDays!发生时间 = "To_Date('" & Format(rsExec!要求时间 & "", "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                                rsDays!NO = strNO
                                                If rsSendDays.RecordCount = 0 Then
                                                    rsSendDays.AddNew
                                                    rsSendDays!医嘱ID = Val(.TextMatrix(i, COL_ID))
                                                    rsSendDays!发送号 = lng发送号Tmp
                                                    rsSendDays!计算时间 = Format(rsExec!要求时间 & "", "yyyy-MM-dd")
                                                    rsSendDays!分解时间 = Get当天分解时间(str分解时间, CDate(rsExec!要求时间))
                                                    If rsSendDays!分解时间 <> "" Then
                                                        rsSendDays!首次时间 = "To_Date('" & Split(rsSendDays!分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                                                        rsSendDays!末次时间 = "To_Date('" & Split(rsSendDays!分解时间, ",")(UBound(Split(rsSendDays!分解时间, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                                                    End If
                                                    rsSendDays!NO = strNO
                                                    rsExec!发送号 = lng发送号Tmp
                                                Else
                                                    rsExec!发送号 = lng发送号Tmp
                                                End If
                                            Else
                                                '替换发送号
                                                rsDayNo.Filter = "计算时间='" & Format(rsExec!要求时间 & "", "yyyy-MM-dd") & "'"
                                                If Not rsDayNo.EOF Then lng发送号Tmp = rsDayNo!发送号
                                                rsDays!数量 = rsDays!数量 + rsExec!数量
                                                rsDays!实收金额 = Format(rsDays!实收金额 + rsExec!数量 / dbl数量 * str实收, gstrDec)
                                                rsDays!应收金额 = Format(rsDays!应收金额 + rsExec!数量 / dbl数量 * cur应收, gstrDec)
                                                rsExec!发送号 = lng发送号Tmp
                                            End If
                                            rsDays.Update
                                            rsSendDays.Update
                                            rsExec.Update
                                        End If
                                        rsDayNo.Filter = 0
                                        rsExec.MoveNext
                                    Loop
                                    rsExec.Filter = 0
                                    rsDays.Filter = 0
                                    '调整 发送号,rsSendDays,rsExec，保证日期越大发送号也越大
                                    rsSendDays.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                                    If rsSendDays.RecordCount > 1 Then
                                        Set rsTmp = zlDatabase.CopyNewRec(rsSendDays)
                                        rsTmp.Sort = "发送号"
                                        rsSendDays.Sort = "计算时间"
                                        rsExec.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                                        rsExec.Sort = "要求时间"
                                        Do While Not rsTmp.EOF
                                            Do While True
                                                If rsExec!发送号 = rsSendDays!发送号 Then
                                                    rsExec!发送号 = rsTmp!发送号
                                                    rsExec.MoveNext
                                                Else
                                                    Exit Do
                                                End If
                                                If rsExec.EOF Then
                                                    Exit Do
                                                End If
                                            Loop
                                            rsSendDays!发送号 = rsTmp!发送号
                                            rsSendDays.MoveNext
                                            rsTmp.MoveNext
                                        Loop
                                    End If
                                    rsExec.Sort = ""
                                    rsExec.Filter = 0
                                End If
                                rsDays.Filter = 0
'                               '如果小于等于一天，产生一张单据
                                If rsDays.RecordCount = 1 Then
                                    rsDays!数量 = dbl数量
                                    rsDays!实收金额 = Format(str实收, gstrDec)
                                    rsDays!应收金额 = Format(cur应收, gstrDec)
                                    rsDays!发生时间 = str发生时间
                                    rsDays.Update
                                ElseIf rsDays.RecordCount = 0 Then
                                    rsDays.AddNew
                                    rsDays!数量 = dbl数量
                                    rsDays!实收金额 = Format(str实收, gstrDec)
                                    rsDays!应收金额 = Format(cur应收, gstrDec)
                                    rsDays!发生时间 = str发生时间
                                    rsDays!NO = strNO
                                    rsDays.Update
                                End If
                                
                                '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                                Do While Not rsDays.EOF
                                    '收集医保上传单据号:mrsBill中的不一定产生了费用
                                    If int划价 = 0 Then
                                        rsUpload.Filter = "NO='" & rsDays!NO & "'"
                                        If rsUpload.EOF Then
                                            rsUpload.AddNew
                                            rsUpload!医嘱ID = lng组ID
                                            rsUpload!NO = rsDays!NO
                                            rsUpload.Update
                                        End If
                                    End If
                                    If rsSQL!Sql <> "" Then rsSQL.AddNew
                                    rsSQL!类型 = 3
                                    rsSQL!医嘱ID = lng组ID
                                    rsSQL!项目ID = rsPrice!ID
                                    rsSQL!序号 = i
                                    rsSQL!NO = rsDays!NO
                                    
                                    If lng病人性质 = 1 Then
                                        rsSQL!Sql = "zl_门诊记帐记录_INSERT(" & _
                                            "'" & rsDays!NO & "'," & lng费用序号 & "," & lng病人ID & "," & _
                                            "'" & rsPati!门诊号 & "','" & .TextMatrix(i, COL_姓名) & "'," & _
                                            "'" & .TextMatrix(i, col_性别) & "','" & .TextMatrix(i, COL_年龄) & "'," & "'" & .TextMatrix(i, COL_费别) & "',0," & Val(.Cell(flexcpData, i, COL_婴儿)) & "," & _
                                            ZVal(.TextMatrix(i, COL_病人科室ID)) & "," & ZVal(.TextMatrix(i, COL_开嘱科室ID)) & "," & _
                                            "'" & .TextMatrix(i, COL_开嘱医生) & "'," & IIF(rsPrice!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                                            rsPrice!ID & ",'" & rsPrice!类别 & "','" & rsPrice!计算单位 & "'," & _
                                            int付数 & "," & rsDays!数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                            IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsPrice!收入项目ID & "," & _
                                            "'" & rsPrice!收据费目 & "'," & dbl单价 & "," & rsDays!应收金额 & "," & rsDays!实收金额 & "," & _
                                            rsDays!发生时间 & "," & str登记时间 & "," & _
                                            "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                            "Null,'" & .TextMatrix(i, col_医嘱内容) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_频率) & "'," & _
                                            ZVal(.TextMatrix(i, COL_单量)) & ",'" & .TextMatrix(i, COL_用法) & "'," & .Cell(flexcpData, i, COL_医嘱期效) & "," & _
                                            IIF(int药品性质 <> 0, int药品性质, Val(.TextMatrix(i, COL_计价特性))) & ",1,Null,0," & ZVal(Val(.TextMatrix(i, COL_检查方法))) & "," & ZVal(lng主页ID) & "," & Val(.TextMatrix(i, COL_病人病区ID)) & ")"
                                    Else
                                        rsSQL!Sql = "ZL_住院记帐记录_Insert(" & _
                                            "'" & rsDays!NO & "'," & lng费用序号 & "," & lng病人ID & "," & ZVal(lng主页ID) & "," & _
                                            IIF(.TextMatrix(i, COL_住院号) = "", "NULL", "'" & .TextMatrix(i, COL_住院号) & "'") & ",'" & .TextMatrix(i, COL_姓名) & "'," & _
                                            "'" & .TextMatrix(i, col_性别) & "','" & .TextMatrix(i, COL_年龄) & "'," & _
                                            "'" & .TextMatrix(i, COL_床号) & "','" & .TextMatrix(i, COL_费别) & "'," & _
                                            Val(.TextMatrix(i, COL_病人病区ID)) & "," & Val(.TextMatrix(i, COL_病人科室ID)) & ",0," & Val(.Cell(flexcpData, i, COL_婴儿)) & "," & _
                                            ZVal(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                                            IIF(rsPrice!从项 = 1, ZVal(int父序号), "NULL") & "," & rsPrice!ID & "," & _
                                            "'" & rsPrice!类别 & "','" & NVL(rsPrice!计算单位) & "'," & _
                                            IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                                            int付数 & "," & rsDays!数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                            IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsPrice!收入项目ID & "," & _
                                            "'" & NVL(rsPrice!收据费目) & "'," & dbl单价 & "," & rsDays!应收金额 & "," & rsDays!实收金额 & "," & _
                                            cur统筹金额 & "," & rsDays!发生时间 & "," & str登记时间 & "," & _
                                            "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & _
                                            IIF(rsPrice!类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                                            "NULL,'" & .TextMatrix(i, col_医嘱内容) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                            "'" & .TextMatrix(i, COL_频率) & "'," & ZVal(.TextMatrix(i, COL_单量)) & "," & _
                                            "'" & .TextMatrix(i, COL_用法) & "'," & .Cell(flexcpData, i, COL_医嘱期效) & "," & _
                                            IIF(int药品性质 <> 0, int药品性质, Val(.TextMatrix(i, COL_计价特性))) & "," & _
                                            "Null,'" & str费用类型 & "',Null," & strTmp & ",NULL,-1,0," & ZVal(Val(.TextMatrix(i, COL_检查方法))) & ")"
                                    End If
                                    rsSQL.Update
                                    
                                
                                    '记录自动发料的SQL
                                    If (gbyt住院自动发料 = 1 Or gbyt住院自动发料 = 2 And lng执行科室ID = Val(.TextMatrix(i, COL_开嘱科室ID))) And int划价 = 0 And lng执行科室ID <> 0 And rsPrice!类别 = "4" And NVL(rsPrice!跟踪在用, 0) = 1 Then
                                        If InStr(str自动发料 & ";", ";" & rsDays!NO & "," & lng执行科室ID & ";") = 0 Then
                                            rsSQL.AddNew
                                            rsSQL!类型 = 6
                                            rsSQL!医嘱ID = lng组ID
                                            rsSQL!项目ID = 0
                                            rsSQL!序号 = i
                                            rsSQL!NO = rsDays!NO
                                            rsSQL!Sql = "zl_材料收发记录_处方发料(" & lng执行科室ID & ",25,'" & rsDays!NO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                                            rsSQL.Update
                                            str自动发料 = str自动发料 & ";" & rsDays!NO & "," & lng执行科室ID
                                        End If
                                    End If
                                    rsDays.MoveNext
                                Loop
                                
                                '医保管控实时监测：生成费用项目记录集,以收费细目汇总
                                If Not IsNull(rsPati!险类) And bln实时监控 Then
                                    rsItems.Filter = "收费细目ID=" & rsPrice!ID
                                    If rsItems.EOF Then
                                        '加入收费项目对应的原始信息
                                        rsItems.AddNew
                                        rsItems!病人ID = rsPati!病人ID
                                        rsItems!主页ID = rsPati!主页ID
                                        rsItems!医嘱ID = Val(.TextMatrix(i, COL_ID))
                                        rsItems!收费类别 = rsPrice!类别
                                        rsItems!收费细目ID = rsPrice!ID
                                        rsItems!开单人 = .TextMatrix(i, COL_开嘱医生)
                                        rsItems!开单科室 = CStr(sys.RowValue("部门表", Val(.TextMatrix(i, COL_开嘱科室ID)), "名称"))
                                        
                                        rsItems!数量 = int付数 * dbl数量
                                        rsItems!单价 = dbl单价
                                    Else
                                        '基于一个医嘱(诊疗项目)的收费对照不会有重复的收费细目
                                        '数量：同一收费项目的不同收入项目记录相同
                                        If rsPrice!费用性质 & "_" & rsPrice!ID <> str收费项目 Then
                                            rsItems!数量 = NVL(rsItems!数量, 0) + int付数 * dbl数量
                                        End If
                                        '单价：同一收费项目的不同收入项目累加
                                        If Val(.TextMatrix(i, COL_ID)) = rsItems!医嘱ID Then
                                            rsItems!单价 = NVL(rsItems!单价, 0) + dbl单价
                                        End If
                                    End If
                                    rsItems!实收金额 = NVL(rsItems!实收金额, 0) + cur实收
                                    rsItems.Update
                                End If
                                    
                                str收费项目 = rsPrice!费用性质 & "_" & rsPrice!ID
                                rsPrice.MoveNext
                            Loop
                        End If
                        
                        '对医嘱金额进行汇总折扣处理
                        If gbln从项汇总折扣 And strHaveSub <> "" Then
                            rsSeek.Filter = 0
                            Do While Not rsSeek.EOF
                                rsSQL.Bookmark = rsSeek!主项标签
                                cur实收 = Format(ActualMoney(.TextMatrix(i, COL_费别), rsSeek!主收入ID, rsSeek!合计), gstrDec)
                                cur实收 = cur实收 - rsSeek!合计 '打折差额
                                
                                '医保管控实时监测：费用项目金额替换
                                If Not IsNull(rsPati!险类) And bln实时监控 Then
                                    rsItems.Filter = "收费细目ID=" & lng父项目ID
                                    If Not rsItems.EOF Then
                                        rsItems!实收金额 = NVL(rsItems!实收金额, 0) + cur实收
                                        rsItems.Update
                                    End If
                                End If
                                
                                '费用SQL生成替换
                                cur实收 = Get实收金额(rsSQL!Sql) + cur实收
                                rsSQL!Sql = Set实收金额(rsSQL!Sql, cur实收)
                                rsSQL.Update
                            
                                rsSeek.MoveNext
                            Loop
                        End If
                                                
                        
                        '产生医嘱发送记录
                        '-----------------------------------------------------------------------------------------
                        If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then  '叮嘱不发送(给药途径，配方煎法、用法、采集方法可能为)
                            '发送了出院,转院,死亡医嘱
                            If .TextMatrix(i, COL_诊疗类别) = "Z" _
                                And InStr(",5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                                mblnRefresh = True
                            End If
                            
                            '一样要产生费用NO
                            Call GetCurBillSet(strNOKey, strNO, -1, lng发送序号)
                                                                    
                            '是否一组医嘱的第一医嘱行
                            blnFirst = False
                            If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                    blnFirst = True '药疗发送时,只有第一药品行才为第一医嘱行
                                End If
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "C" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                    blnFirst = True '检验组合中的第一检验行
                                End If
                            ElseIf Val(.TextMatrix(i, COL_相关ID)) = 0 Then '排开了中药煎法、输血途径
                                If Not (.TextMatrix(i, COL_诊疗类别) = "E" _
                                    And Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(i - 1, COL_相关ID))) Then '排开给药途径、中药用法、采集方法
                                    blnFirst = True
                                End If
                            End If
                            
                            '本科执行的自动执行：特殊医嘱不处理
                            int执行状态 = 0
                            If Val(Mid(mstrAutoExe, IIF(.TextMatrix(i, COL_医嘱期效) = "临嘱", 1, 0) + 1, 1)) <> 0 And Not (.TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) <> 0) _
                                And (Val(.TextMatrix(i, COL_执行科室ID)) = Val(.TextMatrix(i, COL_病人科室ID)) Or Val(.TextMatrix(i, COL_执行科室ID)) = Val(.TextMatrix(i, COL_病人病区ID))) Then
                                If CanAutoExeItem(Val(.TextMatrix(i, COL_执行科室ID)), .TextMatrix(i, COL_诊疗类别), .TextMatrix(i, COL_操作类型), Val(.TextMatrix(i, COL_执行分类))) Then
                                    int执行状态 = 1
                                End If
                                If int执行状态 = 1 Then
                                    '血库相关特殊处理
                                    If gbln血库系统 Then
                                        strTmp = .TextMatrix(i, COL_诊疗类别) & .TextMatrix(i, COL_操作类型) & .TextMatrix(i, COL_执行分类)
                                        If strTmp = "E81" Or strTmp = "E90" Then
                                            If InitObjBlood() Then
                                                int执行状态 = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
    
                            '首末时间(不能用“str分解时间”判断，因为一次性临嘱记录的是首次时间)
                            If .TextMatrix(i, COL_分解时间) <> "" Then
                                str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                                str末次时间 = "To_Date('" & Split(str分解时间, ",")(UBound(Split(str分解时间, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '无法分解或为"一次性"临嘱，填为开始执行时间（74366）
                                str首次时间 = "To_Date('" & .TextMatrix(i, COL_开始时间) & "','YYYY-MM-DD HH24:MI:SS')"
                                str末次时间 = "To_Date('" & .TextMatrix(i, COL_开始时间) & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                           
                            '发送数次:药品为剂量单位的总量,其它为次数
                            If .TextMatrix(i, COL_诊疗类别) = "7" Then
                                dbl发送数次 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量))
                            ElseIf InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                dbl发送数次 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_住院包装)) * Val(.TextMatrix(i, COL_剂量系数))
                            Else
                                dbl发送数次 = Val(.TextMatrix(i, COL_总量))
                            End If
                                       
                            '领药号
                            str领药号 = ""
                            If mbln领药号 And InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                If mstr领药号 = "" Then mstr领药号 = Get领药号
                                str领药号 = mstr领药号
                            End If
                            
                            '普通发送页面都不产生配药记录（64615）
                            If Not gbln发送生成条形码 Then strCuvetteNumber = ""
                            
                            '长嘱非药按天拆分发送记录
                            rsSendDays.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                            If rsSendDays.RecordCount = 1 Then
                                rsSendDays!分解时间 = str分解时间
                                rsSendDays!发送数次 = dbl发送数次
                                rsSendDays.Update
                            ElseIf rsSendDays.RecordCount = 0 Then
                                rsSendDays.AddNew
                                rsSendDays!医嘱ID = Val(.TextMatrix(i, COL_ID))
                                rsSendDays!发送号 = lng发送号
                                rsSendDays!NO = strNO
                                rsSendDays!分解时间 = str分解时间
                                rsSendDays!发送数次 = dbl发送数次
                                rsSendDays.Update
                            Else
                                If rsSendDays.RecordCount > 0 Then
                                    Do While Not rsSendDays.EOF
                                        rsSendDays!发送数次 = dbl发送数次 * ((UBound(Split(Get当天分解时间(rsSendDays!分解时间, rsSendDays!计算时间), ",")) + 1) / (UBound(Split(str分解时间, ",")) + 1))
                                        rsSendDays.MoveNext
                                    Loop
                                    rsSendDays.MoveFirst
                                End If
                            End If
                            
                            rsSendDays.Sort = "计算时间"
                            Do While Not rsSendDays.EOF
                                dbl发送数次 = Format(rsSendDays!发送数次, "0.00000")
                                rsSQL.AddNew
                                rsSQL!类型 = 5
                                rsSQL!医嘱ID = lng组ID
                                rsSQL!项目ID = 0
                                rsSQL!序号 = i
                                rsSQL!NO = rsSendDays!NO
                                
                                If InStr("," & strPlugIn发送号 & ",", "," & rsSendDays!发送号 & ",") = 0 Then
                                    strPlugIn发送号 = strPlugIn发送号 & "," & rsSendDays!发送号
                                End If
                                
                                rsSQL!Sql = "ZL_病人医嘱发送_Insert(" & _
                                    Val(.TextMatrix(i, COL_ID)) & "," & rsSendDays!发送号 & ",2,'" & rsSendDays!NO & "'," & _
                                    lng发送序号 & "," & ZVal(dbl发送数次) & "," & IIF(rsSendDays!首次时间 & "" = "", str首次时间, rsSendDays!首次时间) & "," & IIF(rsSendDays!末次时间 & "" = "", str末次时间, rsSendDays!末次时间) & "," & strCurDate & "," & _
                                    int执行状态 & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & int计费状态 & "," & _
                                    IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.编号 & "'," & _
                                    "'" & UserInfo.姓名 & "','" & str领药号 & "'," & IIF(lng病人性质 = 1, 1, "Null") & ",'" & rsSendDays!分解时间 & "'," & IIF(InStr(str关联药行, "," & Val(.TextMatrix(i, COL_ID)) & "'") > 0, str关联药行, "Null") & ")"
                                rsSQL.Update
                                str关联药行 = "''"
                                rsSendDays.MoveNext
                            Loop
                            If gbln血库系统 And .TextMatrix(i, COL_诊疗类别) = "K" Then
                                rsSQL.AddNew
                                rsSQL!类型 = 9
                                rsSQL!项目ID = 0
                                rsSQL!序号 = 0
                                rsSQL!Sql = "Zl_血液配血费用_Insert(" & Val(.TextMatrix(i, COL_ID)) & ")"
                                rsSQL.Update
                            End If
                            
                        End If
                    
                        '计算中药配方数
                        If .Cell(flexcpData, i, COL_诊疗类别) = 3 Then '中药用法
                            int配方数 = int配方数 + 1
                        End If
                    End If  '要发送和记帐的
                End If  '当前选择的
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
            lngOld组ID = lng组ID
        Next
        
        '提交最后一个病人的数据
        '-----------------------------------------------------------------------------------------
        If lng病人ID <> 0 Then
            If strAudit <> "" Then
                MsgBox "病人""" & rsPati!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, gstrSysName
                GoTo errH
            End If
            
            
            '医嘱执行计价
            If rsExec.RecordCount > 0 Then
                rsExec.MoveFirst
                Do While Not rsExec.EOF
                    rsSQL.AddNew
                    rsSQL!类型 = 8
                    rsSQL!项目ID = 0
                    rsSQL!序号 = 0
                    rsSQL!医嘱ID = lng组ID
                    rsSQL!Sql = "Zl_医嘱执行计价_Insert(" & rsExec!医嘱ID & "," & rsExec!发送号 & ",To_date('" & _
                    rsExec!要求时间 & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!收费细目ID & "")) & "," & rsExec!数量 & "," & rsExec!费用性质 & ")"
                    rsExec.MoveNext
                Loop
            End If
        
            If Not CompletePatiSend(rsPati, rsSQL, rsUpload, rsItems, cur合计, cur记帐合计, str类别, bln划价, blnTran, lng发送号) Then GoTo errH
            SendAdvice = lng发送号 '只要提交成功则标注
        End If
    End With
    mstrRollNotify = Mid(mstrRollNotify, 2)
    SendAdvice = lng发送号
    
    strPlugIn发送号 = Mid(strPlugIn发送号, 2)
    '调用外挂接口
    If CreatePlugInOK(p住院医嘱发送) Then
        On Error Resume Next
        Call gobjPlugIn.AdviceSendEnd(glngSys, p住院医嘱发送, strPlugIn发送号)
        Call zlPlugInErrH(err, "AdviceSendEnd")
        err.Clear: On Error GoTo 0
    End If
    Call Make待执行消息(strCurDateTmp)
FuncEnd:
    '删除所有已成功发送的行
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then
        gcnOracle.RollbackTrans
    End If
    If err.Number <> 0 Then '如医保上传失败退出没有错误
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function Get当天分解时间(ByVal str分解时间 As String, ByVal dat日期 As Date) As String
'功能：获取指定天的分解时间
    Dim arrTmp As Variant
    Dim i As Long
    
    arrTmp = Split(str分解时间, ",")
    For i = 0 To UBound(arrTmp)
        If Format(CDate(arrTmp(i)), "YYYY-MM-DD") = Format(dat日期, "YYYY-MM-DD") Then
            Get当天分解时间 = Get当天分解时间 & "," & arrTmp(i)
        End If
    Next
    Get当天分解时间 = Mid(Get当天分解时间, 2)
End Function

Private Sub ShowSendTotal()
'功能：根据当前选择要发送的医嘱，计算并显示发送的医嘱合计
    Dim curTotal As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) And .Cell(flexcpData, i, COL_选择) = 0 _
                And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                curTotal = curTotal + Val(.TextMatrix(i, COL_金额))
            End If
        Next
    End With
    stbThis.Panels(3).Text = "发送费用：" & Format(curTotal, gstrDec)
    Call Form_Resize
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'功能：设置执行科室输入的的值
    Dim str医嘱IDs As String
    Dim str执行科室IDs As String
    Dim i As Long
        
    With vsAdvice
        If lngCol = COL_执行科室 Then
            '更改显示行的执行科室显示
            .TextMatrix(lngRow, COL_执行科室) = rsInput!名称
            .Cell(flexcpData, lngRow, COL_执行科室) = .TextMatrix(lngRow, COL_执行科室)
            
            '更改主项目的执行科室(排开当前显示行为采集方式的行)
            If Not (.TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID))) Then
                
                str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(lngRow, COL_ID))
                str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(lngRow, COL_执行科室ID))
                
                .TextMatrix(lngRow, COL_执行科室ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_执行科室ID) = 1
                
            End If
            
            '手术或检查组合的附加内容
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If .TextMatrix(i, COL_诊疗类别) <> "G" _
                        And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then  '不更改手术麻醉的执行科室
                        
                        str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                        str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(i, COL_执行科室ID))
                        
                        .TextMatrix(i, COL_执行科室) = rsInput!名称
                        .Cell(flexcpData, i, COL_执行科室) = .TextMatrix(i, COL_执行科室)
                        .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                        .Cell(flexcpData, i, COL_执行科室ID) = 1
                    End If
                Else
                    Exit For
                End If
            Next
            
            '检验组合的内容
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                    
                        str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                        str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(i, COL_执行科室ID))
                        
                        .TextMatrix(i, COL_执行科室) = rsInput!名称
                        .Cell(flexcpData, i, COL_执行科室) = .TextMatrix(i, COL_执行科室)
                        .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                        .Cell(flexcpData, i, COL_执行科室ID) = 1
                        
                    End If
                Else
                    Exit For
                End If
            Next
        ElseIf lngCol = COL_附加执行 Then
            '更改显示行的附加执行科室显示
            .TextMatrix(lngRow, COL_附加执行) = rsInput!名称
            .Cell(flexcpData, lngRow, COL_附加执行) = .TextMatrix(lngRow, COL_附加执行)
            
            '更改附加项目行的执行科室
            If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                '给药途径
                i = .FindRow(CStr(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1, COL_ID)
                
                str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(i, COL_执行科室ID))
                
                .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                .Cell(flexcpData, i, COL_执行科室ID) = 1
                
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                        .TextMatrix(i, COL_附加执行) = rsInput!名称
                        .Cell(flexcpData, i, COL_附加执行) = .TextMatrix(lngRow, COL_附加执行)
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
                And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                
                str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(lngRow, COL_ID))
                str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(lngRow, COL_执行科室ID))
                        
                '中药用法
                .TextMatrix(lngRow, COL_执行科室ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_执行科室ID) = 1
            ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "F" Then
                '手术麻醉
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                        If .TextMatrix(i, COL_诊疗类别) = "G" Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                                
                                str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                                str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(i, COL_执行科室ID))
                                
                                .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                                .Cell(flexcpData, i, COL_执行科室ID) = 1
                                
                            End If
                            Exit For '只有一个麻醉
                        End If
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                '采集方式
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
                    
                    str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(lngRow, COL_ID))
                    str执行科室IDs = str执行科室IDs & "," & Val(.TextMatrix(lngRow, COL_执行科室ID))
                    
                    .TextMatrix(lngRow, COL_执行科室ID) = rsInput!ID
                    .Cell(flexcpData, lngRow, COL_执行科室ID) = 1
                End If
            End If
        End If
        If str医嘱IDs <> "" Then
            str医嘱IDs = Mid(str医嘱IDs, 2)
            str执行科室IDs = Mid(str执行科室IDs, 2)
            
            For i = 0 To UBound(Split(str医嘱IDs, ","))
                '同步更新费用执行科室（只更新和原医嘱执行科室相同的费用执行科室）
                mrsPrice.Filter = "医嘱ID=" & Val(Split(str医嘱IDs, ",")(i))
                If Not mrsPrice.EOF Then mrsPrice.MoveFirst
                Do Until mrsPrice.EOF
                    If Val(mrsPrice!执行科室ID & "") = Val(Split(str执行科室IDs, ",")(i)) And Val(Split(str执行科室IDs, ",")(i)) <> 0 Then
                        mrsPrice!执行科室ID = Val(rsInput!ID & "")
                        mrsPrice.Update
                    End If
                    mrsPrice.MoveNext
                Loop
            Next
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlcommfun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln非本科 As Boolean
    
    If Not CellEditablePrice(Row, Col, bln非本科) Then
        '非本科执行的变价项目不允许定价格
        If bln非本科 Then
            MsgBox "该医嘱非本科执行，不允许对变价项目定价。该计价项目需要手工计价。", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_计价数量 Or Col = COLP_单价 Or Col = COLP_执行科室 Then
            '必须先确定收费项目
            If vsPrice.TextMatrix(Row, COLP_收费项目) = "" Then Cancel = True
        End If
        If Col = COLP_单价 Then
            '输入变价前必须先确定计价医嘱,以决定是否可以输入(本科执行)
            If vsPrice.TextMatrix(Row, COLP_计价医嘱) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_计价数量 Or Col = COLP_单价 Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub GetPatiRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'功能：获取组ID相同的一组医嘱行号范围(注意考虑一并给药中的空行)
    Dim lng病人ID As Long, lng主页ID As Long, lng婴儿 As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lng病人ID = Val(.TextMatrix(lngRow, COL_病人ID))
        lng主页ID = Val(.TextMatrix(lngRow, COL_主页ID))
        lng婴儿 = Val(.TextMatrix(lngRow, COL_婴儿))
        
        For i = lngRow - 1 To .FixedRows Step -1
            If lng病人ID = Val(.TextMatrix(lngRow, COL_病人ID)) And lng主页ID = Val(.TextMatrix(lngRow, COL_主页ID)) And lng婴儿 = Val(.Cell(flexcpData, lngRow, COL_婴儿)) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If Not (lng病人ID = Val(.TextMatrix(lngRow, COL_病人ID)) And lng主页ID = Val(.TextMatrix(lngRow, COL_主页ID)) And lng婴儿 = Val(.Cell(flexcpData, lngRow, COL_婴儿))) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub



Private Sub Del检验申请()
'功能：医嘱发送失败，事务回退后，调用检验申请删除接口
    Dim i As Long, str医嘱IDs As String, strErr As String
        
    '收集采集方法
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_操作类型)) = 6 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    str医嘱IDs = str医嘱IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Call InitObjLis(p住院护士站)
    If str医嘱IDs <> "" Then
        str医嘱IDs = Mid(str医嘱IDs, 2)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(str医嘱IDs, strErr) = False Then
                MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function Get检验医嘱信息(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：获取检验医嘱信息，传递给检验接口程序
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_操作类型)) = 6 And .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_病人ID)) = lng病人ID And Val(.TextMatrix(i, COL_主页ID)) = lng主页ID Then
                '检验医嘱ID1,采集医嘱ID1,执行科室ID1,标本1;.....
                'LIS接口产生的检验，一个采集方式只有一条检验医嘱（没有一并采集的情况）
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    strInfo = strInfo & ";" & .TextMatrix(i - 1, COL_ID) & "," & .TextMatrix(i, COL_ID) & "," & .TextMatrix(i - 1, COL_执行科室ID) & "," & .TextMatrix(i - 1, COL_标本部位)
                End If
            End If
        Next
    End With
    Get检验医嘱信息 = Mid(strInfo, 2)
End Function

Private Function Get输血医嘱信息(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：获取检验医嘱信息，传递给检验接口程序
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_病人ID)) = lng病人ID And Val(.TextMatrix(i, COL_主页ID)) = lng主页ID Then
                '检验医嘱ID1,采集医嘱ID1,执行科室ID1,标本1;.....
                'LIS接口产生的检验，一个采集方式只有一条检验医嘱（没有一并采集的情况）
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    strInfo = strInfo & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Get输血医嘱信息 = Mid(strInfo, 2)
End Function

Private Function CheckAdviceUpdate(ByVal lng医嘱ID As Long, ByVal str新开操作时间 As String) As Boolean
'功能：如果是无需校对模式，则检查是否有并发修改。
    Dim rsTmp As Recordset, strSQL As String
    
    strSQL = "Select 操作时间 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        If Format(rsTmp!操作时间, "yyyy-MM-dd HH:mm:ss") <> str新开操作时间 Then CheckAdviceUpdate = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitExecRecordset(rsExec As Recordset)
'功能：初始化医嘱计价记录集
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "医嘱ID", adBigInt
    rsExec.Fields.Append "发送号", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "要求时间", adDate, , adFldIsNullable
    rsExec.Fields.Append "数量", adDouble, , adFldIsNullable
    rsExec.Fields.Append "费用性质", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
End Sub

Private Function zlPluginAdviceBeforeSend() As Boolean
'功能：医嘱发送前调用外挂号
    Dim i As Long, j As Long
    Dim strAdviceIDs As String, strMsg  As String
    Dim rsDataPlugIn As ADODB.Recordset
    Dim lng数量 As Long
    Dim str分解时间 As String, strTmp As String
    
    zlPluginAdviceBeforeSend = True
    
    '调用外挂接口，医嘱发送前的检查
    If CreatePlugInOK(p住院医嘱发送) Then
        Call InitPlugInRs(rsDataPlugIn)
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    If .TextMatrix(i, COL_分解时间) <> "" Then
                        str分解时间 = .TextMatrix(i, COL_分解时间)
                    Else
                        str分解时间 = .Cell(flexcpData, i, COL_分解时间)    '开始执行时间
                    End If
                    rsDataPlugIn.AddNew
                    rsDataPlugIn!病人ID = Val(.TextMatrix(i, COL_病人ID))
                    rsDataPlugIn!就诊ID = Val(.TextMatrix(i, COL_主页ID))
                    rsDataPlugIn!医嘱ID = Val(.TextMatrix(i, COL_ID))
                    rsDataPlugIn!相关ID = Val(.TextMatrix(i, COL_相关ID))
                    rsDataPlugIn!收费细目ID = Val(.TextMatrix(i, COL_收费细目ID))
                    rsDataPlugIn!分解时间 = str分解时间
                    rsDataPlugIn!次数 = Val(.TextMatrix(i, COL_次数))
                    rsDataPlugIn!单量 = Val(.TextMatrix(i, COL_单量))
                    rsDataPlugIn!单量单位 = .TextMatrix(i, COL_单量单位)
                    rsDataPlugIn!总量 = Val(.TextMatrix(i, COL_总量))
                    rsDataPlugIn!总量单位 = .TextMatrix(i, COL_总量单位)
                    rsDataPlugIn!场合 = 1
                    rsDataPlugIn.Update
                End If
            Next
            If rsDataPlugIn.RecordCount > 0 Then rsDataPlugIn.MoveFirst
            strAdviceIDs = "": strMsg = ""
            On Error Resume Next
            Call gobjPlugIn.AdviceBeforeSend(mstrEnd, rsDataPlugIn, strAdviceIDs, strMsg)
            Call zlPlugInErrH(err, "AdviceBeforeSend")
            err.Clear
            On Error GoTo 0
             
            If strAdviceIDs <> "" Then
                strTmp = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        If InStr("," & strAdviceIDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                                j = Val(.TextMatrix(i, COL_ID))
                            Else
                                j = Val(.TextMatrix(i, COL_相关ID))
                            End If
                            
                            If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                                strTmp = strTmp & "," & j
                            End If
                        End If
                    End If
                Next
                strAdviceIDs = Mid(strTmp, 2)
                lng数量 = 0
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            j = Val(.TextMatrix(i, COL_ID))
                        Else
                            j = Val(.TextMatrix(i, COL_相关ID))
                        End If
                        lng数量 = lng数量 + 1
                        If InStr("," & strAdviceIDs & ",", "," & j & ",") > 0 Then
                            .Cell(flexcpData, i, COL_选择) = 1
                            Set .Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                            lng数量 = lng数量 - 1
                        End If
                    End If
                Next
                
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                If lng数量 = 0 Then
                    MsgBox "当前没有可以发送的医嘱。", vbInformation, gstrSysName
                    zlPluginAdviceBeforeSend = False
                End If
            End If
        End With
    End If
End Function

Private Function GetAdviceRis(ByRef rsData As ADODB.Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：获取发送到RIS的医嘱信息
    Dim i As Long
    
    On Error GoTo errH
    
    Set rsData = New ADODB.Recordset
    
    rsData.Fields.Append "医嘱ID", adBigInt
    rsData.Fields.Append "开嘱科室ID", adBigInt
    rsData.Fields.Append "执行科室ID", adBigInt
    rsData.Fields.Append "诊疗项目ID", adBigInt
    rsData.Fields.Append "病人来源", adInteger '1-门诊;2-住院;
    rsData.Fields.Append "类别", adVarChar, 10
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If Val(.TextMatrix(i, COL_病人ID)) = lng病人ID And Val(.TextMatrix(i, COL_主页ID)) = lng主页ID Then
                    If InStr(",D,F,", .TextMatrix(i, COL_诊疗类别)) > 0 Or _
                        .TextMatrix(i, COL_医嘱期效) = "临嘱" And InStr(",0,5,", Val(.TextMatrix(i, COL_操作类型))) > 0 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            rsData.AddNew
                            rsData!医嘱ID = Val(.TextMatrix(i, COL_ID))
                            rsData!开嘱科室id = Val(.TextMatrix(i, COL_开嘱科室ID))
                            rsData!执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
                            rsData!诊疗项目ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                            rsData!病人来源 = 2
                            rsData!类别 = .TextMatrix(i, COL_诊疗类别)
                            rsData.Update
                        End If
                    End If
                End If
            End If
        Next
    End With
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        GetAdviceRis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDrug零差价(ByVal lngRow As Long, ByRef bln提示 As Boolean) As Boolean
'功能：发送过程中对零差价药品进行检查禁止
    Dim strTmp As String
    Dim blnTmp As Boolean
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        If 0 <> Val(.TextMatrix(lngRow, COL_收费细目ID)) And 0 <> Val(.TextMatrix(lngRow, COL_执行科室ID)) And .Cell(flexcpData, lngRow, COL_选择) <> 1 Then
            If InitObjPublicDrug Then
                blnTmp = gobjPublicDrug.zlCheckPriceAdjustBySell(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), False)
                If Not blnTmp Then
                    strTmp = "在(" & .TextMatrix(lngRow, COL_执行科室) & ")中药品""" & .TextMatrix(lngRow, col_医嘱内容) & """" & vbCrLf & vbCrLf & _
                        "不满足零差价管理的要求：成本价和售价不一致，不能销售出库。" & vbCrLf & vbCrLf & _
                        "请联系药房或药剂科进行调价处理。"
                    
                    If bln提示 Then
                        .Redraw = flexRDDirect:
                        Call .ShowCell(lngRow, COL_选择)
                        Screen.MousePointer = 0
                        vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
                        If vMsg = vbIgnore Then bln提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        Screen.MousePointer = 11
                        .Refresh: .Redraw = flexRDNone
                    Else
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    End If
                    CheckDrug零差价 = True
                End If
            End If
        End If
    End With
End Function

Private Function CheckRISScheduling() As Boolean
'功能：检查项目是否是必须预约
    Dim i As Long
    Dim blnDo As Boolean
    Dim lng医嘱ID As Long
    Dim lng诊疗项目ID As Long
    Dim lngRst As Long
    Dim strMsg As String
    
    CheckRISScheduling = True
    
    If HaveRIS Then
        If gbln启用影像信息系统预约 Then
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Function
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If Val(.TextMatrix(i, COL_紧急标志)) <> 1 Then
                    If InStr(",D,F,", .TextMatrix(i, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_操作类型))) > 0 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            lng医嘱ID = Val(.TextMatrix(i, COL_ID))
                            lng诊疗项目ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                            lngRst = -1
                            lngRst = gobjRis.HISScheduling(2, lng医嘱ID, lng诊疗项目ID, False)
                            If lngRst <> 0 Then
                            '接口返回失败给出提示
                                .Cell(flexcpData, i, COL_选择) = 1 '当前禁止选择
                                Set .Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                                Call RowSelectSame(i, COL_选择)
                                strMsg = IIF("" = strMsg, "", strMsg & "、") & .TextMatrix(i, col_医嘱内容)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
    If strMsg <> "" Then
        MsgBox "由于启用了RIS系统预约流程：" & vbCrLf & "【" & strMsg & "】" & _
                vbCrLf & "医嘱没有预约，预约成功后才能发送。", vbInformation, gstrSysName
        CheckRISScheduling = False
    End If
End Function

Private Function Set阳性用药() As Boolean
'功能：设置药品医嘱行的阳性用药说明
    Dim i As Long
    Dim strMsg As String
    Dim str阳性用药 As String
    Dim strSQL As String
    Dim str医嘱IDs As String
    
    On Error GoTo errH
    If mstrAdDrugIDs = "" Then
        Set阳性用药 = True
        Exit Function
    End If
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If InStr("," & mstrAdDrugIDs & ",", "," & .TextMatrix(i, COL_ID) & ",") > 0 Then
                    strMsg = strMsg & "," & .TextMatrix(i, col_医嘱内容)
                    str医嘱IDs = str医嘱IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    If strMsg = "" Then
        Set阳性用药 = True
        Exit Function
    End If
    Call frmMsgDruExcess.ShowMe(Me, 1, Mid(strMsg, 2), str阳性用药)
    If str阳性用药 = "*NULL*" Then
        Exit Function
    End If
    strSQL = "Zl_病人医嘱记录_阳性用药('" & Mid(str医嘱IDs, 2) & "','" & str阳性用药 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set阳性用药 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
