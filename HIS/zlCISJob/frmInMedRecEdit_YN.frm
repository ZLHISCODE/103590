VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#1.1#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmInMedRecEdit_YN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "首页整理"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   Icon            =   "frmInMedRecEdit_YN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrintdown 
      Caption         =   "↓"
      Height          =   350
      Left            =   2370
      TabIndex        =   383
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   7155
      Width           =   270
   End
   Begin VB.CommandButton cmdPriviewDown 
      Caption         =   "↓"
      Height          =   350
      Left            =   1080
      TabIndex        =   382
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   7155
      Width           =   270
   End
   Begin VB.CommandButton cmdPriview 
      Caption         =   "预览"
      Height          =   350
      Left            =   240
      TabIndex        =   381
      Top             =   7155
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   350
      Left            =   1485
      TabIndex        =   380
      Top             =   7155
      Width           =   900
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   6975
      Left            =   120
      TabIndex        =   328
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   3
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "基本信息"
      TabPicture(0)   =   "frmInMedRecEdit_YN.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraInfo(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "西医诊断"
      TabPicture(1)   =   "frmInMedRecEdit_YN.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraInfo(1)"
      Tab(1).Control(1)=   "cmdInfo(35)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "中医诊断"
      TabPicture(2)   =   "frmInMedRecEdit_YN.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInfo(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "过敏与手术"
      TabPicture(3)   =   "frmInMedRecEdit_YN.frx":05DE
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraInfo(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "住院情况"
      TabPicture(4)   =   "frmInMedRecEdit_YN.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraInfo(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "放疗与化疗"
      TabPicture(5)   =   "frmInMedRecEdit_YN.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraInfo(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "附页1"
      TabPicture(6)   =   "frmInMedRecEdit_YN.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraInfo(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "附页2"
      TabPicture(7)   =   "frmInMedRecEdit_YN.frx":064E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraInfo(7)"
      Tab(7).ControlCount=   1
      Begin VB.Frame fraInfo 
         BorderStyle     =   0  'None
         Height          =   6495
         Index           =   7
         Left            =   -74880
         TabIndex        =   349
         Top             =   360
         Width           =   10455
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   65
            ItemData        =   "frmInMedRecEdit_YN.frx":066A
            Left            =   1920
            List            =   "frmInMedRecEdit_YN.frx":066C
            Style           =   2  'Dropdown List
            TabIndex        =   311
            Top             =   1200
            Width           =   2445
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "非预期的重返重症医学科"
            Height          =   195
            Index           =   25
            Left            =   4440
            TabIndex        =   310
            Top             =   840
            Width           =   2850
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "发生人工气道脱出 "
            Height          =   195
            Index           =   24
            Left            =   1920
            TabIndex        =   309
            Top             =   840
            Width           =   2250
         End
         Begin VB.Frame fra准确度 
            Height          =   75
            Index           =   0
            Left            =   2520
            TabIndex        =   375
            Top             =   173
            Width           =   7695
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   59
            Left            =   4080
            TabIndex        =   308
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   450
            Width           =   270
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   64
            ItemData        =   "frmInMedRecEdit_YN.frx":066E
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":0670
            Style           =   2  'Dropdown List
            TabIndex        =   326
            Top             =   5940
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   63
            ItemData        =   "frmInMedRecEdit_YN.frx":0672
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":0674
            Style           =   2  'Dropdown List
            TabIndex        =   323
            Top             =   4335
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   62
            ItemData        =   "frmInMedRecEdit_YN.frx":0676
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":0678
            Style           =   2  'Dropdown List
            TabIndex        =   324
            Top             =   4710
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   61
            ItemData        =   "frmInMedRecEdit_YN.frx":067A
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":067C
            Style           =   2  'Dropdown List
            TabIndex        =   325
            Top             =   5100
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   58
            Left            =   7890
            MaxLength       =   5
            TabIndex        =   322
            Top             =   3960
            Width           =   885
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "住院期间使用物理约束"
            Height          =   195
            Index           =   21
            Left            =   6240
            TabIndex        =   321
            Top             =   3600
            Width           =   2850
         End
         Begin VB.Frame fra准确度 
            Height          =   75
            Index           =   7
            Left            =   1440
            TabIndex        =   356
            Top             =   4980
            Width           =   4335
         End
         Begin VB.Frame fraInfection 
            Caption         =   "感染因素"
            Height          =   1695
            Left            =   6000
            TabIndex        =   319
            Top             =   1680
            Width           =   4335
            Begin VB.ListBox lstInfection 
               Height          =   1320
               ItemData        =   "frmInMedRecEdit_YN.frx":067E
               Left            =   120
               List            =   "frmInMedRecEdit_YN.frx":0685
               Style           =   1  'Checkbox
               TabIndex        =   320
               Top             =   240
               Width           =   4125
            End
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "1.C&T"
            Height          =   195
            Index           =   12
            Left            =   285
            TabIndex        =   315
            Top             =   5235
            Width           =   675
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "2.&MRI"
            Height          =   195
            Index           =   13
            Left            =   1140
            TabIndex        =   316
            Top             =   5235
            Width           =   765
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "3.彩色多普勒(&R)"
            Height          =   195
            Index           =   14
            Left            =   2100
            TabIndex        =   317
            Top             =   5235
            Width           =   1665
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
            Height          =   930
            Left            =   240
            TabIndex        =   318
            Top             =   5505
            Width           =   5535
            _cx             =   9763
            _cy             =   1640
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":0697
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
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
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   2850
            Left            =   240
            TabIndex        =   314
            Top             =   1920
            Width           =   5565
            _cx             =   9816
            _cy             =   5027
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
            ForeColorSel    =   -2147483642
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   100
            ColWidthMax     =   2400
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
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
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   59
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   307
            Top             =   420
            Width           =   2445
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重返间隔时间"
            Height          =   180
            Index           =   128
            Left            =   780
            TabIndex        =   376
            Top             =   1260
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重症监护室名称"
            Height          =   180
            Index           =   127
            Left            =   600
            TabIndex        =   374
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入住重症监护室（ICU）情况"
            Height          =   180
            Index           =   126
            Left            =   240
            TabIndex        =   373
            Top             =   120
            Width           =   2250
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "离院方式"
            Height          =   180
            Index           =   124
            Left            =   7110
            TabIndex        =   371
            Top             =   6000
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "产科新生儿情况"
            Height          =   180
            Index           =   122
            Left            =   6240
            TabIndex        =   370
            Top             =   5640
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "约束工具"
            Height          =   180
            Index           =   121
            Left            =   7110
            TabIndex        =   369
            Top             =   4770
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "约束方式"
            Height          =   180
            Index           =   120
            Left            =   7110
            TabIndex        =   368
            Top             =   4395
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "约束原因"
            Height          =   180
            Index           =   119
            Left            =   7110
            TabIndex        =   367
            Top             =   5160
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "小时"
            Height          =   180
            Index           =   116
            Left            =   8895
            TabIndex        =   366
            Top             =   4020
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "约束总时间"
            Height          =   180
            Index           =   107
            Left            =   6930
            TabIndex        =   365
            Top             =   4020
            Width           =   900
         End
         Begin VB.Label lbl附加项目 
            AutoSize        =   -1  'True
            Caption         =   "病案附加项目"
            Height          =   180
            Left            =   240
            TabIndex        =   312
            Top             =   1680
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "特殊检查情况"
            Height          =   180
            Index           =   83
            Left            =   285
            TabIndex        =   313
            Top             =   4920
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   52
            Left            =   10680
            TabIndex        =   350
            Top             =   3135
            Width           =   180
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000B&
            Index           =   8
            X1              =   1440
            X2              =   5800
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            Index           =   9
            X1              =   1440
            X2              =   5800
            Y1              =   1785
            Y2              =   1785
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   6
         Left            =   -74880
         TabIndex        =   348
         Top             =   390
         Width           =   10455
         Begin VB.CheckBox chkInfo 
            Caption         =   "住院期间出现危重"
            Height          =   195
            Index           =   1
            Left            =   7230
            TabIndex        =   306
            Top             =   3600
            Width           =   2370
         End
         Begin VB.Frame fraAdvEvent 
            Caption         =   "不良事件"
            Height          =   2955
            Left            =   2760
            TabIndex        =   360
            Top             =   3360
            Width           =   4335
            Begin VB.ListBox lstAdvEvent 
               Height          =   1530
               ItemData        =   "frmInMedRecEdit_YN.frx":0705
               Left            =   120
               List            =   "frmInMedRecEdit_YN.frx":070C
               Style           =   1  'Checkbox
               TabIndex        =   301
               Top             =   240
               Width           =   4125
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   46
               ItemData        =   "frmInMedRecEdit_YN.frx":071E
               Left            =   3315
               List            =   "frmInMedRecEdit_YN.frx":0720
               Style           =   2  'Dropdown List
               TabIndex        =   303
               Top             =   1800
               Width           =   900
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   45
               ItemData        =   "frmInMedRecEdit_YN.frx":0722
               Left            =   1440
               List            =   "frmInMedRecEdit_YN.frx":0724
               Style           =   2  'Dropdown List
               TabIndex        =   302
               Top             =   1800
               Width           =   1335
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   48
               ItemData        =   "frmInMedRecEdit_YN.frx":0726
               Left            =   1440
               List            =   "frmInMedRecEdit_YN.frx":0728
               Style           =   2  'Dropdown List
               TabIndex        =   305
               Top             =   2520
               Width           =   2775
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   47
               ItemData        =   "frmInMedRecEdit_YN.frx":072A
               Left            =   1440
               List            =   "frmInMedRecEdit_YN.frx":072C
               Style           =   2  'Dropdown List
               TabIndex        =   304
               Top             =   2160
               Width           =   2775
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分期"
               Height          =   180
               Index           =   91
               Left            =   2835
               TabIndex        =   364
               Top             =   1860
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "压疮发生期间"
               Height          =   180
               Index           =   89
               Left            =   300
               TabIndex        =   363
               Top             =   1860
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "跌倒或坠床原因"
               Height          =   180
               Index           =   92
               Left            =   120
               TabIndex        =   362
               Top             =   2580
               Width           =   1260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "跌倒或坠床伤害"
               Height          =   180
               Index           =   90
               Left            =   120
               TabIndex        =   361
               Top             =   2220
               Width           =   1260
            End
         End
         Begin VB.Frame fraPath 
            Caption         =   "临床路径信息"
            Height          =   2955
            Left            =   240
            TabIndex        =   295
            Top             =   3360
            Width           =   2415
            Begin VB.CheckBox chkInfo 
               Caption         =   "进入路径"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   296
               Top             =   420
               Width           =   1050
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "完成路径"
               Enabled         =   0   'False
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   297
               TabStop         =   0   'False
               Top             =   720
               Width           =   1050
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "变异"
               Enabled         =   0   'False
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   299
               TabStop         =   0   'False
               Top             =   1680
               Width           =   690
            End
            Begin VB.TextBox txtInfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               Index           =   61
               Left            =   360
               MaxLength       =   100
               TabIndex        =   298
               TabStop         =   0   'False
               Top             =   1275
               Width           =   1965
            End
            Begin VB.TextBox txtInfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               Index           =   62
               Left            =   360
               MaxLength       =   100
               TabIndex        =   300
               TabStop         =   0   'False
               Top             =   2220
               Width           =   1965
            End
            Begin VB.CommandButton cmdPathLoad 
               Caption         =   "自动提取"
               Height          =   350
               Left            =   1320
               TabIndex        =   357
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "退出原因"
               Height          =   180
               Index           =   117
               Left            =   375
               TabIndex        =   359
               Top             =   1020
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "变异原因"
               Height          =   180
               Index           =   118
               Left            =   375
               TabIndex        =   358
               Top             =   1980
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdAutoLoad 
            Caption         =   "自动提取"
            Height          =   350
            Index           =   0
            Left            =   9240
            TabIndex        =   353
            Top             =   120
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vsKSS 
            Height          =   2685
            Left            =   240
            TabIndex        =   294
            Top             =   555
            Width           =   10155
            _cx             =   17912
            _cy             =   4736
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":072E
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抗菌药物使用情况（按DDD数降序排列）"
            Height          =   180
            Index           =   82
            Left            =   360
            TabIndex        =   293
            Top             =   270
            Width           =   3150
         End
      End
      Begin VB.Frame fraInfo 
         BorderStyle     =   0  'None
         Height          =   6345
         Index           =   5
         Left            =   -74880
         TabIndex        =   339
         Top             =   420
         Width           =   10480
         Begin VSFlex8Ctl.VSFlexGrid vs化疗 
            Height          =   2715
            Left            =   45
            TabIndex        =   290
            Top             =   345
            Width           =   10440
            _cx             =   18415
            _cy             =   4789
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483644
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   12632256
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":0844
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
         Begin VSFlex8Ctl.VSFlexGrid vs放疗 
            Height          =   2805
            Left            =   45
            TabIndex        =   292
            Top             =   3480
            Width           =   10440
            _cx             =   18415
            _cy             =   4948
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483644
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   12632256
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":0971
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
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   1680
            TabIndex        =   352
            Top             =   3120
            Width           =   90
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   351
            Top             =   30
            Width           =   90
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "化疗记录信息"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   289
            Top             =   30
            Width           =   1080
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "放疗记录信息"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   291
            Top             =   3240
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   240
         Index           =   35
         Left            =   -64800
         TabIndex        =   149
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   6060
         Width           =   270
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   1
         Left            =   -74880
         TabIndex        =   337
         Top             =   420
         Width           =   10425
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   58
            ItemData        =   "frmInMedRecEdit_YN.frx":0A98
            Left            =   8760
            List            =   "frmInMedRecEdit_YN.frx":0A9A
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   4425
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   31
            ItemData        =   "frmInMedRecEdit_YN.frx":0A9C
            Left            =   1410
            List            =   "frmInMedRecEdit_YN.frx":0A9E
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   4770
            Width           =   1470
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "是否确诊(&B)"
            Height          =   195
            Index           =   0
            Left            =   5160
            TabIndex        =   117
            Top             =   3780
            Width           =   1290
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   3720
            Width           =   1470
         End
         Begin VB.CommandButton cmdInfo 
            Height          =   240
            Index           =   27
            Left            =   9950
            Picture         =   "frmInMedRecEdit_YN.frx":0AA0
            Style           =   1  'Graphical
            TabIndex        =   355
            TabStop         =   0   'False
            ToolTipText     =   "选择(F4)"
            Top             =   3750
            Width           =   240
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   57
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   121
            Top             =   4080
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   34
            ItemData        =   "frmInMedRecEdit_YN.frx":0B96
            Left            =   5010
            List            =   "frmInMedRecEdit_YN.frx":0B98
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   4770
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   32
            ItemData        =   "frmInMedRecEdit_YN.frx":0B9A
            Left            =   4650
            List            =   "frmInMedRecEdit_YN.frx":0B9C
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   5610
            Width           =   1470
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   50
            Left            =   10065
            TabIndex        =   156
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   6150
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   50
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   155
            Top             =   6120
            Width           =   4695
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   28
            Left            =   3180
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   153
            Top             =   6120
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   26
            Left            =   1425
            MaxLength       =   2
            TabIndex        =   151
            Top             =   6120
            Width           =   600
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "新发肿瘤(&Q)"
            Height          =   195
            Index           =   5
            Left            =   2040
            TabIndex        =   144
            Top             =   5670
            Width           =   1290
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   4
            Left            =   4650
            MaxLength       =   100
            TabIndex        =   141
            Top             =   5280
            Width           =   3135
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "死亡患者尸检(&P)"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   143
            Top             =   5670
            Width           =   1770
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "医院感染作病原学检查(&O)"
            Height          =   195
            Index           =   9
            Left            =   7920
            TabIndex        =   142
            Top             =   5280
            Width           =   2370
         End
         Begin VB.TextBox txtInfo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   35
            Left            =   8070
            MaxLength       =   150
            TabIndex        =   148
            Top             =   5610
            Width           =   2295
         End
         Begin VB.ComboBox cboinfo 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   53
            ItemData        =   "frmInMedRecEdit_YN.frx":0B9E
            Left            =   8760
            List            =   "frmInMedRecEdit_YN.frx":0BA0
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   4080
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   52
            ItemData        =   "frmInMedRecEdit_YN.frx":0BA2
            Left            =   5010
            List            =   "frmInMedRecEdit_YN.frx":0BA4
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   4080
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   33
            ItemData        =   "frmInMedRecEdit_YN.frx":0BA6
            Left            =   8760
            List            =   "frmInMedRecEdit_YN.frx":0BA8
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   4770
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   35
            ItemData        =   "frmInMedRecEdit_YN.frx":0BAA
            Left            =   5010
            List            =   "frmInMedRecEdit_YN.frx":0BAC
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   4425
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   36
            ItemData        =   "frmInMedRecEdit_YN.frx":0BAE
            Left            =   1410
            List            =   "frmInMedRecEdit_YN.frx":0BB0
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   4425
            Width           =   1470
         End
         Begin VB.Frame fraInput 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   5355
            TabIndex        =   338
            Top             =   105
            Width           =   4800
            Begin VB.OptionButton optInput 
               Caption         =   "根据诊断标准输入(&1)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   0
               Left            =   600
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "根据疾病编码输入(&2)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   1
               Left            =   2670
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   0
               Width           =   2010
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
            Height          =   3225
            Left            =   45
            TabIndex        =   114
            Top             =   360
            Width           =   10320
            _cx             =   18203
            _cy             =   5689
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit_YN.frx":0BB2
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
         Begin MSMask.MaskEdBox txt死亡时间 
            Height          =   300
            Left            =   1425
            TabIndex        =   139
            Top             =   5280
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   27
            Left            =   8420
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1785
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   13
            X1              =   0
            X2              =   10275
            Y1              =   6015
            Y2              =   6015
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   12
            X1              =   0
            X2              =   10275
            Y1              =   6000
            Y2              =   6000
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与入院(&I)"
            Height          =   180
            Index           =   115
            Left            =   7560
            TabIndex        =   130
            Top             =   4485
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "术前与术后(&J)"
            Height          =   180
            Index           =   70
            Left            =   240
            TabIndex        =   132
            Top             =   4830
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主要诊断确认日期(&C)"
            Height          =   180
            Index           =   37
            Left            =   6690
            TabIndex        =   118
            Top             =   3780
            Width           =   1710
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院情况(&A)"
            Height          =   180
            Index           =   28
            Left            =   420
            TabIndex        =   115
            Top             =   3780
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病理号(&D)"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   114
            Left            =   600
            TabIndex        =   120
            Top             =   4125
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抢救原因(&V)"
            Height          =   180
            Index           =   102
            Left            =   4560
            TabIndex        =   154
            Top             =   6180
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "成功次数(&U)"
            Height          =   180
            Index           =   36
            Left            =   2145
            TabIndex        =   152
            Top             =   6180
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抢救次数(&T)"
            Height          =   180
            Index           =   10
            Left            =   405
            TabIndex        =   150
            Top             =   6180
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "死亡时间(&M)"
            Height          =   180
            Left            =   405
            TabIndex        =   138
            Top             =   5310
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "死亡原因(&N)"
            Height          =   180
            Index           =   69
            Left            =   3630
            TabIndex        =   140
            Top             =   5325
            Width           =   990
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医院感染病原学诊断(&S)"
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   61
            Left            =   6150
            TabIndex        =   147
            Top             =   5670
            Width           =   1890
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最高诊断依据(&F)"
            Height          =   180
            Index           =   104
            Left            =   7380
            TabIndex        =   124
            Top             =   4140
            Width           =   1350
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分化程度(&E)"
            Height          =   180
            Index           =   103
            Left            =   3960
            TabIndex        =   122
            Top             =   4140
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   6
            X1              =   45
            X2              =   10320
            Y1              =   5145
            Y2              =   5145
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   7
            X1              =   45
            X2              =   10320
            Y1              =   5160
            Y2              =   5160
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "临床与尸检(&R)"
            Height          =   180
            Index           =   71
            Left            =   3450
            TabIndex        =   145
            Top             =   5670
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "临床与病理(&L)"
            Height          =   180
            Index           =   72
            Left            =   7545
            TabIndex        =   136
            Top             =   4830
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "放射与病理(&K)"
            Height          =   180
            Index           =   73
            Left            =   3810
            TabIndex        =   134
            Top             =   4830
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院与出院(&H)"
            Height          =   180
            Index           =   74
            Left            =   3810
            TabIndex        =   128
            Top             =   4485
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与出院(&G)"
            Height          =   180
            Index           =   75
            Left            =   225
            TabIndex        =   126
            Top             =   4485
            Width           =   1170
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   3
         Left            =   120
         TabIndex        =   335
         Top             =   420
         Width           =   10395
         Begin VB.CommandButton cmdAutoLoad 
            Caption         =   "自动提取"
            Height          =   350
            Index           =   1
            Left            =   9160
            TabIndex        =   384
            Top             =   3215
            Width           =   1100
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "发生术后猝死"
            Height          =   195
            Index           =   23
            Left            =   4440
            TabIndex        =   193
            Top             =   6240
            Width           =   2010
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "发生围术期死亡"
            Height          =   195
            Index           =   22
            Left            =   2280
            TabIndex        =   192
            Top             =   6240
            Width           =   2130
         End
         Begin VB.Frame fraInput 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   165
            TabIndex        =   336
            Top             =   3300
            Width           =   6360
            Begin VB.CheckBox chkInfo 
               Caption         =   "未找到时允许自由录入"
               ForeColor       =   &H00004000&
               Height          =   195
               Index           =   19
               Left            =   4200
               TabIndex        =   354
               Top             =   0
               Width           =   2145
            End
            Begin VB.OptionButton optInput 
               Caption         =   "根据ICD9-CM3输入(&4)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   5
               Left            =   2070
               TabIndex        =   190
               TabStop         =   0   'False
               Top             =   0
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "根据诊疗项目输入(&3)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   189
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2010
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsOPS 
            Height          =   2520
            Left            =   165
            TabIndex        =   191
            Top             =   3615
            Width           =   10095
            _cx             =   17806
            _cy             =   4445
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   33
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit_YN.frx":0D7F
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
         Begin VSFlex8Ctl.VSFlexGrid vsAller 
            Height          =   2850
            Left            =   165
            TabIndex        =   188
            Top             =   285
            Width           =   10095
            _cx             =   17806
            _cy             =   5027
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":11B8
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手术及操作相关情况："
            Height          =   180
            Index           =   125
            Left            =   240
            TabIndex        =   372
            Top             =   6240
            Width           =   1800
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   2
         Left            =   -74880
         TabIndex        =   333
         Top             =   420
         Width           =   10575
         Begin VB.Frame fraSub 
            Caption         =   " 准确度 "
            Height          =   1635
            Index           =   1
            Left            =   1785
            TabIndex        =   168
            Top             =   4620
            Width           =   2415
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   2
               ItemData        =   "frmInMedRecEdit_YN.frx":1226
               Left            =   825
               List            =   "frmInMedRecEdit_YN.frx":1228
               Style           =   2  'Dropdown List
               TabIndex        =   170
               Top             =   270
               Width           =   1455
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   11
               ItemData        =   "frmInMedRecEdit_YN.frx":122A
               Left            =   825
               List            =   "frmInMedRecEdit_YN.frx":122C
               Style           =   2  'Dropdown List
               TabIndex        =   172
               Top             =   720
               Width           =   1455
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   12
               ItemData        =   "frmInMedRecEdit_YN.frx":122E
               Left            =   825
               List            =   "frmInMedRecEdit_YN.frx":1230
               Style           =   2  'Dropdown List
               TabIndex        =   174
               Top             =   1140
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "辨证(&E)"
               Height          =   180
               Index           =   38
               Left            =   165
               TabIndex        =   169
               Top             =   330
               Width           =   630
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "治法(&F)"
               Height          =   180
               Index           =   39
               Left            =   165
               TabIndex        =   171
               Top             =   765
               Width           =   630
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "方药(&G)"
               Height          =   180
               Index           =   40
               Left            =   165
               TabIndex        =   173
               Top             =   1200
               Width           =   630
            End
         End
         Begin VB.Frame fraSub 
            Caption         =   " 住院期间病情 "
            Height          =   1635
            Index           =   0
            Left            =   180
            TabIndex        =   164
            Top             =   4620
            Width           =   1500
            Begin VB.CheckBox chkInfo 
               Caption         =   "危重(&A)"
               Height          =   195
               Index           =   2
               Left            =   405
               TabIndex        =   165
               Top             =   345
               Width           =   930
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "急症(&B)"
               Height          =   195
               Index           =   3
               Left            =   405
               TabIndex        =   166
               Top             =   765
               Width           =   930
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "疑难(&D)"
               Height          =   195
               Index           =   4
               Left            =   405
               TabIndex        =   167
               Top             =   1185
               Width           =   930
            End
         End
         Begin VB.Frame fraSub 
            Caption         =   " 治疗方法 "
            Height          =   1635
            Index           =   2
            Left            =   4335
            TabIndex        =   175
            Top             =   4620
            Width           =   6090
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   57
               ItemData        =   "frmInMedRecEdit_YN.frx":1232
               Left            =   4530
               List            =   "frmInMedRecEdit_YN.frx":1234
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   1140
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   56
               ItemData        =   "frmInMedRecEdit_YN.frx":1236
               Left            =   4530
               List            =   "frmInMedRecEdit_YN.frx":1238
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   705
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   55
               ItemData        =   "frmInMedRecEdit_YN.frx":123A
               Left            =   4530
               List            =   "frmInMedRecEdit_YN.frx":123C
               Style           =   2  'Dropdown List
               TabIndex        =   183
               Top             =   240
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   13
               ItemData        =   "frmInMedRecEdit_YN.frx":123E
               Left            =   1575
               List            =   "frmInMedRecEdit_YN.frx":1240
               Style           =   2  'Dropdown List
               TabIndex        =   181
               Top             =   1140
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   14
               ItemData        =   "frmInMedRecEdit_YN.frx":1242
               Left            =   1215
               List            =   "frmInMedRecEdit_YN.frx":1244
               Style           =   2  'Dropdown List
               TabIndex        =   179
               Top             =   705
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   15
               ItemData        =   "frmInMedRecEdit_YN.frx":1246
               Left            =   1215
               List            =   "frmInMedRecEdit_YN.frx":1248
               Style           =   2  'Dropdown List
               TabIndex        =   177
               Top             =   270
               Width           =   1410
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "辨证施护(&P)"
               Height          =   180
               Index           =   113
               Left            =   3480
               TabIndex        =   186
               Top             =   1200
               Width           =   990
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "使用中医诊疗技术(&O)"
               Height          =   180
               Index           =   112
               Left            =   2760
               TabIndex        =   184
               Top             =   765
               Width           =   1710
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "使用中医诊疗设备(&N)"
               Height          =   180
               Index           =   111
               Left            =   2760
               TabIndex        =   182
               Top             =   300
               Width           =   1710
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "自制中药制剂(&K)"
               Height          =   180
               Index           =   41
               Left            =   165
               TabIndex        =   180
               Top             =   1200
               Width           =   1350
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抢救方法(&J)"
               Height          =   180
               Index           =   42
               Left            =   165
               TabIndex        =   178
               Top             =   765
               Width           =   990
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "治疗类别(&I)"
               Height          =   180
               Index           =   43
               Left            =   165
               TabIndex        =   176
               Top             =   330
               Width           =   990
            End
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   37
            ItemData        =   "frmInMedRecEdit_YN.frx":124A
            Left            =   4335
            List            =   "frmInMedRecEdit_YN.frx":124C
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   4035
            Width           =   1395
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   38
            ItemData        =   "frmInMedRecEdit_YN.frx":124E
            Left            =   1470
            List            =   "frmInMedRecEdit_YN.frx":1250
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   4035
            Width           =   1395
         End
         Begin VB.Frame fraInput 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   5340
            TabIndex        =   334
            Top             =   105
            Width           =   4800
            Begin VB.OptionButton optInput 
               Caption         =   "根据疾病编码输入(&2)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   3
               Left            =   2760
               TabIndex        =   158
               TabStop         =   0   'False
               Top             =   0
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "根据诊断标准输入(&1)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   2
               Left            =   720
               TabIndex        =   157
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2010
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
            Height          =   3555
            Left            =   165
            TabIndex        =   159
            Top             =   360
            Width           =   10320
            _cx             =   18203
            _cy             =   6271
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit_YN.frx":1252
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
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院与出院(&M)"
            Height          =   180
            Index           =   76
            Left            =   3135
            TabIndex        =   162
            Top             =   4095
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与出院(&L)"
            Height          =   180
            Index           =   77
            Left            =   270
            TabIndex        =   160
            Top             =   4095
            Width           =   1170
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6555
         Index           =   4
         Left            =   -74880
         TabIndex        =   332
         Top             =   360
         Width           =   10455
         Begin VB.CheckBox chkInfo 
            Caption         =   "疑难病例(&X)"
            Height          =   195
            Index           =   20
            Left            =   7785
            TabIndex        =   236
            Top             =   1733
            Width           =   1290
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   65
            Left            =   4815
            MaxLength       =   100
            TabIndex        =   227
            Top             =   2400
            Width           =   2880
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   66
            Left            =   3200
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2430
            Width           =   270
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            Index           =   60
            Left            =   8610
            Style           =   2  'Dropdown List
            TabIndex        =   241
            Top             =   2400
            Width           =   1635
         End
         Begin VB.CommandButton cmdInfo 
            Height          =   240
            Index           =   60
            Left            =   9015
            Picture         =   "frmInMedRecEdit_YN.frx":13EC
            Style           =   1  'Graphical
            TabIndex        =   377
            TabStop         =   0   'False
            ToolTipText     =   "选择(F4)"
            Top             =   5850
            Width           =   240
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   59
            ItemData        =   "frmInMedRecEdit_YN.frx":14E2
            Left            =   7860
            List            =   "frmInMedRecEdit_YN.frx":14E4
            Style           =   2  'Dropdown List
            TabIndex        =   280
            Top             =   5454
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   30
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   210
            Top             =   165
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   31
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   213
            Top             =   543
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   32
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   216
            Top             =   921
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   33
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   219
            Top             =   1299
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   34
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   205
            Top             =   2055
            Width           =   1140
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   41
            ItemData        =   "frmInMedRecEdit_YN.frx":14E6
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":14E8
            Style           =   2  'Dropdown List
            TabIndex        =   203
            Top             =   1680
            Width           =   1425
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "示教病案"
            Height          =   195
            Index           =   8
            Left            =   7785
            TabIndex        =   234
            Top             =   1358
            Width           =   1100
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "随诊(&F)"
            Height          =   195
            Index           =   7
            Left            =   3930
            TabIndex        =   260
            Top             =   4200
            Width           =   930
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            ItemData        =   "frmInMedRecEdit_YN.frx":14EA
            Left            =   7320
            List            =   "frmInMedRecEdit_YN.frx":14EC
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   263
            TabStop         =   0   'False
            Top             =   4140
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   29
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   262
            TabStop         =   0   'False
            Top             =   4140
            Width           =   1020
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   27
            ItemData        =   "frmInMedRecEdit_YN.frx":14EE
            Left            =   8610
            List            =   "frmInMedRecEdit_YN.frx":14F0
            Style           =   2  'Dropdown List
            TabIndex        =   233
            Top             =   921
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   28
            ItemData        =   "frmInMedRecEdit_YN.frx":14F2
            Left            =   8610
            List            =   "frmInMedRecEdit_YN.frx":14F4
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   543
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   29
            ItemData        =   "frmInMedRecEdit_YN.frx":14F6
            Left            =   8610
            List            =   "frmInMedRecEdit_YN.frx":14F8
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   165
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   51
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   222
            Top             =   1680
            Width           =   1140
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   42
            Left            =   4815
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   2055
            Width           =   1170
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   44
            ItemData        =   "frmInMedRecEdit_YN.frx":14FA
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":14FC
            Style           =   2  'Dropdown List
            TabIndex        =   243
            Top             =   2925
            Width           =   1500
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   40
            Left            =   3930
            MaxLength       =   100
            TabIndex        =   245
            Top             =   2925
            Width           =   5055
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   41
            Left            =   3930
            MaxLength       =   100
            TabIndex        =   249
            Top             =   3330
            Width           =   5055
         End
         Begin VB.OptionButton optInput 
            Caption         =   "无"
            Height          =   255
            Index           =   6
            Left            =   2130
            TabIndex        =   247
            Top             =   3360
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optInput 
            Caption         =   "有，目的："
            Height          =   255
            Index           =   7
            Left            =   2700
            TabIndex        =   248
            Top             =   3360
            Width           =   1195
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   54
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   246
            Top             =   3330
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   56
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   251
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   55
            Left            =   6960
            MaxLength       =   4
            TabIndex        =   255
            Top             =   3735
            Width           =   675
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            ItemData        =   "frmInMedRecEdit_YN.frx":14FE
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1500
            TabIndex        =   265
            Top             =   4700
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   30
            ItemData        =   "frmInMedRecEdit_YN.frx":1502
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1504
            Style           =   2  'Dropdown List
            TabIndex        =   201
            Top             =   1299
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            ItemData        =   "frmInMedRecEdit_YN.frx":1506
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1508
            Style           =   2  'Dropdown List
            TabIndex        =   199
            Top             =   921
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   43
            ItemData        =   "frmInMedRecEdit_YN.frx":150A
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":150C
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   165
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   49
            Left            =   1260
            MaxLength       =   4
            TabIndex        =   259
            Top             =   4140
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   48
            Left            =   9060
            MaxLength       =   3
            TabIndex        =   257
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   47
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   256
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   46
            Left            =   5100
            MaxLength       =   3
            TabIndex        =   253
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   45
            Left            =   3930
            MaxLength       =   4
            TabIndex        =   252
            Top             =   3735
            Width           =   675
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   50
            ItemData        =   "frmInMedRecEdit_YN.frx":150E
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1510
            TabIndex        =   288
            Top             =   6210
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   17
            ItemData        =   "frmInMedRecEdit_YN.frx":1512
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1514
            Style           =   2  'Dropdown List
            TabIndex        =   197
            Top             =   543
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   20
            ItemData        =   "frmInMedRecEdit_YN.frx":1516
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1518
            TabIndex        =   267
            Top             =   4700
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            ItemData        =   "frmInMedRecEdit_YN.frx":151A
            Left            =   7860
            List            =   "frmInMedRecEdit_YN.frx":151C
            TabIndex        =   268
            Top             =   4700
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            ItemData        =   "frmInMedRecEdit_YN.frx":151E
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1520
            TabIndex        =   272
            Top             =   5077
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   23
            ItemData        =   "frmInMedRecEdit_YN.frx":1522
            Left            =   7860
            List            =   "frmInMedRecEdit_YN.frx":1524
            TabIndex        =   274
            Top             =   5077
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   24
            ItemData        =   "frmInMedRecEdit_YN.frx":1526
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1528
            TabIndex        =   270
            Top             =   5077
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   25
            ItemData        =   "frmInMedRecEdit_YN.frx":152A
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":152C
            TabIndex        =   276
            Top             =   5454
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   26
            ItemData        =   "frmInMedRecEdit_YN.frx":152E
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1530
            TabIndex        =   278
            Top             =   5454
            Width           =   1425
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "科研病案"
            Height          =   195
            Index           =   11
            Left            =   9120
            TabIndex        =   235
            Top             =   1358
            Width           =   1100
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   5475
            TabIndex        =   340
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "取消"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6030
            TabIndex        =   341
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   5475
            TabIndex        =   347
            Top             =   5070
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "取消"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   6030
            TabIndex        =   346
            Top             =   5070
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "取消"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   9870
            TabIndex        =   343
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   9315
            TabIndex        =   342
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "取消"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   9870
            TabIndex        =   344
            Top             =   5070
            Width           =   555
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   9315
            TabIndex        =   345
            Top             =   5070
            Width           =   555
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   39
            ItemData        =   "frmInMedRecEdit_YN.frx":1532
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1534
            TabIndex        =   284
            Top             =   5831
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   40
            ItemData        =   "frmInMedRecEdit_YN.frx":1536
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1538
            TabIndex        =   282
            Top             =   5831
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   60
            Left            =   7860
            MaxLength       =   16
            TabIndex        =   285
            Top             =   5820
            Width           =   1425
         End
         Begin MSMask.MaskEdBox txt发病日期 
            Height          =   300
            Left            =   8610
            TabIndex        =   238
            Top             =   2040
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt发病时间 
            Height          =   300
            Left            =   9690
            TabIndex        =   239
            Top             =   2040
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   66
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   207
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label lblInfo 
            Caption         =   "医学警示"
            Height          =   180
            Index           =   129
            Left            =   535
            TabIndex        =   206
            Top             =   2460
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Caption         =   "其他医学警示"
            Height          =   180
            Index           =   56
            Left            =   3660
            TabIndex        =   226
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主任(副主任)    医师(&3)"
            Height          =   360
            Index           =   21
            Left            =   6750
            TabIndex        =   379
            Top             =   4680
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发病时间"
            Height          =   180
            Index           =   21
            Left            =   7800
            TabIndex        =   237
            Top             =   2100
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "生育状况"
            Height          =   180
            Index           =   29
            Left            =   7785
            TabIndex        =   240
            Top             =   2460
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "质控日期(&X)"
            Height          =   180
            Index           =   13
            Left            =   6840
            TabIndex        =   286
            Top             =   5880
            Width           =   990
         End
         Begin VB.Label lbl编码 
            AutoSize        =   -1  'True
            Caption         =   "病案质量(&Y)"
            Height          =   180
            Index           =   8
            Left            =   6840
            TabIndex        =   279
            Top             =   5505
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   15
            X1              =   120
            X2              =   10440
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   14
            X1              =   120
            X2              =   10440
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输红细胞(&L)"
            Height          =   180
            Index           =   47
            Left            =   3750
            TabIndex        =   209
            Top             =   225
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位"
            Height          =   180
            Index           =   48
            Left            =   6045
            TabIndex        =   211
            Top             =   225
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血小板(&M)"
            Height          =   180
            Index           =   49
            Left            =   3750
            TabIndex        =   212
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位"
            Height          =   180
            Index           =   50
            Left            =   6045
            TabIndex        =   214
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血浆(&N)"
            Height          =   180
            Index           =   51
            Left            =   3930
            TabIndex        =   215
            Top             =   975
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输全血(&O)"
            Height          =   180
            Index           =   53
            Left            =   3930
            TabIndex        =   218
            Top             =   1365
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   54
            Left            =   6045
            TabIndex        =   220
            Top             =   1365
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输其他(&Q)"
            Height          =   180
            Index           =   55
            Left            =   445
            TabIndex        =   204
            Top             =   2115
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血反应(&K)"
            Height          =   180
            Index           =   60
            Left            =   265
            TabIndex        =   202
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "随诊期限(&G)"
            Height          =   180
            Index           =   44
            Left            =   5265
            TabIndex        =   261
            Top             =   4200
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H&IV-Ab"
            Height          =   180
            Index           =   65
            Left            =   7965
            TabIndex        =   232
            Top             =   981
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HC&V-Ab"
            Height          =   180
            Index           =   66
            Left            =   7965
            TabIndex        =   230
            Top             =   603
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HB&sAg"
            Height          =   180
            Index           =   67
            Left            =   8055
            TabIndex        =   228
            Top             =   225
            Width           =   450
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   105
            Left            =   6045
            TabIndex        =   223
            Top             =   1740
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自体回收(&B)"
            Height          =   180
            Index           =   106
            Left            =   3750
            TabIndex        =   221
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   110
            Left            =   6045
            TabIndex        =   217
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血前的9项检查(&E)"
            Height          =   180
            Index           =   81
            Left            =   3120
            TabIndex        =   224
            Top             =   2115
            Width           =   1620
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出院方式"
            Height          =   180
            Index           =   87
            Left            =   480
            TabIndex        =   242
            Top             =   2985
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "转入"
            Height          =   180
            Index           =   88
            Left            =   3525
            TabIndex        =   244
            Top             =   2985
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊医师(&1)"
            Height          =   180
            Index           =   57
            Left            =   240
            TabIndex        =   264
            Top             =   4760
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病例分型(&D)"
            Height          =   180
            Index           =   86
            Left            =   270
            TabIndex        =   194
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "  入院后        天         小时        分钟"
            Height          =   180
            Index           =   101
            Left            =   6225
            TabIndex        =   254
            Top             =   3795
            Width           =   3870
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "呼吸机使用(&C)         小时"
            Height          =   180
            Index           =   100
            Left            =   90
            TabIndex        =   258
            Top             =   4200
            Width           =   2340
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "颅脑损伤患者昏迷时间(&P) 入院前        天         小时        分钟"
            Height          =   180
            Index           =   99
            Left            =   285
            TabIndex        =   250
            Top             =   3795
            Width           =   5850
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "责任护士(&A)"
            Height          =   180
            Index           =   95
            Left            =   240
            TabIndex        =   287
            Top             =   6270
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "血型(&J)"
            Height          =   180
            Index           =   45
            Left            =   625
            TabIndex        =   196
            Top             =   603
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Rh"
            Height          =   180
            Index           =   46
            Left            =   960
            TabIndex        =   198
            Top             =   975
            Width           =   180
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   10
            X1              =   120
            X2              =   10440
            Y1              =   4575
            Y2              =   4575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   11
            X1              =   120
            X2              =   10440
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科主任(&2)"
            Height          =   180
            Index           =   20
            Left            =   3180
            TabIndex        =   266
            Top             =   4755
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主治医师(&5)"
            Height          =   180
            Index           =   22
            Left            =   3000
            TabIndex        =   271
            Top             =   5137
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院医师(&6)"
            Height          =   180
            Index           =   23
            Left            =   6840
            TabIndex        =   273
            Top             =   5137
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "进修医师(&4)"
            Height          =   180
            Index           =   62
            Left            =   240
            TabIndex        =   269
            Top             =   5137
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "研究生医师(&7)"
            Height          =   180
            Index           =   63
            Left            =   60
            TabIndex        =   275
            Top             =   5514
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "实习医师(&8)"
            Height          =   180
            Index           =   64
            Left            =   3000
            TabIndex        =   277
            Top             =   5514
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输液反应(&S)"
            Height          =   180
            Index           =   68
            Left            =   265
            TabIndex        =   200
            Top             =   1359
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "质控护士(&0)"
            Height          =   180
            Index           =   58
            Left            =   3000
            TabIndex        =   283
            Top             =   5891
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "质控医师(&9)"
            Height          =   180
            Index           =   59
            Left            =   240
            TabIndex        =   281
            Top             =   5891
            Width           =   990
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   0
         Left            =   -74880
         TabIndex        =   330
         Top             =   420
         Width           =   10545
         Begin ZlPatiAddress.PatiAddress PatiAddress籍贯 
            Height          =   360
            Left            =   7755
            TabIndex        =   45
            Top             =   2040
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Items           =   2
            MaxLength       =   50
         End
         Begin ZlPatiAddress.PatiAddress PatiAddress出生地 
            Height          =   360
            Left            =   1290
            TabIndex        =   41
            Top             =   2070
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Items           =   3
            MaxLength       =   50
         End
         Begin ZlPatiAddress.PatiAddress PatiAddress户口地址 
            Height          =   360
            Left            =   1290
            TabIndex        =   61
            Top             =   3240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   50
         End
         Begin ZlPatiAddress.PatiAddress PatiAddress现住址 
            Height          =   360
            Left            =   1290
            TabIndex        =   53
            Top             =   2880
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   50
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   53
            Left            =   5055
            TabIndex        =   63
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3270
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   54
            Left            =   7215
            MaxLength       =   6
            TabIndex        =   65
            Top             =   3240
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   53
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   62
            Top             =   3240
            Width           =   4035
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   52
            Left            =   9990
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2070
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   52
            Left            =   7785
            MaxLength       =   30
            TabIndex        =   46
            Top             =   2040
            Width           =   2490
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "入院前经外院治疗"
            Height          =   195
            Index           =   15
            Left            =   3645
            TabIndex        =   93
            Top             =   5228
            Width           =   2010
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   44
            Left            =   8535
            MaxLength       =   8
            TabIndex        =   32
            Top             =   1350
            Width           =   1710
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   43
            Left            =   4410
            MaxLength       =   8
            TabIndex        =   29
            Top             =   1350
            Width           =   1755
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            Index           =   51
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1357
            Width           =   645
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   42
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   26
            Top             =   1357
            Width           =   1080
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   23
            Left            =   5055
            TabIndex        =   98
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   5550
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   6
            Left            =   6360
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2085
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   15
            Left            =   5055
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   4380
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   10
            Left            =   5055
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3645
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   1
            Left            =   5055
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2910
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   24
            Left            =   7305
            TabIndex        =   101
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   5550
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   25
            Left            =   10005
            TabIndex        =   103
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   5550
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   36
            Left            =   9405
            TabIndex        =   87
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   4380
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   135
            Width           =   1740
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   36
            Left            =   7215
            MaxLength       =   30
            TabIndex        =   84
            Top             =   4350
            Width           =   2490
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            ItemData        =   "frmInMedRecEdit_YN.frx":153A
            Left            =   7800
            List            =   "frmInMedRecEdit_YN.frx":153C
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   135
            Width           =   2475
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   3390
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   135
            Width           =   375
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   3
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   615
            Width           =   1740
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            ItemData        =   "frmInMedRecEdit_YN.frx":153E
            Left            =   4410
            List            =   "frmInMedRecEdit_YN.frx":1540
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   615
            Width           =   1605
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   5
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   15
            Top             =   990
            Width           =   1080
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            ItemData        =   "frmInMedRecEdit_YN.frx":1542
            Left            =   8535
            List            =   "frmInMedRecEdit_YN.frx":1544
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   990
            Width           =   1740
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            ItemData        =   "frmInMedRecEdit_YN.frx":1546
            Left            =   7785
            List            =   "frmInMedRecEdit_YN.frx":1548
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1725
            Width           =   2490
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1725
            Width           =   1740
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            ItemData        =   "frmInMedRecEdit_YN.frx":154A
            Left            =   4410
            List            =   "frmInMedRecEdit_YN.frx":154C
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1725
            Width           =   2250
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   6
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   42
            Top             =   2055
            Width           =   5295
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   1
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   54
            Top             =   2880
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   7215
            MaxLength       =   20
            TabIndex        =   57
            Top             =   2880
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   9
            Left            =   9330
            MaxLength       =   6
            TabIndex        =   59
            Top             =   2880
            Width           =   945
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   10
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   67
            Top             =   3615
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   7215
            MaxLength       =   20
            TabIndex        =   70
            Top             =   3615
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            Left            =   9330
            MaxLength       =   6
            TabIndex        =   72
            Top             =   3615
            Width           =   945
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   13
            Left            =   1320
            MaxLength       =   64
            TabIndex        =   74
            Top             =   3990
            Width           =   1545
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   9
            ItemData        =   "frmInMedRecEdit_YN.frx":154E
            Left            =   3645
            List            =   "frmInMedRecEdit_YN.frx":1550
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   3990
            Width           =   1700
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   7215
            MaxLength       =   20
            TabIndex        =   78
            Top             =   3990
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   15
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   80
            Top             =   4350
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   4815
            Width           =   1545
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   17
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   4815
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   18
            Left            =   6225
            MaxLength       =   100
            TabIndex        =   90
            Top             =   4815
            Width           =   1305
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   5865
            Width           =   1545
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   20
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   5865
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   21
            Left            =   6225
            MaxLength       =   100
            TabIndex        =   109
            Top             =   5865
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   8610
            Locked          =   -1  'True
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   5865
            Width           =   1665
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   23
            Left            =   3645
            MaxLength       =   100
            TabIndex        =   97
            Top             =   5520
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   24
            Left            =   5640
            MaxLength       =   100
            TabIndex        =   100
            Top             =   5520
            Width           =   1965
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   25
            Left            =   7830
            MaxLength       =   100
            TabIndex        =   102
            Top             =   5520
            Width           =   2445
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            Index           =   10
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   990
            Width           =   645
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "再入院"
            Height          =   285
            Index           =   10
            Left            =   4515
            TabIndex        =   4
            Top             =   143
            Width           =   840
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmInMedRecEdit_YN.frx":1552
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1554
            TabIndex        =   49
            Top             =   2415
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   5520
            Width           =   1545
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   37
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   51
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   39
            Left            =   6075
            MaxLength       =   5
            TabIndex        =   21
            Top             =   990
            Width           =   555
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   38
            Left            =   4410
            MaxLength       =   5
            TabIndex        =   18
            Top             =   990
            Width           =   555
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   49
            ItemData        =   "frmInMedRecEdit_YN.frx":1556
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1558
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   5175
            Width           =   1545
         End
         Begin MSMask.MaskEdBox txt出生时间 
            Height          =   300
            Left            =   9660
            TabIndex        =   13
            Top             =   615
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   300
            Left            =   8535
            TabIndex        =   12
            Top             =   615
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "邮编(&U)"
            Height          =   180
            Index           =   109
            Left            =   6525
            TabIndex        =   64
            Top             =   3300
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口地址(&T)"
            Height          =   180
            Index           =   108
            Left            =   300
            TabIndex        =   60
            Top             =   3300
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "籍贯(&M)"
            Height          =   180
            Index           =   93
            Left            =   7125
            TabIndex        =   44
            Top             =   2100
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "克"
            Height          =   180
            Index           =   2
            Left            =   10320
            TabIndex        =   33
            Top             =   1410
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "新生儿入院体重"
            Height          =   180
            Index           =   98
            Left            =   7230
            TabIndex        =   31
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "克"
            Height          =   180
            Index           =   1
            Left            =   6240
            TabIndex        =   30
            Top             =   1410
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "新生儿出生体重"
            Height          =   180
            Index           =   97
            Left            =   3105
            TabIndex        =   28
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婴幼儿年龄(&O)"
            Height          =   180
            Index           =   96
            Left            =   120
            TabIndex        =   25
            Top             =   1417
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号(&A)"
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   0
            Top             =   195
            Width           =   810
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第     次住院"
            Height          =   180
            Index           =   3
            Left            =   3180
            TabIndex        =   2
            Top             =   195
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付费方式(&B)"
            Height          =   180
            Index           =   2
            Left            =   6765
            TabIndex        =   5
            Top             =   195
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名(&D)"
            Height          =   180
            Index           =   4
            Left            =   660
            TabIndex        =   7
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别(&E)"
            Height          =   180
            Index           =   5
            Left            =   3735
            TabIndex        =   9
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期(&F)"
            Height          =   180
            Index           =   6
            Left            =   7485
            TabIndex        =   11
            Top             =   675
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄(&G)"
            Height          =   180
            Index           =   7
            Left            =   660
            TabIndex        =   14
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻(&I)"
            Height          =   180
            Index           =   8
            Left            =   7845
            TabIndex        =   23
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "职业(&J)"
            Height          =   180
            Index           =   9
            Left            =   7125
            TabIndex        =   38
            Top             =   1785
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "区域(&3)"
            Height          =   180
            Index           =   11
            Left            =   6525
            TabIndex        =   83
            Top             =   4410
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "国籍(&K)"
            Height          =   180
            Index           =   12
            Left            =   660
            TabIndex        =   34
            Top             =   1785
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "民族(&L)"
            Height          =   180
            Index           =   13
            Left            =   3735
            TabIndex        =   36
            Top             =   1785
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生地点(&N)"
            Height          =   180
            Index           =   14
            Left            =   300
            TabIndex        =   40
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证号(&P)"
            Height          =   180
            Index           =   15
            Left            =   300
            TabIndex        =   48
            Top             =   2475
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   0
            X1              =   135
            X2              =   10320
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   135
            X2              =   10320
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   135
            X2              =   10320
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   135
            X2              =   10320
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "现住址(&Q)"
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   52
            Top             =   2940
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话(&R)"
            Height          =   180
            Index           =   16
            Left            =   6525
            TabIndex        =   56
            Top             =   2940
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "邮编(&S)"
            Height          =   180
            Index           =   17
            Left            =   8670
            TabIndex        =   58
            Top             =   2940
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "工作单位(&V)"
            Height          =   180
            Index           =   18
            Left            =   300
            TabIndex        =   66
            Top             =   3675
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话(&W)"
            Height          =   180
            Index           =   19
            Left            =   6525
            TabIndex        =   69
            Top             =   3675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "邮编(&X)"
            Height          =   180
            Index           =   123
            Left            =   8670
            TabIndex        =   71
            Top             =   3675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人姓名(&Y)"
            Height          =   180
            Index           =   79
            Left            =   120
            TabIndex        =   73
            Top             =   4050
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "关系(&Z)"
            Height          =   180
            Index           =   78
            Left            =   2985
            TabIndex        =   75
            Top             =   4050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话(&1)"
            Height          =   180
            Index           =   80
            Left            =   6525
            TabIndex        =   77
            Top             =   4050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人地址(&2)"
            Height          =   180
            Index           =   24
            Left            =   120
            TabIndex        =   79
            Top             =   4410
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   4
            X1              =   135
            X2              =   10080
            Y1              =   4740
            Y2              =   4740
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   135
            X2              =   10320
            Y1              =   4725
            Y2              =   4725
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院时间"
            Height          =   180
            Index           =   25
            Left            =   570
            TabIndex        =   82
            Top             =   4875
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科室"
            Height          =   180
            Index           =   26
            Left            =   3255
            TabIndex        =   86
            Top             =   4875
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病房"
            Height          =   180
            Index           =   27
            Left            =   5835
            TabIndex        =   89
            Top             =   4875
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出院时间"
            Height          =   180
            Index           =   29
            Left            =   570
            TabIndex        =   104
            Top             =   5925
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科室"
            Height          =   180
            Index           =   30
            Left            =   3255
            TabIndex        =   106
            Top             =   5925
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病房"
            Height          =   180
            Index           =   31
            Left            =   5835
            TabIndex        =   108
            Top             =   5925
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "转科"
            Height          =   180
            Index           =   33
            Left            =   3255
            TabIndex        =   96
            Top             =   5580
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "→"
            Height          =   180
            Index           =   34
            Left            =   5415
            TabIndex        =   99
            Top             =   5580
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "→"
            Height          =   180
            Index           =   35
            Left            =   7635
            TabIndex        =   331
            Top             =   5580
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院天数"
            Height          =   180
            Index           =   32
            Left            =   7860
            TabIndex        =   110
            Top             =   5925
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入科时间"
            Height          =   180
            Index           =   84
            Left            =   570
            TabIndex        =   94
            Top             =   5580
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "其他证件"
            Height          =   180
            Index           =   85
            Left            =   6195
            TabIndex        =   50
            Top             =   2460
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "kg"
            Height          =   180
            Index           =   0
            Left            =   6675
            TabIndex        =   22
            Top             =   1050
            Width           =   180
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体重(&W)"
            Height          =   180
            Index           =   24
            Left            =   5400
            TabIndex        =   20
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            Height          =   180
            Index           =   1
            Left            =   5040
            TabIndex        =   19
            Top             =   1050
            Width           =   180
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身高(&H)"
            Height          =   180
            Index           =   23
            Left            =   3735
            TabIndex        =   17
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院途径"
            Height          =   180
            Index           =   94
            Left            =   555
            TabIndex        =   91
            Top             =   5235
            Width           =   720
         End
      End
   End
   Begin VB.Timer timThis 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4920
      Top             =   6840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9840
      TabIndex        =   329
      Top             =   7155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8595
      TabIndex        =   327
      Top             =   7155
      Width           =   1100
   End
   Begin MSComCtl2.MonthView dtpInfo 
      Height          =   2160
      Left            =   2760
      TabIndex        =   378
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   115802113
      TitleBackColor  =   8421504
      TitleForeColor  =   16777215
      CurrentDate     =   38003
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   3480
      Picture         =   "frmInMedRecEdit_YN.frx":155A
      Top             =   7320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   4320
      Picture         =   "frmInMedRecEdit_YN.frx":7DAC
      Top             =   7320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu menuPriview 
      Caption         =   "预览首页"
      Visible         =   0   'False
      Begin VB.Menu menuPage 
         Caption         =   "正面(&1)"
         Index           =   1
      End
      Begin VB.Menu menuPage 
         Caption         =   "反面(&2)"
         Index           =   2
      End
      Begin VB.Menu menuPage 
         Caption         =   "附页1(&3)"
         Index           =   3
      End
      Begin VB.Menu menuPage 
         Caption         =   "附页2(&4)"
         Index           =   4
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   "打印首页"
      Visible         =   0   'False
      Begin VB.Menu menuPagePrint 
         Caption         =   "正面(&1)"
         Index           =   1
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "反面(&2)"
         Index           =   2
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "附页1(&3)"
         Index           =   3
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "附页2(&4)"
         Index           =   4
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "正面+附页1(&5)"
         Index           =   5
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "反面+附页2(&6)"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmInMedRecEdit_YN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Closed(ByVal EditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String) '住院首页关闭事件

Private mcol人员SQL As Collection
Private mblnReadOnly As Boolean
Private mstrPrivs As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mbln出院 As Boolean
Private mint险类 As Integer
Private mlngPathState As Long   '路径状态   -1=未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
Private mlngDiagnosisType As Long '导入路径时的诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
Private mstr疾病ID As String   '用于保存疾病ID,在Closed事件中传递给父窗体
Private mstr诊断ID As String   '用于保存诊断ID,在Closed事件中传递给父窗体
Private mstrPathDiag As String '病人一次住院第二条路径以后导入诊断列表：诊断类型1|疾病ID1|诊断ID1,诊断类型2|疾病ID2|诊断ID2```
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mbln护士站 As Boolean

Private mblnIsFirst As Boolean
Private mstrZYDiagInfo As String
Private mstrXYDiagInfo As String
Public mblnDiagChange As Boolean
Private mlng区域 As String      '区域是否检查  0-不检查，2-提示，1-必须填写
Private mlng损伤中毒 As Long
Private mlng病理诊断 As Long
Private mlngSize As Long '记录病案主页从表“信息值”字段长度
Private mstr手术输入情况 As String
Private mobjESign As Object           '签名部件对象
Private mblnIsPathOutTime As Boolean   '完成路径的时间是否比出院诊断记录时间大
Private mlngDateIndex As Long


Private mstr类型 As String
Private mblnDiagnose As Boolean

Private mblnOpen As Boolean
Private mbln中医 As Boolean

Private mbln病案共享 As Boolean

Private mstrLike As String
Private mint简码 As Integer
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mblnNoClick As Boolean
Private mblnReturn As Boolean
Private mstrDelete As String
Private mlngNum As Long
Private mlngSelNum As Long
Private mlngNumBack As Long
Private mbln首页诊断 As Boolean
Private Const GRD_LOSTFOCUS_COLORSEL = &H80000010  '离开焦点时,选择的显示颜色
Private Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '进入控件时,选择显示颜色
Private Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private mbln启用结构化地址 As Boolean
Private mbln不使用西医项目 As Boolean
Private mbln医生护士分填首页 As Boolean

Private mrsXYDiag  As ADODB.Recordset '西医诊断记录集
Private mrsZYDiag  As ADODB.Recordset '中医诊断记录集

Private Enum Tab菜单
    TAB_基本信息 = 0
    TAB_西医诊断 = 1
    TAB_中医诊断 = 2
    TAB_过敏与手术 = 3
    TAB_住院情况 = 4
    TAB_放疗与化疗 = 5
    TAB_特定药品 = 6
    TAB_其他 = 7
End Enum

Private Enum COL诊断情况
    col诊断类型 = 0
    col诊断编码 = 1
    col诊断描述 = 2
    col中医证候 = 3
    col备注 = 4
    col入院病情 = 5
    col出院情况 = 6
    col是否未治 = 7
    col是否疑诊 = 8
    col增加 = 9
    colDel = 10
    col诊断ID = 11
    col疾病ID = 12
    col类型 = 13 '1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
    
    colzy增加 = 7
    colzyDel = 8
    colzy诊断ID = 9
    colzy疾病ID = 10
    colzy证候ID = 11
    colzy类型 = 12
End Enum
Private Enum COL手术情况
    col手术日期 = 0
    COL手术情况 = 1
    col手术编码 = 2
    col手术名称 = 3
    col再次手术 = 4
    col主刀医师 = 5
    col助产护士 = 6
    col助手1 = 7
    col助手2 = 8
    col麻醉类型 = 9
    colASA分级 = 10
    colNNIS分级 = 11
    col手术级别 = 12
    col麻醉医师 = 13
    col切口愈合 = 14
    col预防用抗菌药 = 15
    col抗菌药天数 = 16
    col非预期的二次手术 = 17
    col麻醉并发症 = 18
    col术中异物遗留 = 19
    col手术并发症 = 20
    col术后出血或血肿 = 21
    col手术伤口裂开 = 22
    col术后深静脉血栓 = 23
    col术后生理代谢紊乱 = 24
    col术后呼吸衰竭 = 25
    col术后肺栓塞 = 26
    col术后败血症 = 27
    col术后髋关节骨折 = 28
    col手术操作ID = 29
    col诊疗项目ID = 30
    col麻醉ID = 31
    col麻醉方式 = 32
End Enum

Private Enum 基本信息
    cbo付款方式 = 0
    cbo性别 = 1
    cbo婚姻 = 3
    cbo职业 = 4
    cbo入院病情 = 5
    cbo身份证号 = 6
    txt区域 = 36
    cbo国籍 = 7
    cbo民族 = 8
    cbo联系人关系 = 9
    cbo年龄单位 = 10
    txt住院号 = 0
    txt家庭地址 = 1
    txt住院次数 = 2
    txt姓名 = 3
    'txt出生日期 = 4
    txt年龄 = 5
    txt出生地点 = 6
    txt家庭电话 = 8
    txt家庭邮编 = 9
    txt单位名称 = 10
    txt单位电话 = 11
    txt单位邮编 = 12
    txt联系人姓名 = 13
    txt联系人电话 = 14
    txt联系人地址 = 15
    txt入院时间 = 16
    txt入院科室 = 17
    txt入院病室 = 18
    txt出院时间 = 19
    txt出院科室 = 20
    txt出院病室 = 21
    txt住院天数 = 22
    txt入科时间 = 7
    txt转科1 = 23
    txt转科2 = 24
    txt转科3 = 25
    txt其他证件 = 37
    txt身高 = 38
    txt体重 = 39
    chk再入院 = 10
    cbo入院方式 = 49
    cbo婴儿年龄单位 = 51
    txt婴儿年龄 = 42
    txt新生儿体重 = 43
    txt新生儿入院体重 = 44
    cbo31天和7天再入院 = 54
End Enum
Private Enum 西医诊断
    chk是否确诊 = 0
    txt抢救次数 = 26
    txt确诊日期 = 27
    txt成功次数 = 28
    cbo门诊与出院 = 36
    cbo入院与出院 = 35
    cbo门诊与入院 = 58
    cbo放射与病理 = 34
    cbo临床与病理 = 33
    cbo临床与尸检 = 32
    cbo术前与术后 = 31
    txt抢救原因 = 50
    cbo分化程度 = 52
    cbo最高诊断依据 = 53
    lbl分化程度 = 103
    lbl最高诊断依据 = 104
    txt病理号 = 57
End Enum
Private Enum 中医诊断
    chk危重 = 2
    chk急症 = 3
    chk疑难 = 4
    cbo辨证 = 2
    cbo治法 = 11
    cbo方药 = 12
    cbo自制中药 = 13
    cbo抢救方法 = 14
    cbo治疗类别 = 15
    cbo中医门诊与出院 = 38
    cbo中医入院与出院 = 37
    cbo使用中医诊疗设备 = 55
    cbo使用中医诊疗技术 = 56
    cbo辨证施护 = 57
End Enum
Private Enum 过敏与手术
    cboHBsAg = 29
    cboHCVAb = 28
    cboHIVAb = 27
    chk手术自由录入 = 19
End Enum
Private Enum 住院情况
    chk新发肿瘤 = 5
    chk尸检 = 6
    chk随诊 = 7
    chk示教病案 = 8
    chk科研病案 = 11
    chk疑难病例 = 20
    chk经外院治疗 = 15
    txt死亡原因 = 4
    txt随诊期限 = 29
    txt输红细胞 = 30
    txt输血小板 = 31
    txt输血浆 = 32
    txt输全血 = 33
    txt输其他 = 34
    txt医学警示 = 66
    txt其他医学警示 = 65
    cbo输液反应 = 30
    cbo随诊Ex = 16
    cbo血型 = 17
    cboRh = 18
    cbo门诊医师 = 19
    cbo科主任 = 20
    cbo主任医师 = 21
    cbo主治医师 = 22
    cbo住院医师 = 23
    cbo进修医师 = 24
    cbo研究生医师 = 25
    cbo实习医师 = 26
    cbo质控护士 = 39
    cbo质控医师 = 40
    cbo输血反应 = 41
    cbo出院方式 = 44
    txt出院转入 = 40
    lbl转出去向 = 88
    txt31天目的 = 41
    opt31天无 = 6
    opt31天有 = 7
    cbo责任护士 = 50
    txt入院前天 = 56
    txt入院后天 = 55
    txt入院前小时 = 45
    txt入院前分钟 = 46
    txt入院后小时 = 47
    txt入院后分钟 = 48
    txt呼吸机小时 = 49
    txt籍贯 = 52
    txt户口地址 = 53
    txt户口邮编 = 54
    cbo病案质量 = 59
    txt质控日期 = 60
    cbo生育状况 = 60
End Enum
Private Enum 附加内容
    chk病原学 = 9
    txt病原学 = 35
    lbl病原学 = 61
    cbo输血检查 = 42
    cbo病例分型 = 43
    chkCT = 12
    chkMRI = 13
    chk多普勒 = 14
    pic压疮 = 0
    pic跌倒或坠床 = 1
    cbo压疮发生期间 = 45
    cbo压疮分期 = 46
    cbo跌倒或坠床伤害 = 47
    cbo跌倒或坠床原因 = 48
    chk住院期间告病重或病危 = 1
    chk进入路径 = 16
    chk完成路径 = 17
    chk变异 = 18
    txt退出原因 = 61
    txt变异原因 = 62
    chk是否使用物理约束 = 21
    txt约束总时间 = 58
    cbo约束方式 = 63
    cbo约束工具 = 62
    cbo约束原因 = 61
    cbo新生儿离院方式 = 64
    chk围术期死亡 = 22
    chk术后猝死 = 23
    cbo重返间隔时间 = 65
    chk人工气道脱出 = 24
    chk重返重症医学科 = 25
    txt重症监护室 = 59
End Enum
Private Enum 签名级别
    cmd科主任 = 0
    cmd主任医师 = 1
    cmd主治医师 = 2
    cmd住院医师 = 3
End Enum
Private Enum 抗生素
    kss名称 = 1
    kss用药目的 = 2
    kss使用阶段 = 3
    kss使用天数 = 4
    KSS一类切口预防用 = 5
    KSSDDD数 = 6
    KSS联合用药 = 7
End Enum

Private Enum AllerColS
    AC_过敏时间 = 0
    AC_过敏药物 = 1
    AC_过敏反应 = 2
End Enum

Private Enum 旧的登记项
    txt自体回收 = 51
End Enum


Private Const ColorUnEditCell = &H8000000B  '灰蓝色

Public Function EditMedicalRecord(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, _
                                ByVal lngPathState As Long, ByVal strPrivs As String, frmParent As Object, ByVal blnModal As Boolean, _
                                Optional ByVal str类型 As String, Optional blnDiagnose As Boolean, Optional ByVal blnReadOnly As Boolean, _
                                Optional ByRef str疾病ID As String, Optional ByRef str诊断ID As String, Optional ByVal bln护士站 As Boolean) As Boolean

'参数：str类型=要示录入的诊断类型，如"3,13"格式
'      blnDiagnose=要求录入诊断，并缺省定位到诊断
'返回：blnDiagnose=是否录入了指定类型的诊断
    mstrPrivs = strPrivs
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng科室ID = lng科室ID
    mlngPathState = lngPathState
    mblnDiagChange = False
    mstrXYDiagInfo = ""
    mstrZYDiagInfo = ""
    
    mstr类型 = str类型
    mblnDiagnose = blnDiagnose
    mblnReadOnly = blnReadOnly
    mbln护士站 = bln护士站
    
    mstr疾病ID = ""
    mstr诊断ID = ""

    On Error Resume Next
    If blnModal Then
        Me.Show 1, frmParent
        blnDiagnose = mblnDiagnose
        EditMedicalRecord = mblnOk
        str疾病ID = mstr疾病ID
        str诊断ID = mstr诊断ID

    Else
        Me.Show , frmParent
    End If
End Function

Public Property Let Opened(ByVal vData As Boolean)
    mblnOpen = vData
End Property

Public Property Get Opened() As Boolean
    Opened = mblnOpen
End Property

Private Sub cboInfo_Change(Index As Integer)
    If cboinfo(Index).Style = 0 Then
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    If cboinfo(Index).ItemData(cboinfo(Index).ListIndex) = -1 And Visible Then
        '选择其他内容
        If Index = cbo门诊医师 Or Index = cbo科主任 Or Index = cbo主任医师 Or Index = cbo主治医师 Or Index = cbo住院医师 _
            Or Index = cbo实习医师 Or Index = cbo进修医师 Or Index = cbo研究生医师 Or Index = cbo质控医师 Or Index = cbo质控护士 Or Index = cbo责任护士 Then
            
            StrSQL = mcol人员SQL("_" & Index)

            vRect = GetControlRect(cboinfo(Index).hwnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "医生护士", , , , , , True, vRect.Left, vRect.Top, cboinfo(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = SeekCboIndex(cboinfo(Index), rsTmp!ID)
                If intIdx <> -1 Then
                    cboinfo(Index).ListIndex = intIdx
                Else
                    cboinfo(Index).AddItem rsTmp!姓名, cboinfo(Index).ListCount - 1
                    cboinfo(Index).ItemData(cboinfo(Index).NewIndex) = rsTmp!ID
                    cboinfo(Index).ListIndex = cboinfo(Index).NewIndex
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有住院医生或护士的数据，请先到部门/人员管理中设置。", vbInformation, gstrSysName
                End If
                '恢复成现有的人员(不引发Click)
                intIdx = SeekCboIndex(cboinfo(Index), cboinfo(Index).Tag)
                Call zlControl.CboSetIndex(cboinfo(Index).hwnd, intIdx)
            End If
        End If
    Else
        cboinfo(Index).Tag = cboinfo(Index).Text
    End If
    
    If Index = cbo科主任 Or Index = cbo主任医师 Or Index = cbo主治医师 Or Index = cbo住院医师 Then
        '医师更改,刷新签名状态
        If Visible Then
            mblnReadOnly = SetSignature(False)
            Call SetFaceEditable(mblnReadOnly)
        End If
    ElseIf Index = cbo随诊Ex Then
        If cboinfo(Index).Text = "终身" Then
            txtInfo(txt随诊期限).Text = ""
            txtInfo(txt随诊期限).Locked = True
            txtInfo(txt随诊期限).TabStop = False
            txtInfo(txt随诊期限).BackColor = vbButtonFace
        Else
            txtInfo(txt随诊期限).Locked = False
            txtInfo(txt随诊期限).TabStop = True
            txtInfo(txt随诊期限).BackColor = vbWindowBackground
            If Visible Then txtInfo(txt随诊期限).SetFocus
        End If
    ElseIf Index = cbo出院方式 Then
        If cboinfo(Index).Text = "转院" Or cboinfo(Index).Text = "转社区" Then
            txtInfo(txt出院转入).Enabled = True
            lblInfo(lbl转出去向).Enabled = True
            txtInfo(txt出院转入).TabStop = True
            txtInfo(txt出院转入).BackColor = vbWindowBackground
        Else
            txtInfo(txt出院转入).Enabled = False
            lblInfo(lbl转出去向).Enabled = False
            txtInfo(txt出院转入).TabStop = False
            txtInfo(txt出院转入).BackColor = vbButtonFace
        End If
    End If
End Sub

Private Sub cboInfo_GotFocus(Index As Integer)
    If cboinfo(Index).Style = 0 Then
        Call zlControl.TxtSelAll(cboinfo(Index))
    End If
End Sub

Private Sub cboInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboinfo(Index).Style = 2 And cboinfo(Index).ListIndex <> -1 Then
            cboinfo(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngidx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = cbo辨证施护 Then
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call sstInfo_KeyPress(13)
        ElseIf Index = cbo责任护士 Then
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call sstInfo_KeyPress(13)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
            If Index = cbo重返间隔时间 Then
                If vsfMain.Rows = 1 Then Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    ElseIf KeyAscii >= 32 Then
        If Index = cbo身份证号 Then
            '限制输入长度
            If zlCommFun.ActualLen(cboinfo(Index).Text) > 18 Then
                KeyAscii = 0: Exit Sub
            End If
            
            '限制输入内容
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        ElseIf Not cboinfo(Index).Locked And cboinfo(Index).Style = 2 Then
            lngidx = zlControl.CboMatchIndex(cboinfo(Index).hwnd, KeyAscii)
            If lngidx = -1 And cboinfo(Index).ListCount > 0 Then lngidx = 0
            cboinfo(Index).ListIndex = lngidx
        End If
    End If
End Sub

Private Sub cboInfo_LostFocus(Index As Integer)
    Dim strTmp As String, strMsg As String
    
    On Local Error Resume Next
    
    If Index = cbo年龄单位 Then
        If IsNumeric(txtInfo(txt年龄).Text) And cboinfo(cbo年龄单位).ListIndex <> -1 Then
            Select Case cboinfo(cbo年龄单位).Text
                Case "岁"
                    If Val(txtInfo(txt年龄).Text) > 200 Then
                        MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                        txtInfo(txt年龄).SetFocus: Exit Sub
                    End If
                Case "月"
                    If Val(txtInfo(txt年龄).Text) > 2400 Then
                        MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                        txtInfo(txt年龄).SetFocus: Exit Sub
                    End If
                Case "天"
                    If Val(txtInfo(txt年龄).Text) > 73000 Then
                        MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                        txtInfo(txt年龄).SetFocus: Exit Sub
                    End If
                Case Else
                    Exit Sub
            End Select
            
            '反算或检查出生日期（小时与分钟做单位不进行反算与检查）
            If cboinfo(cbo年龄单位).ListIndex < 3 Then
                If Not IsDate(txt出生日期.Text) Then
                    txt出生日期.Text = ReCalcBirth(txtInfo(txt年龄).Text, cboinfo(cbo年龄单位).Text)
                Else
                    strTmp = PatiAgeCalc(txt出生日期.Text, , txtInfo(txt入院时间).Text)
                    If Right(strTmp, 1) = cboinfo(cbo年龄单位).Text And IsNumeric(Left(strTmp, Len(strTmp) - 1)) _
                        And strTmp <> txtInfo(txt年龄).Text & cboinfo(cbo年龄单位).Text Then
                        
                        strMsg = zlCommFun.ShowMsgBox(gstrSysName, "年龄和出生日期不一致，" & txt出生日期.Text & "出生现在应该是" & strTmp & "。" & _
                            vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，并选择下面相应的操作。", "!重算生日(&R),忽略(&A),?取消(&C)", Me, vbQuestion)
                        If strMsg = "重算生日" Then
                            txt出生日期.Text = ReCalcBirth(txtInfo(txt年龄).Text, cboinfo(cbo年龄单位).Text)
                            txt出生时间.Text = "__:__"
                        ElseIf strMsg = "忽略" Then
                        Else
                            txtInfo(txt年龄).SetFocus: Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboInfo_Validate(Index As Integer, Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If Index = cbo门诊医师 Or Index = cbo科主任 Or Index = cbo主任医师 Or Index = cbo主治医师 Or Index = cbo住院医师 _
        Or Index = cbo实习医师 Or Index = cbo进修医师 Or Index = cbo研究生医师 Or Index = cbo质控医师 Or Index = cbo质控护士 Then
        If cboinfo(Index).ListIndex <> -1 Then Exit Sub '已选中
        If cboinfo(Index).Text = "" Then cboinfo(Index).Tag = "": Exit Sub '无输入
        
        strInput = UCase(NeedName(cboinfo(Index).Text))
        StrSQL = mcol人员SQL("_" & Index)
        StrSQL = Replace(UCase(StrSQL), UCase("Order by"), " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2]) Order by")
        
        On Error GoTo errH
        vRect = GetControlRect(cboinfo(Index).hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "医生护士", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, cboinfo(Index).Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cboinfo(Index), rsTmp!ID)
            If intIdx <> -1 Then
                cboinfo(Index).ListIndex = intIdx
            Else
                cboinfo(Index).AddItem rsTmp!姓名, cboinfo(Index).ListCount - 1
                cboinfo(Index).ItemData(cboinfo(Index).NewIndex) = rsTmp!ID
                cboinfo(Index).ListIndex = cboinfo(Index).NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未找到对应的医生或护士。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    ElseIf Index = cbo术前与术后 Then
        If cboinfo(cbo术前与术后).ListIndex = 0 And vsOPS.TextMatrix(1, col手术名称) <> "" Then
            '如果填了手术，就不允许选未做
            cboinfo(cbo术前与术后).ListIndex = 1
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkInfo_Click(Index As Integer)
    If mblnNoClick Then Exit Sub
    If Visible And mblnReadOnly Then
        mblnNoClick = True
        chkInfo(Index).Value = IIf(chkInfo(Index).Value = 1, 0, 1)
        mblnNoClick = False: Exit Sub
    End If
    
    Select Case Index
        Case chk是否确诊
            If chkInfo(Index).Value = 1 Then
                txtInfo(txt确诊日期).Locked = False
                txtInfo(txt确诊日期).TabStop = True
                txtInfo(txt确诊日期).BackColor = vbWindowBackground
            Else
                txtInfo(txt确诊日期).Text = ""
                txtInfo(txt确诊日期).Locked = True
                txtInfo(txt确诊日期).TabStop = False
                txtInfo(txt确诊日期).BackColor = vbButtonFace
            End If
        Case chk随诊
            If chkInfo(Index).Value = 1 Then
                txtInfo(txt随诊期限).Locked = False
                txtInfo(txt随诊期限).TabStop = True
                txtInfo(txt随诊期限).BackColor = vbWindowBackground
                cboinfo(cbo随诊Ex).Locked = False
                cboinfo(cbo随诊Ex).TabStop = True
                cboinfo(cbo随诊Ex).BackColor = vbWindowBackground
                
                Call cboInfo_Click(cbo随诊Ex)
            Else
                txtInfo(txt随诊期限).Text = ""
                txtInfo(txt随诊期限).Locked = True
                txtInfo(txt随诊期限).TabStop = False
                txtInfo(txt随诊期限).BackColor = vbButtonFace
                cboinfo(cbo随诊Ex).Locked = True
                cboinfo(cbo随诊Ex).TabStop = False
                cboinfo(cbo随诊Ex).BackColor = vbButtonFace
            End If
        Case chk进入路径
            If chkInfo(Index).Value = 0 Then
                chkInfo(chk完成路径).Enabled = False
                chkInfo(chk完成路径).Value = 0
                txtInfo(txt退出原因).Enabled = False
                txtInfo(txt退出原因).Text = ""
                txtInfo(txt退出原因).BackColor = vbButtonFace
                chkInfo(chk变异).Enabled = False
                chkInfo(chk变异).Value = 0
                txtInfo(txt变异原因).Enabled = False
                txtInfo(txt变异原因).Text = ""
                txtInfo(txt变异原因).BackColor = vbButtonFace
            Else
                chkInfo(chk完成路径).Enabled = True
                chkInfo(chk完成路径).TabStop = True
                txtInfo(txt退出原因).TabStop = True
                txtInfo(txt退出原因).Enabled = True
                txtInfo(txt退出原因).BackColor = vbWindowBackground
                chkInfo(chk变异).Enabled = True
                chkInfo(chk变异).TabStop = True
                chkInfo(Index).TabStop = True
            End If
        Case chk完成路径
            If chkInfo(Index).Value = 0 Then
                txtInfo(txt退出原因).Enabled = True
                txtInfo(txt退出原因).TabStop = True
                txtInfo(txt退出原因).BackColor = vbWindowBackground
            Else
                txtInfo(txt退出原因).Enabled = False
                txtInfo(txt退出原因).Text = ""
                txtInfo(txt退出原因).BackColor = vbButtonFace
                chkInfo(Index).TabStop = True
            End If
        Case chk变异
            If chkInfo(Index).Value = 1 Then
                txtInfo(txt变异原因).Enabled = True
                txtInfo(txt变异原因).TabStop = True
                txtInfo(txt变异原因).BackColor = vbWindowBackground
                chkInfo(Index).TabStop = True
            Else
                txtInfo(txt变异原因).Enabled = False
                txtInfo(txt变异原因).Text = ""
                txtInfo(txt变异原因).BackColor = vbButtonFace
            End If
        Case chk尸检
            '设置诊断符合情况
            If Visible Then Call Set诊断符合情况(cbo临床与尸检)
        Case chk病原学
            If chkInfo(chk病原学).Value = 0 Then
                txtInfo(txt病原学).Text = ""
                txtInfo(txt病原学).Tag = ""
                cmdInfo(txt病原学).Tag = ""
                txtInfo(txt病原学).Enabled = False
                cmdInfo(txt病原学).Enabled = False
                lblInfo(lbl病原学).ForeColor = &H808080
            ElseIf Not txtInfo(txt病原学).Enabled Then
                txtInfo(txt病原学).Enabled = True
                cmdInfo(txt病原学).Enabled = True
                lblInfo(lbl病原学).ForeColor = Me.ForeColor
            End If
        Case chk是否使用物理约束
            If chkInfo(Index).Value = 0 Then
                txtInfo(txt约束总时间).Text = ""
                cboinfo(cbo约束方式).ListIndex = -1
                cboinfo(cbo约束工具).ListIndex = -1
                cboinfo(cbo约束原因).ListIndex = -1
                txtInfo(txt约束总时间).BackColor = vbButtonFace
                cboinfo(cbo约束方式).BackColor = vbButtonFace
                cboinfo(cbo约束工具).BackColor = vbButtonFace
                cboinfo(cbo约束原因).BackColor = vbButtonFace
                txtInfo(txt约束总时间).Enabled = False
                cboinfo(cbo约束方式).Enabled = False
                cboinfo(cbo约束工具).Enabled = False
                cboinfo(cbo约束原因).Enabled = False
            Else
                txtInfo(txt约束总时间).BackColor = vbWindowBackground
                cboinfo(cbo约束方式).BackColor = vbWindowBackground
                cboinfo(cbo约束工具).BackColor = vbWindowBackground
                cboinfo(cbo约束原因).BackColor = vbWindowBackground
                txtInfo(txt约束总时间).Enabled = True
                cboinfo(cbo约束方式).Enabled = True
                cboinfo(cbo约束工具).Enabled = True
                cboinfo(cbo约束原因).Enabled = True
            End If
    End Select
    If Visible Then mblnChange = True
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = chk术后猝死 Or Index = chk住院期间告病重或病危 Then
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call sstInfo_KeyPress(13)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function GetStage(ByVal DateUseBegin As Date, ByVal DateUseEnd As Date, ByVal DateSs As Date, ByVal strTime As String) As String
'功能：获得抗生素使用阶段
'参数：DateUseBegin 使用时间,DateUseEnd -结束时间  DateSs-手术时间,strTime 上一次的使用阶段
    
    '如果没有手术，则返回空
    If DateSs = 0 Then GetStage = " ": Exit Function
    '如果已经是围手术期，直接退出
    If strTime = "围手术期" Then GetStage = strTime: Exit Function
    
    If DateUseBegin < DateSs Then
        If DateUseEnd < DateSs Then
            If strTime <> "" Then
                If strTime <> "术前" Then strTime = "围手术期"
            Else
                strTime = "术前"
            End If
        Else
            strTime = "围手术期"
        End If
    ElseIf DateUseBegin > DateSs Then
        If DateUseEnd > DateSs Then
            If strTime <> "" Then
                If strTime <> "术后" Then strTime = "围手术期"
            Else
                strTime = "术后"
            End If
        Else
            strTime = "围手术期"
        End If
    Else
        If DateUseEnd = DateSs Then
            If strTime <> "" Then
                If strTime <> "术中" Then strTime = "围手术期"
            Else
                strTime = "术中"
            End If
        Else
            strTime = "围手术期"
        End If
    End If
    GetStage = strTime
End Function

Private Sub cmdAutoLoad_Click(Index As Integer)
    '自动提取
    Dim StrSQL As String, rsTmp As Recordset
    Dim i As Long, j As Long
    Dim blnAgain As Boolean, blnIsNull As Boolean
    Dim DateSs As Date          '该病人最早的手术时间
    Dim rsTime As New ADODB.Recordset
    Dim blnStage As Boolean
    Dim strOld天数 As String
    Dim lngRow As Long
    Dim blnClear As Boolean
    Dim strPrivs As String
    
    On Error GoTo errH
    Select Case Index
        Case 0
            StrSQL = "Select Min(NVL(to_date(c.标本部位,'yyyy-mm-dd hh24:mi:ss'),c.开始执行时间)) as 使用时间" & vbNewLine & _
                    " From 诊疗项目目录 A, 病人医嘱记录 C" & vbNewLine & _
                    " Where  a.Id = c.诊疗项目id and a.类别='F' And c.病人id = [1] And c.主页id = [2] And c.医嘱状态=8"
        
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
            If rsTmp.RecordCount > 0 Then DateSs = CDate(Format(Nvl(rsTmp!使用时间, 0), "yyyy-MM-dd"))
            
            StrSQL = "Select distinct ID, 医嘱id, 上级id, 编码, 名称, 单位, 执行时间方案, 频率间隔, 间隔单位, 频率次数, 上次执行时间, 开始执行时间, 结束时间," & vbNewLine & _
                    "       Sum(Ddd数) Over(Partition By ID) As Ddd数, Count(1) Over(Partition By 相关id) As 联合用药" & vbNewLine & _
                    "From   (Select Distinct ID, 医嘱id, 上级id, 编码, 名称, 单位, 执行时间方案, 频率间隔, 间隔单位, 频率次数, 上次执行时间, 开始执行时间, 结束时间," & vbNewLine & _
                    "                Sum(数次) Over(Partition By ID, 医嘱id, 相关id) * 剂量系数 / Decode(Ddd值, 0, Null, Ddd值) As Ddd数, 相关id" & vbNewLine & _
                    "         From   (Select z.Id, a.Id As 医嘱id, z.分类id As 上级id, z.编码, z.名称, z.计算单位 As 单位, a.执行时间方案, a.频率间隔, a.间隔单位, a.频率次数," & vbNewLine & _
                    "                         a.上次执行时间, a.开始执行时间, Nvl(a.上次执行时间, Nvl(a.执行终止时间, a.开始执行时间)) As 结束时间, a.相关id, f.数次, h.剂量系数," & vbNewLine & _
                    "                         Nvl((Select e.Ddd值 From 诊疗用法用量 E Where e.项目id = a.诊疗项目id And e.用法id = r.诊疗项目id), h.Ddd值) As Ddd值" & vbNewLine & _
                    "                  From   病人医嘱记录 A, 病人医嘱记录 R, 住院费用记录 F, 药品规格 H, 药品特性 B, 诊疗项目目录 Z" & vbNewLine & _
                    "                  Where  a.诊疗项目id = b.药名id And a.诊疗类别 In ('5', '6') And" & vbNewLine & _
                    "                         (a.医嘱期效 = 0 And a.上次执行时间 Is Not Null Or a.医嘱期效 = 1 And a.医嘱状态 = 8) And Nvl(b.抗生素, 0) <> 0 And" & vbNewLine & _
                    "                         a.相关id = r.Id And a.Id = f.医嘱序号 And f.记录状态 <> 0 And f.收费细目id = h.药品id And b.药名id = z.Id And" & vbNewLine & _
                    "                         a.病人id = [1] And a.主页id = [2]))" & vbNewLine & _
                    "Order  By Ddd数 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
            
            rsTime.Fields.Append "收费时间", adVarChar, 10
            rsTime.Fields.Append "药品ID", adBigInt
            rsTime.CursorLocation = adUseClient
            rsTime.LockType = adLockOptimistic
            rsTime.CursorType = adOpenStatic
            rsTime.Open
                                
            With vsKSS
                
                If rsTmp.RecordCount = 0 Then MsgBox "没有找到该病人的抗菌药物使用记录。", vbInformation, Me.Caption
                Do Until rsTmp.EOF
                    
                    For i = .FixedRows To .Rows - 1
                        '判断是否有重复的
                        If rsTmp!ID = Val(.RowData(i) & "") Then
                            '如果是空才提取出来，如果不是空，则不改变用户的手动选择
                            If .Cell(flexcpData, i, kss使用阶段) = "新增" Or Trim(.TextMatrix(i, kss使用阶段)) = "" Then
                                blnStage = True
                            End If
                            lngRow = i
                            If .TextMatrix(i, KSSDDD数) = "" Then .TextMatrix(i, KSSDDD数) = FormatEx(Val(rsTmp!DDD数 & ""), 2)
                            If Decode(.TextMatrix(i, KSS联合用药), "Ⅰ种", 1, "Ⅱ联", 2, "Ⅲ联", 3, "Ⅳ联", 4, ">Ⅳ联", 999, 0) < Val(rsTmp!联合用药 & "") Then
                                .TextMatrix(i, KSS联合用药) = Decode(Val(rsTmp!联合用药 & ""), 1, "Ⅰ种", 2, "Ⅱ联", 3, "Ⅲ联", 4, "Ⅳ联", ">Ⅳ联")
                            End If
                            Exit For
                        End If
                        '判断填入之前的行有没有空的
                        If .TextMatrix(i, kss名称) & "" = "" Then
                            .TextMatrix(i, kss名称) = rsTmp!名称 & ""
                            .RowData(i) = Val(rsTmp!ID)
                            .Cell(flexcpData, i, kss名称) = rsTmp!名称 & ""
                            .TextMatrix(i, KSSDDD数) = FormatEx(Val(rsTmp!DDD数 & ""), 2)
                            .Cell(flexcpData, i, KSSDDD数) = .TextMatrix(i, KSSDDD数)
                            If Decode(.TextMatrix(i, KSS联合用药), "Ⅰ种", 1, "Ⅱ联", 2, "Ⅲ联", 3, "Ⅳ联", 4, ">Ⅳ联", 999, 0) < Val(rsTmp!联合用药 & "") Then
                                .TextMatrix(i, KSS联合用药) = Decode(Val(rsTmp!联合用药 & ""), 1, "Ⅰ种", 2, "Ⅱ联", 3, "Ⅲ联", 4, "Ⅳ联", ">Ⅳ联")
                            End If
                            lngRow = i
                            blnStage = True
                            Exit For
                        Else
                            If i = .Rows - 1 Then
                                .AddItem "": .TextMatrix(.Rows - 1, 0) = .Rows - 1
                                .TextMatrix(.Rows - 1, kss名称) = rsTmp!名称 & ""
                                .RowData(.Rows - 1) = Val(rsTmp!ID)
                                .Cell(flexcpData, .Rows - 1, kss名称) = rsTmp!名称 & ""
                                .TextMatrix(i, KSSDDD数) = FormatEx(Val(rsTmp!DDD数 & ""), 2)
                                .Cell(flexcpData, i, KSSDDD数) = .TextMatrix(i, KSSDDD数)
                                If Decode(.TextMatrix(i, KSS联合用药), "Ⅰ种", 1, "Ⅱ联", 2, "Ⅲ联", 3, "Ⅳ联", 4, ">Ⅳ联", 999, 0) < Val(rsTmp!联合用药 & "") Then
                                    .TextMatrix(i, KSS联合用药) = Decode(Val(rsTmp!联合用药 & ""), 1, "Ⅰ种", 2, "Ⅱ联", 3, "Ⅲ联", 4, "Ⅳ联", ">Ⅳ联")
                                End If
                                lngRow = .Rows - 1
                                blnStage = True
                                Exit For
                            End If
                        End If
                    Next
                    If blnStage Then
                        '使用阶段
                        .TextMatrix(lngRow, kss使用阶段) = _
                            GetStage(CDate(Format(rsTmp!开始执行时间 & "", "yyyy-MM-dd")), CDate(Format(rsTmp!结束时间 & "", "yyyy-MM-dd")), DateSs, Trim(.TextMatrix(lngRow, kss使用阶段)))
                        .Cell(flexcpData, lngRow, kss使用阶段) = "新增"
                        vsKSS.Tag = "": mblnChange = True
                    End If
                    strOld天数 = Trim(.TextMatrix(lngRow, kss使用天数))
                        '使用天数
                    .TextMatrix(lngRow, kss使用天数) = GetUseDay(Val(rsTmp!医嘱ID), Val(.RowData(lngRow)), Nvl(rsTmp!执行时间方案) & "", CDate(rsTmp!开始执行时间), CDate(rsTmp!结束时间), _
                                Nvl(rsTmp!频率次数, 0), Nvl(rsTmp!频率间隔, 0), Nvl(rsTmp!间隔单位), rsTime) & ""
                    If strOld天数 <> Trim(.TextMatrix(lngRow, kss使用天数)) Then
                        .Cell(flexcpData, lngRow, kss使用天数) = "新增"
                        vsKSS.Tag = "": mblnChange = True
                    End If
                    
                    rsTmp.MoveNext
                    blnStage = False
                    lngRow = 0
                Loop
            End With
        Case 1
            strPrivs = GetInsidePrivs(p手麻接口, , 2400)
            If InStr(strPrivs, "内部接口") > 0 Then
                Set rsTmp = AutoGetOPSInfo(True, mlng病人ID, mlng主页ID)
            Else
                If gbln手术提取手麻 Then
                    If MsgBox("由于你没有【手麻管理系统-手麻接口管理】模块的内部接口权限，" & vbCrLf & "系统默认从医嘱系统中读取手术信息，是否继续 ? " & vbCrLf & "选择是则下次将不再提示。", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                        gbln手术提取手麻 = False
                    Else
                        Exit Sub
                    End If
                End If
                Set rsTmp = AutoGetOPSInfo(False, mlng病人ID, mlng主页ID)
            End If
            If Not rsTmp.EOF Then
                If MsgBox("是否清空原有的手术信息？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    blnClear = True
                End If
        
                With vsOPS
                    
                    If blnClear Then
                        .Rows = .FixedRows
                        .Rows = .FixedRows + rsTmp.RecordCount + 1
                        lngRow = .FixedRows
                    Else
                        If .Rows > .FixedRows Then
                            If .TextMatrix(.Rows - 1, col手术名称) = "" Then
                                lngRow = .Rows - 1
                                .Rows = .Rows + rsTmp.RecordCount
                            Else
                                lngRow = .Rows
                                .Rows = .Rows + rsTmp.RecordCount + 1
                            End If
                        End If
                    End If
                    
                    rsTmp.MoveFirst
                    
                    For i = lngRow To lngRow + rsTmp.RecordCount - 1
                        .TextMatrix(i, col手术日期) = Format(Nvl(rsTmp!手术日期), "yyyy-MM-dd")
                        .TextMatrix(i, col手术编码) = Nvl(rsTmp!手术编码)
                        .TextMatrix(i, col手术名称) = Nvl(rsTmp!已行手术)
                        .TextMatrix(i, col主刀医师) = Nvl(rsTmp!主刀医师)
                        .TextMatrix(i, col助产护士) = Nvl(rsTmp!助产护士)
                        .TextMatrix(i, col助手1) = Nvl(rsTmp!第一助手)
                        .TextMatrix(i, col助手2) = Nvl(rsTmp!第二助手)
                        .TextMatrix(i, col麻醉方式) = GetItemField("诊疗项目目录", Val(Nvl(rsTmp!麻醉方式, 0)), "名称")
                        .TextMatrix(i, col麻醉医师) = Nvl(rsTmp!麻醉医师)
                        If Not IsNull(rsTmp!切口) And Not IsNull(rsTmp!愈合) Then
                            .TextMatrix(i, col切口愈合) = rsTmp!切口 & "/" & rsTmp!愈合
                        End If
                        .TextMatrix(i, col手术操作ID) = Nvl(rsTmp!手术操作ID)
                        .TextMatrix(i, col诊疗项目ID) = Nvl(rsTmp!诊疗项目id)
                        .TextMatrix(i, col麻醉ID) = Nvl(rsTmp!麻醉方式)
                        .TextMatrix(i, col麻醉类型) = Nvl(rsTmp!麻醉类型)
                        .TextMatrix(i, COL手术情况.COL手术情况) = Nvl(rsTmp!手术情况)
                        .TextMatrix(i, colASA分级) = Decode(Nvl(rsTmp!asa分级), "I级", "P1", "II级", "P2", "III级", "P3", "IV级", "P4", "V级", "P5", Nvl(rsTmp!asa分级))
                        .TextMatrix(i, colNNIS分级) = Nvl(rsTmp!NNIS分级)
                        .TextMatrix(i, col手术级别) = Nvl(rsTmp!手术级别)
                        .TextMatrix(i, col再次手术) = IIf(Val(rsTmp!再次手术 & "") = 1, -1, 0)
                        .TextMatrix(i, col抗菌药天数) = rsTmp!抗菌用药天数 & ""
                        .Cell(flexcpChecked, i, col预防用抗菌药) = Val(rsTmp!术前抗菌用药 & "")
                        .Cell(flexcpChecked, i, col非预期的二次手术) = Val(rsTmp!非预期的二次手术 & "")
                        .Cell(flexcpChecked, i, col麻醉并发症) = Val(rsTmp!麻醉并发症 & "")
                        .Cell(flexcpChecked, i, col术中异物遗留) = Val(rsTmp!术中异物遗留 & "")
                        .Cell(flexcpChecked, i, col手术并发症) = Val(rsTmp!手术并发症 & "")
                        .Cell(flexcpChecked, i, col术后出血或血肿) = Val(rsTmp!术后出血或血肿 & "")
                        .Cell(flexcpChecked, i, col手术伤口裂开) = Val(rsTmp!手术伤口裂开 & "")
                        .Cell(flexcpChecked, i, col术后深静脉血栓) = Val(rsTmp!术后深静脉血栓 & "")
                        .Cell(flexcpChecked, i, col术后生理代谢紊乱) = Val(rsTmp!术后生理代谢紊乱 & "")
                        .Cell(flexcpChecked, i, col术后呼吸衰竭) = Val(rsTmp!术后呼吸衰竭 & "")
                        .Cell(flexcpChecked, i, col术后肺栓塞) = Val(rsTmp!术后肺栓塞 & "")
                        .Cell(flexcpChecked, i, col术后败血症) = Val(rsTmp!术后败血症 & "")
                        .Cell(flexcpChecked, i, col术后髋关节骨折) = Val(rsTmp!术后髋关节骨折 & "")
                        '记录用于编辑恢复
                        For j = 0 To .Cols - 1
                            .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                        Next
                        
                        rsTmp.MoveNext
                    Next
                End With
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetUseDay(ByVal AdviceID As Long, ByVal lng药品ID As Long, ByVal str执行时间方案 As String, ByVal Date开始执行时间 As Date, _
            ByVal Date结束时间 As Date, ByVal lng频率次数 As Long, ByVal lng频率间隔 As Long, ByVal str间隔单位 As String, _
            ByRef rsTime As ADODB.Recordset) As Long
'功能：获取抗生素的使用天数

    Dim strPause As String
    Dim j As Long
    Dim StrDecTime As String, arrDecTime As Variant
    Dim DateStart As String
    Dim strTmp As String
        
    strPause = GetAdvicePause(AdviceID)
    If str执行时间方案 <> "" Then
        StrDecTime = Calc段内分解时间(Date开始执行时间, Date结束时间, strPause, str执行时间方案, lng频率次数, lng频率间隔, str间隔单位)
        arrDecTime = Split(StrDecTime, ",")
        For j = 0 To UBound(arrDecTime)
            strTmp = Format(arrDecTime(j), "yyyy-MM-dd")
            rsTime.Filter = "收费时间='" & strTmp & "' And " & "药品id=" & lng药品ID
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!收费时间 = strTmp
                rsTime!药品id = lng药品ID
                rsTime.Update
            End If
        Next
    Else
        DateStart = CDate(Format(Date开始执行时间 & "", "yyyy-MM-dd"))
        Do While DateStart <= CDate(Format(Date结束时间 & "", "yyyy-MM-dd"))
            rsTime.Filter = "收费时间='" & strTmp & "' And " & "药品id=" & lng药品ID
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!收费时间 = Format(CStr(DateStart), "yyyy-MM-dd")
                rsTime!药品id = lng药品ID
                rsTime.Update
            End If
            DateStart = CDate(DateStart) + 1
        Loop
    End If
    rsTime.Filter = "药品id=" & lng药品ID
    GetUseDay = rsTime.RecordCount
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInfo_Click(Index As Integer)
'说明：注意界面上要求CMD和对应TXT的Index相同
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, blnLevel As Boolean
    Dim strResult As String
    
    '使用Lock的方式,不采用Enabled的方式
    If Not cmdInfo(Index).Enabled Or txtInfo(Index).Locked Then
        If txtInfo(Index).Enabled Then txtInfo(Index).SetFocus: Exit Sub
    End If
    
    Select Case Index
        Case txt出生地点, txt家庭地址, txt联系人地址, txt户口地址
            '选择地区数据
            StrSQL = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""地区""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!名称
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt区域, txt籍贯
            '选择区域数据
            On Error GoTo errH
            StrSQL = "Select Nvl(级数,0) as 级数 From 区域 Group by Nvl(级数,0)"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
            If rsTmp.RecordCount > 1 Then blnLevel = True
            
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            If blnLevel Then
                StrSQL = _
                    " Select ID,上级id,编码,名称,简码,末级" & _
                    " From (Select 编码 As ID,RPad(Substr(编码,1,Decode(Nvl(级数,0),0,0,1,2,4)),6,'0') As 上级id," & _
                    "       编码,名称,简码,Decode(Nvl(级数,0),2,1,3,1,0) as 末级" & _
                    "       From 区域 Order By 编码)" & _
                    " Start With 上级ID Is Null Connect By Prior ID=上级id"
                Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 2, "区域", , , , , , , vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            Else
                StrSQL = "Select Rownum as ID,编码,名称,简码 From 区域 Order by 编码"
                Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""区域""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!名称
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt单位名称
            '选择单位信息
            StrSQL = "Select ID,上级ID,末级,编码,名称,简码,地址,电话,开户银行,帐号,联系人" & _
                " From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 2, "合约单位", , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""合约单位""数据，请先到合约单位管理中设置。", vbInformation, gstrSysName
                End If
                txtInfo(Index).Tag = ""
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                txtInfo(Index).Tag = Val(rsTmp!ID)
                If txtInfo(txt单位电话).Text = "" Then
                    txtInfo(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt转科1, txt转科2, txt转科3
            '选择转科科室
            StrSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码,A.位置" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 IN(2,3) And B.工作性质 IN('临床','手术')" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""临床科室""数据，请先到部门管理中设置。", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!名称
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt确诊日期
            If IsDate(txtInfo(txt确诊日期).Text) Then
                dtpInfo.Value = CDate(txtInfo(txt确诊日期).Text)
            Else
                dtpInfo.Value = zlDatabase.Currentdate
            End If
            mlngDateIndex = Index
            dtpInfo.Left = cmdInfo(Index).Left + cmdInfo(Index).Width - dtpInfo.Width + txtInfo(Index).Container.Left + sstInfo.Left
            dtpInfo.Top = cmdInfo(Index).Top - dtpInfo.Height - 20 + txtInfo(Index).Container.Top + sstInfo.Top
            dtpInfo.ZOrder
            dtpInfo.Visible = True
            dtpInfo.SetFocus
        Case txt质控日期
            If IsDate(txtInfo(txt质控日期).Text) Then
                dtpInfo.Value = CDate(txtInfo(txt质控日期).Text)
            Else
                dtpInfo.Value = zlDatabase.Currentdate
            End If
            mlngDateIndex = Index
            dtpInfo.Left = cmdInfo(Index).Left + cmdInfo(Index).Width - dtpInfo.Width + txtInfo(Index).Container.Left + sstInfo.Left
            dtpInfo.Top = cmdInfo(Index).Top - dtpInfo.Height - 20 + txtInfo(Index).Container.Top + sstInfo.Top
            dtpInfo.ZOrder
            dtpInfo.Visible = True
            dtpInfo.SetFocus
        Case txt病原学
            'D-ICD-10疾病编码
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "'D'", mlng科室ID, cboinfo(cbo性别).Text, False)
            If Not rsTmp Is Nothing Then
                txtInfo(txt病原学).Text = IIf(Not IsNull(rsTmp!编码), "(" & rsTmp!编码 & ")", "") & Nvl(rsTmp!名称)
                txtInfo(txt病原学).Tag = txtInfo(txt病原学).Text
                cmdInfo(txt病原学).Tag = rsTmp!项目ID
            End If
            txtInfo(txt病原学).SetFocus
        Case txt抢救原因
             '选择单位信息
            StrSQL = "Select 编码 ID,名称,简码 From 抢救病因分类"
               
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "抢救原因", , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!名称
                txtInfo(Index).SetFocus
            End If
        Case txt重症监护室
            StrSQL = " Select Distinct A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where B.部门ID=A.ID And B.工作性质='ICU'" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "重症监护室", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            
            If rsTmp Is Nothing Then
                If Not blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    MsgBox "没有设置ICU重症监护室。", vbInformation, Me.Caption
                End If
            Else
                txtInfo(Index).Text = rsTmp!名称 & ""
            End If
        Case txt医学警示
            '选择医学警示
            On Error GoTo errH
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            StrSQL = "Select Rownum ID,编码,名称,简码 From 医学警示 Order by 编码"
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, StrSQL, 0, "", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, True, True)

            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""医学警示""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                While Not rsTmp.EOF
                    strResult = strResult & "," & rsTmp!名称
                    rsTmp.MoveNext
                Wend
                txtInfo(Index).Text = Mid(strResult, 2)
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim blnDiagnose As Boolean
    
    If Not CheckPageData(blnDiagnose, False) Then Exit Sub
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("要求的诊断信息还没有输入，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    If Not SavePageData(False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPathLoad_Click()
'功能：自动提取路径信息
    Dim StrSQL As String, rsTmp As Recordset
    
    
    StrSQL = "Select Decode(c.性质, 2, c.名称, '') As 名称,b.状态" & vbNewLine & _
            "From 病人路径评估 A, 病人临床路径 B, 变异常见原因 C" & vbNewLine & _
            "Where a.路径记录id(+) = b.Id And b.当前天数 = a.天数(+) And Nvl(b.当前阶段id, b.前一阶段id) = a.阶段id(+) And b.状态 <> 0 And a.变异原因 = c.编码(+) And b.病人id = [1] And b.主页id = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then
        chkInfo(chk进入路径).Value = 1
        If Val(rsTmp!状态 & "") = 3 Then
            chkInfo(chk完成路径).Value = 0
            txtInfo(txt退出原因).Text = rsTmp!名称 & ""
        ElseIf Val(rsTmp!状态 & "") = 2 Then
            chkInfo(chk完成路径).Value = 1
        End If
    Else
        chkInfo(chk进入路径).Value = 0
    End If
    '提取变异情况
    StrSQL = "Select Count(1) Over(Partition By b.病人id, b.主页id) As 变异数, c.名称 As 变异原因" & vbNewLine & _
            "From 病人路径评估 A, 病人临床路径 B, 变异常见原因 C" & vbNewLine & _
            "Where a.路径记录id = b.Id And c.编码(+) = a.变异原因 And a.评估结果 = -1 And b.病人id = [1] And b.主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then
        chkInfo(chk变异).Value = 1
        If Val(rsTmp!变异数 & "") = 1 Then
            txtInfo(txt变异原因).Text = rsTmp!变异原因 & ""
        End If
    Else
        chkInfo(chk变异).Value = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdPrint_Click()
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(2, mlng病人ID, mlng主页ID, mobjReport, mlng科室ID, Me)
End Sub

Private Sub cmdPriview_Click()
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(1, mlng病人ID, mlng主页ID, mobjReport, mlng科室ID, Me)
End Sub

Private Sub cmdPrintdown_Click()
    PopupMenu menuPrint, , cmdPrint.Left, cmdPrint.Top + cmdPrint.Height
End Sub

Private Sub cmdPriviewDown_Click()
    PopupMenu menuPriview, , cmdPriview.Left, cmdPriview.Top + cmdPriview.Height
End Sub

Private Sub menuPage_Click(Index As Integer)
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(1, mlng病人ID, mlng主页ID, mobjReport, mlng科室ID, Me, Index)
End Sub

Private Sub menuPagePrint_Click(Index As Integer)
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(2, mlng病人ID, mlng主页ID, mobjReport, mlng科室ID, Me, Index)
End Sub

Private Sub cmdSign_Click(Index As Integer)
'功能：签名
    Dim StrSQL As String
    Dim rsTmp As Recordset, i As Long
    Dim bln手术 As Boolean    '是否填写了手术记录
    
    '判断是否启用数字签名
    If gintCA > 0 And Mid(gstrESign, 2, 1) = "1" Then
        If mobjESign Is Nothing Then
            On Error Resume Next
            Set mobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not mobjESign Is Nothing Then
                Call mobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If mobjESign Is Nothing Then
                MsgBox "电子签名部件未能正确安装，签名操作不能继续。", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not mobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '需要确定更高签名级别的人
    If Index = cmd住院医师 Or Index = cmd主治医师 Or Index = cmd主任医师 Then
        If cboinfo(cbo科主任).Text = "" Then
            Call ShowMessage(cboinfo(cbo科主任), "没有确定科主任。")
            Exit Sub
        End If
    End If
    If Index = cmd住院医师 Or Index = cmd主治医师 Then
        If cboinfo(cbo主任医师).Text = "" Then
            Call ShowMessage(cboinfo(cbo主任医师), "没有确定主任医师。")
            Exit Sub
        End If
    End If
    If Index = cmd住院医师 Then
        If cboinfo(cbo主治医师).Text = "" Then
            Call ShowMessage(cboinfo(cbo主治医师), "没有确定主治医师。")
            Exit Sub
        End If
    End If
    
    '签名前自动保存
    If mblnChange Then
        If Not CheckPageData(False, True) Then Exit Sub
        If Not SavePageData(True) Then Exit Sub
    End If
    
    On Error GoTo errH
    
    '如果有手术记录，则提示是否继续
    bln手术 = False
    For i = 1 To vsOPS.Rows - 1
        If Trim(vsOPS.TextMatrix(i, col手术名称)) <> "" Then
            bln手术 = True
        End If
    Next
    
    StrSQL = "Select Count(1) As 手术 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] And 医嘱状态=8 And 诊疗类别='F'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Val(rsTmp!手术 & "") > 0 And Not bln手术 Then
        vsOPS.Row = 1: vsOPS.Col = col手术编码
        If ShowMessage(vsOPS, "该病人存在手术医嘱，但首页中没有添加手术记录，是否继续？", True) = vbNo Then Exit Sub
    End If
    
    If Index = cmd科主任 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'科主任签名','" & UserInfo.姓名 & "')"
    ElseIf Index = cmd主任医师 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'主任医师签名','" & UserInfo.姓名 & "')"
    ElseIf Index = cmd主治医师 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'主治医师签名','" & UserInfo.姓名 & "')"
    ElseIf Index = cmd住院医师 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'住院医师签名','" & UserInfo.姓名 & "')"
    End If
    Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    
    mblnReadOnly = SetSignature()
    Call SetFaceEditable(mblnReadOnly)
    If cmdOK.Enabled Then cmdOK.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdUnSign_Click(Index As Integer)
'功能：取消签名
    Dim StrSQL As String
    
    If gintCA > 0 And Mid(gstrESign, 2, 1) = "1" Then
        If mobjESign Is Nothing Then
            On Error Resume Next
            Set mobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not mobjESign Is Nothing Then
                Call mobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If mobjESign Is Nothing Then
                MsgBox "电子签名部件未能正确安装，签名操作不能继续。", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not mobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '并发检查病案是否编目或首页处于锁定状态
    If Not CheckMecRed(mlng病人ID, mlng主页ID, Me.Caption, "取消签名") Then Exit Sub
        
    On Error GoTo errH
    
    If Index = cmd科主任 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'科主任签名',Null)"
    ElseIf Index = cmd主任医师 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'主任医师签名',Null)"
    ElseIf Index = cmd主治医师 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'主治医师签名',Null)"
    ElseIf Index = cmd住院医师 Then
        StrSQL = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'住院医师签名',Null)"
    End If
    Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    
    mblnReadOnly = SetSignature()
    Call SetFaceEditable(mblnReadOnly)
    If cmdOK.Enabled Then cmdOK.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dpkInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpInfo_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If mlngDateIndex = txt确诊日期 Then
        If IsDate(txtInfo(txt确诊日期).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt确诊日期).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        If Not CheckDateRange(strDate, True) Then
            MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, Me.Caption
            Exit Sub
        End If
    ElseIf mlngDateIndex = txt质控日期 Then
        strDate = Format(DateClicked, "yyyy-MM-dd")
    End If
    txtInfo(mlngDateIndex).Text = strDate
    dtpInfo.Visible = False
    txtInfo(mlngDateIndex).SetFocus
    
    If Visible Then mblnChange = True
End Sub


Private Sub dtpInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call dtpInfo_DateClick(dtpInfo.Value)
    End If
End Sub

Private Sub dtpInfo_Validate(Cancel As Boolean)
    dtpInfo.Visible = False
End Sub

Private Sub Form_Activate()
    mblnIsFirst = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dtpInfo.Visible Then
            dtpInfo.Visible = False
        Else
            Call cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF1 Then
        '###
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim lngW As Long, lngH As Long
    Dim ctlTmp As Control
    Dim StrSQL As String
    Dim rsTmp As Recordset
    
    Me.Opened = True
    mblnIsFirst = False
    mstrPathDiag = ""
    '个性化设置保存后，以前的宽和高可能太小
    lngW = Me.Width
    lngH = Me.Height
    Call RestoreWinState(Me, App.ProductName)
    If lngW <> Me.Width Then Me.Width = lngW
    If lngH <> Me.Height Then Me.Height = lngH
    
    On Error Resume Next
    If Val(zlDatabase.GetPara("西医诊断输入", glngSys, p住院医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0)) = 0 Then
        optInput(0).Value = True
    Else
        optInput(1).Value = True
    End If
    If Val(zlDatabase.GetPara("中医诊断输入", glngSys, p住院医生站, 0, Array(optInput(2), optInput(3)), InStr(mstrPrivs, "参数设置") > 0)) = 0 Then
        optInput(2).Value = True
    Else
        optInput(3).Value = True
    End If
    mstr手术输入情况 = zlDatabase.GetPara("手术情况输入", glngSys, p住院医生站, 0, Array(optInput(4), optInput(5), chkInfo(chk手术自由录入)), InStr(mstrPrivs, "参数设置") > 0)
    If Mid(mstr手术输入情况, 1, 1) = "0" Then
        optInput(4).Value = True
    Else
        optInput(5).Value = True
    End If
    chkInfo(chk手术自由录入).Value = Val(Mid(mstr手术输入情况, 2, 1))
    
    mlng损伤中毒 = Val(zlDatabase.GetPara("损伤中毒检查", glngSys, p住院医生站, 2) & "")
    mlng病理诊断 = Val(zlDatabase.GetPara("病理诊断检查", glngSys, p住院医生站, 2) & "")
    mlng区域 = Val(zlDatabase.GetPara("区域检查", glngSys, p住院医生站, 1) & "")
    If InStr(mstrPrivs, "修改医疗付款方式") > 0 Then
        cboinfo(cbo付款方式).Enabled = True
    Else
        cboinfo(cbo付款方式).Enabled = False
    End If
    
    mbln启用结构化地址 = Val(zlDatabase.GetPara("病人地址结构化录入", glngSys, p住院医生站, 0)) <> 0
    mbln不使用西医项目 = Val(zlDatabase.GetPara("中医科室不使用西医病案首页项目", glngSys, p住院医生站, 0)) <> 0
    mbln医生护士分填首页 = Val(zlDatabase.GetPara("医生和护士分别填写病案首页", glngSys, p住院医生站, 0)) = 1
    
    If mbln启用结构化地址 Then
        txtInfo(txt出生地点).Visible = False
        txtInfo(txt籍贯).Visible = False
        txtInfo(txt户口地址).Visible = False
        txtInfo(txt家庭地址).Visible = False
        cmdInfo(txt出生地点).Visible = False
        cmdInfo(txt籍贯).Visible = False
        cmdInfo(txt户口地址).Visible = False
        cmdInfo(txt家庭地址).Visible = False
    Else
        PatiAddress出生地.Visible = False
        PatiAddress户口地址.Visible = False
        PatiAddress籍贯.Visible = False
        PatiAddress现住址.Visible = False
    End If
    

    Call optInput_Click(0)
    On Error GoTo 0
    
    '诊断输入来源
    If gint诊断来源 > 1 Then
        optInput(0).Enabled = False
        optInput(1).Enabled = False
        optInput(2).Enabled = False
        optInput(3).Enabled = False
        If gint诊断来源 = 2 Then
            optInput(0).Value = True
            optInput(2).Value = True
        ElseIf gint诊断来源 = 3 Then
            optInput(1).Value = True
            optInput(3).Value = True
        End If
    End If
    
    mblnOk = False
    mblnChange = False
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    
    '卡片设置
    If Not Have部门性质(mlng科室ID, "产科") Then
        vsOPS.ColHidden(col助产护士) = True
    End If
    mbln中医 = Have部门性质(mlng科室ID, "中医科")
    If Not mbln中医 Then
        sstInfo.TabVisible(TAB_中医诊断) = False
    End If
    For i = 0 To sstInfo.Tabs - 1
        fraInfo(i).BackColor = Me.BackColor
    Next
    If mbln护士站 Then
        sstInfo.Tab = TAB_其他
    Else
        sstInfo.Tab = TAB_基本信息
    End If
    
    '放疗化疗
    mbln病案共享 = CheckShare(300) '病案系统
    If Not mbln病案共享 Then
        sstInfo.TabVisible(TAB_放疗与化疗) = False
        lblInfo(107).Visible = False
    End If
    StrSQL = "select 信息值 from 病案主页从表 where 病人id=0 and 主页id=0"
    On Error Resume Next
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    mlngSize = rsTmp.Fields("信息值").DefinedSize
    
    
    vs放疗.Tag = "未修改"
    vs化疗.Tag = "未修改"
    vsfMain.Tag = "未修改"
    
    '初始化数据
    If Not InitPageData Then Unload Me: Exit Sub
    '读取首页内容
    If Not LoadPageData Then Unload Me: Exit Sub
    
    Call SetEditableFrom出院情况
    Call Set病原学
    
    '缺省定位到诊断页
    If mblnDiagnose Then
        If mbln中医 Then
            sstInfo.Tab = TAB_中医诊断
        Else
            sstInfo.Tab = TAB_西医诊断
        End If
    End If

    '设置签名级别情况、只读情况
    If mblnReadOnly Then
        Call SetSignature
        Call SetFaceEditable(True)
        '签名处理部份单独设置
        cboinfo(cbo科主任).Locked = True: cboinfo(cbo主任医师).Locked = True
        cboinfo(cbo主治医师).Locked = True: cboinfo(cbo住院医师).Locked = True
        cboinfo(cbo科主任).BackColor = vbButtonFace: cboinfo(cbo主任医师).BackColor = vbButtonFace
        cboinfo(cbo主治医师).BackColor = vbButtonFace: cboinfo(cbo住院医师).BackColor = vbButtonFace
        For i = 0 To cmdSign.UBound
            cmdSign(i).Visible = False: cmdUnSign(i).Visible = False
        Next
    Else
        '医生站需判断签名级别
        If Not mbln护士站 Then
            mblnReadOnly = SetSignature
        End If
        Call SetFaceEditable(mblnReadOnly)
    End If
        
    '没有年龄有出生日期时计算一下年龄,只读或已出院时不重算年龄
    If txtInfo(txt年龄).Text = "" And IsDate(txt出生日期.Text) _
        And Not (mblnReadOnly Or IsDate(txtInfo(txt出院时间).Text)) Then '只读或已出院时不自动重算年龄
        txt出生日期.Tag = "": Call txt出生日期_Validate(False)
    End If
End Sub

Private Sub PatiAddress出生地_Validate(Cancel As Boolean)
    If PatiAddress出生地.Tag <> PatiAddress出生地.Value Then mblnChange = True
End Sub

Private Sub PatiAddress户口地址_Validate(Cancel As Boolean)
    If PatiAddress户口地址.Tag <> PatiAddress户口地址.Value Then mblnChange = True
End Sub

Private Sub PatiAddress籍贯_Validate(Cancel As Boolean)
    If PatiAddress籍贯.Tag <> PatiAddress籍贯.Value Then mblnChange = True
End Sub

Private Sub PatiAddress现住址_Validate(Cancel As Boolean)
    If PatiAddress现住址.Tag <> PatiAddress现住址.Value Then mblnChange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("如果退出，刚才所修改的内容将不会被保存。确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    If Not mobjESign Is Nothing Then Set mobjESign = Nothing
    Set mcol人员SQL = Nothing
    
    Call zlDatabase.SetPara("西医诊断输入", IIf(optInput(0).Value, 0, 1), glngSys, p住院医生站, InStr(mstrPrivs, "参数设置") > 0)
    Call zlDatabase.SetPara("中医诊断输入", IIf(optInput(2).Value, 0, 1), glngSys, p住院医生站, InStr(mstrPrivs, "参数设置") > 0)
    Call zlDatabase.SetPara("手术情况输入", IIf(optInput(4).Value, "0", "1") & IIf(chkInfo(chk手术自由录入).Value = 1, "1", "0"), glngSys, p住院医生站, InStr(mstrPrivs, "参数设置") > 0)
    Call SaveWinState(Me, App.ProductName)
        
    Me.Opened = False
    RaiseEvent Closed(Not mblnOk, mstr疾病ID, mstr诊断ID)
End Sub

Private Sub fra主页_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

End Sub

Private Sub lstAdvEvent_GotFocus()
    lstAdvEvent.ListIndex = 0
End Sub

Private Sub lstAdvEvent_ItemCheck(Item As Integer)
    If lstAdvEvent.List(Item) = "压疮" Then
        cboinfo(cbo压疮发生期间).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cbo压疮发生期间).TabStop = cboinfo(cbo压疮发生期间).Enabled
        cboinfo(cbo压疮分期).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cbo压疮分期).TabStop = cboinfo(cbo压疮分期).Enabled
        If cboinfo(cbo压疮发生期间).Enabled Then
            cboinfo(cbo压疮发生期间).BackColor = vbWindowBackground
            cboinfo(cbo压疮分期).BackColor = vbWindowBackground
        Else
            cboinfo(cbo压疮发生期间).BackColor = vbButtonFace
            cboinfo(cbo压疮分期).BackColor = vbButtonFace
        End If
    ElseIf lstAdvEvent.List(Item) = "医院内跌倒/坠床" Then
        cboinfo(cbo跌倒或坠床伤害).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cbo跌倒或坠床伤害).TabStop = cboinfo(cbo跌倒或坠床伤害).Enabled
        cboinfo(cbo跌倒或坠床原因).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cbo跌倒或坠床原因).TabStop = cboinfo(cbo跌倒或坠床原因).Enabled
        If cboinfo(cbo跌倒或坠床伤害).Enabled Then
            cboinfo(cbo跌倒或坠床原因).BackColor = vbWindowBackground
            cboinfo(cbo跌倒或坠床伤害).BackColor = vbWindowBackground
        Else
            cboinfo(cbo跌倒或坠床原因).BackColor = vbButtonFace
            cboinfo(cbo跌倒或坠床伤害).BackColor = vbButtonFace
        End If
    End If
    If mblnIsFirst Then mblnChange = True
End Sub

Private Sub lstAdvEvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lstAdvEvent.ListIndex = lstAdvEvent.ListCount - 1 Then
            If cboinfo(cbo跌倒或坠床伤害).Enabled Or cboinfo(cbo压疮发生期间).Enabled Then
                If cboinfo(cbo压疮发生期间).Enabled Then
                    cboinfo(cbo压疮发生期间).SetFocus
                Else
                    cboinfo(cbo跌倒或坠床伤害).SetFocus
                End If
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            lstAdvEvent.ListIndex = lstAdvEvent.ListIndex + 1
        End If
    End If
End Sub

Private Sub lstInfection_GotFocus()
    lstInfection.ListIndex = 0
End Sub

Private Sub lstInfection_ItemCheck(Item As Integer)
    If mblnIsFirst Then mblnChange = True
End Sub

Private Sub lstInfection_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lstInfection.ListIndex = lstInfection.ListCount - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        lstInfection.ListIndex = lstInfection.ListIndex + 1
    End If
End Sub

Private Sub optInput_Click(Index As Integer)
    Dim i As Integer

    If Index = opt31天无 Then
        txtInfo(txt31天目的).Enabled = False
        txtInfo(txt31天目的).BackColor = &H8000000F
    ElseIf Index = opt31天有 Then
        txtInfo(txt31天目的).Enabled = True
        txtInfo(txt31天目的).BackColor = &H80000005
    End If
End Sub

Private Sub optInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub sstInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Index = txt年龄 Then
        '数字年龄才带标准年龄单位
        If IsNumeric(txtInfo(Index).Text) Or txtInfo(Index).Text = "" Then
            cboinfo(cbo年龄单位).Visible = True
            If cboinfo(cbo年龄单位).ListIndex = -1 Then cboinfo(cbo年龄单位).ListIndex = 0
        Else
            cboinfo(cbo年龄单位).Visible = False
            cboinfo(cbo年龄单位).ListIndex = -1
        End If
    ElseIf Index = txt婴儿年龄 Then
        '数字年龄才带标准年龄单位
        If IsNumeric(txtInfo(Index).Text) Or txtInfo(Index).Text = "" Then
            cboinfo(cbo婴儿年龄单位).Visible = True
            If cboinfo(cbo婴儿年龄单位).ListIndex = -1 Then cboinfo(cbo婴儿年龄单位).ListIndex = 0
        Else
            cboinfo(cbo婴儿年龄单位).Visible = False
            cboinfo(cbo婴儿年龄单位).ListIndex = -1
        End If
    ElseIf Index = txt抢救次数 Then
        If Val(txtInfo(Index).Text) > 0 Then
            txtInfo(txt成功次数).Locked = False
            txtInfo(txt成功次数).TabStop = True
            txtInfo(txt成功次数).BackColor = vbWindowBackground
            
            '主要诊断的出院情况不为死亡时,缺省：成功次数=抢救次数
            If Visible Then
                If vsDiagXY.TextMatrix(GetRow(3), col出院情况) <> "死亡" Then
                    txtInfo(txt成功次数).Text = txtInfo(txt抢救次数).Text
                ElseIf Val(txtInfo(txt抢救次数).Text) > 1 Then
                    txtInfo(txt成功次数).Text = Val(txtInfo(txt抢救次数).Text) - 1
                End If
            End If
        Else
            txtInfo(txt成功次数).Text = ""
            txtInfo(txt成功次数).Locked = True
            txtInfo(txt成功次数).TabStop = False
            txtInfo(txt成功次数).BackColor = vbButtonFace
        End If
    ElseIf Index = txt转科1 Then
        If txtInfo(Index).Text = "" Then
            txtInfo(txt转科2).Text = ""
            txtInfo(txt转科3).Text = ""
        End If
    ElseIf Index = txt转科2 Then
        If txtInfo(Index).Text = "" Then
            txtInfo(txt转科3).Text = ""
        End If
    End If
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Or (KeyCode = vbKeyDown And Shift = vbAltMask) Then
        If Index = txt确诊日期 Then
            Call cmdInfo_Click(Index)
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Index = txt医学警示 Then
            txtInfo(txt医学警示) = ""
        End If
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt出生地点 Or Index = txt家庭地址 Or Index = txt联系人地址 Or Index = txt户口地址) And txtInfo(Index).Text <> "" Then
            '输入地区数据
            StrSQL = "Select Rownum as ID,编码,名称,简码 From 地区 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "地区", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!名称
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf (Index = txt区域 Or Index = txt籍贯) And txtInfo(Index).Text <> "" Then
            '输入区域数据
            StrSQL = "Select Rownum as ID,编码,名称,简码 From 区域 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(Index = txt区域, "区域", "籍贯"), False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!名称
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt单位名称 And txtInfo(Index).Text <> "" Then
            '输入工作单位
            StrSQL = "Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " And (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "工作单位", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                txtInfo(Index).Tag = Val(rsTmp!ID)
                If txtInfo(txt单位电话).Text = "" Then
                    txtInfo(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
            Else
                txtInfo(Index).Tag = ""
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf (Index = txt转科1 Or Index = txt转科2 Or Index = txt转科3) And txtInfo(Index).Text <> "" Then
            '输入转科科室
            StrSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码,A.位置" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 IN(2,3) And B.工作性质 IN('临床','手术')" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "转科科室", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!名称
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt住院号 Then
            If txtInfo(Index).Text = "" Then
                txtInfo(Index).Text = zlDatabase.GetNextNo(2)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt出院病室 Or Index = txt抢救原因 Then
            If Index = txt抢救原因 Then
                 '选择单位信息
                If txtInfo(Index).Text <> "" Then
                    StrSQL = "Select 编码 ID,名称,简码 From 抢救病因分类 where 名称 like [1] or 简码 like [2] or to_number(编码)=[3]"
                       
                    vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "抢救原因", True, 1, "抢救病因", False, False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, gstrLike & txtInfo(Index).Text & "%", gstrLike & txtInfo(Index).Text & "%", Val(txtInfo(Index).Text))
                    If Not rsTmp Is Nothing Then
                        txtInfo(Index).Text = rsTmp!名称
                        txtInfo(Index).SetFocus
                    Else
                        Exit Sub
                    End If
                End If
            End If
            '跳到下一个卡片
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            If Index = txt抢救原因 And Not mbln中医 Then
                vsAller.SetFocus
            Else
                Call sstInfo_KeyPress(13)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Index = txt医学警示 Then
            txtInfo(txt医学警示).Text = ""
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '非控制按键
        If Index = txt医学警示 Then
            KeyAscii = 0
        End If
        '选择快捷键
        If KeyAscii = Asc("*") Then
            '注意界面上要求CMD和对应TXT的Index相同
            On Error Resume Next
            StrSQL = ""
            StrSQL = cmdInfo(Index).Name
            Err.Clear: On Error GoTo 0
            If StrSQL <> "" Then
                KeyAscii = 0
                Call cmdInfo_Click(Index)
                Exit Sub
            End If
        End If
        
        '限制输入长度
        If txtInfo(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtInfo(Index).Text) > txtInfo(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '限制输入内容
        Select Case Index
'            Case txt年龄 '允许自由录入了
'                strMask = "1234567890"
            'Case txt出生日期 'MaskEdit限制了
                'strMask = "1234567890-"
            Case txt确诊日期, txt质控日期
                strMask = "1234567890-: "
            Case txt家庭电话, txt单位电话, txt联系人电话
                strMask = "1234567890-()"
            Case txt住院号, txt户口邮编, txt家庭邮编, txt单位邮编, txt抢救次数, txt成功次数, txt随诊期限, txt入院前小时, txt入院前分钟, txt入院后小时, txt入院后分钟, txt呼吸机小时
                strMask = "1234567890"
            Case txt输红细胞, txt输血小板, txt输血浆, txt输全血, txt自体回收, txt约束总时间
                strMask = "1234567890."
            Case txt身高, txt体重, txt新生儿体重, txt新生儿入院体重
                strMask = "1234567890."
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    Else
        If Index = txt医学警示 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox索引数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        cboinfo(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboinfo(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboinfo(arrCboIdx(i)).ListIndex = intDefault '缺省为未选中
    Next
End Sub

Private Sub SetCboFromSQL(ByVal StrSQL As String, ByVal arrCboIdx As Variant, Optional ByVal strSQLExt As String, Optional colsql As Collection)
'功能：将指定数据源中的数据装入指定索引的一个或多个ComboBox
'参数：strSQL=包含"ID,简码,名称/姓名,缺省标志/缺省"字段，包含Order by，主表别名为A
'      strSQLExt=附加的SQL条件
'      colSQL=要加入SQL的集合
    Dim rsTmp As New ADODB.Recordset
    Dim str名称 As String, str缺省 As String
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        '清除原有数据
        cboinfo(arrCboIdx(i)).Clear

        '记录原始SQL
        If Not colsql Is Nothing Then
            colsql.Add StrSQL, "_" & arrCboIdx(i)
        End If
    Next
    
    If strSQLExt <> "" Then
        StrSQL = Replace(UCase(StrSQL), UCase("Order by"), strSQLExt & " Order by")
    End If
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, StrSQL, Me.Caption)
    
    '装入数据
    If Not rsTmp.EOF Then
        For i = 0 To rsTmp.Fields.Count - 1
            If rsTmp.Fields(i).Name = "名称" Or rsTmp.Fields(i).Name = "姓名" Then
                str名称 = rsTmp.Fields(i).Name
            ElseIf rsTmp.Fields(i).Name = "缺省标志" Or rsTmp.Fields(i).Name = "缺省" Then
                str缺省 = rsTmp.Fields(i).Name
            End If
        Next
        For i = 1 To rsTmp.RecordCount
            For j = 0 To UBound(arrCboIdx)
                If IsNull(rsTmp!简码) Then
                    cboinfo(arrCboIdx(j)).AddItem rsTmp.Fields(str名称).Value
                Else
                    cboinfo(arrCboIdx(j)).AddItem rsTmp!简码 & "-" & Chr(13) & rsTmp.Fields(str名称).Value
                End If
                cboinfo(arrCboIdx(j)).ItemData(cboinfo(arrCboIdx(j)).NewIndex) = Nvl(rsTmp!ID, 0)
                If str缺省 <> "" Then
                    If Nvl(rsTmp.Fields(str缺省).Value, 0) = 1 Then
                        Call zlControl.CboSetIndex(cboinfo(arrCboIdx(j)).hwnd, cboinfo(arrCboIdx(j)).NewIndex)
                    End If
                End If
            Next
            rsTmp.MoveNext
        Next
    End If
    
    '无缺省时,为未选中
    For i = 0 To UBound(arrCboIdx)
        If cboinfo(arrCboIdx(i)).Style = 0 Then
            cboinfo(arrCboIdx(i)).AddItem "[其他...]"
            cboinfo(arrCboIdx(i)).ItemData(cboinfo(arrCboIdx(i)).NewIndex) = -1
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function SetCboFromName(ByVal strName As String, objCbo As Object) As Boolean
'功能：将指定姓名的人员加入到下拉框中
    Static rsTmp As ADODB.Recordset
    Dim StrSQL As String, intIdx As Integer
    
    On Error GoTo errH
    
    If rsTmp Is Nothing Then
        StrSQL = "Select A.ID,A.编号,A.姓名,Null As 简码" & _
            " From 人员表 A,人员性质说明 B" & _
            " Where A.ID=B.人员ID And B.人员性质 IN('医生','护士')" & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.姓名"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, StrSQL, "SetCboFromName")
    End If
    
    rsTmp.Filter = "姓名='" & strName & "'"
    If Not rsTmp.EOF Then
        intIdx = objCbo.ListCount
        If objCbo.ListCount > 0 Then
            If objCbo.ItemData(objCbo.ListCount - 1) = -1 Then
                intIdx = objCbo.ListCount - 1
            End If
        End If
        
        If IsNull(rsTmp!简码) Then
            objCbo.AddItem rsTmp!姓名, intIdx
        Else
            objCbo.AddItem rsTmp!简码 & "-" & Chr(13) & rsTmp!姓名, intIdx
        End If
        objCbo.ItemData(objCbo.NewIndex) = Val(rsTmp!ID)
        
        objCbo.ListIndex = objCbo.NewIndex
    End If
    
    SetCboFromName = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitPageData() As Boolean
'功能：初始化首页编辑时所需要的一些数据
    Dim StrSQL As String, strSQLExt As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    '设置部分下拉框的高度的宽度
    Call zlControl.CboSetWidth(cboinfo(cbo职业).hwnd, cboinfo(cbo职业).Width + 500)
    Call zlControl.CboSetWidth(cboinfo(cbo国籍).hwnd, cboinfo(cbo国籍).Width * 2)
    Call zlControl.CboSetHeight(cboinfo(cbo民族), cboinfo(cbo民族).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo国籍), cboinfo(cbo国籍).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo职业), cboinfo(cbo职业).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo门诊医师), cboinfo(cbo门诊医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo科主任), cboinfo(cbo科主任).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo主任医师), cboinfo(cbo主任医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo主治医师), cboinfo(cbo主治医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo住院医师), cboinfo(cbo住院医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo进修医师), cboinfo(cbo进修医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo研究生医师), cboinfo(cbo研究生医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo实习医师), cboinfo(cbo实习医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo质控医师), cboinfo(cbo质控医师).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo质控护士), cboinfo(cbo质控护士).Height * 16)
    
    '部分固定内容的下拉框
    Call SetCboFromList(Array("未带", "遗失待办", "未办"), Array(cbo身份证号), 0)
    Call SetCboFromList(Array("岁", "月", "天", "小时", "分钟"), Array(cbo年龄单位), 0) '添加项目时请注意cboInfo(cbo年龄单位).listIndex<3的判断
    Call SetCboFromList(Array("天", "周", "月", "年", "终身"), Array(cbo随诊Ex), 0)
    Call SetCboFromList(Array("0-未查", "1-阴", "2-阳", "3-不详"), Array(cboRh))
    Call SetCboFromList(Array("1.1-中", "1.2-民族", "2-中西", "3-西"), Array(cbo治疗类别, cbo抢救方法))
    Call SetCboFromList(Array("0-未知", "1-有", "2-无"), Array(cbo自制中药))
    Call SetCboFromList(Array(" ", "1-是", "2-否"), Array(cbo使用中医诊疗设备))
    Call SetCboFromList(Array(" ", "1-是", "2-否"), Array(cbo使用中医诊疗技术))
    Call SetCboFromList(Array(" ", "1-是", "2-否"), Array(cbo辨证施护))
    Call SetCboFromList(Array("0-未做", "1-准确", "2-基本准确", "3-重大缺陷", "4-错误"), Array(cbo辨证, cbo治法, cbo方药))
    Call SetCboFromList(Array("0-未做", "1-阴性", "2-阳性", "3-弱阳性"), Array(cboHBsAg))
    Call SetCboFromList(Array("0-未做", "1-阴性", "2-阳性"), Array(cboHCVAb, cboHIVAb))
    Call SetCboFromList(Array("1-有", "2-无", "3-未输"), Array(cbo输液反应))
    Call SetCboFromList(Array("0-无", "1-有", "2-未输"), Array(cbo输血反应))
    Call SetCboFromList(Array("1-是", "2-否", "3-部分"), Array(cbo输血检查))
    Call SetCboFromList(Array("0-未做", "1-符合", "2-不符合", "3-不肯定"), Array(cbo门诊与出院, cbo门诊与入院, cbo入院与出院, cbo放射与病理, cbo临床与病理, cbo临床与尸检, cbo术前与术后, cbo中医门诊与出院, cbo中医入院与出院))
    Call SetCboFromList(Array(" ", "0-入院前", "1-住院期间"), Array(cbo压疮发生期间))
    Call SetCboFromList(Array(" ", "1期", "2期", "3期", "4期"), Array(cbo压疮分期))
    Call SetCboFromList(Array(" ", "一级", "二级", "三级", "未造成伤害"), Array(cbo跌倒或坠床伤害))
    Call SetCboFromList(Array(" ", "健康原因", "治疗、药物、麻醉原因", "环境因素", "其他原因"), Array(cbo跌倒或坠床原因))
    Call SetCboFromList(Array("月", "天", "小时", "分钟"), Array(cbo婴儿年龄单位))
    Call SetCboFromList(Array("31天内再住院计划", "7天内再住院计划"), Array(cbo31天和7天再入院))
    Call SetCboFromList(Array("", "1-甲", "2-乙", "3-丙"), Array(cbo病案质量))
    Call SetCboFromList(Array("", "一处", "两处", "三处", "其他"), Array(cbo约束方式))
    Call SetCboFromList(Array("", "软式管", "硬式管", "背心", "老人椅", "约束带", "其他"), Array(cbo约束工具))
    Call SetCboFromList(Array("", "认知障碍", "可能跌倒", "行为紊乱", "治疗需要", "躁动", "医疗限制", "其他"), Array(cbo约束原因))
    Call SetCboFromList(Array("", "医嘱出院", "转儿科", "转院", "非医嘱出院", "死亡"), Array(cbo新生儿离院方式))
    Call SetCboFromList(Array("非重返", "24h内", "24-48h", "＞48h"), Array(cbo重返间隔时间))
    Call SetCboFromList(Array("", "0-未生育", "1-生育1胎", "2-生育2胎及以上", "4-不详"), Array(cbo生育状况), 0)
    cboinfo(cbo31天和7天再入院).ListIndex = 0
    cboinfo(cbo婴儿年龄单位).ListIndex = 0
    cboinfo(cbo婴儿年龄单位).ListIndex = 0
    
    '根据一些字典设置下拉框内容
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 医疗付款方式 Order by 编码", Array(cbo付款方式))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 性别 Order by 编码", Array(cbo性别))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 婚姻状况 Order by 编码", Array(cbo婚姻))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 职业 Order by 编码", Array(cbo职业))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 民族 Order by 编码", Array(cbo民族))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 国籍 Order by 编码", Array(cbo国籍))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 血型 Order by 编码", Array(cbo血型))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 社会关系 Order by 编码", Array(cbo联系人关系))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,0 as 缺省标志 From 病情 Order by 编码", Array(cbo入院病情))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,0 as 缺省标志 From 临床病例分型 Order by 编码", Array(cbo病例分型))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 入院方式 Order by 编码", Array(cbo入院方式))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 分化程度 Order by 编码", Array(cbo分化程度))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 最高诊断依据 Order by 编码", Array(cbo最高诊断依据))
    cboinfo(cbo病例分型).AddItem " "
    
    '医生、护士数据----------------------------------------------------------------
    Set mcol人员SQL = New Collection
    strSQLExt = " And Exists(Select 1 From 部门人员 Where 人员ID=A.ID And 部门ID IN(Select B.部门ID From 上机人员表 A,部门人员 B Where A.用户名=User And A.人员ID=B.人员ID))"
    
    '门诊医生
    StrSQL = "Select Distinct A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B,部门人员 C,部门性质说明 D" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And A.ID=C.人员ID And C.部门ID=D.部门ID And D.服务对象 IN(1,2,3)" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo门诊医师), , mcol人员SQL)
    
    '医生
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo进修医师, cbo研究生医师, cbo实习医师, cbo质控医师), strSQLExt, mcol人员SQL)
    
    '住院医师
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And A.专业技术职务 IN('主任医师','副主任医师','主治医师','医师','医士')" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo住院医师), strSQLExt, mcol人员SQL)
    
    '主治医师
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And A.专业技术职务 IN('主任医师','副主任医师','主治医师')" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo主治医师), strSQLExt, mcol人员SQL)
    
    '主任医师
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And A.专业技术职务 IN('主任医师','副主任医师')" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo主任医师), strSQLExt, mcol人员SQL)
    
    '科主任
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And A.管理职务 IN('科室主任','科室副主任')" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo科主任), strSQLExt, mcol人员SQL)
    
    '质控护士
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='护士'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo质控护士), strSQLExt, mcol人员SQL)
    
    '责任护士
    StrSQL = "Select A.ID,A.编号,A.姓名,Null as 简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='护士'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.姓名"
    Call SetCboFromSQL(StrSQL, Array(cbo责任护士), strSQLExt, mcol人员SQL)
    
    '出院方式
    cboinfo(cbo出院方式).Clear
    StrSQL = "select 编码 AS ID,名称,'' 简码,缺省标志 from 出院方式 order by 编码"
    Call SetCboFromSQL(StrSQL, Array(cbo出院方式))
    
    '-------------------
    Call SetKSSSerial
    Call vsKSS_AfterRowColChange(-1, -1, vsKSS.Row, vsKSS.Col)
    Call vsTSJC_AfterRowColChange(-1, -1, vsTSJC.Row, vsTSJC.Col)
    
    If mbln病案共享 Then Call Init化疗与放疗Grid
    Call FillVsf

    Screen.MousePointer = 0
    InitPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Load附页内容(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:加载附页内容
    '参数:lng病人id-病人id
    '     lng主页id -主页id
    '返回:加载成功,返回true,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
     
    
    '加重症情况
    StrSQL = "" & _
        " Select 监护室名称,人工气道脱出,重返重症医学科," & _
        "      重返间隔时间 " & _
        " From 病案重症监护情况 " & _
        " where 病人id=[1] and 主页id=[2] " & _
        " order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTemp.RecordCount > 0 Then
        txtInfo(txt重症监护室).Text = rsTemp!监护室名称 & ""
        chkInfo(chk人工气道脱出).Value = Val(rsTemp!人工气道脱出 & "")
        chkInfo(chk重返重症医学科).Value = Val(rsTemp!重返重症医学科 & "")
        Call GetCboIndex(cboinfo(cbo重返间隔时间), Nvl(rsTemp!重返间隔时间))
    End If
    
    Load附页内容 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get治疗结果() As String
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
        
    On Error GoTo errH
    StrSQL = "Select 编码,名称,简码 From 治疗结果 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, StrSQL, Me.Caption)
    
    StrSQL = ""
    Do While Not rsTmp.EOF
        StrSQL = StrSQL & "|" & rsTmp!编码 & "-" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If StrSQL = "" Then
        Get治疗结果 = "1-治愈|2-好转|3-未愈|4-死亡|5-其他"
    Else
        Get治疗结果 = Mid(StrSQL, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Init化疗与放疗Grid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化化疗与化疗网格控件的默认属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-10-21 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = Get化疗与放疗(True)
        
    With vs化疗
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        .Editable = flexEDKbdMouse

        .ColComboList(.ColIndex("化学治疗编码")) = .BuildComboList(rsTemp, "疾病信息", "ID")
        If rsTemp.RecordCount = 1 Then
            .ColData(.ColIndex("化学治疗编码")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!疾病信息)
        Else
            rsTemp.Filter = "缺省标志=1"
            If rsTemp.EOF = False Then
                .ColData(.ColIndex("化学治疗编码")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!疾病信息)
            Else
                .ColData(.ColIndex("化学治疗编码")) = ";"
                lblEdit(2).Caption = "没有可用的化疗治疗编码，请到病案系统中设置。"
            End If
        End If
        Call vs化疗_LostFocus
        zl_vsGrid_Para_Restore glngModul, vs化疗, Me.Caption, "化疗"
    End With
    Set rsTemp = Get化疗与放疗(False)
    With vs放疗
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("放射治疗编码")) = .BuildComboList(rsTemp, "疾病信息", "ID")
        If rsTemp.RecordCount = 1 Then
            .ColData(.ColIndex("放射治疗编码")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!疾病信息)
        Else
            rsTemp.Filter = "缺省标志=1"
            If rsTemp.EOF = False Then
                .ColData(.ColIndex("放射治疗编码")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!疾病信息)
            Else
                .ColData(.ColIndex("放射治疗编码")) = ";"
                lblEdit(2).Caption = "没有可用的放疗治疗编码，请到病案系统中设置。"
            End If
        End If
        Call vs放疗_LostFocus
        zl_vsGrid_Para_Restore glngModul, vs放疗, Me.Caption, "放疗"
    End With
End Sub

Private Function Get化疗与放疗(ByVal bln化疗 As Boolean, Optional ByVal arrControl As Variant, Optional blnSetup As Boolean = False) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化化疗与放疗的参数
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-10-21 10:37:13
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String, StrSQL As String
    Dim arrData As Variant, strDefaultCode As String, strCodeIN As String
    Dim rsTemp As New ADODB.Recordset, i As Long
    
    
    '获取放疗和化疗
    '   zlDatabase.SetPara IIf(bln放疗, "放疗项目", "化疗项目"), strSaveData, glngSys, mlngModule, False
    strTemp = zlDatabase.GetPara(IIf(Not bln化疗, "放疗项目", "化疗项目"), glngSys * 3, 200, , arrControl, blnSetup)
    If strTemp <> "" Then
        arrData = Split(strTemp, ";")
        For i = 0 To UBound(arrData)
            If InStr(1, arrData(i), ",") > 0 Then
                If Val(Split(arrData(i), ",")(1)) = 1 Then
                    strDefaultCode = Split(arrData(i), ",")(0)
                End If
                strCodeIN = strCodeIN & "," & Split(arrData(i), ",")(0)
            Else
                strCodeIN = strCodeIN & "," & arrData(i)
            End If
        Next
    End If
    If strCodeIN <> "" Then
        strCodeIN = Mid(strCodeIN, 2)
    Else
        strCodeIN = ";-"
    End If
    StrSQL = "" & _
    "   Select /*+ Rule*/ A.id,A.编码,A.编码||'-'||A.名称 as 疾病信息,decode(A.编码,[2],1,0) as 缺省标志 " & _
    "   From 疾病编码目录 A, " & _
    "       Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
    "   Where A.编码 = B.Column_Value"
    On Error GoTo errH
    Set Get化疗与放疗 = zlDatabase.OpenSQLRecord(StrSQL, "获取化疗与放疗信息", strCodeIN, strDefaultCode)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal strKEY As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从数据库中恢复网格的宽度等信息
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '     blnSaveToDataBase-是否是往数据库中保存参数(如果是往数据库中保存,则强制保存为true,否则根据是否使用个性化风格来确定)
    '     bln强制恢复保存-决定是否将保存注册表的参数值,进行强制恢复
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        '只有在本地注册表中才会处理个性化设置
        zl_vsGrid_Para_Restore = True
        If bln强制恢复保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g私有模块, strCaption, strKEY, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKEY, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function

Private Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKEY As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKEY, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKEY, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKEY, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKEY, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKEY, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKEY, "")
    End Select
Errhand:
End Sub

Private Function GetIDTmp(ByVal strName As String) As Long
'功能：由于现在将病案主页从表的抗生素 移到了新表 病人抗生素记录中，以前没有记录药品id，现在根据名称将id查出来
    Dim rsTmp As Recordset, StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select Distinct a.Id" & vbNewLine & _
                "From 诊疗项目目录 A, 诊疗项目别名 B, 药品特性 C" & vbNewLine & _
                "Where a.Id = b.诊疗项目id And a.Id = c.药名id And Nvl(c.抗生素, 0) <> 0 And A.名称=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strName)
    If rsTmp.RecordCount > 0 Then
        GetIDTmp = Val(rsTmp!ID)
    Else
        GetIDTmp = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadPageData() As Boolean
'功能：读取病人的首页信息
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim str治疗结果 As String, blnDo As Boolean
    Dim lngCol As Long
    Dim bln分化程度 As Boolean
    Dim strTmp As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If mlngPathState <> -1 Then
        '只处理首页中输入的诊断，以前没填的，缺省当作来自于“西医入院诊断”
        StrSQL = "Select Nvl(诊断类型,2) as 诊断类型,NVL(疾病ID,0) As 疾病ID,NVL(诊断ID,0) as 诊断ID,状态 From 病人临床路径 Where 病人ID=[1] And 主页ID=[2] And (诊断来源 = 3 or 诊断来源 is null) Order By 导入时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.RecordCount > 0 Then
            mlngDiagnosisType = rsTmp!诊断类型
            '如果有多条路径，则取第一条的状态
            If rsTmp.RecordCount >= 2 Then mlngPathState = Val(rsTmp!状态 & "")
            rsTmp.MoveNext
            Do While Not rsTmp.EOF
                mstrPathDiag = mstrPathDiag & "," & rsTmp!诊断类型 & "|" & rsTmp!疾病id & "|" & rsTmp!诊断id
                rsTmp.MoveNext
            Loop
            mstrPathDiag = Mid(mstrPathDiag, 2)
        Else
            mlngDiagnosisType = 0
        End If
        '完成路径的时间是否比出院诊断记录时间大()取第一条路径
        If mlngPathState = 2 Then
            StrSQL = "Select Sign(Nvl(a.结束时间, Null)-Nvl(b.记录日期, Sysdate)) As 判断" & vbNewLine & _
                    "From 病人临床路径 A, (Select 病人id, 主页id, 记录日期 From 病人诊断记录 Where 记录来源 = 3 And 诊断次序 = 1 And 诊断类型 = [3]) B" & vbNewLine & _
                    " Where a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人ID=[1] And A.主页ID=[2]" & _
                    " and a.导入时间=(Select Min(导入时间) From 病人临床路径 Where 病人ID=[1] and 主页ID=[2])"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID, IIf(mlngDiagnosisType > 10, 13, 3))
            If rsTmp.RecordCount > 0 Then
                mblnIsPathOutTime = Val(rsTmp!判断 & "") = 1
            Else
                mblnIsPathOutTime = False
            End If
        End If
    End If
    
    '病人信息部份
    '---------------------------------------------------------------
    StrSQL = "Select 姓名,性别,出生日期,出生地点,身份证号,其他证件,民族,区域,住院号,籍贯 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID)
        
    txtInfo(txt住院次数).Text = mlng主页ID
    txtInfo(txt姓名).Text = Nvl(rsTmp!姓名)
    Call GetCboIndex(cboinfo(cbo性别), Nvl(rsTmp!性别))
    
    If Not IsNull(rsTmp!出生日期) Then
        txt出生日期.Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
        If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
            txt出生时间.Text = Format(rsTmp!出生日期, "HH:mm")
        End If
    End If
    txt出生日期.Tag = txt出生日期.Text '用于记录输入变化
    If mbln启用结构化地址 Then
        '出生地
        Call SetStrucAddress(PatiAddress出生地, GetStrucAddress(mlng病人ID, mlng主页ID, 1), Nvl(rsTmp!出生地点))
        PatiAddress出生地.Tag = PatiAddress出生地.Value
        '籍贯
        Call SetStrucAddress(PatiAddress籍贯, GetStrucAddress(mlng病人ID, mlng主页ID, 2), Nvl(rsTmp!籍贯))
        PatiAddress籍贯.Tag = PatiAddress籍贯.Value
    Else
        txtInfo(txt出生地点).Text = Nvl(rsTmp!出生地点)
        txtInfo(txt籍贯).Text = Nvl(rsTmp!籍贯)
    End If
    cboinfo(cbo身份证号).Text = Nvl(rsTmp!身份证号)
    txtInfo(txt其他证件).Text = Nvl(rsTmp!其他证件)
    Call GetCboIndex(cboinfo(cbo民族), Nvl(rsTmp!民族))
    
    txtInfo(txt区域).Text = Nvl(rsTmp!区域)
    
    '病案主页部份
    '---------------------------------------------------------------
    StrSQL = "Select A.*,B.名称 as 入院科室,C.名称 as 出院科室" & _
        " From 病案主页 A,部门表 B,部门表 C" & _
        " Where A.入院科室ID=B.ID And A.出院科室ID=C.ID" & _
        " And A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    mint险类 = Nvl(rsTmp!险类, 0)
    
    '留观病人无住院号
    If Nvl(rsTmp!病人性质, 0) <> 0 Then
        lblInfo(0).Visible = False
        txtInfo(txt住院号).Visible = False
        txtInfo(txt住院号).Enabled = False '标志为不检查
    End If
    
    Call GetCboIndex(cboinfo(cbo付款方式), Nvl(rsTmp!医疗付款方式))
    
    Call LoadOldData("" & rsTmp!年龄)
    
    Call GetCboIndex(cboinfo(cbo婚姻), Nvl(rsTmp!婚姻状况))
    Call GetCboIndex(cboinfo(cbo职业), Nvl(rsTmp!职业))
    
    Call GetCboIndex(cboinfo(cbo国籍), Nvl(rsTmp!国籍))
    If Not IsNull(rsTmp!区域) Then
        txtInfo(txt区域).Text = Nvl(rsTmp!区域)
    End If
    txtInfo(txt住院号).Text = Nvl(rsTmp!住院号)
    If mbln启用结构化地址 Then
        '现住址
        Call SetStrucAddress(PatiAddress现住址, GetStrucAddress(mlng病人ID, mlng主页ID, 3), Nvl(rsTmp!家庭地址))
        PatiAddress现住址.Tag = PatiAddress现住址.Value
        '户口地址
        Call SetStrucAddress(PatiAddress户口地址, GetStrucAddress(mlng病人ID, mlng主页ID, 4), Nvl(rsTmp!户口地址))
        PatiAddress户口地址.Tag = PatiAddress户口地址.Value
    Else
        txtInfo(txt家庭地址).Text = Nvl(rsTmp!家庭地址)
        txtInfo(txt户口地址).Text = Nvl(rsTmp!户口地址)
    End If
    txtInfo(txt家庭电话).Text = Nvl(rsTmp!家庭电话)
    txtInfo(txt家庭邮编).Text = Nvl(rsTmp!家庭地址邮编)
    txtInfo(txt单位名称).Text = Nvl(rsTmp!单位地址)
    txtInfo(txt单位电话).Text = Nvl(rsTmp!单位电话)
    txtInfo(txt单位邮编).Text = Nvl(rsTmp!单位邮编)
    txtInfo(txt户口邮编).Text = Nvl(rsTmp!户口地址邮编)
    txtInfo(txt联系人姓名).Text = Nvl(rsTmp!联系人姓名)
    Call GetCboIndex(cboinfo(cbo联系人关系), Nvl(rsTmp!联系人关系))
    txtInfo(txt联系人电话).Text = Nvl(rsTmp!联系人电话)
    txtInfo(txt联系人地址).Text = Nvl(rsTmp!联系人地址)
    
    chkInfo(chk再入院).Value = Nvl(rsTmp!再入院, 0)
    txtInfo(txt入院时间).Text = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
    
    txtInfo(txt入院科室).Text = rsTmp!入院科室
    Call GetCboIndex(cboinfo(cbo入院病情), Nvl(rsTmp!入院病况))
    
    txtInfo(txt出院时间).Text = Format(Nvl(rsTmp!出院日期), "yyyy-MM-dd HH:mm")
    
    txtInfo(txt出院科室).Text = rsTmp!出院科室
    
    Call GetCboIndex(cboinfo(cbo出院方式), Nvl(rsTmp!出院方式))
    
    If Not IsNull(rsTmp!出院日期) Then
        txtInfo(txt住院天数).Text = DateDiff("d", rsTmp!入院日期, rsTmp!出院日期)
    Else
        txtInfo(txt住院天数).Text = DateDiff("d", rsTmp!入院日期, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txt住院天数).Text) = 0 Then txtInfo(txt住院天数).Text = "1"
    
    chkInfo(chk是否确诊).Value = Nvl(rsTmp!是否确诊, 0)
    If chkInfo(chk是否确诊).Value = 1 Then
        txtInfo(txt确诊日期).Text = Format(Nvl(rsTmp!确诊日期), "yyyy-MM-dd HH:mm")
    End If
    txtInfo(txt抢救次数).Text = Nvl(rsTmp!抢救次数)
    If Val(txtInfo(txt抢救次数).Text) <> 0 Then
        txtInfo(txt成功次数).Text = Nvl(rsTmp!成功次数)
    End If
    chkInfo(chk新发肿瘤).Value = Nvl(rsTmp!新发肿瘤, 0)
    Call GetCboIndex(cboinfo(cbo治疗类别), Nvl(rsTmp!中医治疗类别))
    chkInfo(chk尸检).Value = Nvl(rsTmp!尸检标志, 0)
    
    chkInfo(chk随诊).Value = IIf(Nvl(rsTmp!随诊标志, 0) = 0, 0, 1)
    If chkInfo(chk随诊).Value = 1 Then
        cboinfo(cbo随诊Ex).Text = Decode(Nvl(rsTmp!随诊标志, 0), 1, "月", 2, "年", 3, "周", 4, "天", 9, "终身")
        txtInfo(txt随诊期限).Text = IIf(Nvl(rsTmp!随诊标志, 0) = 9, "", Nvl(rsTmp!随诊期限, 0))
    End If
    
    Call GetCboIndex(cboinfo(cbo门诊医师), Nvl(rsTmp!门诊医师))
    If Not IsNull(rsTmp!门诊医师) And cboinfo(cbo门诊医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!门诊医师, cboinfo(cbo门诊医师))
    
    Call GetCboIndex(cboinfo(cbo住院医师), Nvl(rsTmp!住院医师))
    If Not IsNull(rsTmp!住院医师) And cboinfo(cbo住院医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!住院医师, cboinfo(cbo住院医师))
    
    Call GetCboIndex(cboinfo(cbo责任护士), Nvl(rsTmp!责任护士))
    If Not IsNull(rsTmp!责任护士) And cboinfo(cbo责任护士).ListIndex = -1 Then Call SetCboFromName(rsTmp!责任护士, cboinfo(cbo责任护士))
    
    '兼容老数据  未知 读为 不详
    Call GetCboIndex(cboinfo(cbo血型), IIf(Nvl(rsTmp!血型) = "未知", "不详", Nvl(rsTmp!血型)))
    '身高体重
    txtInfo(txt身高).Text = IIf(rsTmp!身高 & "" = "0", "", rsTmp!身高 & "")
    txtInfo(txt体重).Text = IIf(rsTmp!体重 & "" = "0", "", rsTmp!体重 & "")
    
    '入院方式
    Call GetCboIndex(cboinfo(cbo入院方式), Nvl(rsTmp!入院方式))
    
    '入科时间
    If Nvl(rsTmp!状态, 0) = 1 Then
        txtInfo(txt入科时间).Text = "尚未入科"
    Else
        StrSQL = "Select 开始时间 From 病人变动记录" & _
            " Where 病人ID=[1] And 主页ID=[2] And 开始原因 IN(2,1) And 开始时间 is Not Null Order by 开始原因 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then
            txtInfo(txt入科时间).Text = Format(rsTmp!开始时间, "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '病案从表部份
    '---------------------------------------------------------------
    StrSQL = "Select a.病人ID,a.主页ID,a.信息名,a.信息值,b.编码 From 病案主页从表 a " & _
            ",病案项目 b" & " where a.信息名=b.名称(+) And a.病人ID=[1] And a.主页ID=[2] Order by a.信息名"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(Nvl(rsTmp!信息名))
            Case "病例分型"
                Call GetCboIndex(cboinfo(cbo病例分型), Nvl(rsTmp!信息值))
                If cboinfo(cbo病例分型).ListIndex = -1 And Not IsNull(rsTmp!信息值) Then    '病案系统以前可能定义有不规范的值
                    cboinfo(cbo病例分型).AddItem rsTmp!信息值
                    cboinfo(cbo病例分型).ListIndex = cboinfo(cbo病例分型).NewIndex
                End If
            Case "入院病室"
                txtInfo(txt入院病室).Text = Nvl(rsTmp!信息值)
            Case "出院病室"
                txtInfo(txt出院病室).Text = Nvl(rsTmp!信息值)
            Case "转科记录"
                varTmp = Split(Nvl(rsTmp!信息值), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txt转科1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txt转科2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txt转科3).Text = varTmp(2)
            Case UCase("HBsAg")
                Call GetCboIndex(cboinfo(cboHBsAg), Nvl(rsTmp!信息值))
            Case UCase("HCV-Ab")
                Call GetCboIndex(cboinfo(cboHCVAb), Nvl(rsTmp!信息值))
            Case UCase("HIV-Ab")
                Call GetCboIndex(cboinfo(cboHIVAb), Nvl(rsTmp!信息值))
            Case "中医危重"
                chkInfo(chk危重).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "中医急症"
                chkInfo(chk急症).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "中医疑难"
                chkInfo(chk疑难).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "中医抢救方法"
                Call GetCboIndex(cboinfo(cbo抢救方法), Nvl(rsTmp!信息值))
            Case "自制中药制剂"
                Call GetCboIndex(cboinfo(cbo自制中药), Nvl(rsTmp!信息值))
            Case "死亡根本原因"
                txtInfo(txt死亡原因).Text = Nvl(rsTmp!信息值)
            Case "死亡时间"
                If IsNull(rsTmp!信息值) Then
                    txt死亡时间.Text = "____-__-__ __:__:__"
                ElseIf Not IsDate(rsTmp!信息值) Then
                    txt死亡时间.Text = "____-__-__ __:__:__"
                Else
                    txt死亡时间.Text = rsTmp!信息值
                End If
            Case "入院前经外院治疗"
                chkInfo(chk经外院治疗).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "示教病案"
                chkInfo(chk示教病案).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "科研病案"
                chkInfo(chk科研病案).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "疑难病历"
                chkInfo(chk疑难病例).Value = Val(Nvl(rsTmp!信息值))
            Case UCase("Rh")
                '兼容老数据，未做 改为 未查
                Call GetCboIndex(cboinfo(cboRh), IIf(Nvl(rsTmp!信息值) = "未做", "未查", Nvl(rsTmp!信息值)))
            Case "输血反应"
                cboinfo(cbo输血反应).ListIndex = Val(Nvl(rsTmp!信息值, 0))
            Case "输红细胞"
                txtInfo(txt输红细胞).Text = Nvl(rsTmp!信息值)
            Case "输血小板"
                txtInfo(txt输血小板).Text = Nvl(rsTmp!信息值)
            Case "输血浆"
                txtInfo(txt输血浆).Text = Nvl(rsTmp!信息值)
            Case "输全血"
                txtInfo(txt输全血).Text = Nvl(rsTmp!信息值)
            Case "输其他"
                txtInfo(txt输其他).Text = Nvl(rsTmp!信息值)
            Case "输液反应"
                Call GetCboIndex(cboinfo(cbo输液反应), Nvl(rsTmp!信息值))
            Case "医学警示"
                txtInfo(txt医学警示).Text = Nvl(rsTmp!信息值)
            Case "其他医学警示"
                txtInfo(txt其他医学警示).Text = Nvl(rsTmp!信息值)
            Case "科主任"
                Call GetCboIndex(cboinfo(cbo科主任), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo科主任).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo科主任))
            Case "主任医师"
                Call GetCboIndex(cboinfo(cbo主任医师), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo主任医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo主任医师))
            Case "主治医师"
                Call GetCboIndex(cboinfo(cbo主治医师), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo主治医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo主治医师))
            Case "进修医师"
                Call GetCboIndex(cboinfo(cbo进修医师), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo进修医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo进修医师))
            Case "研究生实习医师"
                Call GetCboIndex(cboinfo(cbo研究生医师), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo研究生医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo研究生医师))
            Case "实习医师"
                Call GetCboIndex(cboinfo(cbo实习医师), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo实习医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo实习医师))
            Case "质控医师"
                Call GetCboIndex(cboinfo(cbo质控医师), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo质控医师).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo质控医师))
            Case "质控护士"
                Call GetCboIndex(cboinfo(cbo质控护士), Nvl(rsTmp!信息值))
                If Not IsNull(rsTmp!信息值) And cboinfo(cbo质控护士).ListIndex = -1 Then Call SetCboFromName(rsTmp!信息值, cboinfo(cbo质控护士))
            Case "病原学检查"
                chkInfo(chk病原学).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "输血检查"
                Call GetCboIndex(cboinfo(cbo输血检查), Nvl(rsTmp!信息值))
            Case "CT"
                chkInfo(chkCT).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "MRI"
                chkInfo(chkMRI).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "彩色多普勒"
                chkInfo(chk多普勒).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "特殊检查4"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1) = Nvl(rsTmp!信息值)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 0, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1)
            Case "特殊检查5"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1) = Nvl(rsTmp!信息值)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 1, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1)
            Case "特殊检查6"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1) = Nvl(rsTmp!信息值)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 2, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1)
            Case "出院转入"
                txtInfo(txt出院转入).Text = Nvl(rsTmp!信息值)
            Case "压疮发生期间"
                cboinfo(cbo压疮发生期间).Text = Nvl(rsTmp!信息值, " ")
            Case "压疮分期"
                cboinfo(cbo压疮分期).Text = Nvl(rsTmp!信息值, " ")
            Case "跌倒或坠床伤害"
                cboinfo(cbo跌倒或坠床伤害).Text = Nvl(rsTmp!信息值, " ")
            Case "跌倒或坠床原因"
                cboinfo(cbo跌倒或坠床原因).Text = Nvl(rsTmp!信息值, " ")
            Case "31天内再住院"
                If Nvl(rsTmp!信息值) <> "" Then
                    optInput(opt31天有).Value = True
                    txtInfo(txt31天目的).Text = Nvl(rsTmp!信息值)
                    txtInfo(txt31天目的).Enabled = True
                End If
            Case "再入院计划天数"
                cboinfo(cbo31天和7天再入院).ListIndex = Val(Nvl(rsTmp!信息值))
            Case "不足周岁年龄"
                Call LoadOldData("" & rsTmp!信息值, txt婴儿年龄)
            Case "新生儿出生体重"
                txtInfo(txt新生儿体重).Text = Nvl(rsTmp!信息值)
            Case "新生儿入院体重"
                txtInfo(txt新生儿入院体重).Text = Nvl(rsTmp!信息值)
            Case "呼吸机使用时间"
                txtInfo(txt呼吸机小时).Text = Nvl(rsTmp!信息值)
            Case "昏迷时间"
                '保存格式:入院前(天，小时,分钟)|入院后(天，小时,分钟)
                txtInfo(txt入院前天).Text = Split(Split(Nvl(rsTmp!信息值), "|")(0) & ",", ",")(0)
                txtInfo(txt入院前小时).Text = Split(Split(Nvl(rsTmp!信息值), "|")(0) & ",", ",")(1)
                txtInfo(txt入院前分钟).Text = Split(Split(Nvl(rsTmp!信息值), "|")(0) & ",", ",")(2)
                txtInfo(txt入院后天).Text = Split(Split(Nvl(rsTmp!信息值), "|")(1) & ",", ",")(0)
                txtInfo(txt入院后小时).Text = Split(Split(Nvl(rsTmp!信息值) & "|", "|")(1) & ",", ",")(1)
                txtInfo(txt入院后分钟).Text = Split(Split(Nvl(rsTmp!信息值) & "|", "|")(1) & ",", ",")(2)
            Case "抢救病因"
                txtInfo(txt抢救原因).Text = Nvl(rsTmp!信息值)
            Case "自体回收"
                txtInfo(txt自体回收).Text = Nvl(rsTmp!信息值)
            Case "籍贯"
                txtInfo(txt籍贯).Text = Nvl(rsTmp!信息值)
            Case "最高诊断依据"
                If Nvl(rsTmp!信息值) <> "" Then
                    Call GetCboIndex(cboinfo(cbo最高诊断依据), Nvl(rsTmp!信息值))
                End If
            Case "分化程度"
                If Nvl(rsTmp!信息值) <> "" Then
                    Call GetCboIndex(cboinfo(cbo分化程度), Nvl(rsTmp!信息值))
                End If
            Case "中医设备"
                Call GetCboIndex(cboinfo(cbo使用中医诊疗设备), Nvl(rsTmp!信息值))
            Case "中医技术"
                Call GetCboIndex(cboinfo(cbo使用中医诊疗技术), Nvl(rsTmp!信息值))
            Case "辨证施护"
                Call GetCboIndex(cboinfo(cbo辨证施护), Nvl(rsTmp!信息值))
            Case "病理号"
                txtInfo(txt病理号).Text = Nvl(rsTmp!信息值)
            Case "病案质量"
                Call GetCboIndex(cboinfo(cbo病案质量), Nvl(rsTmp!信息值))
            Case "主页质量日期"
                txtInfo(txt质控日期).Text = Nvl(rsTmp!信息值)
            Case "告病重病危"
                chkInfo(chk住院期间告病重或病危).Value = Val(Nvl(rsTmp!信息值))
            Case "临床路径"
                chkInfo(chk进入路径).Value = Val(Nvl(rsTmp!信息值))
            Case "退出原因"
                If Nvl(rsTmp!信息值) = "1" Then
                    chkInfo(chk完成路径).Value = 1
                Else
                    chkInfo(chk完成路径).Value = 0
                    txtInfo(txt退出原因).Text = Nvl(rsTmp!信息值)
                End If
            Case "变异原因"
                If Nvl(rsTmp!信息值) = "0" Then
                    chkInfo(chk变异).Value = 0
                Else
                    chkInfo(chk变异).Value = 1
                    txtInfo(txt变异原因).Text = Trim(Nvl(rsTmp!信息值))
                End If
            Case "身体约束"
                chkInfo(chk是否使用物理约束).Value = Val(Nvl(rsTmp!信息值))
            Case "约束总时间"
                txtInfo(txt约束总时间).Text = Nvl(rsTmp!信息值)
            Case "约束方式"
                Call GetCboIndex(cboinfo(cbo约束方式), Nvl(rsTmp!信息值))
            Case "约束工具"
                Call GetCboIndex(cboinfo(cbo约束工具), Nvl(rsTmp!信息值))
            Case "约束原因"
                Call GetCboIndex(cboinfo(cbo约束原因), Nvl(rsTmp!信息值))
            Case "新生儿离院方式"
                Call GetCboIndex(cboinfo(cbo新生儿离院方式), Nvl(rsTmp!信息值))
            Case "围术期死亡"
                chkInfo(chk围术期死亡).Value = Val(Nvl(rsTmp!信息值))
            Case "术后猝死"
                chkInfo(chk术后猝死).Value = Val(Nvl(rsTmp!信息值))
            Case "生育状况"
                Call GetCboIndex(cboinfo(cbo生育状况), Nvl(rsTmp!信息值))
            Case "发病时间"
                If Nvl(rsTmp!信息值) <> "" Then
                    txt发病日期.Text = Format(rsTmp!信息值, "yyyy-MM-dd")
                    If Format(rsTmp!信息值, "HH:mm") <> "00:00" Then
                        txt发病时间.Text = Format(rsTmp!信息值, "HH:mm")
                    End If
                End If
            Case Else
                '多个抗生素名称
                If Left(Nvl(rsTmp!信息名), 3) = "抗生素" And Not IsNull(rsTmp!信息值) Then
                    With vsKSS
                        For j = .FixedRows To .Rows - 1
                            If .TextMatrix(j, 1) = "" Then
                                '兼容老数据，在主页从表里先读数据
                                .RowData(j) = GetIDTmp(rsTmp!信息值)
                                If .RowData(j) <> 0 Then
                                    .TextMatrix(j, 1) = rsTmp!信息值
                                    .Cell(flexcpData, j, 1) = .TextMatrix(j, 1)
                                End If
                                Exit For
                            End If
                        Next
                        If j > .Rows - 1 Then
                            .AddItem ""
                             '兼容老数据，在主页从表里先读数据
                            .RowData(.Rows - 1) = GetIDTmp(rsTmp!信息值)
                            If .RowData(.Rows - 1) <> 0 Then
                                .TextMatrix(.Rows - 1, 1) = rsTmp!信息值
                                .Cell(flexcpData, .Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
                            End If
                        End If
                    End With
                Else
                    '附加项目
                    If Not IsNull(rsTmp("编码")) Then
                        With vsfMain
                            For lngCol = 0 To vsfMain.Cols - 1 Step 3
                                lngRow = vsfMain.FindRow(rsTmp("信息名"), , lngCol)
                                If lngRow >= 0 Then
                                    If vsfMain.TextMatrix(lngRow, lngCol) = rsTmp("信息名") Then
                                        If vsfMain.TextMatrix(lngRow, lngCol + 2) = "是否" Then
                                            vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = IIf(rsTmp("信息值") = 0, 2, 1)
                                            Exit For
                                        Else
                                            vsfMain.TextMatrix(lngRow, lngCol + 1) = rsTmp("信息值") & ""
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next lngCol
                        End With
                    End If
                End If
                Call SetKSSSerial
        End Select
        rsTmp.MoveNext
    Next
    
    '自动提取转科科室及入出病室(房间号)
    '---------------------------------------------------------------
    If txtInfo(txt转科1).Text = "" And txtInfo(txt转科2).Text = "" And txtInfo(txt转科3).Text = "" Then
        StrSQL = _
            " Select B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.科室ID=B.ID And A.开始原因=3 And A.开始时间 is Not NULL" & _
            " Order by A.开始时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        For i = 1 To rsTmp.RecordCount
            If i = 1 Then
                txtInfo(txt转科1).Text = rsTmp!名称
            ElseIf i = 2 Then
                txtInfo(txt转科2).Text = rsTmp!名称
            ElseIf i = 3 Then
                txtInfo(txt转科3).Text = rsTmp!名称
                Exit For
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If txtInfo(txt入院病室).Text = "" Then
        StrSQL = "Select B.房间号" & _
            " From 病案主页 A,床位状况记录 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院病区ID=B.病区ID And A.入院病床=B.床号"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then txtInfo(txt入院病室).Text = Nvl(rsTmp!房间号)
    End If
    If txtInfo(txt出院病室).Text = "" Then
        StrSQL = "Select B.房间号" & _
            " From 病案主页 A,床位状况记录 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.当前病区ID=B.病区ID And A.出院病床=B.床号"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then txtInfo(txt出院病室).Text = Nvl(rsTmp!房间号)
    End If
    
    '过敏信息:本次住院的,过敏的
    '---------------------------------------------------------------
    StrSQL = "Select 记录来源,NVL(过敏时间,记录时间) as 过敏时间,药物ID,药物名,过敏反应 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by NVL(过敏时间,记录时间),药物名"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 2 '固定行+新行
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , AC_过敏药物)
                End If
                If lngRow = -1 Then
                    .TextMatrix(i, AC_过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, AC_过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")  '用于保存
                    .TextMatrix(i, AC_过敏药物) = Nvl(rsTmp!药物名)
                    .Cell(flexcpData, i, AC_过敏药物) = .TextMatrix(i, AC_过敏药物) '用于输入恢复
                    .TextMatrix(i, AC_过敏反应) = Nvl(rsTmp!过敏反应)
                    .Cell(flexcpData, i, AC_过敏反应) = .TextMatrix(i, AC_过敏反应)   '用于输入恢复

                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    vsAller.Row = 1: vsAller.Col = AC_过敏药物
    vsAller.Tag = "未修改"
    
    '西医诊断
    '---------------------------------------------------------------
    str治疗结果 = Get治疗结果
    vsDiagXY.ColData(col出院情况) = str治疗结果
    
    '判断首页是否填过诊断
    StrSQL = "Select 1 From 病人诊断记录 Where 病人ID=[1] And 主页ID=[2] And 记录来源=3  And RowNum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    mbln首页诊断 = rsTmp.RecordCount > 0
    If mbln首页诊断 Then
        strTmp = " and a.记录来源=3 "
    Else
        strTmp = " And a.记录来源 IN(1,2,3,4) "
    End If
    '缺省表格初始化
    With vsDiagXY
        '1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
        .TextMatrix(1, col类型) = 1
        .TextMatrix(2, col类型) = 2
        .TextMatrix(3, col类型) = 3
        .TextMatrix(4, col类型) = 3
        .TextMatrix(5, col类型) = 5
        .TextMatrix(6, col类型) = 10
        .TextMatrix(7, col类型) = 6
        .TextMatrix(8, col类型) = 7
    End With
    
    '读取各种来源的诊断
    StrSQL = "Select a.备注,a.ID,a.病人ID,a.主页ID,a.医嘱ID,a.记录来源,a.诊断次序,a.编码序号,a.病历ID,a.诊断类型,a.疾病ID,a.入院病情," & _
        " a.诊断ID,a.证候ID,a.诊断描述,a.出院情况,a.是否未治,a.是否疑诊,a.记录日期,a.记录人,a.取消时间,a.取消人,a.病例ID, b.编码 As 疾病编码, c.编码 As 诊断编码 " & _
        " From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & _
        " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+)  And a.诊断类型 IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=[2]" & _
        " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Set mrsXYDiag = zlDatabase.CopyNewRec(rsTmp)
    If Not rsTmp.EOF Then
        With vsDiagXY
            StrSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(StrSQL, ",")(i)
                If Val(Split(StrSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(StrSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(StrSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If
                
                If Val(Split(StrSQL, ",")(i)) = 21 Then
                    '21-病原学诊断
                    If Not rsTmp.EOF Then
                        txtInfo(txt病原学).Text = Nvl(rsTmp!诊断描述)
                        txtInfo(txt病原学).Tag = txtInfo(txt病原学).Text
                        cmdInfo(txt病原学).Tag = Nvl(rsTmp!疾病id, 0)
                    End If
                Else
                    Do While Not rsTmp.EOF
                        If Val("" & rsTmp!记录来源) = 3 And Val("" & rsTmp!诊断类型) = 2 And Val("" & rsTmp!诊断次序) = 1 Then
                            mstrXYDiagInfo = "" & rsTmp!诊断描述
                        End If
                        '确定当前显示行
                        lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , col类型)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col类型)) = Val(Split(StrSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col类型) = Split(StrSQL, ",")(i)
                        End If
                        
                        If IsNull(rsTmp!诊断描述) Then
                            .TextMatrix(lngRow, col诊断编码) = ""
                            .TextMatrix(lngRow, col诊断描述) = ""
                        Else
                            If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                                '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                                If Val(rsTmp!疾病id & "") <> 0 Then
                                    .TextMatrix(lngRow, col诊断编码) = Nvl(rsTmp!疾病编码)
                                ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                                    .TextMatrix(lngRow, col诊断编码) = Nvl(rsTmp!诊断编码)
                                Else
                                    .TextMatrix(lngRow, col诊断编码) = ""
                                End If
                                .TextMatrix(lngRow, col诊断描述) = rsTmp!诊断描述
                            Else
                                .TextMatrix(lngRow, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                                .TextMatrix(lngRow, col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                            End If
                        End If
                        If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                            .Cell(flexcpData, lngRow, col诊断描述) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                        Else
                            .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
                        End If
                        
                        '分化程度和最高诊断依据
                        If Val("" & rsTmp!诊断类型) = 3 And Val("" & rsTmp!诊断次序) = 1 Then
                            If Trim(Nvl(rsTmp!疾病编码)) = "" Then
                                bln分化程度 = False
                            Else
                                bln分化程度 = ((InStr("C", UCase(Left(Nvl(rsTmp!疾病编码), 1)))) > 0) Or ((InStr("D0", UCase(Left(Nvl(rsTmp!疾病编码), 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(Nvl(rsTmp!疾病编码), 4)))) > 0)
                            End If
                        End If
                        
                        cboinfo(cbo分化程度).Enabled = bln分化程度
                        lblInfo(lbl分化程度).Enabled = bln分化程度
                        lblInfo(lbl最高诊断依据).Enabled = bln分化程度
                        cboinfo(cbo最高诊断依据).Enabled = bln分化程度
                        .TextMatrix(lngRow, col备注) = Nvl(rsTmp!备注)
                       .Cell(flexcpData, lngRow, col是否疑诊) = Val(rsTmp!ID & "")
                        .TextMatrix(lngRow, col出院情况) = Nvl(rsTmp!出院情况)
                        .TextMatrix(lngRow, col入院病情) = Nvl(rsTmp!入院病情)
                        .TextMatrix(lngRow, col是否未治) = IIf(Nvl(rsTmp!是否未治, 0) = 1, "√", "")
                        .TextMatrix(lngRow, col是否疑诊) = IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                        .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!诊断id, 0)
                        .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If
    
    vsDiagXY.Cell(flexcpForeColor, 1, col是否疑诊, vsDiagXY.Rows - 1, col是否疑诊) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Cell(flexcpBackColor, 1, col诊断编码, vsDiagXY.Rows - 1, col诊断编码) = ColorUnEditCell      '灰蓝色
    vsDiagXY.Row = 1: vsDiagXY.Col = col诊断描述
    Call vsDiagXY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
    vsDiagXY.Tag = "未修改"
    If vsDiagXY.TextMatrix(GetRow(6), col诊断描述) <> "" Then
        txtInfo(txt病理号).Enabled = True
        txtInfo(txt病理号).BackColor = vbWindowBackground
    End If
        
    '中医诊断
    '---------------------------------------------------------------
    vsDiagZY.ColData(col出院情况) = str治疗结果
    
    '缺省表格初始化
    With vsDiagZY
        '11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断(主要诊断、其它诊断)
        .TextMatrix(1, colzy类型) = 11
        .TextMatrix(2, colzy类型) = 12
        .TextMatrix(3, colzy类型) = 13
        .TextMatrix(4, colzy类型) = 13
    End With
    
    If mbln首页诊断 Then
        strTmp = " and a.记录来源=3 "
    Else
        strTmp = " And a.记录来源 IN(1,2,3,4) "
    End If
    
    '读取各种来源的诊断
    StrSQL = "Select a.备注, a.Id, a.病人id, a.主页id, a.医嘱id, a.记录来源, a.诊断次序, a.编码序号, a.病历id, a.诊断类型,a.入院病情," & _
        " a.疾病id, a.诊断id, a.证候id, a.诊断描述,a.出院情况, a.是否未治, a.是否疑诊, a.记录日期, a.记录人, a.取消时间," & _
        " a.取消人, a.病例id, b.编码 As 疾病编码, c.编码 As 诊断编码,d.编码 as 证候编码 From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
        " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.证候ID=d.ID(+) And a.诊断类型 IN(11,12,13)" & _
        strTmp & _
        " And 取消时间 Is Null And 病人ID=[1] And 主页ID=[2]" & _
        " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.编码序号,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    strTmp = ""
    Set mrsZYDiag = zlDatabase.CopyNewRec(rsTmp)
    If Not rsTmp.EOF Then
        With vsDiagZY
            StrSQL = "11,12,13"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(StrSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If
                
                Do While Not rsTmp.EOF
                    If Val("" & rsTmp!记录来源) = 3 And Val("" & rsTmp!诊断类型) = 12 And Val("" & rsTmp!诊断次序) = 1 Then
                        mstrZYDiagInfo = "" & rsTmp!诊断描述
                    End If
                    '确定当前显示行
                    lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , colzy类型)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, colzy类型)) = Val(Split(StrSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, col诊断描述) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, col诊断描述) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy类型) = Split(StrSQL, ",")(i)
                    End If
                    
                    If IsNull(rsTmp!诊断描述) Then
                        .TextMatrix(lngRow, col诊断编码) = ""
                        .TextMatrix(lngRow, col诊断描述) = ""
                    Else
                        If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then     '中医的诊断描述后面加了（候症），所以只判断第一个字符
                            '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                            If Val(rsTmp!疾病id & "") <> 0 Then
                                .TextMatrix(lngRow, col诊断编码) = Nvl(rsTmp!疾病编码)
                            ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                                .TextMatrix(lngRow, col诊断编码) = Nvl(rsTmp!诊断编码)
                            Else
                                .TextMatrix(lngRow, col诊断编码) = ""
                            End If
                            .TextMatrix(lngRow, col诊断描述) = rsTmp!诊断描述
                        Else
                            .TextMatrix(lngRow, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                            .TextMatrix(lngRow, col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                        End If
                    End If
                    If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                        .Cell(flexcpData, lngRow, col诊断描述) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                    Else
                        .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
                    End If
                        
                    .TextMatrix(lngRow, col备注) = Nvl(rsTmp!备注)
                    .Cell(flexcpData, lngRow, col是否疑诊) = Val(rsTmp!ID & "")
                    .TextMatrix(lngRow, col出院情况) = Nvl(rsTmp!出院情况)
                    .TextMatrix(lngRow, col入院病情) = Nvl(rsTmp!入院病情)
                    .TextMatrix(lngRow, colzy诊断ID) = Nvl(rsTmp!诊断id, 0)
                    .TextMatrix(lngRow, colzy疾病ID) = Nvl(rsTmp!疾病id, 0)
                    .TextMatrix(lngRow, colzy证候ID) = Nvl(rsTmp!证候id, 0)
                    '取证候名称
                    If InStr(.TextMatrix(lngRow, col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断描述), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(lngRow, col诊断描述), InStrRev(.TextMatrix(lngRow, col诊断描述), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '先取证候
                        .TextMatrix(lngRow, col中医证候) = strTmp
                        '去掉诊断描述的证候
                        .TextMatrix(lngRow, col诊断描述) = Mid(.TextMatrix(lngRow, col诊断描述), 1, InStrRev(.TextMatrix(lngRow, col诊断描述), "(") - 1)
                    Else
                       .TextMatrix(lngRow, col中医证候) = ""
                    End If
                    
                    
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
    vsDiagZY.Cell(flexcpBackColor, 1, col诊断编码, vsDiagZY.Rows - 1, col诊断编码) = ColorUnEditCell      '灰蓝色
    vsDiagZY.Row = 1: vsDiagZY.Col = col诊断描述
    Call vsDiagZY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
    vsDiagZY.Tag = "未修改"
    
    If Not mbln首页诊断 Then
        vsDiagZY.Tag = ""
        vsDiagXY.Tag = ""
    End If
    
    '手术情况
    '---------------------------------------------------------------
    StrSQL = "Select 编码,名称 From 手术切口愈合"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        strTmp = " "
        Do While Not rsTmp.EOF
            strTmp = strTmp & "|" & rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Loop
        vsOPS.ColComboList(col切口愈合) = strTmp
    Else
        vsOPS.ColComboList(col切口愈合) = " |0-0 / |1-Ⅰ/甲|2-Ⅰ/乙|3-Ⅰ/丙|4-Ⅰ/其他|5-Ⅱ/甲|6-Ⅱ/乙|7-Ⅱ/丙|8-Ⅱ/其他|9-Ⅲ/甲|10-Ⅲ/乙|11-Ⅲ/丙|12-Ⅲ/其他|13-IV/甲|14-IV/乙|15-IV/丙|16-IV/其他"
    End If
    'col麻醉类型
    StrSQL = "Select 编码,名称 From 诊疗麻醉类型"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        strTmp = " "
        Do While Not rsTmp.EOF
            strTmp = strTmp & "|" & rsTmp!名称
            rsTmp.MoveNext
        Loop
        vsOPS.ColComboList(col麻醉类型) = strTmp
    Else
        vsOPS.ColComboList(col麻醉类型) = " |局麻|全麻|持硬|其他|静脉|臂丛|颈丛"
    End If
    '手术情况
    vsOPS.ColComboList(COL手术情况.COL手术情况) = " |择期|急诊|限期"
    'ASA分级
    vsOPS.ColComboList(COL手术情况.colASA分级) = " |P1|P2|P3|P4|P5|P6"
    'colNNIS分级
    vsOPS.ColComboList(colNNIS分级) = " |NNIS0级|NNIS1级|NNIS2级|NNIS3级"
    '手术分级
    vsOPS.ColComboList(col手术级别) = " |一级手术|二级手术|三级手术|四级手术"
    vsOPS.ColDataType(col再次手术) = flexDTBoolean
    
    '首读取首页整理保存的内容
    StrSQL = "Select a.助产护士,a.手术情况,a.切口,a.愈合,a.记录日期,a.记录人,a.取消时间,a.取消人,NVl(B.编码,C.编码) as 手术编码,a.ID,a.病人ID,a.主页ID,a.记录来源,a.手术日期,a.手术开始时间,a.手术结束时间,a.拟行手术,a.手术操作ID,a.诊疗项目ID,a.已行手术,a.主刀医师,a.第一助手," & _
    "a.第二助手,a.手术护士,a.麻醉开始时间,a.麻醉结束时间,a.麻醉方式,a.麻醉类型,a.麻醉质量,a.输液总量,a.麻醉医师,a.输氧开始时间,a.输氧结束时间,a.手术情况,a.ASA分级,a.再次手术,a.NNIS分级,decode(a.手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别" & _
    ",a.术前抗菌用药,a.抗菌用药天数,a.非预期的二次手术,a.麻醉并发症,a.术中异物遗留,a.手术并发症,a.术后出血或血肿,a.手术伤口裂开,a.术后深静脉血栓,a.术后生理代谢紊乱,a.术后呼吸衰竭,a.术后肺栓塞,a.术后败血症,a.术后髋关节骨折" & _
            " From 病人手麻记录  A,疾病编码目录 B,诊疗项目目录 C Where c.ID(+)=a.诊疗项目ID And A.手术操作ID=B.ID(+) and 病人ID=[1] And 主页ID=[2] And 记录来源=3 Order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.EOF Then '没有时读取其它来源的诊断
        '病历：病历作废时填写取消
        StrSQL = "Select Max(记录日期) From 病人手麻记录" & _
            " Where 病人ID=" & mlng病人ID & " And 主页ID=" & mlng主页ID & _
            " And 记录来源=1 And 取消时间 is NULL"
         StrSQL = "Select a.助产护士,a.手术情况,a.切口,a.愈合,a.记录日期,a.记录人,a.取消时间,a.取消人,NVl(B.编码,C.编码) as 手术编码,a.ID,a.病人ID,a.主页ID,a.记录来源,a.手术日期,a.手术开始时间,a.手术结束时间,a.拟行手术,a.手术操作ID,a.诊疗项目ID,a.已行手术,a.主刀医师,a.第一助手," & _
         "a.第二助手,a.手术护士,a.麻醉开始时间,a.麻醉结束时间,a.麻醉方式,a.麻醉类型,a.麻醉质量,a.输液总量,a.麻醉医师,a.输氧开始时间,a.输氧结束时间,a.手术情况,a.ASA分级,a.再次手术,a.NNIS分级,decode(a.手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别" & _
         ",a.术前抗菌用药,a.抗菌用药天数,a.非预期的二次手术,a.麻醉并发症,a.术中异物遗留,a.手术并发症,a.术后出血或血肿,a.手术伤口裂开,a.术后深静脉血栓,a.术后生理代谢紊乱,a.术后呼吸衰竭,a.术后肺栓塞,a.术后败血症,a.术后髋关节骨折" & _
            " From 病人手麻记录  A,疾病编码目录 B,诊疗项目目录 C Where c.ID(+)=a.诊疗项目ID And " & _
            " A.手术操作ID=B.ID(+) and 病人ID=[1] And 主页ID=[2]" & _
            " And 记录来源=1 And 取消时间 is NULL And 记录日期=(" & StrSQL & ")" & _
            " Order by ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.EOF Then '病案
            StrSQL = "Select a.助产护士,a.手术情况,a.切口,a.愈合,a.记录日期,a.记录人,a.取消时间,a.取消人,NVl(B.编码,C.编码) as 手术编码,a.ID,a.病人ID,a.主页ID,a.记录来源,a.手术日期,a.手术开始时间,a.手术结束时间,a.拟行手术,a.手术操作ID,a.诊疗项目ID,a.已行手术,a.主刀医师,a.第一助手," & _
            "a.第二助手,a.手术护士,a.麻醉开始时间,a.麻醉结束时间,a.麻醉方式,a.麻醉类型,a.麻醉质量,a.输液总量,a.麻醉医师,a.输氧开始时间,a.输氧结束时间,a.手术情况,a.ASA分级,a.再次手术,a.NNIS分级,decode(a.手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别" & _
            ",a.术前抗菌用药,a.抗菌用药天数,a.非预期的二次手术,a.麻醉并发症,a.术中异物遗留,a.手术并发症,a.术后出血或血肿,a.手术伤口裂开,a.术后深静脉血栓,a.术后生理代谢紊乱,a.术后呼吸衰竭,a.术后肺栓塞,a.术后败血症,a.术后髋关节骨折" & _
                " From 病人手麻记录  A,疾病编码目录 B,诊疗项目目录 C Where c.ID(+)=a.诊疗项目ID And  A.手术操作ID=B.ID(+) and 病人ID=[1] And 主页ID=[2] And 记录来源=4 Order by ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        End If
    End If
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col手术日期) = Format(Nvl(rsTmp!手术日期), "yyyy-MM-dd")
                .TextMatrix(i, col手术编码) = Nvl(rsTmp!手术编码)
                .TextMatrix(i, col手术名称) = Nvl(rsTmp!已行手术)
                .TextMatrix(i, col主刀医师) = Nvl(rsTmp!主刀医师)
                .TextMatrix(i, col助产护士) = Nvl(rsTmp!助产护士)
                .TextMatrix(i, col助手1) = Nvl(rsTmp!第一助手)
                .TextMatrix(i, col助手2) = Nvl(rsTmp!第二助手)
                .TextMatrix(i, col麻醉方式) = GetItemField("诊疗项目目录", Val(Nvl(rsTmp!麻醉方式, 0)), "名称")
                .TextMatrix(i, col麻醉医师) = Nvl(rsTmp!麻醉医师)
                If Not IsNull(rsTmp!切口) And Not IsNull(rsTmp!愈合) Then
                    .TextMatrix(i, col切口愈合) = rsTmp!切口 & "/" & rsTmp!愈合
                End If
                .TextMatrix(i, col手术操作ID) = Nvl(rsTmp!手术操作ID)
                .TextMatrix(i, col诊疗项目ID) = Nvl(rsTmp!诊疗项目id)
                .TextMatrix(i, col麻醉ID) = Nvl(rsTmp!麻醉方式)
                .TextMatrix(i, col麻醉类型) = Nvl(rsTmp!麻醉类型)
                .TextMatrix(i, COL手术情况.COL手术情况) = Nvl(rsTmp!手术情况)
                .TextMatrix(i, colASA分级) = Decode(Nvl(rsTmp!asa分级), "I级", "P1", "II级", "P2", "III级", "P3", "IV级", "P4", "V级", "P5", Nvl(rsTmp!asa分级))
                .TextMatrix(i, colNNIS分级) = Nvl(rsTmp!NNIS分级)
                .TextMatrix(i, col手术级别) = Nvl(rsTmp!手术级别)
                .TextMatrix(i, col再次手术) = IIf(Val(rsTmp!再次手术 & "") = 1, -1, 0)
                .TextMatrix(i, col抗菌药天数) = rsTmp!抗菌用药天数 & ""
                .Cell(flexcpChecked, i, col预防用抗菌药) = Val(rsTmp!术前抗菌用药 & "")
                .Cell(flexcpChecked, i, col非预期的二次手术) = Val(rsTmp!非预期的二次手术 & "")
                .Cell(flexcpChecked, i, col麻醉并发症) = Val(rsTmp!麻醉并发症 & "")
                .Cell(flexcpChecked, i, col术中异物遗留) = Val(rsTmp!术中异物遗留 & "")
                .Cell(flexcpChecked, i, col手术并发症) = Val(rsTmp!手术并发症 & "")
                .Cell(flexcpChecked, i, col术后出血或血肿) = Val(rsTmp!术后出血或血肿 & "")
                .Cell(flexcpChecked, i, col手术伤口裂开) = Val(rsTmp!手术伤口裂开 & "")
                .Cell(flexcpChecked, i, col术后深静脉血栓) = Val(rsTmp!术后深静脉血栓 & "")
                .Cell(flexcpChecked, i, col术后生理代谢紊乱) = Val(rsTmp!术后生理代谢紊乱 & "")
                .Cell(flexcpChecked, i, col术后呼吸衰竭) = Val(rsTmp!术后呼吸衰竭 & "")
                .Cell(flexcpChecked, i, col术后肺栓塞) = Val(rsTmp!术后肺栓塞 & "")
                .Cell(flexcpChecked, i, col术后败血症) = Val(rsTmp!术后败血症 & "")
                .Cell(flexcpChecked, i, col术后髋关节骨折) = Val(rsTmp!术后髋关节骨折 & "")
                '记录用于编辑恢复
                For j = 0 To .Cols - 1
                    .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                Next
                
                rsTmp.MoveNext
            Next
        End With
    End If
    vsOPS.Tag = "未修改"
    vsKSS.Tag = "未修改"
    
    '诊断符合情况
    '---------------------------------------------------------------
    '处理诊断符合情况缺省值
    Call Set诊断符合情况(cbo门诊与出院)
    Call Set诊断符合情况(cbo入院与出院)
    Call Set诊断符合情况(cbo门诊与入院)
    Call Set诊断符合情况(cbo放射与病理)
    Call Set诊断符合情况(cbo临床与病理)
    Call Set诊断符合情况(cbo临床与尸检)
    Call Set诊断符合情况(cbo术前与术后)
    Call Set诊断符合情况(cbo中医门诊与出院)
    Call Set诊断符合情况(cbo中医入院与出院)
    
    StrSQL = "Select 符合类型,符合情况 From 诊断符合情况 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!符合类型
        Case 1 '门诊与出院
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo门诊与出院).ListIndex = rsTmp!符合情况
        Case 2 '入院与出院
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo入院与出院).ListIndex = rsTmp!符合情况
        Case 3 '放射与病理
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo放射与病理).ListIndex = rsTmp!符合情况
        Case 4 '临床与病理
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo临床与病理).ListIndex = rsTmp!符合情况
        Case 5 '临床与尸检
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo临床与尸检).ListIndex = rsTmp!符合情况
        Case 6 '术前与术后
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo术前与术后).ListIndex = rsTmp!符合情况
        Case 7 '门诊与入院
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo门诊与入院).ListIndex = rsTmp!符合情况
        Case 11 '中医门诊与出院
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo中医门诊与出院).ListIndex = rsTmp!符合情况
        Case 12 '中医入院与出院
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo中医入院与出院).ListIndex = rsTmp!符合情况
        Case 13 '中医辨证
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo辨证).ListIndex = rsTmp!符合情况
        Case 14 '中医治法
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo治法).ListIndex = rsTmp!符合情况
        Case 15 '中医方药
            If Nvl(rsTmp!符合情况, 0) >= 0 Then cboinfo(cbo方药).ListIndex = rsTmp!符合情况
        End Select
        rsTmp.MoveNext
    Loop
    
    '附加信息
    '---------------------------------------------------------------
    '不良事件
    lstAdvEvent.Clear
    StrSQL = "Select 编码,名称 From 不良事件 order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If Nvl(rsTmp!名称) = "新生儿产伤" Or Nvl(rsTmp!名称) = "阴道分娩产妇产伤" Then
            If Have部门性质(mlng科室ID, "产科") Then
                lstAdvEvent.AddItem Nvl(rsTmp!名称)
                lstAdvEvent.ItemData(lstAdvEvent.NewIndex) = Val(rsTmp!编码)
            End If
        Else
            lstAdvEvent.AddItem Nvl(rsTmp!名称)
            lstAdvEvent.ItemData(lstAdvEvent.NewIndex) = Val(rsTmp!编码)
        End If
        rsTmp.MoveNext
    Next
    If lstAdvEvent.ListCount > 0 Then lstAdvEvent.ListIndex = 0
    StrSQL = "Select 信息值 From 病案主页从表 Where 病人id=[1] And 主页ID=[2] And 信息名='不良事件'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then
        StrSQL = "Select /*+ Rule*/  * From  Table(f_Str2list([1]))"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, rsTmp!信息值 & "")
        For i = 1 To rsTmp.RecordCount
            For j = 0 To lstAdvEvent.ListCount - 1
                If lstAdvEvent.ItemData(j) = Val(rsTmp!COLUMN_VALUE & "") Then
                    lstAdvEvent.Selected(j) = True
                    If lstAdvEvent.List(j) = "压疮" Or lstAdvEvent.List(j) = "医院内跌倒/坠床" Then Call lstAdvEvent_ItemCheck(CInt(j))
                End If
            Next
        rsTmp.MoveNext
        Next
    End If
    If lstAdvEvent.ListCount > 0 Then lstAdvEvent.ListIndex = 0
    '感染因素
    lstInfection.Clear
    StrSQL = "Select 编码,名称 From 感染因素 order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstInfection.AddItem Nvl(rsTmp!名称)
        lstInfection.ItemData(lstInfection.NewIndex) = Val(rsTmp!编码)
        rsTmp.MoveNext
    Next
    If lstInfection.ListCount > 0 Then lstInfection.ListIndex = 0
    StrSQL = "Select 信息值 From 病案主页从表 Where 病人id=[1] And 主页ID=[2] And 信息名='感染因素'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then
        StrSQL = "Select /*+ Rule*/  * From  Table(f_Str2list([1]))"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, rsTmp!信息值 & "")
        For i = 1 To rsTmp.RecordCount
            For j = 0 To lstInfection.ListCount - 1
                If lstInfection.ItemData(j) = Val(rsTmp!COLUMN_VALUE & "") Then
                    lstInfection.Selected(j) = True
                End If
            Next
        rsTmp.MoveNext
        Next
    End If
    If lstInfection.ListCount > 0 Then lstInfection.ListIndex = 0
    '--------------------------------------------------------------
    '抗菌药物
    StrSQL = "Select a.药名id, a.用药目的, a.使用阶段, a.使用天数,a.药品名称 名称,一类切口预防用,DDD数,联合用药 " & vbNewLine & _
            " From 病人抗生素记录 A" & vbNewLine & _
            " Where a.病人id = [1] And a.主页id = [2] Order By DDD数 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    Do While Not rsTmp.EOF
        With vsKSS
            For j = .FixedRows To .Rows - 1
                If .TextMatrix(j, 1) = "" Then
                    .RowData(j) = Val(rsTmp!药名id & "")
                    If .RowData(j) <> 0 Then
                        .TextMatrix(j, 1) = Nvl(rsTmp!名称)
                        .Cell(flexcpData, j, 1) = .TextMatrix(j, 1)
                        .TextMatrix(j, kss用药目的) = Nvl(rsTmp!用药目的)
                        .TextMatrix(j, kss使用阶段) = Nvl(rsTmp!使用阶段)
                        .TextMatrix(j, kss使用天数) = IIf(Val(rsTmp!使用天数 & "") = 0, "", Val(rsTmp!使用天数 & "") & "")
                        .Cell(flexcpChecked, j, KSS一类切口预防用) = Val(rsTmp!一类切口预防用 & "")
                        .TextMatrix(j, KSSDDD数) = IIf(Val(rsTmp!DDD数 & "") > 0 And Val(rsTmp!DDD数 & "") < 1, "0", "") & Val(rsTmp!DDD数 & "")
                        .TextMatrix(j, KSS联合用药) = rsTmp!联合用药 & ""
                    End If
                    Exit For
                ElseIf .RowData(j) = Val(rsTmp!药名id & "") Then
                '排除重复值，如果有重复的，则将后面的列的信息填上。
                    If .RowData(j) <> 0 Then
                        .TextMatrix(j, 1) = Nvl(rsTmp!名称)
                        .Cell(flexcpData, j, 1) = .TextMatrix(j, 1)
                        .TextMatrix(j, kss用药目的) = Nvl(rsTmp!用药目的)
                        .TextMatrix(j, kss使用阶段) = Nvl(rsTmp!使用阶段)
                        .TextMatrix(j, kss使用天数) = IIf(Val(rsTmp!使用天数 & "") = 0, "", Val(rsTmp!使用天数 & "") & "")
                        .Cell(flexcpChecked, j, KSS一类切口预防用) = Val(rsTmp!一类切口预防用 & "")
                        .TextMatrix(j, KSSDDD数) = IIf(Val(rsTmp!DDD数 & "") > 0 And Val(rsTmp!DDD数 & "") < 1, "0", "") & Val(rsTmp!DDD数 & "")
                        .TextMatrix(j, KSS联合用药) = rsTmp!联合用药 & ""
                    End If
                    Exit For
                End If
            Next
            '如果没界面上没有空行了，则增加一行
            If j > .Rows - 1 Then
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!药名id & "")
                If .RowData(.Rows - 1) <> 0 Then
                    .TextMatrix(.Rows - 1, 1) = rsTmp!名称
                    .Cell(flexcpData, .Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
                    .TextMatrix(.Rows - 1, kss用药目的) = Nvl(rsTmp!用药目的)
                    .TextMatrix(.Rows - 1, kss使用阶段) = Nvl(rsTmp!使用阶段)
                    .TextMatrix(.Rows - 1, kss使用天数) = IIf(Val(rsTmp!使用天数 & "") = 0, "", Val(rsTmp!使用天数 & "") & "")
                    .Cell(flexcpChecked, .Rows - 1, KSS一类切口预防用) = Val(rsTmp!一类切口预防用 & "")
                    .TextMatrix(.Rows - 1, KSSDDD数) = IIf(Val(rsTmp!DDD数 & "") > 0 And Val(rsTmp!DDD数 & "") < 1, "0", "") & Val(rsTmp!DDD数 & "")
                    .TextMatrix(.Rows - 1, KSS联合用药) = rsTmp!联合用药 & ""
                End If
            End If
        End With
        rsTmp.MoveNext
    Loop
    Call SetKSSSerial
    
    If mbln病案共享 Then
        '放疗化疗
        Call Load化疗与放疗(mlng病人ID, mlng主页ID)
        
    End If
    Call Load附页内容(mlng病人ID, mlng主页ID)
        
    Screen.MousePointer = 0
    LoadPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetKSSSerial()
    Dim i As Long
    
    With vsKSS
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
    End With
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtInfo(Index).Locked Then
        glngTXTProc = GetWindowLong(txtInfo(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtInfo(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtInfo(Index).Locked Then
        Call SetWindowLong(txtInfo(Index).hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str性别 As String, int诊断输入 As Integer
    Dim strInput As String, vPoint As POINTAPI

    Select Case Index
        Case txt年龄
            '没有年龄有出生日期时计算一下年龄
            If txtInfo(txt年龄).Text = "" And IsDate(txt出生日期.Text) Then
                txt出生日期.Tag = "": Call txt出生日期_Validate(False)
            End If
        Case txt确诊日期
            txtInfo(Index).Text = GetFullDate(txtInfo(Index).Text)
            If Not IsDate(txtInfo(Index).Text) Then
                txtInfo(Index).Text = ""
            ElseIf Not CheckDateRange(txtInfo(Index).Text, True) Then
                txtInfo(Index).Text = ""
            End If
        Case txt抢救次数, txt成功次数, txt随诊期限, txt输红细胞, txt输血小板, txt输血浆, txt输全血, txt自体回收
            If txtInfo(Index).Text <> "" Then
                If Not IsNumeric(txtInfo(Index).Text) Then
                    txtInfo(Index).Text = ""
                ElseIf Val(txtInfo(Index).Text) <= 0 Then
                    txtInfo(Index).Text = ""
                End If
            End If
            
            If Index = txt抢救次数 Or Index = txt成功次数 Or Index = txt随诊期限 Then
                If IsNumeric(txtInfo(Index).Text) Then
                    txtInfo(Index).Text = Int(Val(txtInfo(Index).Text))
                End If
            End If
        Case txt病原学
            If txtInfo(txt病原学).Text = "" Then
                txtInfo(txt病原学).Tag = ""
                cmdInfo(txt病原学).Tag = ""
            ElseIf txtInfo(txt病原学).Text = txtInfo(txt病原学).Tag Then
                'Nothing
            Else
                int诊断输入 = Val(Mid(gstr诊断输入, 2, 1))
                If int诊断输入 = 0 Then int诊断输入 = 1
                
                strInput = UCase(txtInfo(txt病原学).Text)
                
                If cboinfo(cbo性别).Text Like "*男*" Then
                    str性别 = "男"
                ElseIf cboinfo(cbo性别).Text Like "*女*" Then
                    str性别 = "女"
                End If
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "名称 Like [2]" '输入汉字时只匹配名称
                Else
                    StrSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(mint简码 = 0, "简码", "五笔码") & " Like [2]"
                End If
                StrSQL = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录 Where Instr([3],类别)>0 And (" & StrSQL & ")" & _
                    IIf(str性别 <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"

                If int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                    '损伤中毒码：Y-损伤中毒的外部原因；病理诊断允许：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", "'D'", str性别)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = GetCoordPos(txtInfo(txt病原学).hwnd, 0, 0)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "病原学诊断", False, "", "", False, False, True, vPoint.X, vPoint.Y, txtInfo(txt病原学).Height, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", "'D'", str性别)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        End If
                    End If
                End If
                
                If Not Cancel Then
                    If rsTmp Is Nothing Then
                        cmdInfo(txt病原学).Tag = ""
                    Else
                        txtInfo(txt病原学).Text = IIf(Not IsNull(rsTmp!编码), "(" & rsTmp!编码 & ")", "") & Nvl(rsTmp!名称)
                        txtInfo(txt病原学).Tag = txtInfo(txt病原学).Text
                        cmdInfo(txt病原学).Tag = rsTmp!项目ID
                    End If
                End If
            End If
        Case txt重症监护室
            strInput = UCase(txtInfo(Index).Text)
            If strInput = "" Then Exit Sub
            StrSQL = " Select Distinct A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where B.部门ID=A.ID And B.工作性质='ICU'" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And (A.编码 Like [1] Or A.简码 Like [2] Or A.名称 Like [2])" & _
                    " Order by A.编码"
            vPoint = GetCoordPos(txtInfo(Index).hwnd, 0, 0)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "重症监护室", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, True, _
                strInput & "%", mstrLike & strInput & "%")
            
            If rsTmp Is Nothing Then
                If Not blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    MsgBox "没有找到指定的ICU重症监护室。", vbInformation, Me.Caption
                End If
                Cancel = True
                Exit Sub
            Else
                txtInfo(Index).Text = rsTmp!名称 & ""
            End If

    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt出生日期_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt出生日期_GotFocus()
    Call zlControl.TxtSelAll(txt出生日期)
End Sub

Private Sub txt出生日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt出生日期_Validate(Cancel As Boolean)
    Dim str年龄 As String
    
    If IsDate(txt出生日期.Text) Then
        If txt出生日期.Tag = txt出生日期.Text Then Exit Sub
        txt出生日期.Tag = txt出生日期.Text '用于记录输入变化
        '小时与分钟做单位则不进行反算
        If cboinfo(cbo年龄单位).ListIndex < 3 Then
            str年龄 = PatiAgeCalc(txt出生日期.Text, , txtInfo(txt入院时间).Text)
            Call LoadOldData(str年龄)
        End If
    ElseIf txt出生日期.Text = "____-__-__" Then
        txt出生时间.Text = "__:__"
    Else
        txt出生日期.Text = "____-__-__"
        txt出生时间.Text = "__:__"
        Cancel = True
    End If
End Sub

Private Sub txt出生时间_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt出生时间_GotFocus()
    Call zlControl.TxtSelAll(txt出生时间)
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not IsDate(txt出生日期.Text) Then
        KeyAscii = 0: txt出生时间.Text = "__:__"
    End If
End Sub

Private Sub txt出生时间_Validate(Cancel As Boolean)
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        txt出生时间.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub txt发病日期_Change()
    If Visible Then mblnChange = True
    
    If IsDate(txt发病日期.Text) Then
        txt发病时间.Enabled = True
    Else
        txt发病时间.Enabled = False
    End If
End Sub

Private Sub txt发病日期_GotFocus()
    Call zlControl.TxtSelAll(txt发病日期)
End Sub

Private Sub txt发病日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt发病日期_Validate(Cancel As Boolean)
    If txt发病日期.Text <> "____-__-__" And Not IsDate(txt发病日期.Text) Then
        txt发病日期.Text = "____-__-__": Cancel = True
    End If
End Sub

Private Sub txt发病时间_GotFocus()
    Call zlControl.TxtSelAll(txt发病时间)
End Sub

Private Sub txt发病时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt发病时间_Validate(Cancel As Boolean)
    If txt发病时间.Text <> "__:__" And Not IsDate(txt发病时间.Text) Then
        txt发病时间.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub txt死亡时间_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt死亡时间_GotFocus()
    Call zlControl.TxtSelAll(txt死亡时间)
End Sub

Private Sub txt死亡时间_Validate(Cancel As Boolean)
    If Not IsDate(txt死亡时间.Text) And txt死亡时间.Text <> "____-__-__ __:__:__" Then
        Cancel = True
    End If
End Sub

Private Sub timThis_Timer()
    Dim lngSelNum As Long
    
    If vsAller.Col = AC_过敏时间 Then
        lngSelNum = vsAller.EditSelStart
        If lngSelNum <> mlngSelNum And lngSelNum <> 16 And lngSelNum <> 0 Then
            Call Vs_EditSelChange(lngSelNum)
            mlngSelNum = lngSelNum
        End If
    End If
End Sub

Private Sub Vs_EditSelChange(ByVal lngSelNum As Long)
'当用户切换光标的时候触发
    With vsAller
        If lngSelNum <= 4 Then
            .EditSelStart = 0
            .EditSelLength = 4
            mlngNum = 0
            mlngNumBack = 4
        ElseIf lngSelNum <= 7 Then
            .EditSelStart = 5
            .EditSelLength = 2
            mlngNum = 5
            mlngNumBack = 7
        ElseIf lngSelNum <= 10 Then
            .EditSelStart = 8
            .EditSelLength = 2
            mlngNum = 8
            mlngNumBack = 10
        ElseIf lngSelNum <= 13 Then
            .EditSelStart = 11
            .EditSelLength = 2
            mlngNum = 11
            mlngNumBack = 13
        ElseIf lngSelNum < 16 Then
            .EditSelStart = 14
            .EditSelLength = 2
            mlngNum = 14
            mlngNumBack = 16
        End If
    End With
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsAller_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAller
        If NewCol = AC_过敏药物 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = IIf(Trim(vsAller.TextMatrix(NewRow, AC_过敏药物)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AC_过敏时间 And Trim(vsAller.Cell(flexcpData, Row, AC_过敏药物)) = "" Then Cancel = True
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim int性别 As Integer
    
    With vsAller
        If cboinfo(cbo性别).Text Like "*男*" Then
            int性别 = 1
        ElseIf cboinfo(cbo性别).Text Like "*女*" Then
            int性别 = 2
        End If
        
        StrSQL = _
            " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
            " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
            " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
            " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
            " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
            " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
            " Union All" & _
            " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
            " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
            " From 诊疗项目目录 A,药品特性 B" & _
            " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
            IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetAllerInput(Row, rsTmp)
            Call AllerEnterNextCell
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AC_过敏药物) <> "" Then
                If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    vsAller.Tag = ""
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyLeft Then
        If mlngNum <= 4 Then Exit Sub
        If mlngNum <= 7 Then Vs_EditSelChange (4): Exit Sub
        If mlngNum <= 10 Then Vs_EditSelChange (7): Exit Sub
        If mlngNum <= 13 Then Vs_EditSelChange (10): Exit Sub
        If mlngNum <= 16 Then Vs_EditSelChange (13): Exit Sub
    End If
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub

    With vsAller
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AC_过敏药物 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim blnIsNextchr As Boolean
    Dim strChr As String
    
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    With vsAller
        If Col = AC_过敏时间 Then
            If KeyAscii = 13 Then .Col = .Col + 1: .ShowCell Row, Col: Exit Sub
            If KeyAscii = vbKeyBack Then
                If mlngNumBack <= 16 Then
                    If mlngNumBack = 0 Then KeyAscii = 0: Exit Sub
                    blnIsNextchr = InStr("1234567890", Mid(.TextMatrix(.Row, .Col), mlngNumBack, 1)) = 0
                    strChr = Mid(.TextMatrix(.Row, .Col), mlngNumBack - IIf(blnIsNextchr, 1, 0), 1)
                    mlngNumBack = mlngNumBack - IIf(blnIsNextchr, 2, 1)
                    .EditText = Mid(.EditText, 1, mlngNumBack) & strChr & Mid(.EditText, mlngNumBack + 2)
                    mlngNum = mlngNumBack
                    KeyAscii = 0
                    If mlngNum <= 4 Then
                        .EditSelStart = 0
                        .EditSelLength = 4
                    ElseIf mlngNum <= 8 Then
                        .EditSelStart = 5
                        .EditSelLength = 2
                    ElseIf mlngNum <= 11 Then
                        .EditSelStart = 8
                        .EditSelLength = 2
                    ElseIf mlngNum <= 14 Then
                        .EditSelStart = 11
                        .EditSelLength = 2
                    ElseIf mlngNum <= 16 Then
                        .EditSelStart = 14
                        .EditSelLength = 2
                    End If
                End If
            Else
                If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
                If Len(.EditText) <= 16 And mlngNum <> 16 Then
                    blnIsNextchr = InStr("1234567890", Mid(.TextMatrix(.Row, .Col), mlngNum + 2, 1)) = 0
                    strChr = Chr(KeyAscii)
                    .EditText = Mid(.EditText, 1, mlngNum) & strChr & Mid(.EditText, mlngNum + 2)
                    mlngNum = mlngNum + IIf(blnIsNextchr, 2, 1)
                    mlngNumBack = mlngNum
                End If
                KeyAscii = 0
                If mlngNum <= 4 Then
                    .EditSelStart = 0
                    .EditSelLength = 4
                ElseIf mlngNum <= 7 Then
                    .EditSelStart = 5
                    .EditSelLength = 2
                ElseIf mlngNum <= 10 Then
                    .EditSelStart = 8
                    .EditSelLength = 2
                ElseIf mlngNum <= 13 Then
                    .EditSelStart = 11
                    .EditSelLength = 2
                ElseIf mlngNum <= 16 Then
                    .EditSelStart = 14
                    .EditSelLength = 2
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = AC_过敏药物 Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
    ElseIf Col = AC_过敏时间 Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = 4
        mlngNum = 0
        timThis.Enabled = True
    End If
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AC_过敏反应 And Trim(vsAller.TextMatrix(Row, AC_过敏药物)) = "" Then Cancel = True
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int性别  As Integer
    Dim curDate As Date
    
    With vsAller
        If Col = AC_过敏药物 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call AllerEnterNextCell
            Else
                If LenB(StrConv(.EditText, vbFromUnicode)) > 60 Then
                    MsgBox "药物名称不能超过30个汉字的长度。", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
                strInput = UCase(.EditText)
                If cboinfo(cbo性别).Text Like "*男*" Then
                    int性别 = 1
                ElseIf cboinfo(cbo性别).Text Like "*女*" Then
                    int性别 = 2
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                    " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                    " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                    " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                    IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                    Decode(mint简码, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
                
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "过敏药物", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", int性别, mint简码 + 1)
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    Call SetAllerInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call AllerEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = AC_过敏时间 Then
            If Not IsDate(.EditText) And .EditText <> "" Then
                MsgBox "您输入的日期格式不正确。格式如：2010-10-10 18:30。"
                Cancel = True
                .EditText = vsAller.TextMatrix(Row, Col)
            Else
                If .EditText <> "" Then
                    curDate = zlDatabase.Currentdate
                    If CDate(.EditText) > curDate Then
                        MsgBox "您输入的日期不能大于当前时间。当前时间：" & curDate & "。"
                        Cancel = True
                        .EditText = .TextMatrix(Row, Col)
                    End If
                End If
                timThis.Enabled = False
                If .Cell(flexcpData, Row, Col) <> .EditText Then
                    .Cell(flexcpData, Row, Col) = .EditText
                    mblnChange = True
                End If
                .Tag = ""
            End If
        Else
            If LenB(StrConv(.EditText, vbFromUnicode)) > 100 Then
                MsgBox "药物名称不能超过50个汉字的长度。", vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col出院情况 Then
            '主要处理非回车离开:不用ComboIndex,取消编辑时不对
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            If Not XYCellEditable(Row, col是否未治) Then
                .TextMatrix(Row, col是否未治) = ""
            End If
            Call SetEditableFrom出院情况
            mblnChange = True
            .Tag = ""
        End If
        
        If Col = col诊断描述 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagXY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        
        Call vsDiagXY_AfterRowColChange(-1, -1, .Row, .Col)
        '判断是否做了修改
        If vsDiagXY.Tag = "未修改" Then
            vsDiagXY.Tag = ""
        End If
    End With
End Sub

Private Sub SetEditableFrom出院情况()
    With vsDiagXY
'        '主要诊断的出院情况为其他则不可输入抢救次数
'        If .TextMatrix(GetRow(3), col出院情况) = "其他" Then
'            txtInfo(txt抢救次数).Text = ""
'            txtInfo(txt抢救次数).Locked = True
'            txtInfo(txt抢救次数).TabStop = False
'            txtInfo(txt抢救次数).BackColor = vbButtonFace
'        Else
'            txtInfo(txt抢救次数).Locked = False
'            txtInfo(txt抢救次数).TabStop = True
'            txtInfo(txt抢救次数).BackColor = vbWindowBackground
'        End If
'        Call txtInfo_Change(txt抢救次数)
        
        '主要诊断的出院情况为死亡时才可以尸检
        If .TextMatrix(GetRow(3), col出院情况) = "死亡" Then
            txt死亡时间.Enabled = True: txt死亡时间.TabStop = True: txt死亡时间.BackColor = vbWindowBackground
            txtInfo(txt死亡原因).Enabled = True
            txtInfo(txt死亡原因).TabStop = True
            txtInfo(txt死亡原因).BackColor = vbWindowBackground
            chkInfo(chk尸检).Enabled = True
            chkInfo(chk尸检).TabStop = True
        Else
            txt死亡时间.Text = "____-__-__ __:__:__"
            txt死亡时间.Enabled = False: txt死亡时间.TabStop = False: txt死亡时间.BackColor = vbButtonFace
            txtInfo(txt死亡原因).Text = ""
            txtInfo(txt死亡原因).Enabled = False
            txtInfo(txt死亡原因).TabStop = False
            txtInfo(txt死亡原因).BackColor = vbButtonFace
            chkInfo(chk尸检).Value = 0
            chkInfo(chk尸检).Enabled = False
            chkInfo(chk尸检).TabStop = False
        End If
        
        '主要诊断的出院情况不为死亡时才可以随诊
        If .TextMatrix(GetRow(3), col出院情况) <> "死亡" Then
            chkInfo(chk随诊).Enabled = True
            chkInfo(chk随诊).TabStop = True
            cboinfo(cbo出院方式).Enabled = True
        Else
            '如果是死亡，则出院情况必须为死亡
            cboinfo(cbo出院方式).Text = "死亡"
            cboinfo(cbo出院方式).Enabled = False
            
            chkInfo(chk随诊).Value = 0
            chkInfo(chk随诊).Enabled = False
            chkInfo(chk随诊).TabStop = False
        End If
        Call chkInfo_Click(chk随诊)
    End With
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsDiagXY
        '清除图片
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, col增加) Is Nothing Then
                Set .Cell(flexcpPicture, i, col增加) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colDel) = Nothing
            End If
        Next
        
        If Not XYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
            
            If NewCol = col诊断描述 Then
                .ComboList = "..."
            ElseIf NewCol = col出院情况 Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col入院病情 Then
                If .TextMatrix(NewRow, 0) = "出院诊断" Or .TextMatrix(NewRow, 0) = "其他诊断" Or .TextMatrix(NewRow, 0) = "" Then
                    .ComboList = "有|临床未确定|情况不明|无"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = col增加 Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '显示图片
            If NewCol <> col增加 And .TextMatrix(NewRow, col诊断描述) <> "" And .TextMatrix(NewRow, 0) <> "出院诊断" Then
                Set .Cell(flexcpPicture, NewRow, col增加) = imgButtonNew.Picture
            End If
            '显示图片
            If NewCol <> colDel Then
                Set .Cell(flexcpPicture, NewRow, colDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col增加 Then Cancel = True
End Sub

Private Sub vsDiagXY_Click()
    With vsDiagXY
        If (.MouseCol = col增加 Or .MouseCol = colDel) And .MouseRow >= .FixedRows Then
            
            If .MouseCol = col增加 Then
                If .TextMatrix(.MouseRow, col诊断描述) = "" Or .TextMatrix(.MouseRow, 0) = "出院诊断" Then Exit Sub
            End If
            
            .Select .MouseRow, .MouseCol
            Call vsDiagXY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagXY
        If Col = col出院情况 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
    '设置为已修改
    If vsDiagXY.Col = col是否未治 Or vsDiagXY.Col = col是否疑诊 Then
        If vsDiagXY.Tag = "未修改" Then vsDiagXY.Tag = "": mblnChange = True
    End If
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断描述 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col诊断描述) <> "" Then
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col诊断类型) = "入院诊断" And mlngDiagnosisType = 2 Or .TextMatrix(.Row, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 1 Then
                        If .TextMatrix(.Row, col诊断类型) <> .TextMatrix(.Row - 1, col诊断类型) Then
                            '首要诊断不允许改
                            Exit Sub
                        End If
                    End If
                End If
                '合并路径
                If Not CheckMergePath(mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, col类型)), Val(.TextMatrix(.Row, col疾病ID))) Then Exit Sub

                '两条路径以上
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                        '导入诊断不允许该
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col诊断类型) = "出院诊断" And mlngDiagnosisType <= 2 Then
                        '正常完成的出院诊断不允许改
                        Exit Sub
                    End If
                End If
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, col类型))
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, col类型) = i
                    
                    '下面的同类诊断数据上移
                    If .TextMatrix(.Row, col诊断类型) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col诊断类型) = "" Then
                                '下一行为无标题的增加行时，数据才上移，否则当前行为有标题时只清空行
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, col类型)) = Val(.TextMatrix(.Row, col类型)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, col类型) = Val(.TextMatrix(.Row, col类型))
                                        
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, col类型)) <> Val(.TextMatrix(i, col类型)) Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    
'                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'                    .Cell(flexcpBackColor, GetRow(3), .FixedRows, GetRow(3), .Cols - 1) = &HC0FFC0
                    
                    '设置诊断符合情况
                    Call Set西医诊断相关(.Row)
                    Call Set病原学
                    
                    mblnChange = True
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col诊断类型) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call XYEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = col是否未治 Or .Col = col是否疑诊) Then
            If XYCellEditable(.Row, .Col) Then
                KeyAscii = 0
                If .Col = col是否疑诊 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "？", "")
                ElseIf .Col = col是否未治 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "√", "")
                End If
            End If
        Else
            If .Col = col诊断描述 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not XYCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = col是否未治 Or Col = col是否疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    Dim str性别 As String, lngRow As Long
    
    With vsDiagXY
        If Col = col诊断描述 Then
            If optInput(0).Value Then
                '按诊断输入:西医部份，一个诊断可能属于多个分类
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", mlng科室ID, , True, False)
            Else
                '7-损伤中毒：Y-损伤中毒的外部原因；6-病理诊断：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(Me, Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), mlng科室ID, cboinfo(cbo性别).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call XYSetDiagInput(Row, rsTmp)
                Call XYEnterNextCell
            End If
        ElseIf Col = col增加 Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, col类型) = .TextMatrix(Row, col类型)
            .Cell(flexcpBackColor, lngRow, col诊断编码) = ColorUnEditCell      '灰蓝色
            
'            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'            .Cell(flexcpBackColor, GetRow(3), .FixedRows, GetRow(3), .Cols - 1) = &HC0FFC0
            
            .Row = lngRow: .Col = col诊断描述
            .ShowCell .Row, .Col
        ElseIf Col = colDel Then
            Call vsDiagXY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
        
        With vsDiagXY
            If Col = col出院情况 Then
                KeyAscii = 0
                If .ComboIndex <> -1 Then
                    '此时.TextMatrix尚未更新,所以取ComboItem
                    .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                    If Not XYCellEditable(Row, col是否未治) Then
                        .TextMatrix(Row, col是否未治) = ""
                    End If
                    .Tag = ""
                    mblnChange = True
                    Call XYEnterNextCell
                End If
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理西医诊断项目的输入
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim bln分化程度 As Boolean
    
    With vsDiagXY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '损伤中毒选择多条时的处理
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, col类型) = .TextMatrix(lngRow, col类型)
                    End If
                    '确定当前显示行
                    If Val(.TextMatrix(lngRow + 1, col类型)) = Val(.TextMatrix(lngRow, col类型)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, col类型)) = Val(.TextMatrix(lngRow, col类型)) Then
                                lngRow = j
                                If .TextMatrix(j, col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col类型) = .TextMatrix(lngRow - 1, col类型)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, col类型) = .TextMatrix(lngRow - 1, col类型)
                    End If
                End If
                
                If .TextMatrix(lngRow, col诊断类型) = "出院诊断" Then
                    If Nvl(rsInput!编码) = "" Then
                        bln分化程度 = False
                    Else
                        bln分化程度 = ((InStr("C", UCase(Left(rsInput!编码, 1)))) > 0) Or ((InStr("D0", UCase(Left(rsInput!编码, 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(rsInput!编码, 4)))) > 0)
                    End If
                    cboinfo(cbo分化程度).Enabled = bln分化程度
                    lblInfo(lbl分化程度).Enabled = bln分化程度
                    lblInfo(lbl最高诊断依据).Enabled = bln分化程度
                    cboinfo(cbo最高诊断依据).Enabled = bln分化程度
                End If
                .TextMatrix(lngRow, col诊断编码) = "" & rsInput!编码
                .TextMatrix(lngRow, col诊断描述) = "" & rsInput!名称
                .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
                
                '根据诊断确定疾病,或根据疾病确定诊断
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col疾病ID) = ""
                    StrSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col诊断ID) = ""
                    StrSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If optInput(0).Value Then
                        .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col诊断编码) = ""
            .TextMatrix(lngRow, col诊断描述) = .EditText
            .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
        End If
        
        .Cell(flexcpForeColor, 1, col是否疑诊, .Rows - 1, col是否疑诊) = vbRed
        
        '设置诊断符合情况
        Call Set西医诊断相关(lngRow)
        Call Set病原学
        
        .Tag = ""
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set西医诊断相关(ByVal lngRow As Long)
    With vsDiagXY
        If lngRow > .Rows - 1 Then Exit Sub
        
        '诊断符合情况
        If .TextMatrix(lngRow, 0) = "门诊诊断" Or .TextMatrix(lngRow, 0) = "出院诊断" Then
            Call Set诊断符合情况(cbo门诊与出院)
            Call Set诊断符合情况(cbo门诊与入院)
        End If
        If .TextMatrix(lngRow, 0) = "入院诊断" Or .TextMatrix(lngRow, 0) = "出院诊断" Then
            Call Set诊断符合情况(cbo入院与出院)
            Call Set诊断符合情况(cbo门诊与入院)
        End If
        If .TextMatrix(lngRow, 0) = "病理诊断" Then
            Call Set诊断符合情况(cbo放射与病理)
            Call Set诊断符合情况(cbo临床与病理)
            If vsDiagXY.TextMatrix(lngRow, col诊断描述) <> "" Then
                txtInfo(txt病理号).Enabled = True
                txtInfo(txt病理号).BackColor = vbWindowBackground
            Else
                txtInfo(txt病理号).Enabled = False
                txtInfo(txt病理号).BackColor = &H8000000F
            End If
        End If
    End With
End Sub

Private Sub Set中医诊断相关(ByVal lngRow As Long)
    With vsDiagZY
        If lngRow > .Rows - 1 Then Exit Sub
        If .TextMatrix(lngRow, 0) = "门诊诊断" Or .TextMatrix(lngRow, 0) = "主要诊断" Then
            Call Set诊断符合情况(cbo中医门诊与出院)
        End If
        If .TextMatrix(lngRow, 0) = "入院诊断" Or .TextMatrix(lngRow, 0) = "主要诊断" Then
            Call Set诊断符合情况(cbo中医入院与出院)
        End If
    End With
End Sub

Private Sub Set病原学()
    With vsDiagXY
        '院内感染与病原学诊断
        If Trim(.TextMatrix(GetRow(5), col诊断描述)) = "" Then
            chkInfo(chk病原学).Value = 0
            chkInfo(chk病原学).Enabled = False
            chkInfo(chk病原学).TabStop = False
            Call chkInfo_Click(chk病原学)
        ElseIf Not chkInfo(chk病原学).Enabled Then
            chkInfo(chk病原学).Enabled = True
            chkInfo(chk病原学).TabStop = True
        End If
    End With
End Sub

Private Function GetRow(ByVal lng诊断类型 As Long) As Long
'功能：返回指定诊断类型的第一诊断行
    If InStr(",11,12,13,", "," & lng诊断类型 & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng诊断类型), , colzy类型)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng诊断类型), , col类型)
    End If
End Function

Private Sub Set诊断符合情况(ByVal intIdx As Integer)
'功能：对诊断符合情况进行缺省值设置以及检查是否可以输入
'参数：intIdx=要设置的符合情况控件
    Dim i As Long
    
    With vsDiagXY
        '门诊与出院：门诊诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        If intIdx = cbo门诊与出院 Then
            If Trim(.TextMatrix(GetRow(1), col诊断描述)) = "" And Trim(.TextMatrix(GetRow(3), col诊断描述)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '可以改时缺省为符合
            Else
                If Trim(.TextMatrix(GetRow(1), col诊断描述)) = "" Or Trim(.TextMatrix(GetRow(3), col诊断描述)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(1), col诊断描述) <> .TextMatrix(GetRow(3), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(1), col诊断描述) = .TextMatrix(GetRow(3), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '入院与出院：入院诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        If intIdx = cbo入院与出院 Then
            If Trim(.TextMatrix(GetRow(2), col诊断描述)) = "" And Trim(.TextMatrix(GetRow(3), col诊断描述)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '可以改时缺省为符合
            Else
                If Trim(.TextMatrix(GetRow(2), col诊断描述)) = "" Or Trim(.TextMatrix(GetRow(3), col诊断描述)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(2), col诊断描述) <> .TextMatrix(GetRow(3), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(2), col诊断描述) = .TextMatrix(GetRow(3), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '门诊与入院：门诊诊断和入院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        If intIdx = cbo门诊与入院 Then
            If Trim(.TextMatrix(GetRow(1), col诊断描述)) = "" And Trim(.TextMatrix(GetRow(2), col诊断描述)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '可以改时缺省为符合
            Else
                If Trim(.TextMatrix(GetRow(1), col诊断描述)) = "" Or Trim(.TextMatrix(GetRow(2), col诊断描述)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(1), col诊断描述) <> .TextMatrix(GetRow(2), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(1), col诊断描述) = .TextMatrix(GetRow(2), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '放射与病理、临床与病理：录入病理诊断后可以录入，缺省为符合。
        If intIdx = cbo放射与病理 Or intIdx = cbo临床与病理 Then
            cboinfo(intIdx).Enabled = .TextMatrix(GetRow(6), col诊断描述) <> ""
            If Not cboinfo(intIdx).Enabled Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 0) '不可以改时缺省为未做
                cboinfo(intIdx).BackColor = vbButtonFace
            Else
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                cboinfo(intIdx).BackColor = vbWindowBackground
            End If
        End If
    End With
    
    '临床与尸检：勾选尸检后可以录入，缺省为符合。
    If intIdx = cbo临床与尸检 Then
        cboinfo(intIdx).Enabled = chkInfo(chk尸检).Value = 1
        If Not cboinfo(intIdx).Enabled Then
            Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 0) '不可以改时缺省为未做
            cboinfo(intIdx).BackColor = vbButtonFace
        Else
            Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
            cboinfo(intIdx).BackColor = vbWindowBackground
        End If
    End If
    
    '术前与术后：输入手术情况后可以录入，缺省为符合。
    If intIdx = cbo术前与术后 Then
        With vsOPS
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col手术名称)) <> "" Then Exit For
            Next
            If Not i <= .Rows - 1 Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 0) '不可以改时缺省为未做
            Else
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
            End If
        End With
    End If
    
    With vsDiagZY
        '中医门诊与出院：门诊诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        If intIdx = cbo中医门诊与出院 Then
            If Trim(.TextMatrix(GetRow(11), col诊断描述)) = "" And Trim(.TextMatrix(GetRow(13), col诊断描述)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '可以改时缺省为符合
            Else
                If Trim(.TextMatrix(GetRow(11), col诊断描述)) = "" Or Trim(.TextMatrix(GetRow(13), col诊断描述)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(11), col诊断描述) <> .TextMatrix(GetRow(13), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(11), col诊断描述) = .TextMatrix(GetRow(13), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '中医入院与出院：入院诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        If intIdx = cbo中医入院与出院 Then
            If Trim(.TextMatrix(GetRow(12), col诊断描述)) = "" And Trim(.TextMatrix(GetRow(13), col诊断描述)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '可以改时缺省为符合
            Else
                If Trim(.TextMatrix(GetRow(12), col诊断描述)) = "" Or Trim(.TextMatrix(GetRow(13), col诊断描述)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(12), col诊断描述) <> .TextMatrix(GetRow(13), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(12), col诊断描述) = .TextMatrix(GetRow(13), col诊断描述) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
    End With
End Sub

Private Sub KSSEnterNextCell()
    With vsKSS
        If .Row = .Rows - 1 And .Col = .Cols - 1 And .TextMatrix(.Row, .FixedCols) = "" Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        Else
            If .Row + 1 > .Rows - 1 And .Col = .Cols - 1 Then
                If .Rows - .FixedRows >= 10 Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                End If
                .AddItem "": Call SetKSSSerial
            End If
            If .Col = .Cols - 1 Then
                .Row = .Row + 1
                .Col = .FixedCols
                .ShowCell .Row, .Col
            Else
                .Col = .Col + 1
                .ShowCell .Row, .Col
            End If
        End If
    End With
End Sub

Private Sub KSSSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理抗生素项目的输入
    With vsKSS
        If Not rsInput Is Nothing Then
            '判断是否是修改
            If .RowData(lngRow) & "" <> "" Then
                If InStr(mstrDelete, .RowData(lngRow) & "") <= 0 Then
                    mstrDelete = mstrDelete & IIf(mstrDelete <> "", ",", "") & .RowData(lngRow)
                End If
            End If
            .TextMatrix(lngRow, 1) = Nvl(rsInput!名称)
            .RowData(lngRow) = Val(rsInput!ID)
        Else
            .TextMatrix(lngRow, 1) = .EditText
        End If
        .Cell(flexcpData, lngRow, 1) = .TextMatrix(lngRow, 1)
        mblnChange = True
        .Tag = ""
    End With
End Sub

Private Sub TSJCEnterNextCell()
    With vsTSJC
        If .Row = .Rows - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If .Row + 1 > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub TSJCSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理特殊检查项目的输入
    With vsTSJC
        If Not rsInput Is Nothing Then
            .TextMatrix(lngRow, 1) = Nvl(rsInput!名称)
        Else
            .TextMatrix(lngRow, 1) = .EditText
        End If
        .Cell(flexcpData, lngRow, 1) = .TextMatrix(lngRow, 1)
        mblnChange = True
    End With
End Sub

Private Sub XYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagXY
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col诊断描述) To col增加
                If XYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col增加 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function XYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagXY
        '隐藏列不可编辑
        If .ColHidden(lngCol) Then Exit Function
        
        If lngCol = col诊断描述 And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col诊断类型) = "入院诊断" And mlngDiagnosisType = 2 Or .TextMatrix(lngRow, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 1 Then
                If .TextMatrix(lngRow, col诊断描述) <> "" And .TextMatrix(lngRow, col诊断类型) <> .TextMatrix(lngRow - 1, col诊断类型) Then
                    '首要诊断不允许改
                    Exit Function
                End If
            End If
            '合并路径
            If Not CheckMergePath(mlng病人ID, mlng主页ID, Val(.TextMatrix(lngRow, col类型)), Val(.TextMatrix(lngRow, col疾病ID))) Then Exit Function
        End If
        If lngCol = col诊断描述 Then
            '两条路径以上
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                    '导入诊断不允许该
                    Exit Function
                End If
            End If
        End If
        If lngCol = col诊断描述 And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col诊断类型) = "出院诊断" And mlngDiagnosisType <= 2 Then
                '正常完成的出院诊断不允许改
                Exit Function
            End If
        End If
        '必须先输入诊断
        If .TextMatrix(lngRow, col诊断描述) = "" Then
            If lngCol = col出院情况 Or lngCol = col备注 Or lngCol = col是否未治 Or lngCol = col是否疑诊 Or lngCol = col增加 Then
                Exit Function
            End If
        End If
        If lngCol = col诊断编码 Then Exit Function
        
        If lngCol = col增加 Then
            If Val(.TextMatrix(lngRow, col类型)) = 3 Then
                If .TextMatrix(lngRow, col诊断类型) = "出院诊断" Then Exit Function
            End If
        End If
        
        '出院诊断和院内感染允许输入出院情况(因为可能院内感染在出院时已经好转或治愈了)
        If Val(.TextMatrix(lngRow, col类型)) = 3 Or Val(.TextMatrix(lngRow, col类型)) = 5 Or Val(.TextMatrix(lngRow, col类型)) = 10 Then
            '出院诊断必须依次输入(尚未输入时)
            If .TextMatrix(lngRow, col诊断描述) = "" And Val(.TextMatrix(lngRow, col类型)) = 3 Then
                If Val(.TextMatrix(lngRow - 1, col类型)) = 3 And .TextMatrix(lngRow - 1, col诊断描述) = "" Then
                    Exit Function
                End If
            End If

            '出院情况为"其他"时才可以设置是否未治
            If .TextMatrix(lngRow, col出院情况) <> "其他" And lngCol = col是否未治 Then
                Exit Function
            End If
        ElseIf lngCol = col出院情况 Or lngCol = col是否未治 Then
            Exit Function
        End If
        
        '入院病情只能在出院诊断和其他诊断行填写
        If lngCol = col入院病情 Then
            If .TextMatrix(lngRow, col类型) <> "3" Then
                Exit Function
            End If
        End If
    End With
    XYCellEditable = True
End Function

Private Function ZYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagZY
        '隐藏列不可编辑
        If .ColHidden(lngCol) Then Exit Function
        
        If lngCol = col诊断描述 And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col诊断类型) = "入院诊断" And mlngDiagnosisType = 12 Or .TextMatrix(lngRow, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 11 Then
                If .TextMatrix(lngRow, col诊断描述) <> "" And .TextMatrix(lngRow, col诊断类型) <> .TextMatrix(lngRow - 1, col诊断类型) Then
                    '首要诊断不允许改
                    Exit Function
                End If
            End If
            '合并路径
            If Not CheckMergePath(mlng病人ID, mlng主页ID, Val(.TextMatrix(lngRow, colzy类型)), Val(.TextMatrix(lngRow, colzy疾病ID))) Then Exit Function
        End If
        If lngCol = col诊断描述 Then
            '两条路径以上
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                    '导入诊断不允许该
                    Exit Function
                End If
            End If
        End If
        If lngCol = col诊断描述 And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col诊断类型) = "主要诊断" And mlngDiagnosisType > 10 Then
                '正常完成的出院诊断不允许改
                Exit Function
            End If
        End If
        '必须先输入诊断
        If .TextMatrix(lngRow, col诊断描述) = "" Then
            If lngCol = col出院情况 Or lngCol = col备注 Or lngCol = colzy增加 Then Exit Function
        End If
        If lngCol = col诊断编码 Then Exit Function
        
        If lngCol = colzy增加 Then
            If Val(.TextMatrix(lngRow, colzy类型)) = 13 Then
                If .TextMatrix(lngRow, col诊断类型) = "主要诊断" Then Exit Function
            End If
        End If
        
        If Val(.TextMatrix(lngRow, colzy类型)) = 13 Then
            '出院诊断必须依次输入(尚未输入时)
            If .TextMatrix(lngRow, col诊断描述) = "" Then
                If Val(.TextMatrix(lngRow - 1, colzy类型)) = 13 And .TextMatrix(lngRow - 1, col诊断描述) = "" Then
                    Exit Function
                End If
            End If
        ElseIf lngCol = col出院情况 Then
            '非出院诊断时不允许输入
            If Val(.TextMatrix(lngRow, colzy类型)) <> 13 Then Exit Function
        End If
        '入院病情只能在主要诊断和其他诊断行填写
        If lngCol = col入院病情 Then
            If .TextMatrix(lngRow, colzy类型) <> "13" Then
                Exit Function
            End If
        End If
        '必须先输诊断再输证候
        If lngCol = col中医证候 Then
            If .TextMatrix(lngRow, col诊断描述) = "" Then Exit Function
        End If
    End With
    ZYCellEditable = True
End Function

Private Sub ZYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagZY
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col诊断描述) To colzy增加
                If ZYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= colzy增加 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理中医诊断项目的输入
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '其他诊断选择多条时的处理
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, colzy类型) = .TextMatrix(lngRow, colzy类型)
                    End If
                    '确定当前显示行
                    If Val(.TextMatrix(lngRow + 1, colzy类型)) = Val(.TextMatrix(lngRow, colzy类型)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, colzy类型)) = Val(.TextMatrix(lngRow, colzy类型)) Then
                                lngRow = j
                                If .TextMatrix(j, col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, colzy类型) = .TextMatrix(lngRow - 1, colzy类型)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy类型) = .TextMatrix(lngRow - 1, colzy类型)
                    End If
                End If
                
                If InStr(.TextMatrix(lngRow, col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断描述), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, col诊断描述), InStrRev(.TextMatrix(lngRow, col诊断描述), "("))
                End If
                                        
                .TextMatrix(lngRow, col诊断编码) = "" & rsInput!编码
                .TextMatrix(lngRow, col诊断描述) = "" & rsInput!名称 & strTmp
                .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
                                
                
                '根据诊断确定疾病,或根据疾病确定诊断
                If optInput(2).Value Then
                    .TextMatrix(lngRow, colzy诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, colzy疾病ID) = ""
                    StrSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, colzy疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, colzy诊断ID) = ""
                    StrSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If optInput(2).Value Then
                        .TextMatrix(lngRow, colzy疾病ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, colzy诊断ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                '中医根据疾病诊断参考取证候
                Call Set中医证候(lngRow, Val(.TextMatrix(lngRow, colzy诊断ID)))
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col诊断编码) = ""
            .TextMatrix(lngRow, col诊断描述) = .EditText
            .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
            .TextMatrix(lngRow, colzy诊断ID) = ""
            .TextMatrix(lngRow, colzy疾病ID) = ""
            .TextMatrix(lngRow, colzy证候ID) = ""
        End If
        
        '设置诊断符合情况
        Call Set中医诊断相关(lngRow)
        
        .Tag = ""
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Set中医证候(ByVal lngRow As Long, ByVal lng诊断ID As Long, Optional ByVal rsInput As Recordset) As Boolean
'功能：中医根据疾病诊断参考取证候
'参数：rsInput-如果不为空，则输出指定的中药证候记录集
'返回：是否有对应关系
    Dim rsTmp As Recordset
    Dim StrSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    With vsDiagZY
        '去掉已有的证候
        If InStr(.TextMatrix(lngRow, col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断描述), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, col诊断描述), 1, InStrRev(.TextMatrix(lngRow, col诊断描述), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, col诊断描述)
        End If
        If rsInput Is Nothing Then
            If lng诊断ID <> 0 Then
                StrSQL = "Select Distinct a.证候序号 as ID,a.证候ID,a.证候名称,b.编码 as 证候编码" & _
                    " From 疾病诊断参考 A,疾病编码目录 B" & _
                    " Where a.证候ID=b.ID(+) And a.诊断ID=[1] And a.证候名称 is Not NULL" & _
                    " Order by a.证候序号"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "中医证候", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng诊断ID)
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, colzy证候ID) = Nvl(rsTmp!证候id)
                    If Not IsNull(rsTmp!证候名称) Then
                        .TextMatrix(lngRow, col诊断描述) = strTmp
                        .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
                        .TextMatrix(lngRow, col中医证候) = Nvl(rsTmp!证候名称)
                        .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
                        If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
                        mblnChange = True
                        .Tag = ""
                    End If
                    Set中医证候 = True
                Else
                    If blnCancel Then
                        Set中医证候 = True
                        If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col中医证候)
                    Else
                        Set中医证候 = False
                    End If
                End If
            Else
                Set中医证候 = False
            End If
        Else
            .TextMatrix(lngRow, colzy证候ID) = Nvl(rsInput!项目ID)
            .TextMatrix(lngRow, col诊断描述) = strTmp
            .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
            .TextMatrix(lngRow, col中医证候) = Nvl(rsInput!名称)
            .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
            .Tag = ""
            mblnChange = True
        End If
    End With
End Function

Private Function GetSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str性别 As String, Optional ByVal strOtherInfo As String) As String
'功能：获得查询西医诊断的SQL
'参数：intType:获取的SQL类型,0-西医诊断，1-中医诊断，2-手术操作
'    strInput-查询条件，str性别--病人的性别
'   strOtherInfo:中医诊断-疾病编码种类
'返回：strsql--查询诊断的SQL
    Dim StrSQL As String
    
    If cboinfo(cbo性别).Text Like "*男*" Then
        str性别 = "男"
    ElseIf cboinfo(cbo性别).Text Like "*女*" Then
        str性别 = "女"
    End If
    
    Select Case intType
        Case 0 '西医诊断
            If optInput(0).Value Then
                '按诊断输入:西医部份，一个诊断可能属于多个分类
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    StrSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                    " From 疾病诊断目录 A,疾病诊断别名 B" & _
                    " Where A.ID=B.诊断ID And A.类别=1" & _
                    " And B.码类=[5] And (" & StrSQL & ")" & _
                    " Order by A.编码"
            Else
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "名称 Like [2]" '输入汉字时只匹配名称
                Else
                    StrSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(mint简码 = 0, "简码", "五笔码") & " Like [2]"
                End If
                StrSQL = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录 Where Instr([3],类别)>0 And (" & StrSQL & ")" & _
                    IIf(str性别 <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"
            End If
        
        Case 1 '中医诊断
            If optInput(2).Value And strOtherInfo <> "Z" Then
                '按诊断输入:中医部份，一个诊断可能属于多个分类
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    StrSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                    " From 疾病诊断目录 A,疾病诊断别名 B" & _
                    " Where A.ID=B.诊断ID And A.类别=2" & _
                    " And B.码类=[4] And (" & StrSQL & ")" & _
                    " Order by A.编码"
            Else
                'B-中医疾病编码
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "名称 Like [2]" '输入汉字时只匹配名称
                Else
                    StrSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(mint简码 = 0, "简码", "五笔码") & " Like [2]"
                End If
                StrSQL = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录" & _
                    " Where 类别='" & IIf(strOtherInfo = "", "B", strOtherInfo) & "' And (" & StrSQL & ")" & _
                    IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"
            End If
        Case 2 '手术操作
            If optInput(4).Value Then
                '按诊疗项目输入
                StrSQL = "Select distinct A.ID,A.编码,A.名称,A.操作类型 as 规模" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.类别='F' And A.服务对象 IN(2,3) And A.ID=B.诊疗项目ID" & _
                    IIf(str性别 <> "", " And Nvl(A.适用性别,0) IN(0,[4])", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or B.简码 Like [2] Or B.名称 Like [2])" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
            Else
                '按ICD9-CM3输入
                StrSQL = _
                    " Select distinct ID,编码,附码,名称,简码,说明" & _
                    " From 疾病编码目录 Where 类别='S'" & _
                    IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (编码 Like [1] Or 名称 Like [2] Or 简码 Like [2])" & _
                    " Order by 编码"
            End If
    End Select
    GetSQL = StrSQL
End Function

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str性别 As String, int诊断输入 As Integer
    Dim strInput As String, vPoint As POINTAPI
    
    With vsDiagXY
        If Col = col诊断描述 Then
            '.Cell(flexcpData, Row, Col) <> ""排除空行回车
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call XYEnterNextCell
            ElseIf .TextMatrix(Row, col诊断编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '判断加了前缀后的名称是否存在其他的诊断编码
                strInput = UCase(.EditText)
                StrSQL = GetSQL(0, strInput, str性别)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                If rsTmp.RecordCount <> 1 Then
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, col诊断描述) = .EditText
                Else
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                End If
                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
                .Tag = ""
                mblnChange = True
            Else
                If Val(.TextMatrix(Row, col类型)) = 1 Then
                    int诊断输入 = Val(Mid(gstr诊断输入, 1, 1))
                Else
                    int诊断输入 = Val(Mid(gstr诊断输入, 2, 1))
                End If
                If int诊断输入 = 0 Then int诊断输入 = 1
                
                strInput = UCase(.EditText)
                StrSQL = GetSQL(0, strInput, str性别)
                If int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                    '损伤中毒码：Y-损伤中毒的外部原因；病理诊断允许：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", _
                        Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                        If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn And rsTmp Is Nothing Then Call XYEnterNextCell '不是自由录入时，暂不跳到下一行，因为可能还要改描述内容
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(optInput(0).Value, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And ((int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0)) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                            'If mblnReturn Then Call XYEnterNextCell    '暂不跳到下一行，因为可能还要改描述内容
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col出院情况 Then
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            .Tag = ""
            mblnChange = True
        End If
        If Col = col诊断描述 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagZY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagZY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        Call vsDiagZY_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    If vsDiagZY.Tag = "未修改" Then vsDiagZY.Tag = "": mblnChange = True
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsDiagZY
        '清除图片
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, colzy增加) Is Nothing Then
                Set .Cell(flexcpPicture, i, colzy增加) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colzyDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colzyDel) = Nothing
            End If
        Next
        
        If Not ZYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
            
            If NewCol = col诊断描述 Then
                .ComboList = "..."
            ElseIf NewCol = col中医证候 Then
                If .TextMatrix(NewRow, col诊断描述) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            ElseIf NewCol = col出院情况 Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col入院病情 Then
                If .TextMatrix(NewRow, colzy类型) = "13" Then
                    .ComboList = "有|临床未确定|情况不明|无"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = colzy增加 Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colzyDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '显示图片
            If NewCol <> colzy增加 And .TextMatrix(NewRow, col诊断描述) <> "" And .TextMatrix(NewRow, 0) <> "主要诊断" Then
                Set .Cell(flexcpPicture, NewRow, colzy增加) = imgButtonNew.Picture
            End If
            '显示图片
            If NewCol <> colzyDel Then
                Set .Cell(flexcpPicture, NewRow, colzyDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colzy增加 Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    Dim str性别 As String, lngRow As Long
    Dim blnCancle As Boolean
    
    With vsDiagZY
        If Col = col诊断描述 Then
            If optInput(2).Value Then
                '按诊断输入:中医部份，一个诊断可能属于多个分类
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", mlng科室ID, , True, False)
            Else
                'B-中医疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", mlng科室ID, cboinfo(cbo性别).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call ZYSetDiagInput(Row, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = col中医证候 Then
            If optInput(2).Value Then
                '按诊断输入:先查是否有对应
                If Not Set中医证候(Row, Val(.TextMatrix(Row, colzy诊断ID))) Then
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, cboinfo(cbo性别).Text, True)
                Else
                    Exit Sub
                End If
            Else
                'Z-中医疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, cboinfo(cbo性别).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call Set中医证候(Row, 0, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = colzy增加 Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, colzy类型) = .TextMatrix(Row, colzy类型)
            .Cell(flexcpBackColor, lngRow, col诊断编码) = ColorUnEditCell      '灰蓝色
            
'            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'            .Cell(flexcpBackColor, GetRow(13), .FixedRows, GetRow(13), .Cols - 1) = &HC0FFC0
            
            .Row = lngRow: .Col = col诊断描述
            .ShowCell .Row, .Col
        ElseIf Col = colzyDel Then
            Call vsDiagZY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagZY_Click()
    With vsDiagZY
        If (.MouseCol = colzy增加 Or .MouseCol = colzyDel) And .MouseRow >= .FixedRows Then
            If .MouseCol = colzy增加 Then
                If .TextMatrix(.MouseRow, col诊断描述) = "" Or .TextMatrix(.MouseRow, 0) = "主要诊断" Then Exit Sub
            End If
        
            .Select .MouseRow, .MouseCol
            Call vsDiagZY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagZY
        If Col = col出院情况 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断描述 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col诊断描述) <> "" Then
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col诊断类型) = "入院诊断" And mlngDiagnosisType = 12 Or .TextMatrix(.Row, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 11 Then
                        If .TextMatrix(.Row, col诊断类型) <> .TextMatrix(.Row - 1, col诊断类型) Then
                            '首要诊断不允许改
                            Exit Sub
                        End If
                    End If
                End If
                '合并路径
                If Not CheckMergePath(mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, colzy类型)), Val(.TextMatrix(.Row, colzy疾病ID))) Then Exit Sub
                '两条路径以上
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                        '导入诊断不允许该
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col诊断类型) = "主要诊断" And mlngDiagnosisType > 10 Then
                        '正常完成的出院诊断不允许改
                        Exit Sub
                    End If
                End If
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, colzy类型))
                    .Cell(flexcpText, .Row, .FixedRows, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedRows, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, colzy类型) = i
                    
                    '下面的同类诊断数据上移
                    If .TextMatrix(.Row, col诊断类型) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col诊断类型) = "" Then
                                '下一行为无标题的增加行时，数据才上移，否则当前行为有标题时只清空行
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, colzy类型)) = Val(.TextMatrix(.Row, colzy类型)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, colzy类型) = Val(.TextMatrix(.Row, colzy类型))
                                        
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, colzy类型)) <> Val(.TextMatrix(i, colzy类型)) Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    
'                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'                    .Cell(flexcpBackColor, GetRow(13), .FixedRows, GetRow(13), .Cols - 1) = &HC0FFC0
                    
                    '设置诊断符合情况
                    Call Set中医诊断相关(.Row)

                    mblnChange = True
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col诊断类型) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ZYEnterNextCell
        Else
            If .Col = col诊断描述 Or .Col = col中医证候 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
        
        With vsDiagZY
            If Col = col出院情况 Then
                KeyAscii = 0
                
                '此时.TextMatrix尚未更新,所以取ComboItem
                .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                mblnChange = True
                .Tag = ""
                Call ZYEnterNextCell
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ZYCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim str性别 As String, int诊断输入 As Integer
    
    With vsDiagZY
        If Col = col诊断描述 Or Col = col中医证候 Then
            '.Cell(flexcpData, Row, Col) <> ""排除空行回车
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
                '中医症候则清除备份数据
                If Col = col中医证候 Then
                    .Cell(flexcpData, Row, Col) = ""
                End If
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ZYEnterNextCell
            ElseIf Col = col诊断描述 And .TextMatrix(Row, col诊断编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                StrSQL = GetSQL(1, strInput, str性别)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, str性别, mint简码 + 1)
                If rsTmp.RecordCount = 1 Then
                    Call ZYSetDiagInput(Row, rsTmp):
                    .EditText = .Text
                Else
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, col诊断描述) = .EditText
                End If
                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
                .Tag = ""
                mblnChange = True
            Else
                If Val(.TextMatrix(Row, colzy类型)) = 11 Then
                    int诊断输入 = Val(Mid(gstr诊断输入, 1, 1))
                Else
                    int诊断输入 = Val(Mid(gstr诊断输入, 2, 1))
                End If
                If int诊断输入 = 0 Then int诊断输入 = 1
                
                strInput = UCase(.EditText)
                StrSQL = GetSQL(1, strInput, str性别, IIf(Col = col诊断描述, "B", "Z"))
                If Col = col诊断描述 Then
                    If int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, str性别, mint简码 + 1)
                            If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                        End If
                        Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                        If mblnReturn And rsTmp Is Nothing Then Call ZYEnterNextCell '不是自由录入时，暂不跳到下一行，因为可能还要改描述内容
                    Else
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(optInput(2).Value, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1)
                        If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                            Cancel = True
                        Else
                            '检查诊断输入方式
                            If rsTmp Is Nothing And ((int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0)) Then
                                MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                                Cancel = True
                            Else
                                Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                                'If mblnReturn Then Call ZYEnterNextCell '暂不跳到下一行，因为可能还要改描述内容
                            End If
                        End If
                    End If
                ElseIf Col = col中医证候 Then
                    If optInput(2).Value Then
                        '按诊断输入:先查是否有对应
                        If Set中医证候(Row, Val(.TextMatrix(Row, colzy诊断ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call Set中医证候(Row, 0, rsTmp)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub VsGriedFocuesMove(ByVal vsBill As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal KeyCode As Integer, _
        Optional lngFiexCol As Long = 0, Optional lngFiexCol1 As Long = -1)
    '------------------------------------------------------------------------------------------------------------
    '功能:按一定规则移动单元格
    '参数:vsBill-表格控件
    '       lngRow-当前行
    '       lngCol-当前列
    '       KeyCode-按键
    '       lngFiexCol-判断是否移到或加入行的固定列
    '       lngFiexCol1-判断是否移到或加入行的固定列(但同时要满足lngFiexCol列)
    '编制:刘兴宏
    '日期:2007/05/18
    '------------------------------------------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim strCurrValue As String
    If lngCol = lngFiexCol Then
        strCurrValue = vsBill.EditText
    Else
        strCurrValue = ""
    End If
    
    With vsBill
        
        Select Case lngCol
        Case 0
            If Trim(.TextMatrix(lngRow, lngFiexCol)) = "" And strCurrValue = "" Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            .Col = lngCol + 1
            GoTo ShowCell:
        Case Else
            If lngCol >= .Cols - 1 Then
                If lngRow < .Rows - 1 Then
                    .Row = lngRow + 1
                    .Col = 0
                    GoTo ShowCell:
                    Exit Sub
                End If
                If Trim(.TextMatrix(lngRow, lngFiexCol)) <> "" Then
                    If lngFiexCol1 > 0 Then
                        If Trim(.TextMatrix(lngRow, lngFiexCol1)) <> "" Then
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 0
                        End If
                    Else
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                        .Col = 0
                    End If
                End If
                GoTo ShowCell:
                Exit Sub
            End If
            .Col = lngCol + 1
         End Select
ShowCell:
        .ShowCell .Row, .Col
    End With
End Sub


Private Function CheckInPutIsDate(ByVal vsObj As Object, lngRow As Long, lngCol As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:检查所输入的日期是否合法
    '参数:lngRow -行,lngCol -列
    '返回:日期合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/05/21
    '---------------------------------------------------------------------------------------------------------
    Dim strKEY As String
    Dim str进入时间 As String, str退出时间 As String
    Dim str入院时间 As String
    str入院时间 = txtInfo(txt入院时间).Text
    
        
    strKEY = Trim(vsObj.EditText)
    strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
    strKEY = Replace(strKEY, Chr(10), "")
    If strKEY <> "" Then
        
        If Not IsDate(strKEY) Then
            MsgBox vsObj.TextMatrix(0, lngCol) & "必须为日期型,请重新输入！", vbInformation + vbDefaultButton1, Me.Caption
             vsObj.EditSelStart = 0
             vsObj.EditSelLength = 1000
            Exit Function
        End If
        Select Case lngCol
        Case 1
            str进入时间 = strKEY
            str退出时间 = Trim(vsObj.TextMatrix(lngRow, 2))
            If str退出时间 <> "" And str进入时间 > str退出时间 Then
                MsgBox "注:" & vbCrLf & "  进入时间大于了退出时间,请检查！", vbInformation + vbDefaultButton1, Me.Caption
                Exit Function
            End If
        Case Else
            str进入时间 = Trim(vsObj.TextMatrix(lngRow, 1))
            str退出时间 = strKEY
 
            If str进入时间 <> "" And CDate(str进入时间) >= CDate(str退出时间) Then
                MsgBox "注:" & vbCrLf & "  退出时间小于了进入时间,请检查！", vbInformation + vbDefaultButton1, Me.Caption
                Exit Function
            End If
        End Select
    End If
    CheckInPutIsDate = True
End Function

Private Sub vsfMain_EnterCell()
    Select Case vsfMain.Col
        Case 0, 3, 6
            vsfMain.Editable = flexEDNone
        Case 1, 4
            If vsfMain.BackColor = vbButtonFace Then Exit Sub
            If InStr(vsfMain.TextMatrix(vsfMain.Row, vsfMain.Col + 1), ",") > 0 Then
                vsfMain.ColComboList(vsfMain.Col) = Replace(vsfMain.TextMatrix(vsfMain.Row, vsfMain.Col + 1), ",", "|")
            Else
                vsfMain.ColComboList(vsfMain.Col) = ""
            End If
            If vsfMain.TextMatrix(vsfMain.Row, vsfMain.Col - 1) <> "" Then
                vsfMain.Editable = flexEDKbdMouse
            Else
                vsfMain.Editable = flexEDNone
            End If
    End Select
End Sub
Private Sub vsfMain_KeyPress(KeyAscii As Integer)
    If vsfMain.Rows <= 1 Then Exit Sub
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case vsfMain.Col
            Case 0, 3, 6
                vsfMain.Col = vsfMain.Col + 1
            Case 1, 4
                If vsfMain.Col = 4 And vsfMain.Row <> vsfMain.Rows - 1 Then
                    vsfMain.Col = 0
                    vsfMain.Row = vsfMain.Row + 1
                ElseIf vsfMain.Col = 4 And vsfMain.Row = vsfMain.Rows - 1 Then
                    Call zlCommFun.PressKey(vbKeyTab)
                Else
                    vsfMain.Col = vsfMain.Col + 3
                End If
        End Select
        vsfMain.ShowCell vsfMain.Row, vsfMain.Col
    End If
End Sub

Private Sub vsfMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim sngNum1, sngNum2 As Single
    If InStr(vsfMain.TextMatrix(Row, Col + 1), "...") > 0 Then
        sngNum1 = Mid(vsfMain.TextMatrix(Row, Col + 1), 1, InStr(vsfMain.TextMatrix(Row, Col + 1), "...") - 1)
        sngNum2 = Mid(vsfMain.TextMatrix(Row, Col + 1), InStr(vsfMain.TextMatrix(Row, Col + 1), "...") + 3)
        If Not IsNumeric(vsfMain.EditText) Then
            Cancel = True
        ElseIf CSng(vsfMain.EditText) < sngNum1 Or CSng(vsfMain.EditText) > sngNum2 Then
            MsgBox "数据应该在" & vsfMain.TextMatrix(Row, Col + 1) & "的范围以内!", vbInformation, gstrSysName
            Cancel = True
        End If
    ElseIf InStr(vsfMain.TextMatrix(Row, Col + 1), "-") > 0 Then
        If InStr(vsfMain.TextMatrix(Row, Col + 1), "-") = 1 Then
            sngNum1 = Mid(vsfMain.TextMatrix(Row, Col + 1), 2, InStr(2, vsfMain.TextMatrix(Row, Col + 1), "-") - 1)
            sngNum2 = Mid(vsfMain.TextMatrix(Row, Col + 1), InStr(2, vsfMain.TextMatrix(Row, Col + 1), "-") + 1)
        Else
            sngNum1 = Mid(vsfMain.TextMatrix(Row, Col + 1), 1, InStr(1, vsfMain.TextMatrix(Row, Col + 1), "-") - 1)
            sngNum2 = Mid(vsfMain.TextMatrix(Row, Col + 1), InStr(1, vsfMain.TextMatrix(Row, Col + 1), "-") + 1)
        End If
        If Not IsNumeric(vsfMain.EditText) Then
            Cancel = True
        ElseIf CSng(vsfMain.EditText) < sngNum1 Or CSng(vsfMain.EditText) > sngNum2 Then
            MsgBox "数据应该在" & vsfMain.TextMatrix(Row, Col + 1) & "的范围以内!", vbInformation, gstrSysName
            Cancel = True
        End If
    ElseIf vsfMain.TextMatrix(Row, Col + 1) = "" Then
        If zlCommFun.ActualLen(vsfMain.EditText) > mlngSize Then
            MsgBox "输入长度不能大于" & "[" & mlngSize & "]", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
    If Cancel = False Then mblnChange = True: vsfMain.Tag = ""
    
End Sub

Private Sub vsKSS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsKSS_AfterRowColChange(-1, -1, vsKSS.Row, vsKSS.Col)
End Sub

Private Sub vsKSS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsKSS.ColComboList(kss名称) = "..."
    vsKSS.ColComboList(kss使用阶段) = " |术前|术中|术后|围手术期"
    vsKSS.ColComboList(KSS联合用药) = "Ⅰ种|Ⅱ联|Ⅲ联|Ⅳ联|>Ⅳ联"
    vsKSS.ColComboList(kss用药目的) = " |预防|治疗|预防和治疗"
End Sub

Private Sub vsKSS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    With vsKSS
        If Col = kss名称 Then
            strSQLItem = _
                " From 诊疗项目目录 A,药品特性 B" & _
                " Where A.ID=B.药名ID And A.类别='5' And A.服务对象 IN(2,3) And Nvl(b.抗生素, 0) <> 0" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) "
            StrSQL = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位" & _
                " From 诊疗分类目录 Where 类型=1 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With ID In (Select A.分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
                " Group by ID,上级ID,编码,名称"
            StrSQL = StrSQL & " Union ALL" & _
                " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位" & _
                strSQLItem & " Order By 末级,级ID Desc,编码"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, "抗菌药物", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有抗菌药物数据可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call KSSSetDiagInput(Row, rsTmp)
                Call KSSEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub vsKSS_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    If KeyCode = vbKeyF4 Then
        Call zlCommFun.PressKey(vbKeySpace)
    ElseIf KeyCode = vbKeyDelete Then
        If MsgBox("确实要删除该行内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsKSS
                '判断是否是修改
                If .RowData(.Row) & "" <> "" Then
                    If InStr(mstrDelete, .RowData(.Row) & "") <= 0 Then
                        mstrDelete = mstrDelete & IIf(mstrDelete <> "", ",", "") & .RowData(.Row)
                    End If
                End If
                .RemoveItem .Row
                If .Rows < 4 Then .Rows = 4
                Call SetKSSSerial
                vsKSS.Tag = ""
            End With
            mblnChange = True
        End If
    ElseIf KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsKSS_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsKSS_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsKSS
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call KSSEnterNextCell
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsKSS_CellButtonClick(.Row, .Col)
            Else
                .ColComboList(kss名称) = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsKSS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    If Col = kss使用天数 And Len(vsKSS.EditText) > 18 And KeyAscii <> vbKeyBack And vsKSS.EditSelLength = 0 Then KeyAscii = 0
    If Col = kss用药目的 And LenB(StrConv(vsKSS.EditText, vbFromUnicode)) >= 200 And KeyAscii <> vbKeyBack And vsKSS.EditSelLength = 0 Then KeyAscii = 0
End Sub

Private Sub vsKSS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsKSS.EditSelStart = 0
    vsKSS.EditSelLength = zlCommFun.ActualLen(vsKSS.EditText)
End Sub

Private Sub vsKSS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsKSS
        If Col = kss名称 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call KSSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call KSSEnterNextCell
            Else
                strInput = UCase(.EditText)
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    StrSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B,药品特性 C" & _
                    " Where A.ID=B.诊疗项目ID And A.ID=C.药名ID And Nvl(c.抗生素, 0) <> 0" & _
                    " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And A.类别='5' And A.服务对象 IN(2,3) And B.码类=[3] And (" & StrSQL & ")" & _
                    " Order by A.编码"
                If zlCommFun.IsCharChinese(strInput) Then
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint简码 + 1)
                    '判断是否有数据
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "没有找到指定的抗菌药物。", vbInformation, gstrSysName
                        Cancel = True: .EditText = "": Exit Sub
                    End If
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                    Call KSSSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then Call KSSEnterNextCell
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "抗菌药物", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '判断是否有数据
                        If rsTmp Is Nothing Then
                            MsgBox "没有找到指定的抗菌药物。", vbInformation, gstrSysName
                            Cancel = True: .EditText = "": Exit Sub
                        End If
                        Call KSSSetDiagInput(Row, rsTmp)
                        .EditText = .Text
                        If mblnReturn Then Call KSSEnterNextCell
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = kss使用天数 Or Col = KSSDDD数 Then
            If (Not IsNumeric(.EditText) Or InStr(.EditText, "-") > 0 Or InStr(.EditText, "+") > 0) And .EditText <> "" Then
                MsgBox "请输入有效的数字。", vbInformation, Me.Caption
                Cancel = True
            Else
                If Len(.EditText) > 12 Then
                    MsgBox "请输入12位以下的数字。", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
                If .TextMatrix(Row, Col) <> .EditText Then .Tag = "": mblnChange = True
            End If
        Else
            '如果用户修改了，则提取的时候不影响这一项
            If .Cell(flexcpData, Row, Col) = "新增" Then .Cell(flexcpData, Row, Col) = ""
            If .TextMatrix(Row, Col) <> .EditText Then .Tag = "": mblnChange = True
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInput As String
    
    With vsOPS
        If Col = col手术日期 Then
            strInput = GetFullDate(.TextMatrix(Row, Col), False)
            If Not IsDate(strInput) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Else
                .TextMatrix(Row, Col) = strInput
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                mblnChange = True
                .Tag = ""
            End If
        ElseIf Col = col切口愈合 Then
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            mblnChange = True
            .Tag = ""
        ElseIf Col = col手术级别 Or Col = col麻醉类型 Then
            mblnChange = True
            .Tag = ""
        End If
    End With
    Call vsOPS_AfterRowColChange(-1, -1, vsOPS.Row, vsOPS.Col)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsOPS
        If Not OPSCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col手术编码 Or NewCol = col主刀医师 _
                Or NewCol = col助产护士 Or NewCol = col助手1 Or NewCol = col助手2 _
                Or NewCol = col麻醉方式 Or NewCol = col麻醉医师 Or (NewCol = col手术名称 And chkInfo(chk手术自由录入).Value) Then
                .ComboList = "..."
            ElseIf NewCol = col切口愈合 Then
                .ComboList = .ColData(NewCol)
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsOPS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str性别 As String, int性别 As Integer
    Dim vPoint As POINTAPI
    
    With vsOPS
        If Col = col手术编码 Or Col = col手术名称 Then
            If optInput(4).Value Then
                '按诊疗项目输入
                If cboinfo(cbo性别).Text Like "*男*" Then
                    int性别 = 1
                ElseIf cboinfo(cbo性别).Text Like "*女*" Then
                    int性别 = 2
                End If
                            
                StrSQL = "Select 0 as 末级,ID,上级ID,编码,名称,NULL as 规模" & _
                    " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                    " Union ALL " & _
                    " Select 1 as 末级,ID,分类ID as 上级ID,编码,名称,操作类型 as 规模" & _
                    " From 诊疗项目目录" & _
                    " Where 类别='F' And 服务对象 IN(2,3) And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
                    IIf(int性别 <> 0, " And Nvl(适用性别,0) IN(0,[2])", "") & _
                    " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)"
            Else
                '按ICD9-CM3输入
                If cboinfo(cbo性别).Text Like "*男*" Then
                    str性别 = "男"
                ElseIf cboinfo(cbo性别).Text Like "*女*" Then
                    str性别 = "女"
                End If
                StrSQL = _
                    " Select 0 as 末级,ID,上级ID," & _
                    " 类别||LPAD(序号,3,'0') as 编码," & _
                    " NULL as 附码,名称,简码,NULL as 说明" & _
                    " From 疾病编码分类 Where 类别='S'" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                    " Union ALL " & _
                    " Select 1 as 末级,ID,分类ID as 上级ID,编码,附码,名称,简码,说明" & _
                    " From 疾病编码目录 Where 类别='S'" & _
                    IIf(str性别 <> "", " And (性别限制=[1] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
            End If
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, IIf(optInput(4).Value, "手术项目", "手术编码"), _
                False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, str性别, int性别)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有" & IIf(optInput(4).Value, "手术项目", "手术编码") & "可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        ElseIf Col = col麻醉方式 Then
            StrSQL = "Select 0 as 末级,ID,上级ID,编码,名称,NULL as 麻醉类型" & _
                " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                " Union ALL " & _
                " Select 1 as 末级,ID,分类ID as 上级ID,编码,名称,操作类型 as 麻醉类型" & _
                " From 诊疗项目目录 Where 类别='G'" & _
                " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " And (站点='" & gstrNodeNo & "' Or 站点 is Null)"
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 2, "麻醉项目", , , , , True, , , , , blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有麻醉项目可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        ElseIf Col = col主刀医师 Or Col = col助手1 Or Col = col助手2 Or Col = col麻醉医师 Then
            StrSQL = "Select A.ID,A.编号,A.姓名,A.简码" & _
                " From 人员表 A,人员性质说明 B" & _
                " Where A.ID=B.人员ID And B.人员性质='医生'" & _
                " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编号"
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "医生", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有医生可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        ElseIf Col = col助产护士 Then
            StrSQL = "Select A.ID,A.编号,A.姓名,A.简码" & _
                " From 人员表 A,人员性质说明 B" & _
                " Where A.ID=B.人员ID And B.人员性质='护士'" & _
                " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编号"
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "护士", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有护士可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理过敏药物的输入
    Dim StrSQL As String, curDate As Date
    
    With vsAller
        If Not rsInput Is Nothing Then
            .RowData(lngRow) = CLng(rsInput!ID)
            .TextMatrix(lngRow, AC_过敏药物) = Nvl(rsInput!名称)
        Else
            .RowData(lngRow) = 0
            .TextMatrix(lngRow, AC_过敏药物) = .EditText
        End If
        .Cell(flexcpData, lngRow, AC_过敏药物) = .TextMatrix(lngRow, AC_过敏药物)
        
        If .Cell(flexcpData, lngRow, AC_过敏时间) = "" Then
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, AC_过敏时间) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, AC_过敏时间) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
        
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        .Tag = ""
        mblnChange = True
    End With
End Sub

Private Sub OPSSetInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'功能：根据手术情况输入的情况，设置表格数据
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    
    With vsOPS
        If lngCol = col手术编码 Or lngCol = col手术名称 Then
            If Not rsInput Is Nothing Then
                .TextMatrix(lngRow, col手术名称) = rsInput!名称
                 .Cell(flexcpData, lngRow, col手术名称) = .TextMatrix(lngRow, col手术名称)
                .TextMatrix(lngRow, col手术编码) = rsInput!编码
                If optInput(4).Value Then
                    .TextMatrix(lngRow, col诊疗项目ID) = rsInput!ID
                    .TextMatrix(lngRow, col手术操作ID) = ""
                    StrSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 手术ID=[1]"
                Else
                    .TextMatrix(lngRow, col手术操作ID) = rsInput!ID
                    .TextMatrix(lngRow, col诊疗项目ID) = ""
                    StrSQL = "Select 手术ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(rsInput!ID))
                If Not rsTmp.EOF Then
                    If optInput(4).Value Then
                        .TextMatrix(lngRow, col手术操作ID) = Val(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col诊疗项目ID) = Val(rsTmp!ID)
                    End If
                End If
            Else
                .TextMatrix(lngRow, lngCol) = .EditText
                .TextMatrix(lngRow, col手术操作ID) = ""
                .TextMatrix(lngRow, col诊疗项目ID) = ""
            End If
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            
            '手术日期相同时，其他输入内容默认与上一行相同
            If Not rsInput Is Nothing And lngRow > .FixedRows And lngRow = .Rows - 1 Then
                If .TextMatrix(lngRow, col手术日期) = .TextMatrix(lngRow - 1, col手术日期) Then
                    .TextMatrix(lngRow, col主刀医师) = .TextMatrix(lngRow - 1, col主刀医师)
                    .TextMatrix(lngRow, col助产护士) = .TextMatrix(lngRow - 1, col助产护士)
                    .TextMatrix(lngRow, col助手1) = .TextMatrix(lngRow - 1, col助手1)
                    .TextMatrix(lngRow, col助手2) = .TextMatrix(lngRow - 1, col助手2)
                    .TextMatrix(lngRow, col麻醉方式) = .TextMatrix(lngRow - 1, col麻醉方式)
                    .TextMatrix(lngRow, col麻醉医师) = .TextMatrix(lngRow - 1, col麻醉医师)
                    .TextMatrix(lngRow, col切口愈合) = .TextMatrix(lngRow - 1, col切口愈合)
                    .TextMatrix(lngRow, col麻醉ID) = .TextMatrix(lngRow - 1, col麻醉ID)
                    .TextMatrix(lngRow, col麻醉类型) = .TextMatrix(lngRow - 1, col麻醉类型)
                    
                    For i = col主刀医师 To .Cols - 1
                        .Cell(flexcpData, lngRow, i) = .TextMatrix(lngRow, i)
                    Next
                End If
            End If
            
            '输入后始终保持一新行
            If lngRow = .Rows - 1 Then .AddItem ""
        ElseIf lngCol = col麻醉方式 Then
            .TextMatrix(lngRow, lngCol) = rsInput!名称
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            .TextMatrix(lngRow, col麻醉ID) = rsInput!ID
            .TextMatrix(lngRow, col麻醉类型) = Nvl(rsInput!麻醉类型)
        ElseIf lngCol = col主刀医师 Or lngCol = col助产护士 Or lngCol = col助手1 Or lngCol = col助手2 Or lngCol = col麻醉医师 Then
            .TextMatrix(lngRow, lngCol) = rsInput!姓名
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
        End If
        
        '设置诊断符合情况
        Call Set诊断符合情况(cbo术前与术后)
        
        .Tag = ""
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsOPS_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsOPS
        If Col = col切口愈合 Then
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsOPS_DblClick()
    Call vsOPS_KeyPress(32)
End Sub

Private Sub vsOPS_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsOPS
        If KeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col手术名称) <> "" Then
                If MsgBox("确实要删除该行手术吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    
                    '设置诊断符合情况
                    Call Set诊断符合情况(cbo术前与术后)

                    mblnChange = True
                    .Tag = ""
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsOPS_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsOPS_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsOPS
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call OPSEnterNextCell
        Else
            If .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsOPS_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsOPS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strInput As String
    
    With vsOPS
        If KeyAscii = 13 Then
            mblnReturn = True
            
            If Col = col手术日期 Then
                KeyAscii = 0
                strInput = GetFullDate(.EditText, False)
                If IsDate(strInput) Then
                    .TextMatrix(Row, Col) = strInput
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    mblnChange = True
                    .Tag = ""
                    Call OPSEnterNextCell
                End If
            ElseIf Col = col切口愈合 Then
                KeyAscii = 0
                If .ComboIndex <> -1 Then
                    .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                    mblnChange = True
                    .Tag = ""
                    Call OPSEnterNextCell
                End If
            End If
        Else
            mblnReturn = False
            
            If Col = col手术日期 Then
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            ElseIf Col = col抗菌药天数 Then
                If InStr("0123456789" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsOPS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsOPS.EditSelStart = 0
    vsOPS.EditSelLength = zlCommFun.ActualLen(vsOPS.EditText)
End Sub

Private Function OPSCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsOPS
        If .ColHidden(lngCol) Then Exit Function
        
        '必须先输入手术日期,手术名称
        If Not IsDate(.TextMatrix(lngRow, col手术日期)) Then
            If lngCol > col手术日期 Then Exit Function
        End If
        If .TextMatrix(lngRow, col手术名称) = "" Then
            If lngCol > col手术名称 Then Exit Function
        End If
        
        '必须先输入主刀医师
        If .TextMatrix(lngRow, col主刀医师) = "" Then
            If lngCol = col助手1 Or lngCol = col助手2 Then Exit Function
        End If
        
        '必须先输入第1助手
        If .TextMatrix(lngRow, col助手1) = "" Then
            If lngCol = col助手2 Then Exit Function
        End If
        
        '必须先输入麻醉方式
        If Trim(.TextMatrix(lngRow, col麻醉类型)) = "" Then
            If lngCol = col麻醉医师 Then Exit Function
        End If
        
        '手术名称不能输入
        If lngCol = col手术名称 And chkInfo(chk手术自由录入).Value = 0 Then Exit Function
    End With
    OPSCellEditable = True
End Function

Private Sub OPSEnterNextCell()
    Dim i As Long, j As Long
    
    With vsOPS
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col手术日期) To col术后髋关节骨折
                If OPSCellEditable(i, j) Then Exit For
            Next
            If j <= col术后髋关节骨折 Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAller
        If .Col = AC_过敏反应 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AC_过敏药物
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long

    If UCase(objTmp.Container.Name) <> UCase("fraInfo") Then
        If UCase(objTmp.Container.Container.Name) = UCase("fraInfo") Then sstInfo.Tab = objTmp.Container.Container.Index
    Else
        sstInfo.Tab = objTmp.Container.Index
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then objTmp.SetFocus
    Me.Refresh
End Function

Private Function CheckPageData(ByRef blnDiagnose As Boolean, ByVal blnBeforSign As Boolean) As Boolean
'功能：检查首页输入数据合法性
'返回：blnDiagnose=是否填写了诊断
'参数：blnBeforSign-是否签名时保存前调用
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str身份证 As String, str出生日期 As String, lng性别 As Long
    Dim lng手术次数 As Long, str年龄 As String
    Dim str疾病IDs As String, str诊断IDs As String
    Dim i As Long, j As Long
    Dim StrSQL As String, rsTmp As Recordset
    
    
    blnDiagnose = False
    
    '项目输入的长度检查
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtInfo
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "输入内容过长，请检查。(该项目最多允许 " & objTmp.MaxLength & " 个字符或 " & objTmp.MaxLength \ 2 & " 个汉字)")
                Exit Function
            End If
        End If
    Next
    If Not mbln护士站 Then
        curDate = zlDatabase.Currentdate
        
        '必须要输入的内容检查
        '-----------------------------------------------------------------------------------------
        arrInfo = Array(txt住院号, txt姓名, txt年龄, txt区域)
        arrName = Array("住院号", "姓名", "年龄", "区域")
        For i = 0 To UBound(arrInfo)
            If txtInfo(arrInfo(i)).Enabled And Not txtInfo(arrInfo(i)).Locked And txtInfo(arrInfo(i)).Text = "" Then
                If arrName(i) <> "区域" Or mlng区域 = 1 Then
                    Call ShowMessage(txtInfo(arrInfo(i)), "必须输入病人的" & arrName(i) & "。")
                    Exit Function
                ElseIf mlng区域 = 2 Then
                    If ShowMessage(txtInfo(arrInfo(i)), "没有输入病人的" & arrName(i) & ",是否继续？", True) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        Next
        
        Select Case cboinfo(cbo年龄单位).Text
            Case "岁"
                If Val(txtInfo(txt年龄).Text) > 200 Then
                    Call ShowMessage(txtInfo(txt年龄), "年龄值超过了最大限制200岁，请检查输入是否正确。")
                    Exit Function
                End If
            Case "月"
                If Val(txtInfo(txt年龄).Text) > 2400 Then
                    Call ShowMessage(txtInfo(txt年龄), "年龄值超过了最大限制2400月，请检查输入是否正确。")
                    Exit Function
                End If
            Case "天"
                If Val(txtInfo(txt年龄).Text) > 73000 Then
                    Call ShowMessage(txtInfo(txt年龄), "年龄值超过最大限制73000天，请检查输入是否正确。")
                    Exit Function
                End If
            Case "小时" '不能大于30天即720小时
                If Val(txtInfo(txt年龄).Text) > 720 Then
                    Call ShowMessage(txtInfo(txt年龄), "年龄值超过了最大限制720小时，请使用合适的年龄单位。")
                    Exit Function
                End If
            Case "分钟" '不能大于24小时即1440分钟
                If Val(txtInfo(txt年龄).Text) > 1440 Then
                    Call ShowMessage(txtInfo(txt年龄), "年龄值超过了最大限制1440分钟，请使用合适的年龄单位。")
                    Exit Function
                End If
        End Select
        
        Select Case cboinfo(cbo婴儿年龄单位).Text
            Case "月" '一年
                If Val(txtInfo(txt婴儿年龄).Text) > 12 Then
                    Call ShowMessage(txtInfo(txt婴儿年龄), "婴儿年龄值超过了最大限制12月，请检查输入是否正确。")
                    Exit Function
                End If
            Case "天" '365天
                If Val(txtInfo(txt婴儿年龄).Text) > 365 Then
                    Call ShowMessage(txtInfo(txt婴儿年龄), "婴儿年龄值超过了最大限制365天，请使用合适的年龄单位。")
                    Exit Function
                End If
            Case "小时" '不能大于30天即720小时
                If Val(txtInfo(txt婴儿年龄).Text) > 720 Then
                    Call ShowMessage(txtInfo(txt婴儿年龄), "婴儿年龄值超过了最大限制720小时，请使用合适的年龄单位。")
                    Exit Function
                End If
            Case "分钟" '不能大于24小时即1440分钟
                If Val(txtInfo(txt婴儿年龄).Text) > 1440 Then
                    Call ShowMessage(txtInfo(txt婴儿年龄), "婴儿年龄值超过了最大限制1440分钟，请使用合适的年龄单位。")
                    Exit Function
                End If
        End Select
        If Not IsDate(txt出生日期.Text) Then
            Call ShowMessage(txt出生日期, "必须输入病人的出生日期。")
            Exit Function
        ElseIf txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
            Call ShowMessage(txt出生时间, "请输入正确的病人出生时间。")
            Exit Function
        End If
        
        arrInfo = Array(cbo付款方式, cbo性别, cbo民族, cbo职业, cbo入院病情)
        arrName = Array("付款方式", "性别", "民族", "职业", "入院病情")
        For i = 0 To UBound(arrInfo)
            If cboinfo(arrInfo(i)).Enabled And Not cboinfo(arrInfo(i)).Locked And cboinfo(arrInfo(i)).ListIndex = -1 Then
                Call ShowMessage(cboinfo(arrInfo(i)), "必须输入病人的" & arrName(i) & "。")
                Exit Function
            End If
        Next
        
        If txtInfo(txt转科3).Text <> "" And txtInfo(txt转科2).Text = "" Then
            Call ShowMessage(txtInfo(txt转科2), "请依次输入转科科室。")
            Exit Function
        End If
        If txtInfo(txt转科2).Text <> "" And txtInfo(txt转科1).Text = "" Then
            Call ShowMessage(txtInfo(txt转科1), "请依次输入转科科室。")
            Exit Function
        End If
        If txtInfo(txt转科1).Text = txtInfo(txt转科2).Text And txtInfo(txt转科1).Text <> "" Then
            Call ShowMessage(txtInfo(txt转科2), "转科的两个科室不应该相同。")
            Exit Function
        End If
        If txtInfo(txt转科2).Text = txtInfo(txt转科3).Text And txtInfo(txt转科2).Text <> "" Then
            Call ShowMessage(txtInfo(txt转科3), "转科的两个科室不应该相同。")
            Exit Function
        End If
        
        If cboinfo(cbo科主任).ListIndex = -1 And cboinfo(cbo主任医师).ListIndex = -1 _
            And cboinfo(cbo主治医师).ListIndex = -1 And cboinfo(cbo住院医师).ListIndex = -1 Then
            Call ShowMessage(cboinfo(cbo科主任), "请在科主任、主任医师、主治医师和住院医师之间至少选择一位。")
            Exit Function
        End If
            
    
        '年龄长度要带上单位
        If zlCommFun.ActualLen(txtInfo(txt年龄).Text & cboinfo(cbo年龄单位).Text) > txtInfo(txt年龄).MaxLength Then
            Call ShowMessage(txtInfo(txt年龄), "输入内容过长，请检查。(该项目最多允许 " & txtInfo(txt年龄).MaxLength & " 个字符或 " & txtInfo(txt年龄).MaxLength \ 2 & " 个汉字)")
        End If
        
        If txt死亡时间.Text <> "____-__-__ __:__:__" Then
            If Not IsDate(txt死亡时间.Text) Then
                Call ShowMessage(txt死亡时间, "死亡时间不是有效的日期格式。")
                Exit Function
            End If
            If Format(txt死亡时间.Text, "yyyy-MM-dd HH:mm:ss") <= Format(txtInfo(txt入院时间).Text, "yyyy-MM-dd HH:mm:ss") Then
                Call ShowMessage(txt死亡时间, "死亡时间应比入院时间晚。")
                Exit Function
            End If
        End If
        
        '输入内容的有效性检查
        '-----------------------------------------------------------------------------------------
        '出生日期必须早于当前时间
        If Format(txt出生日期.Text, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
            Call ShowMessage(txt出生日期, "出生日期不应该比当前日期还晚。")
            Exit Function
        End If
        
        '出生日期必须早于入院时间
        If Trim(txtInfo(txt入院时间).Text) <> "" Then
            If Format(txt出生日期.Text, "yyyy-MM-dd") > Format(txtInfo(txt入院时间).Text, "yyyy-MM-dd") Then
                Call ShowMessage(txt出生日期, "出生日期不应该比入院时间还晚。")
                Exit Function
            End If
        End If
        
        '年龄与出生日期的匹配性
        If IsNumeric(txtInfo(txt年龄).Text) And cboinfo(cbo年龄单位).ListIndex <> -1 And cboinfo(cbo年龄单位).ListIndex < 3 Then
            str年龄 = PatiAgeCalc(txt出生日期.Text, , txtInfo(txt入院时间).Text)
            If Right(str年龄, 1) = cboinfo(cbo年龄单位).Text And IsNumeric(Left(str年龄, Len(str年龄) - 1)) _
                And str年龄 <> txtInfo(txt年龄).Text & cboinfo(cbo年龄单位).Text Then
                If ShowMessage(txt出生日期, "年龄和出生日期不一致，" & txt出生日期.Text & "出生现在应该是" & str年龄 & "。" & _
                    vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", True) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        
        '15岁以下应为未婚
        If DateDiff("yyyy", CDate(txt出生日期.Text), curDate) < 15 Then
            If InStr(cboinfo(cbo婚姻).Text, "未婚") = 0 Then
                Call ShowMessage(cboinfo(cbo婚姻), "该病人年龄小于15岁，婚姻状况应该写为未婚。")
                Exit Function
            End If
        End If
                
        '身份证号码检查
        '对身份证号进行验证
        str身份证 = cboinfo(cbo身份证号).Text
        If str身份证 <> "" Then
            If zlCommFun.ActualLen(str身份证) = Len(str身份证) Then
                If Len(str身份证) <> 15 And Len(str身份证) <> 18 Then
                    Call ShowMessage(cboinfo(cbo身份证号), "身份证号码的长度不正确，应为15位或18位。")
                    Exit Function
                End If
                
                If Len(str身份证) = 15 Then
                    str出生日期 = Mid(str身份证, 7, 6)
                    str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
                    lng性别 = Val(Right(str身份证, 1))
                Else
                    str出生日期 = Mid(str身份证, 7, 8)
                    str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
                    lng性别 = Val(Mid(str身份证, 17, 1))
                End If
                If Not IsDate(str出生日期) Then
                    If ShowMessage(cboinfo(cbo身份证号), "身份证号码中的出生日期信息不正确，是否继续？", True) = vbNo Then Exit Function
                Else
                    If Format(str出生日期, "yyyy-MM-dd") <> Format(txt出生日期.Text, "yyyy-MM-dd") Then
                        If ShowMessage(cboinfo(cbo身份证号), "身份证号码中的出生日期信息与病人的出生日期不符，是否继续？", True) = vbNo Then Exit Function
                    End If
                End If
                If (lng性别 Mod 2 = 1 And InStr(cboinfo(cbo性别).Text, "女") > 0) Or (lng性别 Mod 2 = 0 And InStr(cboinfo(cbo性别).Text, "男") > 0) Then
                    If ShowMessage(cboinfo(cbo身份证号), "身份证号码中的性别信息与病人的性别不符，是否继续？", True) = vbNo Then Exit Function
                End If
            Else
                If zlCommFun.ActualLen(str身份证) > 18 Then
                    Call ShowMessage(cboinfo(cbo身份证号), "不能超过9个汉字的长度，请检查。")
                    Exit Function
                End If
            End If
        End If
        
        '确诊时间必须在入院时间和出院时间之间
        If IsDate(txtInfo(txt确诊日期).Text) Then
            If Not Between(Format(txtInfo(txt确诊日期).Text, "yyyy-MM-dd"), Format(txtInfo(txt入院时间).Tag, "yyyy-MM-dd"), _
                Format(IIf(txtInfo(txt出院时间).Text = "", zlDatabase.Currentdate, txtInfo(txt出院时间).Text), "yyyy-MM-dd")) Then
                Call ShowMessage(txtInfo(txt确诊日期), "确诊时间必须在入院时间和出院时间之间。")
                Exit Function
            End If
        ElseIf chkInfo(chk是否确诊).Value = 1 Then
            Call ShowMessage(txtInfo(txt确诊日期), "确诊时间输入错误。")
            Exit Function
        End If
        
        '入院病情为危时需要进行抢救
        If InStr(cboinfo(cbo入院病情).Text, "危") > 0 And Val(txtInfo(txt抢救次数).Text) = 0 Then
            If ShowMessage(txtInfo(txt抢救次数), "该病人入院病情为危，但没有进行抢救，是否继续？", True) = vbNo Then Exit Function
        End If
        
        '成功次数不能超过抢救次数
        If Val(txtInfo(txt成功次数).Text) > Val(txtInfo(txt抢救次数).Text) Then
            Call ShowMessage(txtInfo(txt成功次数), "成功次数不能超过抢救次数。")
            Exit Function
        End If
        '成功次数小于抢救次数时出院情况应为死亡 2010-03-23 27224 死的时候，成功次数可以等于抢救次数，因为有   病人没有抢救就死了。
        If InStr(vsDiagXY.TextMatrix(GetRow(3), col出院情况), "死亡") > 0 Then
            If Val(txtInfo(txt成功次数).Text) > Val(txtInfo(txt抢救次数).Text) And txtInfo(txt抢救次数).Text <> "" Then
                Call ShowMessage(txtInfo(txt成功次数), "病人出院情况为死亡，成功次数不能大于抢救次数。")
                Exit Function
            End If
        Else
            If Val(txtInfo(txt成功次数).Text) <> Val(txtInfo(txt抢救次数).Text) And txtInfo(txt抢救次数).Text <> "" Then
                If InStr(vsDiagXY.TextMatrix(GetRow(3), col出院情况), "其他") > 0 Then
                    If ShowMessage(txtInfo(txt成功次数), "病人出院情况不为死亡，成功次数应等于抢救次数，是否继续？", True) = vbNo Then Exit Function
                Else
                    Call ShowMessage(txtInfo(txt成功次数), "病人出院情况不为死亡，成功次数应等于抢救次数。")
                    Exit Function
                End If
            End If
        End If
        '成功次数最多比抢救次数少一次
        If Val(txtInfo(txt抢救次数).Text) - Val(txtInfo(txt成功次数).Text) > 1 And txtInfo(txt抢救次数).Text <> "" Then
            Call ShowMessage(txtInfo(txt成功次数), "成功次数最多比抢救次数少一次。")
            Exit Function
        End If
        
        '随诊检查
        If chkInfo(chk随诊).Value = 1 Then
            If Val(txtInfo(txt随诊期限).Text) <= 0 And cboinfo(cbo随诊Ex).Text <> "终身" Then
                Call ShowMessage(txtInfo(txt随诊期限), "请输入正确的随诊期限。")
                Exit Function
            End If
        End If
        
        '31天内再住院计划的目的
        If optInput(opt31天有).Value Then
            If Trim(txtInfo(txt31天目的).Text) = "" Then
                Call ShowMessage(txtInfo(txt31天目的), "请填写" & cboinfo(cbo31天和7天再入院).Text & "的目的。")
                Exit Function
            End If
        End If
        '填写了术前与术后，必须填写手术情况
        If vsOPS.TextMatrix(1, col手术名称) = "" And cboinfo(cbo术前与术后).ListIndex > 0 Then
            Call ShowMessage(cboinfo(cbo术前与术后), "没有填写手术情况,术前与术后只能选择""未做""。")
            Exit Function
        End If
        
        '发病时间检查
        If txt发病日期.Text <> "____-__-__" Then
            If Not IsDate(txt发病日期.Text) Then
                Call ShowMessage(txt发病日期, "请输入正确的发病日期。")
                Exit Function
            Else
                If txt发病时间.Text <> "__:__" Then
                    If Not IsDate(txt发病时间.Text) Then
                        Call ShowMessage(txt发病时间, "请输入正确的发病时间。")
                        Exit Function
                    End If
                End If
                
                If txt发病日期.Text & IIf(txt发病时间.Text = "__:__", "", " " & txt发病时间.Text) _
                    >= Format(curDate, txt发病日期.Format & IIf(txt发病时间.Text = "__:__", "", " " & txt发病时间.Format)) Then
                    Call ShowMessage(txt发病日期, "发病时间应该早于当前时间。")
                    Exit Function
                End If
            End If
        End If
        
        '表格的检查
        '-----------------------------------------------------------------------------------------
        With vsOPS
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col手术名称)) <> "" Then
                    lng手术次数 = lng手术次数 + 1
                End If
            Next
        End With
        
        With vsDiagXY
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, col诊断描述) <> "" And .TextMatrix(i - 1, col诊断描述) = "" _
                    And Val(.TextMatrix(i, col类型)) = Val(.TextMatrix(i - 1, col类型)) Then
                    .Row = i - 1: .Col = col诊断描述
                    Call ShowMessage(vsDiagXY, "请依次输入诊断信息。")
                    Exit Function
                End If
                
                If Trim(.TextMatrix(i, col诊断描述)) <> "" Then
                    If zlCommFun.ActualLen(.TextMatrix(i, col诊断描述)) > 200 Then
                        .Row = i: .Col = col诊断描述
                        Call ShowMessage(vsDiagXY, IIf(.TextMatrix(i, col诊断类型) = "", "出院诊断", .TextMatrix(i, col诊断类型)) & "内容太长，只允许200个字符或100个汉字。")
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, col备注)) > 50 Then
                        .Row = i: .Col = col备注
                        Call ShowMessage(vsDiagXY, """" & .TextMatrix(i, col诊断描述) & """的备注内容太长，只允许50个字符或25个汉字。")
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, col类型)) = 5 Then    '院内感染
                        If .TextMatrix(i, col出院情况) = "" Then
                            .Row = i: .Col = col出院情况
                            If ShowMessage(vsDiagXY, "院内感染的出院情况没有填写，是否继续？", True) = vbNo Then Exit Function
                        End If
                    End If
                    If Val(.TextMatrix(i, col类型)) = 3 Then
                        If .TextMatrix(i, col出院情况) = "" Then
                            .Row = i: .Col = col出院情况
                            Call ShowMessage(vsDiagXY, "请填写出院诊断的出院情况。")
                            Exit Function
                        ElseIf Val(.TextMatrix(i - 1, col类型)) <> 3 And InStr(.TextMatrix(i, col出院情况), "其他") > 0 And lng手术次数 > 0 Then
                            .Row = i: .Col = col出院情况
                            If ShowMessage(vsDiagXY, "该病人进行了手术，但出院情况选择为其他。是否继续？", True) = vbNo Then Exit Function
    '                    ElseIf Val(.TextMatrix(i - 1, col类型)) = 3 And InStr(.TextMatrix(GetRow(3), col出院情况), "其他") > 0 And InStr(.TextMatrix(i, col出院情况), "其他") = 0 Then
    '                        .Row = i: .Col = col出院情况
    '                        Call ShowMessage(vsDiagXY, "主要诊断的出院情况为其他，但其他诊断的出院情况却出现""" & .TextMatrix(i, col出院情况) & """。")
    '                        Exit Function
                        ElseIf Val(.TextMatrix(i - 1, col类型)) = 3 And InStr(.TextMatrix(GetRow(3), col出院情况), "死亡") = 0 And InStr(.TextMatrix(i, col出院情况), "死亡") > 0 Then
                            .Row = i: .Col = col出院情况
                            Call ShowMessage(vsDiagXY, "主要诊断的出院情况不为死亡，但其它诊断的出院情况却为死亡。")
                            Exit Function
                        ElseIf Val(.TextMatrix(i - 1, col类型)) <> 3 And InStr(.TextMatrix(i, col出院情况), "治愈") > 0 And Val(txtInfo(txt住院天数).Text) < 3 Then
                            .Row = i: .Col = col出院情况
                            If ShowMessage(vsDiagXY, "该病人住院天院为 " & Val(txtInfo(txt住院天数).Text) & " 天，出院情况却为治愈，是否继续？", True) = vbNo Then Exit Function
                        ElseIf .TextMatrix(i, col诊断类型) = "出院诊断" Then
                            If mlng损伤中毒 <> 0 Then
                                '主要诊断需要有损伤的外部原因
                                If InStr("ST", Left(.TextMatrix(i, col诊断编码), 1)) > 0 And Left(.TextMatrix(i, col诊断编码), 1) <> "" Then
                                    '需要损伤中毒外部原因
                                    If .TextMatrix(GetRow(7), col诊断描述) = "" Then
                                        If Not sstInfo.TabVisible(TAB_中医诊断) Then
                                            .Row = GetRow(7): .Col = col诊断描述
                                            If mlng损伤中毒 = 1 Then
                                                Call ShowMessage(vsDiagXY, "请填写损伤中毒的原因。")
                                                Exit Function
                                            Else
                                                If ShowMessage(vsDiagXY, "没有填写损伤中毒的原因,是否继续？", True) = vbNo Then Exit Function
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(GetRow(7), col诊断描述) <> "" Then
                                        .Row = GetRow(7): .Col = col诊断描述
                                        If mlng损伤中毒 = 1 Then
                                            Call ShowMessage(vsDiagXY, "不能填写损伤中毒的原因。")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "出院诊断与损伤中毒的原因不符,是否继续？", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                            If mlng病理诊断 <> 0 Then
                                '主要诊断需要填写病理诊断的外部原因
                                If InStr("CD", Left(.TextMatrix(i, col诊断编码), 1)) > 0 And Left(.TextMatrix(i, col诊断编码), 1) <> "" Then
                                    '需要病理诊断的外部原因
                                    If .TextMatrix(GetRow(6), col诊断描述) = "" Then
                                        If Not sstInfo.TabVisible(TAB_中医诊断) Then
                                            .Row = GetRow(6): .Col = col诊断描述
                                            If mlng病理诊断 = 1 Then
                                                Call ShowMessage(vsDiagXY, "请填写病理诊断。")
                                                Exit Function
                                            Else
                                                If ShowMessage(vsDiagXY, "没有填写病理诊断,是否继续？", True) = vbNo Then Exit Function
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(GetRow(6), col诊断描述) <> "" Then
                                        .Row = GetRow(6): .Col = col诊断描述
                                        If mlng病理诊断 = 1 Then
                                            Call ShowMessage(vsDiagXY, "不能填写病理诊断。")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "出院诊断与病理诊断不符,是否继续？", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        For j = GetRow(3) To .Rows - 1
                            If Val(.TextMatrix(j, col类型)) = 3 Then
                                If j <> i And .TextMatrix(j, col诊断描述) <> "" Then
                                    If .TextMatrix(j, col诊断描述) = .TextMatrix(i, col诊断描述) Then
                                        .Row = i: .Col = col诊断描述
                                        Call ShowMessage(vsDiagXY, "发现存在两行相同的出院诊断信息。")
                                        Exit Function
                                    ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                                        If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                            .Row = i: .Col = col诊断描述
                                            Call ShowMessage(vsDiagXY, "发现存在两行相同的出院诊断信息。")
                                            Exit Function
                                        End If
                                    ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 Then
                                        If Val(.TextMatrix(j, col诊断ID)) = Val(.TextMatrix(i, col诊断ID)) Then
                                            .Row = i: .Col = col诊断描述
                                            Call ShowMessage(vsDiagXY, "发现存在两行相同的出院诊断信息。")
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                    If Val(.TextMatrix(i, col疾病ID)) <> 0 Then str疾病IDs = str疾病IDs & "," & Val(.TextMatrix(i, col疾病ID))
                    If Val(.TextMatrix(i, col诊断ID)) <> 0 Then str诊断IDs = str诊断IDs & "," & Val(.TextMatrix(i, col诊断ID))
                    
                    '是否输入了要求的诊断类型
                    If InStr("," & mstr类型 & ",", "," & Val(.TextMatrix(i, col类型)) & ",") > 0 Then
                        blnDiagnose = True
                    End If
                End If
            Next
        End With
            
        If mbln中医 Then
            With vsDiagZY
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, col诊断描述) <> "" And .TextMatrix(i - 1, col诊断描述) = "" _
                        And Val(.TextMatrix(i, colzy类型)) = Val(.TextMatrix(i - 1, colzy类型)) Then
                        .Row = i - 1: .Col = col诊断描述
                        Call ShowMessage(vsDiagZY, "请依次输入诊断信息。")
                        Exit Function
                    End If
                
                    If Trim(.TextMatrix(i, col诊断描述)) <> "" Then
                        If zlCommFun.ActualLen(.TextMatrix(i, col诊断描述)) > 200 Then
                            .Row = i: .Col = col诊断描述
                            Call ShowMessage(vsDiagZY, IIf(.TextMatrix(i, col诊断类型) = "", "出院诊断", .TextMatrix(i, col诊断类型)) & "内容太长，只允许200个字符或100个汉字。")
                            Exit Function
                        End If
                        If zlCommFun.ActualLen(.TextMatrix(i, col备注)) > 50 Then
                            .Row = i: .Col = col备注
                            Call ShowMessage(vsDiagZY, """" & .TextMatrix(i, col诊断描述) & """的备注内容太长，只允许50个字符或25个汉字。")
                            Exit Function
                        End If
                        If Val(.TextMatrix(i, colzy类型)) = 13 Then
                            If .TextMatrix(i, col出院情况) = "" Then
                                .Row = i: .Col = col出院情况
                                Call ShowMessage(vsDiagZY, "请填写出院诊断的出院情况。")
                                Exit Function
    '                        ElseIf Val(.TextMatrix(i - 1, colzy类型)) = 13 And InStr(.TextMatrix(GetRow(13), col出院情况), "其他") > 0 And InStr(.TextMatrix(i, col出院情况), "其他") = 0 Then
    '                            .Row = i: .Col = col出院情况
    '                            Call ShowMessage(vsDiagZY, "主要诊断的出院情况为其他，但其他诊断的出院情况却出现""" & .TextMatrix(i, col出院情况) & """。")
    '                            Exit Function
                            ElseIf Val(.TextMatrix(i - 1, colzy类型)) = 13 And InStr(.TextMatrix(GetRow(13), col出院情况), "死亡") = 0 And InStr(.TextMatrix(i, col出院情况), "死亡") > 0 Then
                                .Row = i: .Col = col出院情况
                                Call ShowMessage(vsDiagZY, "主要诊断的出院情况不为死亡，但其它诊断的出院情况却为死亡。")
                                Exit Function
                            End If
                            
                            For j = GetRow(13) To .Rows - 1
                                If j <> i And .TextMatrix(j, col诊断描述) <> "" Then
                                    If .TextMatrix(j, col诊断描述) = .TextMatrix(i, col诊断描述) Then
                                        .Row = i: .Col = col诊断描述
                                        Call ShowMessage(vsDiagZY, "发现存在两行相同的出院诊断信息。")
                                        Exit Function
                                    ElseIf Val(.TextMatrix(i, colzy疾病ID)) <> 0 Then
                                        If Val(.TextMatrix(j, colzy疾病ID)) = Val(.TextMatrix(i, colzy疾病ID)) Then
                                            .Row = i: .Col = col诊断描述
                                            Call ShowMessage(vsDiagZY, "发现存在两行相同的出院诊断信息。")
                                            Exit Function
                                        End If
                                    ElseIf Val(.TextMatrix(i, colzy诊断ID)) <> 0 Then
                                        '因中医诊断带证候,可能无对应证候ID,诊断ID又相同
    '                                    If Val(.TextMatrix(j, colzy诊断ID)) = Val(.TextMatrix(i, colzy诊断ID)) Then
    '                                        .Row = i: .Col = col诊断描述
    '                                        Call ShowMessage(vsDiagZY, "发现存在两行相同的出院诊断信息。")
    '                                        Exit Function
    '                                    End If
                                    End If
                                End If
                            Next
                        End If
                        
                        If Val(.TextMatrix(i, colzy疾病ID)) <> 0 Then str疾病IDs = str疾病IDs & "," & Val(.TextMatrix(i, colzy疾病ID))
                        If Val(.TextMatrix(i, colzy诊断ID)) <> 0 Then str诊断IDs = str诊断IDs & "," & Val(.TextMatrix(i, colzy诊断ID))
                        
                        '是否输入了要求的诊断类型
                        If InStr("," & mstr类型 & ",", "," & Val(.TextMatrix(i, colzy类型)) & ",") > 0 Then
                            blnDiagnose = True
                        End If
                    End If
                Next
            End With
        End If
        
        With vsOPS
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col手术名称)) <> "" Then
                    If Not IsDate(.TextMatrix(i, col手术日期)) Then
                        .Row = i: .Col = col手术日期
                        Call ShowMessage(vsOPS, "手术日期输入不正确。")
                        Exit Function
                    ElseIf txtInfo(txt出院时间).Text <> "" And Format(.TextMatrix(i, col手术日期), "yyyy-MM-dd") > Format(txtInfo(txt出院时间).Text, "yyyy-MM-dd") Or _
                        Format(.TextMatrix(i, col手术日期), "yyyy-MM-dd") < Format(txtInfo(txt入院时间).Text, "yyyy-MM-dd") Then
                        .Row = i: .Col = col手术日期    '手术日期没有精确到时间
                        Call ShowMessage(vsOPS, "手术日期不在入出院日期范围内。")
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, col手术名称)) > 100 Then
                        .Row = i: .Col = col手术名称
                        Call ShowMessage(vsOPS, "手术名称内容太长，只允许100个字符或50个汉字。")
                        Exit Function
                    End If
                    If .ColHidden(col助产护士) Then
                        If .TextMatrix(i, col主刀医师) = "" Then
                            .Row = i: .Col = col主刀医师
                            Call ShowMessage(vsOPS, "请输入主刀医师。")
                            Exit Function
                        End If
                    Else
                        If .TextMatrix(i, col主刀医师) = "" And .TextMatrix(i, col助产护士) = "" Then
                            .Row = i: .Col = col主刀医师
                            Call ShowMessage(vsOPS, "请输入主刀医师或助产护士。")
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
            
        With vsKSS
            For i = .FixedRows To .Rows - 1
                If i > .FixedRows Then
                    If Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i - 1, 1)) = "" Then
                        .Row = i - 1: .Col = 1
                        Call ShowMessage(vsKSS, "请依次输入抗菌药物内容。")
                        Exit Function
                    End If
                End If
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    For j = .FixedRows To .Rows - 1
                        If j <> i And Trim(.TextMatrix(j, 1)) = Trim(.TextMatrix(i, 1)) Then
                            .Row = j: .Col = 1
                            Call ShowMessage(vsKSS, "发现存在两行相同的抗菌药物信息。")
                            Exit Function
                        End If
                    Next
                End If
            Next
        End With
        
        With vsTSJC
            For i = .FixedRows To .Rows - 1
                If i > .FixedRows Then
                    If Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i - 1, 1)) = "" Then
                        .Row = i - 1: .Col = 1
                        Call ShowMessage(vsTSJC, "请依次输入特殊检查内容。")
                        Exit Function
                    End If
                End If
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    For j = .FixedRows To .Rows - 1
                        If j <> i And Trim(.TextMatrix(j, 1)) = Trim(.TextMatrix(i, 1)) Then
                            .Row = j: .Col = 1
                            Call ShowMessage(vsTSJC, "发现存在两行相同的特殊检查信息。")
                            Exit Function
                        End If
                    Next
                End If
            Next
        End With
            
        '抗菌药物
        With vsKSS
        
            For i = .FixedRows To .Rows - 1
                If (Len(.TextMatrix(i, kss使用天数)) > 18 Or Val(.TextMatrix(i, kss使用天数)) = 0) And Len(.TextMatrix(i, kss使用天数)) > 0 Then
                    .Row = i: .Col = kss使用天数
                    Call ShowMessage(vsKSS, "请填写十八位数以内的数字天数。")
                    Exit Function
                End If
                If LenB(StrConv(.TextMatrix(i, kss用药目的), vbFromUnicode)) > 200 And LenB(StrConv(.TextMatrix(i, kss用药目的), vbFromUnicode)) > 0 Then
                    .Row = i: .Col = kss用药目的
                    Call ShowMessage(vsKSS, "请填写100个汉字以内的用药目的。")
                    Exit Function
                End If
            Next
        
        End With
        
         
        mstr疾病ID = Mid(str疾病IDs, 2)
        mstr诊断ID = Mid(str诊断IDs, 2)
    End If
    '并发检查病案是否编目或首页处于锁定状态
    If Not CheckMecRed(mlng病人ID, mlng主页ID, Me.Caption, "修改首页") Then Exit Function
         
    CheckPageData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadOldData(strOld As String, Optional ByVal lngIndex As Long)
'功能:将数据库中保存的年龄按估计的格式加载到界面
'参数:lngIndex-传入的年龄索引，分为婴儿的和病人本身的
    Dim strTmp As Long
    Dim lng单位 As Long

    If lngIndex = 0 Then lngIndex = txt年龄: lng单位 = cbo年龄单位
    If lngIndex = txt婴儿年龄 Then lng单位 = cbo婴儿年龄单位
    If InStr(strOld, "岁") > 0 And lngIndex = txt年龄 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng单位).ListCount > 0 Then cboinfo(lng单位).ListIndex = 0
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng单位).ListIndex = -1
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng单位).ListCount > 1 Then
                cboinfo(lng单位).ListIndex = IIf(lngIndex = txt年龄, 1, 0)
            End If
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng单位).ListIndex = -1
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng单位).ListCount > 1 Then cboinfo(lng单位).ListIndex = IIf(lngIndex = txt年龄, 2, 1)
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng单位).ListIndex = -1
        End If
    ElseIf InStr(strOld, "小时") > 0 Then
        If InStr(strOld, "小时") + 1 = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "小时") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng单位).ListCount > 1 Then cboinfo(lng单位).ListIndex = IIf(lngIndex = txt年龄, 3, 2)
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng单位).ListIndex = -1
        End If
    ElseIf InStr(strOld, "分钟") > 0 Then
        If InStr(strOld, "分钟") + 1 = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "分钟") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng单位).ListCount > 1 Then cboinfo(lng单位).ListIndex = IIf(lngIndex = txt年龄, 4, 3)
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng单位).ListIndex = -1
        End If
    ElseIf IsNumeric(strOld) Then
        txtInfo(lngIndex).Text = strOld
        If cboinfo(lng单位).ListCount > 0 Then cboinfo(lng单位).ListIndex = 0
    Else
        txtInfo(lngIndex).Text = strOld
        cboinfo(lng单位).ListIndex = -1
    End If
End Sub

Private Function SavePageData(ByVal blnBeforSign As Boolean) As Boolean
'功能：保存病人首页数据
'参数：blnBeforSign-是否签名时保存前调用
    Dim arrSQL() As Variant, i As Long
    Dim str确诊日期 As String, str随诊标志 As String
    Dim arrField医生() As Variant, arrValue医生() As Variant
    Dim arrField护士() As Variant, arrValue护士() As Variant
    Dim str转科科室 As String, curDate As Date
    Dim str切口 As String, str愈合 As String
    Dim lng手术分级 As Long
    Dim intIdx As Integer, str生日 As String, str病例分型 As String
    Dim lng单位ID As Long, str诊断描述 As String
    Dim str感染因素 As String
    Dim str出院去向 As String
    Dim str不良事件 As String
    Dim str生育状况 As String
    Dim str发病 As String
    Dim ArrDel As Variant
    Dim blnIsYCcheck As Boolean
    Dim blnIsDDcheck As Boolean
    Dim StrSQL As String
    Dim strTemp As String
    Dim lngCol As Long
    Dim lngRow As Long
    Dim blnTrans As Boolean, blnDiagChange As Boolean
    Dim strFilter As String, strTmp As String
    
    arrSQL = Array()
    arrField医生 = Array()
    arrField护士 = Array()
    
    '病案主页从表
    
    If txtInfo(txt转科1).Text <> "" Then
        str转科科室 = txtInfo(txt转科1).Text
        If txtInfo(txt转科2).Text <> "" Then
            str转科科室 = str转科科室 & "," & txtInfo(txt转科2).Text
            If txtInfo(txt转科3).Text <> "" Then
                str转科科室 = str转科科室 & "," & txtInfo(txt转科3).Text
            End If
        End If
    End If
    str病例分型 = Trim(cboinfo(cbo病例分型).Text)
    If InStr(str病例分型, "-") > 0 Then  '如果是规范的数据，则只存编码
        str病例分型 = Mid(str病例分型, 1, InStr(str病例分型, "-") - 1)
    End If
    
    '感染因素
    For i = 0 To lstInfection.ListCount - 1
        If lstInfection.Selected(i) = True Then
            str感染因素 = str感染因素 & IIf(i <> 0, ",", "") & lstInfection.ItemData(i)
        End If
    Next
    '不良事件
    For i = 0 To lstAdvEvent.ListCount - 1
        If lstAdvEvent.Selected(i) = True Then
            str不良事件 = str不良事件 & IIf(i <> 0, ",", "") & lstAdvEvent.ItemData(i)
            If lstAdvEvent.List(i) = "压疮" Then blnIsYCcheck = True
            If lstAdvEvent.List(i) = "医院内跌倒/坠床" Then blnIsDDcheck = True
        End If
    Next
    
    '出院去向
    str出院去向 = cboinfo(cbo出院方式).Text
    
    '发病时间
    str发病 = ""
    If IsDate(txt发病日期.Text) Then
        If IsDate(txt发病时间.Text) Then
            str发病 = txt发病日期.Text & " " & txt发病时间.Text
        Else
            str发病 = txt发病日期.Text
        End If
    End If
    '生育状况
    If cboinfo(cbo生育状况).ListIndex > 0 Then
        str生育状况 = Mid(cboinfo(cbo生育状况), 1, InStr(cboinfo(cbo生育状况), "-") - 1)
    End If
    '护士站分填时保存护士信息部分，或者医生站不分填时护士信息部
    If mbln护士站 And mbln医生护士分填首页 Or Not mbln护士站 And Not mbln医生护士分填首页 Then
        arrField护士 = Array("不良事件", "压疮发生期间", "压疮分期", "跌倒或坠床伤害", "跌倒或坠床原因", _
                    "身体约束", "约束总时间", "约束方式", "约束工具", "约束原因")
        arrValue护士 = Array(str不良事件, IIf(blnIsYCcheck, cboinfo(cbo压疮发生期间).Text, ""), IIf(blnIsYCcheck, cboinfo(cbo压疮分期).Text, ""), _
                    IIf(blnIsDDcheck, cboinfo(cbo跌倒或坠床伤害).Text, ""), IIf(blnIsDDcheck, cboinfo(cbo跌倒或坠床原因).Text, ""), _
                    chkInfo(chk是否使用物理约束).Value, txtInfo(txt约束总时间).Text, NeedName(cboinfo(cbo约束方式).Text), _
                    NeedName(cboinfo(cbo约束工具).Text), NeedName(cboinfo(cbo约束原因).Text))
    End If
    '医生站保存医生需填信息部分
    If Not mbln护士站 Then
        arrField医生 = Array("入院病室", "出院病室", "转科记录", "HBsAg", "HCV-Ab", "HIV-Ab", _
                    "中医危重", "中医急症", "中医疑难", "中医抢救方法", "自制中药制剂", "死亡根本原因", "死亡时间", _
                    "入院前经外院治疗", "示教病案", "科研病案", "疑难病历", "Rh", "输血反应", "输红细胞", "输血小板", "输血浆", "输全血", "输其他", _
                    "输液反应", "科主任", "主任医师", "主治医师", "进修医师", "研究生实习医师", "实习医师", _
                    "质控医师", "质控护士", "病原学检查", "输血检查", "CT", "MRI", "彩色多普勒", "特殊检查4", "特殊检查5", "特殊检查6", _
                    "病例分型", "感染因素", "出院方式", "出院转入", _
                    "再入院计划天数", "31天内再住院", "不足周岁年龄", "新生儿出生体重", "新生儿入院体重", "呼吸机使用时间", "昏迷时间", "抢救病因", _
                    "自体回收", "分化程度", "最高诊断依据", "中医设备", "中医技术", "辨证施护", "病理号", "籍贯", "病案质量", "主页质量日期", _
                    "告病重病危", "临床路径", "退出原因", "变异原因", _
                    "新生儿离院方式", "围术期死亡", "术后猝死", "生育状况", "发病时间", "医学警示", "其他医学警示")
                
    
    
        arrValue医生 = Array(txtInfo(txt入院病室).Text, txtInfo(txt出院病室).Text, str转科科室, _
                    NeedName(cboinfo(cboHBsAg).Text), NeedName(cboinfo(cboHCVAb).Text), NeedName(cboinfo(cboHIVAb).Text), _
                    chkInfo(chk危重).Value, chkInfo(chk急症).Value, chkInfo(chk疑难).Value, NeedName(cboinfo(cbo抢救方法).Text), _
                    NeedName(cboinfo(cbo自制中药).Text), txtInfo(txt死亡原因).Text, IIf(txt死亡时间.Text = "____-__-__ __:__:__", "", txt死亡时间.Text), chkInfo(chk经外院治疗).Value, chkInfo(chk示教病案).Value, _
                    chkInfo(chk科研病案).Value, chkInfo(chk疑难病例).Value, NeedName(cboinfo(cboRh).Text), IIf(cboinfo(cbo输血反应).ListIndex = -1, "", cboinfo(cbo输血反应).ListIndex), _
                    txtInfo(txt输红细胞).Text, txtInfo(txt输血小板).Text, txtInfo(txt输血浆).Text, _
                    txtInfo(txt输全血).Text, txtInfo(txt输其他).Text, NeedName(cboinfo(cbo输液反应).Text), _
                    NeedName(cboinfo(cbo科主任).Text), NeedName(cboinfo(cbo主任医师).Text), _
                    NeedName(cboinfo(cbo主治医师).Text), NeedName(cboinfo(cbo进修医师).Text), _
                    NeedName(cboinfo(cbo研究生医师).Text), NeedName(cboinfo(cbo实习医师).Text), _
                    NeedName(cboinfo(cbo质控医师).Text), NeedName(cboinfo(cbo质控护士).Text), _
                    chkInfo(chk病原学).Value, NeedName(cboinfo(cbo输血检查).Text), _
                    chkInfo(chkCT).Value, chkInfo(chkMRI).Value, chkInfo(chk多普勒).Value, _
                    vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1), vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1), _
                    vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1), str病例分型, str感染因素, str出院去向, IIf(cboinfo(cbo出院方式).Text = "转院" Or cboinfo(cbo出院方式).Text = "转社区", txtInfo(txt出院转入).Text, ""), _
                    cboinfo(cbo31天和7天再入院).ListIndex, IIf(txtInfo(txt31天目的).Enabled, Trim(txtInfo(txt31天目的).Text), ""), _
                    IIf(Trim(txtInfo(txt婴儿年龄).Text) <> "", txtInfo(txt婴儿年龄).Text & IIf(cboinfo(cbo婴儿年龄单位).Visible, cboinfo(cbo婴儿年龄单位).Text, ""), ""), txtInfo(txt新生儿体重).Text, txtInfo(txt新生儿入院体重).Text, _
                    txtInfo(txt呼吸机小时).Text, txtInfo(txt入院前天).Text & "," & txtInfo(txt入院前小时).Text & "," & txtInfo(txt入院前分钟).Text & "|" & txtInfo(txt入院后天).Text & "," & txtInfo(txt入院后小时).Text & "," & txtInfo(txt入院后分钟).Text, _
                    txtInfo(txt抢救原因).Text, txtInfo(txt自体回收).Text, IIf(cboinfo(cbo分化程度).Enabled, NeedName(cboinfo(cbo分化程度).Text), ""), IIf(cboinfo(cbo最高诊断依据).Enabled, NeedName(cboinfo(cbo最高诊断依据).Text), ""), _
                    NeedName(cboinfo(cbo使用中医诊疗设备).Text), NeedName(cboinfo(cbo使用中医诊疗技术).Text), NeedName(cboinfo(cbo辨证施护).Text), IIf(txtInfo(txt病理号).Enabled, txtInfo(txt病理号).Text, ""), "", NeedName(cboinfo(cbo病案质量).Text), txtInfo(txt质控日期).Text _
                    , chkInfo(chk住院期间告病重或病危).Value, chkInfo(chk进入路径).Value, IIf(chkInfo(chk完成路径).Value = 1, "1", txtInfo(txt退出原因).Text), IIf(chkInfo(chk变异).Value = 0, "0", IIf(txtInfo(txt变异原因).Text = "", " ", txtInfo(txt变异原因).Text)) _
                    , NeedName(cboinfo(cbo新生儿离院方式).Text), chkInfo(chk围术期死亡).Value, chkInfo(chk术后猝死).Value _
                    , str生育状况, str发病, txtInfo(txt医学警示).Text, txtInfo(txt其他医学警示).Text)
    End If

    '护士部分
    For i = 0 To UBound(arrField护士)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & _
            mlng病人ID & "," & mlng主页ID & ",'" & arrField护士(i) & "','" & arrValue护士(i) & "')"
    Next
    '医生部分
    For i = 0 To UBound(arrField医生)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & _
            mlng病人ID & "," & mlng主页ID & ",'" & arrField医生(i) & "','" & arrValue医生(i) & "')"
    Next
    
    If Not mbln护士站 Then
        curDate = zlDatabase.Currentdate
        
        If IsDate(txt出生时间.Text) Then
            str生日 = "To_Date('" & Format(txt出生日期.Text & " " & txt出生时间.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str生日 = "To_Date('" & Format(txt出生日期.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        End If
        If Trim(txtInfo(txt单位名称).Text) <> "" Then
            lng单位ID = Val(txtInfo(txt单位名称).Tag)
        End If
        
        '病人信息
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人信息_首页整理(" & _
            mlng病人ID & "," & IIf(txtInfo(txt住院号).Text = "", "NULL", "'" & txtInfo(txt住院号).Text & "'") & "," & _
            "'" & txtInfo(txt姓名).Text & "','" & NeedName(cboinfo(cbo性别).Text) & "','" & txtInfo(txt年龄).Text & cboinfo(cbo年龄单位).Text & "'," & _
            str生日 & ",'" & IIf(mbln启用结构化地址, PatiAddress出生地.Value, txtInfo(txt出生地点).Text) & "','" & cboinfo(cbo身份证号).Text & "'," & _
            "'" & NeedName(cboinfo(cbo民族).Text) & "','" & NeedName(cboinfo(cbo国籍).Text) & "','" & txtInfo(txt区域).Text & "'," & _
            "'" & NeedName(cboinfo(cbo婚姻).Text) & "','" & NeedName(cboinfo(cbo职业).Text) & "'," & _
            "'" & NeedName(cboinfo(cbo付款方式).Text) & "','" & IIf(mbln启用结构化地址, PatiAddress现住址.Value, txtInfo(txt家庭地址).Text) & "'," & _
            "'" & txtInfo(txt家庭电话).Text & "','" & txtInfo(txt家庭邮编).Text & "'," & _
            "'" & txtInfo(txt单位名称).Text & "','" & txtInfo(txt单位电话).Text & "'," & _
            "'" & txtInfo(txt单位邮编).Text & "','" & txtInfo(txt联系人姓名).Text & "'," & _
            "'" & NeedName(cboinfo(cbo联系人关系).Text) & "','" & txtInfo(txt联系人电话).Text & "'," & _
            "'" & txtInfo(txt联系人地址).Text & "',null,null,null,null,null,null,'" & Trim(txtInfo(txt其他证件).Text) & "'," & _
            ZVal(lng单位ID) & ",'" & IIf(mbln启用结构化地址, PatiAddress户口地址.Value, txtInfo(txt户口地址).Text) & "','" & txtInfo(txt户口邮编).Text & "','" & IIf(mbln启用结构化地址, PatiAddress籍贯.Value, txtInfo(txt籍贯).Text) & "')"
    
        '结构化地址
        If mbln启用结构化地址 Then
            '出生地
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress出生地.value省 <> "" Or PatiAddress出生地.value市 <> "" Or PatiAddress出生地.value区县 <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(1," & mlng病人ID & "," & mlng主页ID & ",1,'" & PatiAddress出生地.value省 & "','" & _
                    PatiAddress出生地.value市 & "','" & PatiAddress出生地.value区县 & "','" & PatiAddress出生地.value详细地址 & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(2," & mlng病人ID & "," & mlng主页ID & ",1)"
            End If
            '籍贯
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress籍贯.value省 <> "" Or PatiAddress籍贯.value市 <> "" Or PatiAddress籍贯.value区县 <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(1," & mlng病人ID & "," & mlng主页ID & ",2,'" & PatiAddress籍贯.value省 & "','" & _
                    PatiAddress籍贯.value市 & "','" & PatiAddress籍贯.value区县 & "','" & PatiAddress籍贯.value详细地址 & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(2," & mlng病人ID & "," & mlng主页ID & ",2)"
            End If
        End If
    
        '病案主页
        If IsDate(txtInfo(txt确诊日期).Text) Then
            str确诊日期 = "To_Date('" & Format(txtInfo(txt确诊日期).Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str确诊日期 = "NULL"
        End If
        If chkInfo(chk随诊).Value = 1 Then
            str随诊标志 = Decode(NeedName(cboinfo(cbo随诊Ex).Text), "月", 1, "年", 2, "周", 3, "天", 4, "终身", 9)
        Else
            str随诊标志 = "NULL"
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病案主页_首页整理(" & _
            mlng病人ID & "," & mlng主页ID & ",'" & NeedName(cboinfo(cbo婚姻).Text) & "'," & _
            "'" & txtInfo(txt年龄).Text & cboinfo(cbo年龄单位).Text & "','" & NeedName(cboinfo(cbo职业).Text) & "'," & _
            "'" & NeedName(cboinfo(cbo国籍).Text) & "','" & txtInfo(txt区域).Text & "'," & _
            "'" & NeedName(cboinfo(cbo付款方式).Text) & "','" & IIf(mbln启用结构化地址, PatiAddress现住址.Value, txtInfo(txt家庭地址).Text) & "'," & _
            "'" & txtInfo(txt家庭电话).Text & "','" & txtInfo(txt家庭邮编).Text & "'," & _
            "'" & txtInfo(txt单位名称).Text & "','" & txtInfo(txt单位电话).Text & "'," & _
            "'" & txtInfo(txt单位邮编).Text & "','" & txtInfo(txt联系人姓名).Text & "'," & _
            "'" & NeedName(cboinfo(cbo联系人关系).Text) & "','" & txtInfo(txt联系人电话).Text & "'," & _
            "'" & txtInfo(txt联系人地址).Text & "','" & NeedName(cboinfo(cbo入院病情).Text) & "'," & _
            "'" & chkInfo(chk是否确诊).Value & "'," & str确诊日期 & "," & _
            IIf(Val(txtInfo(txt抢救次数).Text) <> 0, Val(txtInfo(txt抢救次数).Text), "NULL") & "," & _
            IIf(Val(txtInfo(txt成功次数).Text) <> 0, Val(txtInfo(txt成功次数).Text), "NULL") & "," & _
            IIf(chkInfo(chk尸检).Enabled, chkInfo(chk尸检).Value, "NULL") & "," & _
            str随诊标志 & "," & IIf(Val(txtInfo(txt随诊期限).Text) <> 0, Val(txtInfo(txt随诊期限).Text), "NULL") & "," & _
            "'" & NeedName(cboinfo(cbo血型).Text) & "','" & NeedName(cboinfo(cbo门诊医师).Text) & "'," & _
            "'" & NeedName(cboinfo(cbo住院医师).Text) & "','" & NeedName(cboinfo(cbo主治医师).Text) & "','" & NeedName(cboinfo(cbo主任医师).Text) & _
            "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & chkInfo(chk新发肿瘤).Value & "," & _
            "'" & NeedName(cboinfo(cbo治疗类别).Text) & "'," & chkInfo(chk再入院).Value & "," & Val(txtInfo(txt身高).Text) & "," & Val(txtInfo(txt体重).Text) & ",'" & _
            NeedName(cboinfo(cbo出院方式).Text) & "','" & NeedName(cboinfo(cbo入院方式).Text) & "','" & NeedName(cboinfo(cbo责任护士).Text) & "','" & IIf(mbln启用结构化地址, PatiAddress户口地址.Value, txtInfo(txt户口地址).Text) & "','" & txtInfo(txt户口邮编).Text & "')"
    
        '结构化地址
        If mbln启用结构化地址 Then
            '现住址
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress现住址.value省 <> "" Or PatiAddress现住址.value市 <> "" Or PatiAddress现住址.value区县 <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(1," & mlng病人ID & "," & mlng主页ID & ",3,'" & PatiAddress现住址.value省 & "','" & _
                    PatiAddress现住址.value市 & "','" & PatiAddress现住址.value区县 & "','" & PatiAddress现住址.value详细地址 & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(2," & mlng病人ID & "," & mlng主页ID & ",3)"
            End If
            '户口地址
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress户口地址.value省 <> "" Or PatiAddress户口地址.value市 <> "" Or PatiAddress户口地址.value区县 <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(1," & mlng病人ID & "," & mlng主页ID & ",4,'" & PatiAddress户口地址.value省 & "','" & _
                    PatiAddress户口地址.value市 & "','" & PatiAddress户口地址.value区县 & "','" & PatiAddress户口地址.value详细地址 & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(2," & mlng病人ID & "," & mlng主页ID & ",4)"
            End If
        End If
        
        '诊断符合情况
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",1," & _
            IIf(cboinfo(cbo门诊与出院).ListIndex = -1, "NULL", cboinfo(cbo门诊与出院).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",2," & _
            IIf(cboinfo(cbo入院与出院).ListIndex = -1, "NULL", cboinfo(cbo入院与出院).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",3," & _
            IIf(cboinfo(cbo放射与病理).ListIndex = -1, "NULL", cboinfo(cbo放射与病理).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",4," & _
            IIf(cboinfo(cbo临床与病理).ListIndex = -1, "NULL", cboinfo(cbo临床与病理).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",5," & _
            IIf(cboinfo(cbo临床与尸检).ListIndex = -1, "NULL", cboinfo(cbo临床与尸检).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",6," & _
            IIf(cboinfo(cbo术前与术后).ListIndex = -1, "NULL", cboinfo(cbo术前与术后).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",7," & _
            IIf(cboinfo(cbo门诊与入院).ListIndex = -1, "NULL", cboinfo(cbo门诊与入院).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",11," & _
            IIf(cboinfo(cbo中医门诊与出院).ListIndex = -1, "NULL", cboinfo(cbo中医门诊与出院).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",12," & _
            IIf(cboinfo(cbo中医入院与出院).ListIndex = -1, "NULL", cboinfo(cbo中医入院与出院).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",13," & _
            IIf(cboinfo(cbo辨证).ListIndex = -1, "NULL", cboinfo(cbo辨证).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",14," & _
            IIf(cboinfo(cbo治法).ListIndex = -1, "NULL", cboinfo(cbo治法).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",15," & _
            IIf(cboinfo(cbo方药).ListIndex = -1, "NULL", cboinfo(cbo方药).ListIndex) & ")"
        '病人信息从表
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'血型','" & IIf(InStr(";A型;B型;O型;AB型;不详;", ";" & NeedName(cboinfo(cbo血型).Text) & ";") > 0, NeedName(cboinfo(cbo血型).Text), "") & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'RH','" & NeedName(cboinfo(cboRh).Text) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'医学警示','" & txtInfo(txt医学警示) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'其他医学警示','" & txtInfo(txt其他医学警示) & "')"
        
        '过敏药物
        If vsAller.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Delete(" & mlng病人ID & "," & mlng主页ID & ",3)"
            With vsAller
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, AC_过敏药物)) <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "zl_病人过敏记录_Insert(" & mlng病人ID & "," & mlng主页ID & "," & _
                            "3," & ZVal(.RowData(i)) & ",'" & .TextMatrix(i, AC_过敏药物) & "',1," & _
                            "To_Date('" & .Cell(flexcpData, i, AC_过敏时间) & "','YYYY-MM-DD HH24:MI:SS')," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & .TextMatrix(i, AC_过敏反应) & "')"
                    End If
                Next
            End With
        End If
        
        '西医诊断
        If vsDiagXY.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_DELETE(" & mlng病人ID & "," & mlng主页ID & ",3,NULL,'1,2,3,5,6,7,10')"
            With vsDiagXY
                intIdx = 0
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, col诊断描述)) <> "" Then
                        If Trim(.TextMatrix(i, col诊断编码)) = "" Then
                            str诊断描述 = .TextMatrix(i, col诊断描述)
                        Else
                            str诊断描述 = "(" & .TextMatrix(i, col诊断编码) & ")" & .TextMatrix(i, col诊断描述)
                        End If
                        blnDiagChange = True
                        If Val(.Cell(flexcpData, i, col是否疑诊) & "") > 0 Then
                            strFilter = "诊断类型=" & Val(.TextMatrix(i, col类型)) & " And 记录来源=3 And 疾病id=" & ZVal(.TextMatrix(i, col疾病ID)) & " And 诊断id=" & ZVal(.TextMatrix(i, col诊断ID))

                            strTmp = IIf(str诊断描述 = "", "Null", "'" & str诊断描述 & "'")
                            strFilter = strFilter & " And 诊断描述= " & strTmp

                            strTmp = .TextMatrix(i, col入院病情)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  入院病情= " & strTmp

                            strTmp = NeedName(.TextMatrix(i, col出院情况))
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  出院情况= " & strTmp
                            
                            strTmp = .TextMatrix(i, col备注)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  备注= " & strTmp
                            
                            strFilter = strFilter & " And 是否未治=" & IIf(.TextMatrix(i, col是否未治) = "", 0, 1) & " And 是否疑诊=" & IIf(.TextMatrix(i, col是否疑诊) = "", 0, 1)
                            mrsXYDiag.Filter = strFilter
                            blnDiagChange = mrsXYDiag.EOF
                        End If
                        
                        
                        If Val(.TextMatrix(i, col类型)) <> Val(.TextMatrix(i - 1, col类型)) Then intIdx = 0
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                        If mblnChange Then
                            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng主页ID & ",3,NULL," & _
                                Val(.TextMatrix(i, col类型)) & "," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & "," & _
                                "NULL,'" & str诊断描述 & "','" & NeedName(.TextMatrix(i, col出院情况)) & "'," & _
                                IIf(.TextMatrix(i, col是否未治) = "", 0, 1) & "," & IIf(.TextMatrix(i, col是否疑诊) = "", 0, 1) & "," & _
                                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col入院病情) & "',Null,'" & UserInfo.姓名 & "')"
                        Else
                            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng主页ID & ",3,NULL," & _
                                Val(.TextMatrix(i, col类型)) & "," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & "," & _
                                "NULL,'" & str诊断描述 & "','" & NeedName(.TextMatrix(i, col出院情况)) & "'," & _
                                IIf(.TextMatrix(i, col是否未治) = "", 0, 1) & "," & IIf(.TextMatrix(i, col是否疑诊) = "", 0, 1) & "," & _
                                "To_Date('" & Format(CDate(mrsXYDiag!记录日期), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col入院病情) & "',Null,'" & mrsXYDiag!记录人 & "')"
                        End If
                        If Val(.TextMatrix(i, col类型)) = 2 And intIdx = 1 Then mblnDiagChange = mstrXYDiagInfo <> str诊断描述
                    End If
                Next
            End With
        End If
        
        '病原学诊断
        If txtInfo(txt病原学).Enabled Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_DELETE(" & mlng病人ID & "," & mlng主页ID & ",3,NULL,'21')"
            If txtInfo(txt病原学).Text <> "" Then
                blnDiagChange = True
                If Not mrsXYDiag Is Nothing Then
                    strFilter = "诊断类型=21 And 记录来源=3 And 疾病id=" & ZVal(cmdInfo(txt病原学).Tag)
                    strTmp = IIf(txtInfo(txt病原学).Text = "", "Null", "'" & txtInfo(txt病原学).Text & "'")
                    strFilter = strFilter & " And 诊断描述= " & strTmp
                    
                    mrsXYDiag.Filter = strFilter
                    blnDiagChange = mrsXYDiag.EOF
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If blnDiagChange Then
                    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng主页ID & ",3,NULL,21," & _
                        ZVal(cmdInfo(txt病原学).Tag) & ",NULL,NULL,'" & txtInfo(txt病原学).Text & "',NULL,NULL,NULL," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,1,Null,Null,Null,'" & UserInfo.姓名 & "')"
                Else
                    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng主页ID & ",3,NULL,21," & _
                        ZVal(cmdInfo(txt病原学).Tag) & ",NULL,NULL,'" & txtInfo(txt病原学).Text & "',NULL,NULL,NULL," & _
                        "To_Date('" & Format(CDate(mrsXYDiag!记录日期), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,1,Null,Null,Null, '" & mrsXYDiag!记录人 & "')"
                End If
            End If
        End If
        
        '中医诊断
        If mbln中医 And vsDiagZY.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_DELETE(" & mlng病人ID & "," & mlng主页ID & ",3,NULL,'11,12,13')"
            With vsDiagZY
                intIdx = 0
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, col诊断描述)) <> "" Then
                        If Trim(.TextMatrix(i, col诊断编码)) = "" Then
                            str诊断描述 = .TextMatrix(i, col诊断描述) & IIf(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "")
                        Else
                            str诊断描述 = "(" & .TextMatrix(i, col诊断编码) & ")" & .TextMatrix(i, col诊断描述) & IIf(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "")
                        End If
                        blnDiagChange = True
                        If Val(.Cell(flexcpData, i, col是否疑诊) & "") > 0 Then
                            strFilter = "诊断类型=" & Val(.TextMatrix(i, colzy类型)) & " And 记录来源=3 And 疾病id=" & ZVal(.TextMatrix(i, colzy疾病ID)) & _
                                        " And 诊断id=" & ZVal(.TextMatrix(i, colzy诊断ID)) & " And 证候ID=" & ZVal(.TextMatrix(i, colzy证候ID))

                            strTmp = IIf(str诊断描述 = "", "Null", "'" & str诊断描述 & "'")
                            strFilter = strFilter & " And 诊断描述= " & strTmp

                            strTmp = .TextMatrix(i, col入院病情)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  入院病情= " & strTmp

                            strTmp = NeedName(.TextMatrix(i, col出院情况))
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  出院情况= " & strTmp
                                                        
                            strTmp = .TextMatrix(i, col备注)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  备注= " & strTmp
                            
                            mrsZYDiag.Filter = strFilter
                            blnDiagChange = mrsZYDiag.EOF
                        End If
                        
                        If Val(.TextMatrix(i, colzy类型)) <> Val(.TextMatrix(i - 1, colzy类型)) Then intIdx = 0
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                        If blnDiagChange Then
                            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng主页ID & ",3,NULL," & _
                                Val(.TextMatrix(i, colzy类型)) & "," & ZVal(.TextMatrix(i, colzy疾病ID)) & "," & ZVal(.TextMatrix(i, colzy诊断ID)) & "," & _
                                ZVal(.TextMatrix(i, colzy证候ID)) & ",'" & str诊断描述 & "','" & NeedName(.TextMatrix(i, col出院情况)) & "'," & _
                                "NULL,NULL,To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col入院病情) & "',Null,'" & UserInfo.姓名 & "')"
                        Else
                            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng主页ID & ",3,NULL," & _
                                Val(.TextMatrix(i, colzy类型)) & "," & ZVal(.TextMatrix(i, colzy疾病ID)) & "," & ZVal(.TextMatrix(i, colzy诊断ID)) & "," & _
                                ZVal(.TextMatrix(i, colzy证候ID)) & ",'" & str诊断描述 & "','" & NeedName(.TextMatrix(i, col出院情况)) & "'," & _
                                "NULL,NULL,To_Date('" & Format(CDate(mrsZYDiag!记录日期), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col入院病情) & "',Null,'" & mrsZYDiag!记录人 & "')"
                        
                        End If
                        If Val(.TextMatrix(i, colzy类型)) = 12 And intIdx = 1 Then mblnDiagChange = mstrZYDiagInfo <> str诊断描述
                    End If
                Next
            End With
        End If
        
        '手术情况
        If vsOPS.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人手麻记录_DELETE(" & mlng病人ID & "," & mlng主页ID & ",3)"
            
            With vsOPS
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, col手术名称)) <> "" Then
                        If Trim(.TextMatrix(i, col切口愈合)) = "" Then
                            str切口 = "NULL": str愈合 = "NULL"
                        Else
                            str切口 = "'" & Split(.TextMatrix(i, col切口愈合), "/")(0) & "'"
                            str愈合 = "'" & Split(.TextMatrix(i, col切口愈合), "/")(1) & "'"
                        End If
                        If .TextMatrix(i, col手术级别) = "一级手术" Then
                            lng手术分级 = 1
                        ElseIf .TextMatrix(i, col手术级别) = "二级手术" Then
                            lng手术分级 = 2
                        ElseIf .TextMatrix(i, col手术级别) = "三级手术" Then
                            lng手术分级 = 3
                        ElseIf .TextMatrix(i, col手术级别) = "四级手术" Then
                            lng手术分级 = 4
                        Else
                            lng手术分级 = 0
                        End If
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人手麻记录_Insert(" & _
                            zlDatabase.GetNextId("病人手麻记录") & "," & mlng病人ID & "," & mlng主页ID & ",3," & _
                            "To_Date('" & Format(.TextMatrix(i, col手术日期), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                            "NULL,NULL,NULL," & ZVal(.TextMatrix(i, col手术操作ID)) & "," & ZVal(.TextMatrix(i, col诊疗项目ID)) & "," & _
                            "'" & .TextMatrix(i, col手术名称) & "','" & .TextMatrix(i, col主刀医师) & "','" & .TextMatrix(i, col助产护士) & "'," & _
                            "'" & .TextMatrix(i, col助手1) & "','" & .TextMatrix(i, col助手2) & "',NULL,NULL,NULL," & _
                            ZVal(.TextMatrix(i, col麻醉ID)) & ",'" & .TextMatrix(i, col麻醉类型) & "',NULL,NULL," & _
                            "'" & .TextMatrix(i, col麻醉医师) & "',NULL,NULL," & str切口 & "," & str愈合 & "," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                            .TextMatrix(i, COL手术情况.COL手术情况) & "','" & .TextMatrix(i, colASA分级) & "'," & Abs(Val(.TextMatrix(i, col再次手术))) & ",'" & .TextMatrix(i, colNNIS分级) & "'," & _
                            lng手术分级 & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL," & .Cell(flexcpChecked, i, col预防用抗菌药) & "," & ZVal(Val(.TextMatrix(i, col抗菌药天数))) & "," & .Cell(flexcpChecked, i, col非预期的二次手术) & _
                            "," & .Cell(flexcpChecked, i, col麻醉并发症) & "," & .Cell(flexcpChecked, i, col术中异物遗留) & "," & .Cell(flexcpChecked, i, col手术并发症) & "," & .Cell(flexcpChecked, i, col术后出血或血肿) & _
                            "," & .Cell(flexcpChecked, i, col手术伤口裂开) & "," & .Cell(flexcpChecked, i, col术后深静脉血栓) & "," & .Cell(flexcpChecked, i, col术后生理代谢紊乱) & "," & .Cell(flexcpChecked, i, col术后呼吸衰竭) & _
                            "," & .Cell(flexcpChecked, i, col术后肺栓塞) & "," & .Cell(flexcpChecked, i, col术后败血症) & "," & .Cell(flexcpChecked, i, col术后髋关节骨折) & ")"
                    End If
                Next
            End With
        End If
            
        '使用抗生素的记录
        If vsKSS.Tag = "" Then
            With vsKSS
                '先删除用户操作过的记录
                ArrDel = Split(mstrDelete, ",")
                mstrDelete = ""
                For i = 0 To UBound(ArrDel)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人抗生素记录_Update(" & _
                            "2," & mlng病人ID & "," & mlng主页ID & "," & ArrDel(i) & ",'" & .TextMatrix(i, kss名称) & "',NULL,NULL,NULL,'" & UserInfo.姓名 & "',To_Date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'))"
                Next
                '插入、如果存在则修改界面上的数据
                For i = 1 To .Rows - 1
                    If Val(.RowData(i) & "") <> 0 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人抗生素记录_Update(" & _
                                "1," & mlng病人ID & "," & mlng主页ID & "," & Val(.RowData(i) & "") & ",'" & .TextMatrix(i, kss名称) & "','" & .TextMatrix(i, kss用药目的) & _
                                "','" & .TextMatrix(i, kss使用阶段) & "'," & Val(.TextMatrix(i, kss使用天数)) & ",'" & UserInfo.姓名 & "',To_Date('" & curDate & "','YYYY-MM-DD HH24:MI:SS')" & _
                                "," & IIf(.Cell(flexcpChecked, i, KSS一类切口预防用) = "", "Null", .Cell(flexcpChecked, i, KSS一类切口预防用)) & "," & ZVal(.TextMatrix(i, KSSDDD数)) & ",'" & .TextMatrix(i, KSS联合用药) & "')"
                    End If
                Next
                '将老数据-病案主页从表的数据一起删去
                For i = 1 To 10
                    If .FixedRows + i - 1 <= .Rows - 1 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        '删除
                        arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'抗生素" & i & "',NULL)"
                    End If
                Next
            End With
        End If
        
        
        If mbln病案共享 Then
            '放疗化疗
            '先删除信息
            'Zl_病案化疗记录_Delete
            If vs化疗.Tag = "" Then
                StrSQL = "Zl_病案化疗记录_Delete("
                '  病人id_In In 病案化疗记录.病人id%Type,
                StrSQL = StrSQL & "" & mlng病人ID & ","
                '  主页id_In In 病案化疗记录.主页id%Type
                StrSQL = StrSQL & "" & mlng主页ID & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = StrSQL
                
                With vs化疗
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("开始日期"))) <> "" And _
                              Val(.Cell(flexcpData, i, .ColIndex("化学治疗编码"))) <> 0 Then
                            'Zl_病案化疗记录_Insert
                            StrSQL = "Zl_病案化疗记录_Insert("
                            '  病人id_In   In 病案化疗记录.病人id%Type,
                            StrSQL = StrSQL & "" & mlng病人ID & ","
                            '  主页id_In   In 病案化疗记录.主页id%Type,
                            StrSQL = StrSQL & "" & mlng主页ID & ","
                            '  序号_In     In 病案化疗记录.序号%Type,
                            StrSQL = StrSQL & "" & i & ","
                            '  疾病id_In   In 病案化疗记录.疾病id%Type,
                            StrSQL = StrSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("化学治疗编码"))) & ","
                            '  开始日期_In In 病案化疗记录.开始日期%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("开始日期")) & "','yyyy-mm-dd'),"
                            '  结束日期_In In 病案化疗记录.结束日期%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("结束日期")) & "','yyyy-mm-dd'),"
                            '  疗程数_In   In 病案化疗记录.疗程数%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("疗程数"))) & ","
                            '  总量_In     In 病案化疗记录.总量%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("总量"))) & ","
                            '  化疗方案_In In 病案化疗记录.化疗方案%Type,
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("化疗方案"))) & "',"
                            '  化疗效果_In In 病案化疗记录.化疗效果%Type
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("化疗效果"))) & "')"
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = StrSQL
                        End If
                    Next
                End With
            End If
            
            If vs放疗.Tag = "" Then
                '先删除信息
                'Zl_病案放疗记录_Delete
                StrSQL = "Zl_病案放疗记录_Delete("
                '  病人id_In In 病案放疗记录.病人id%Type,
                StrSQL = StrSQL & "" & mlng病人ID & ","
                '  主页id_In In 病案放疗记录.主页id%Type
                StrSQL = StrSQL & "" & mlng主页ID & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = StrSQL
                With vs放疗
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("开始日期"))) <> "" And _
                              Val(.Cell(flexcpData, i, .ColIndex("放射治疗编码"))) <> 0 Then
                            'Zl_病案放疗记录_Insert
                            StrSQL = "Zl_病案放疗记录_Insert("
                            '  病人id_In   In 病案放疗记录.病人id%Type,
                            StrSQL = StrSQL & "" & mlng病人ID & ","
                            '  主页id_In   In 病案放疗记录.主页id%Type,
                            StrSQL = StrSQL & "" & mlng主页ID & ","
                            '  序号_In     In 病案放疗记录.序号%Type,
                            StrSQL = StrSQL & "" & i & ","
                            '  疾病id_In   In 病案放疗记录.疾病id%Type,
                            StrSQL = StrSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("放射治疗编码"))) & ","
                            '  开始日期_In In 病案放疗记录.开始日期%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("开始日期")) & "','yyyy-mm-dd'),"
                            '  结束日期_In In 病案放疗记录.结束日期%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("结束日期")) & "','yyyy-mm-dd'),"
                            '  设野部位_In In 病案放疗记录.设野部位%Type,
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("设野部位"))) & "',"
                            '  放射剂量_In In 病案放疗记录.放射剂量%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("放射剂量"))) & ","
                            '  累计量_In   In 病案放疗记录.累计量%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("累计量"))) & ","
                            '  放疗效果_In In 病案放疗记录.放疗效果%Type
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("放疗效果"))) & "')"
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = StrSQL
                        End If
                    Next
                End With
            End If
            
            
        End If
        '病案附加信息
        If vsfMain.Tag = "" Then
            For lngRow = 1 To vsfMain.Rows - 1
                For lngCol = 0 To vsfMain.Cols - 1 Step 3
                    If vsfMain.TextMatrix(lngRow, lngCol + 2) = "是否" Then
                        strTemp = IIf(vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = 2, 0, 1)
                    Else
                        strTemp = vsfMain.TextMatrix(lngRow, lngCol + 1)
                    End If
                    If vsfMain.TextMatrix(lngRow, lngCol) <> "" And strTemp <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'" & vsfMain.TextMatrix(lngRow, lngCol) & "','" & strTemp & "')"
                    ElseIf vsfMain.TextMatrix(lngRow, lngCol) <> "" And strTemp = "" Then
                        '刘兴宏:11557:2007/09/14:增加
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'" & vsfMain.TextMatrix(lngRow, lngCol) & "',NULL)"
                    End If
                Next lngCol
            Next lngRow
        End If
        '重症监护记录
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病案重症监护情况_Delete(" & mlng病人ID & "," & mlng主页ID & ")"
        
        StrSQL = "Zl_病案重症监护情况_Insert("
        StrSQL = StrSQL & "" & mlng病人ID & ","
        StrSQL = StrSQL & "" & mlng主页ID & ","
        StrSQL = StrSQL & "1,"
        StrSQL = StrSQL & "'" & txtInfo(txt重症监护室).Text & "',NULL,NULL,null,null,"
        StrSQL = StrSQL & chkInfo(chk人工气道脱出).Value & ","
        StrSQL = StrSQL & chkInfo(chk重返重症医学科).Value & ","
        StrSQL = StrSQL & "'" & cboinfo(cbo重返间隔时间).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = StrSQL
    End If
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    '调用医保病人信息修改接口
    If mint险类 <> 0 Then
        If Not gclsInsure.ModiPatiSwap(mlng病人ID, mlng主页ID, mint险类, "2") Then
            gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    
    On Error GoTo 0
    Screen.MousePointer = 0
    mblnChange = False
    SavePageData = True
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsOPS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not OPSCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub

Private Sub SetFaceEditable(ByVal blnReadOnly As Boolean)
'功能：根据当前是否只读，设置界面的可编辑属性
    Dim objControl As Object, blnTmp As Boolean
    Dim bln首页 As Boolean, strTypeName As String
    
    bln首页 = InStr(mstrPrivs, "首页基本信息") = 0
    For Each objControl In Me.Controls
        blnTmp = blnReadOnly
        strTypeName = TypeName(objControl)
        If InStr("TextBox;MaskEdBox;ComboBox;CheckBox;VSFlexGrid;ListBox;OptionButton;CommandButton;DTPicker;PatiAddress", TypeName(objControl)) > 0 Then
            'TabStop=False表示当前确实不可编辑的
            If TypeName(objControl.Container) = "Frame" And (objControl.TabStop = True Or TypeName(objControl) = "OptionButton" And objControl.TabStop = False) Then
                '首页可编辑
                If Not blnTmp Then
                    If mbln医生护士分填首页 And mbln护士站 Then '护士站调用首页时只能填写不良事件
                        If TypeName(objControl) <> "PatiAddress" Then
                            If objControl.Container.hwnd <> fraAdvEvent.hwnd And InStr("," & chkInfo(chk是否使用物理约束).hwnd & "," & txtInfo(txt约束总时间).hwnd & _
                                    "," & cboinfo(cbo约束方式).hwnd & "," & cboinfo(cbo约束工具).hwnd & "," & cboinfo(cbo约束原因).hwnd & ",", "," & objControl.hwnd & ",") = 0 Then
                                blnTmp = True
                                objControl.TabStop = False
                            End If
                        Else
                            blnTmp = True
                            objControl.TabStop = False
                        End If
                    Else
                        '判断用户是否有首页基本信息权限
                        If bln首页 Then
                            If objControl.Container.hwnd = fraInfo(0).hwnd Then
                                '入院时间之后的不受控制
                                If objControl.TabIndex < 85 Then
                                    Select Case strTypeName
                                        Case "TextBox", "ComboBox"
                                            If objControl.Text <> "" Then
                                                blnTmp = True
                                            End If
                                        Case "MaskEdBox"
                                            If objControl.Name = "txt出生日期" And IsDate(objControl.Text) Then
                                                blnTmp = True
                                            ElseIf objControl.Name = "txt出生时间" And objControl.Text <> "__:__" Then
                                                blnTmp = True
                                            End If
                                        Case "CheckBox"
                                            If objControl.Value = 1 Then
                                                blnTmp = True
                                            End If
                                        Case "CommandButton"
                                            If txtInfo(objControl.Index).Text <> "" Then
                                                blnTmp = True
                                            End If
                                        Case "PatiAddress"
                                            If objControl.value区县 & objControl.value省 & objControl.value市 & objControl.value详细地址 <> "" Then
                                                blnTmp = True
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                        '中医科室不使用西医病案首页项目
                        If mbln不使用西医项目 And mbln中医 Then
                            If InStr("," & cboinfo(cbo病例分型).hwnd & "," & cboinfo(cbo输液反应).hwnd & "," & cboinfo(cbo输血反应).hwnd & _
                            "," & cboinfo(cbo输血检查).hwnd & "," & cboinfo(cboHBsAg).hwnd & "," & cboinfo(cboHCVAb).hwnd & "," & cboinfo(cboHIVAb).hwnd & _
                             "," & cboinfo(cbo随诊Ex).hwnd & "," & cboinfo(cbo研究生医师).hwnd & "," & txtInfo(txt输红细胞).hwnd & "," & txtInfo(txt输其他).hwnd & _
                              "," & txtInfo(txt输全血).hwnd & "," & txtInfo(txt输血浆).hwnd & "," & txtInfo(txt自体回收).hwnd & "," & txtInfo(txt输血小板).hwnd & _
                              "," & txtInfo(txt随诊期限).hwnd & "," & txtInfo(txt呼吸机小时).hwnd & "," & chkInfo(chk科研病案).hwnd & "," & chkInfo(chk示教病案).hwnd & _
                              "," & chkInfo(chk随诊).hwnd & ",", "," & objControl.hwnd & ",") > 0 Then
                                blnTmp = True
                                objControl.TabStop = False
                            End If
                        End If
                        If mbln医生护士分填首页 And Not mbln护士站 And TypeName(objControl) <> "PatiAddress" Then '医生站不能填写不良事件
                            If objControl.Container.hwnd = fraAdvEvent.hwnd Or InStr("," & chkInfo(chk是否使用物理约束).hwnd & "," & txtInfo(txt约束总时间).hwnd & _
                                     "," & cboinfo(cbo约束方式).hwnd & "," & cboinfo(cbo约束工具).hwnd & "," & cboinfo(cbo约束原因).hwnd & ",", "," & objControl.hwnd & ",") > 0 Then
                                blnTmp = True
                                objControl.TabStop = False
                            End If
                        End If
                    End If
                End If
                
                If TypeName(objControl) = "TextBox" And objControl.Enabled Then
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                    objControl.Locked = blnTmp
                ElseIf TypeName(objControl) = "MaskEdBox" Then
                    '没有Locked属性,用Enabled实现
                    objControl.Enabled = Not blnTmp
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "ComboBox" And objControl.Enabled Then
                    If Not ((objControl Is cboinfo(cbo科主任) Or objControl Is cboinfo(cbo主任医师) _
                        Or objControl Is cboinfo(cbo主治医师) Or objControl Is cboinfo(cbo住院医师)) And Not mbln护士站) Then
                        objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                        objControl.Locked = blnTmp
                    End If
                ElseIf TypeName(objControl) = "DTPicker" Then
                    objControl.Enabled = Not blnTmp
                ElseIf TypeName(objControl) = "CheckBox" Then
                    '没有Locked属性,用Enabled实现
                    objControl.Enabled = Not blnTmp
                ElseIf TypeName(objControl) = "VSFlexGrid" Then
                    '同时注意要在键盘鼠标事件中进行一些控制
                    objControl.Editable = IIf(blnTmp, flexEDNone, flexEDKbdMouse)
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                    objControl.BackColorBkg = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "ListBox" Then
                    objControl.Enabled = IIf(blnTmp, False, True)
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "CommandButton" Then
                    If objControl.Name = "cmdAutoLoad" Or objControl.Name = "cmdPathLoad" Then
                        objControl.Enabled = IIf(blnTmp, False, True)
                    End If
                ElseIf TypeName(objControl) = "PatiAddress" Then
                    objControl.ControlLock = blnTmp
                End If
            End If
            '"OptionButton"这种用Enabled判断
            If TypeName(objControl) = "OptionButton" And TypeName(objControl.Container) = "Frame" And objControl.Enabled = True Then
                objControl.Enabled = IIf(blnTmp, False, True)
                objControl.BackColor = IIf(blnTmp, vbButtonFace, &H8000000F)
            End If
        End If
    Next
End Sub

Public Function BinToDec(ByVal strBin As String) As Long
'功能：将二进制串转换为十进制数字
    Dim i As Byte, X As Long
    
    For i = 1 To Len(strBin)
        X = X * 2 + Val(Mid(strBin, i, 1))
    Next i
    
    BinToDec = X
End Function

Private Function SetSignature(Optional ByVal blnReload As Boolean = True) As Boolean
'功能：根据当前病人的医师及签名情况，确定签名及界面数据的可编辑性
'返回：界面是否已签名只读不能编辑
    Static rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim StrSQL As String, blnReadOnly As Boolean
    Dim i As Integer
    
    '初始化签名相关界面
    blnReadOnly = False
    For i = 0 To cmdSign.UBound
        cmdSign(i).Enabled = False: cmdUnSign(i).Enabled = False
    Next
    cboinfo(cbo科主任).ForeColor = Me.ForeColor: lblInfo(cbo科主任).ForeColor = Me.ForeColor
    cboinfo(cbo主任医师).ForeColor = Me.ForeColor: lblInfo(cbo主任医师).ForeColor = Me.ForeColor
    cboinfo(cbo主治医师).ForeColor = Me.ForeColor: lblInfo(cbo主治医师).ForeColor = Me.ForeColor
    cboinfo(cbo住院医师).ForeColor = Me.ForeColor: lblInfo(cbo住院医师).ForeColor = Me.ForeColor
    cboinfo(cbo科主任).Locked = False: cboinfo(cbo科主任).BackColor = vbWindowBackground
    cboinfo(cbo主任医师).Locked = False: cboinfo(cbo主任医师).BackColor = vbWindowBackground
    cboinfo(cbo主治医师).Locked = False: cboinfo(cbo主治医师).BackColor = vbWindowBackground
    cboinfo(cbo住院医师).Locked = False: cboinfo(cbo住院医师).BackColor = vbWindowBackground
    
    '获取当前人员最高签名级别
    If NeedName(cboinfo(cbo住院医师).Text) = UserInfo.姓名 Then
        '有该级别签名权限初始
        intCurr = 1: cmdSign(cmd住院医师).Enabled = True: cmdUnSign(cmd住院医师).Enabled = False
    End If
    If NeedName(cboinfo(cbo主治医师).Text) = UserInfo.姓名 Then
        intCurr = 2: cmdSign(cmd主治医师).Enabled = True: cmdUnSign(cmd主治医师).Enabled = False
    End If
    If NeedName(cboinfo(cbo主任医师).Text) = UserInfo.姓名 Then
        intCurr = 3: cmdSign(cmd主任医师).Enabled = True: cmdUnSign(cmd主任医师).Enabled = False
    End If
    If NeedName(cboinfo(cbo科主任).Text) = UserInfo.姓名 Then
        intCurr = 4: cmdSign(cmd科主任).Enabled = True: cmdUnSign(cmd科主任).Enabled = False
    End If
    
    '获取首页已经签名最高级别
    If rsTmp Is Nothing Or blnReload Then
        On Error GoTo errH
        Set rsTmp = Nothing
        StrSQL = "Select 信息名,信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2] And 信息值 is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    End If
    rsTmp.Filter = "信息名='住院医师签名'"
    If Not rsTmp.EOF Then
        intHave = 1
        
        '已签名用蓝色字表示
        cboinfo(cbo住院医师).ForeColor = vbBlue: lblInfo(cbo住院医师).ForeColor = vbBlue
        
        '签名按钮可操作状态
        If rsTmp!信息值 = UserInfo.姓名 Then
            cmdSign(cmd住院医师).Enabled = False: cmdUnSign(cmd住院医师).Enabled = True
        Else '非自已的签名不能取消
            cmdSign(cmd住院医师).Enabled = False: cmdUnSign(cmd住院医师).Enabled = False
        End If
    End If
    rsTmp.Filter = "信息名='主治医师签名'"
    If Not rsTmp.EOF Then
        intHave = 2
        
        '已签名用蓝色字表示
        cboinfo(cbo主治医师).ForeColor = vbBlue: lblInfo(cbo主治医师).ForeColor = vbBlue
        
        '签名按钮可操作状态
        If rsTmp!信息值 = UserInfo.姓名 Then
            cmdSign(cmd主治医师).Enabled = False: cmdUnSign(cmd主治医师).Enabled = True
        Else '非自已的签名不能取消
            cmdSign(cmd主治医师).Enabled = False: cmdUnSign(cmd主治医师).Enabled = False
        End If
        
        '低级别签名不能变更
        cmdSign(cmd住院医师).Enabled = False: cmdUnSign(cmd住院医师).Enabled = False
    End If
    rsTmp.Filter = "信息名='主任医师签名'"
    If Not rsTmp.EOF Then
        intHave = 3
        
        '已签名用蓝色字表示
        cboinfo(cbo主任医师).ForeColor = vbBlue: lblInfo(cbo主任医师).ForeColor = vbBlue
        
        '签名按钮可操作状态
        If rsTmp!信息值 = UserInfo.姓名 Then
            cmdSign(cmd主任医师).Enabled = False: cmdUnSign(cmd主任医师).Enabled = True
        Else '非自已的签名不能取消
            cmdSign(cmd主任医师).Enabled = False: cmdUnSign(cmd主任医师).Enabled = False
        End If
        
        '低级别签名不能变更
        cmdSign(cmd住院医师).Enabled = False: cmdUnSign(cmd住院医师).Enabled = False
        cmdSign(cmd主治医师).Enabled = False: cmdUnSign(cmd主治医师).Enabled = False
    End If
    rsTmp.Filter = "信息名='科主任签名'"
    If Not rsTmp.EOF Then
        intHave = 4
        
        '已签名用蓝色字表示
        cboinfo(cbo科主任).ForeColor = vbBlue
        lblInfo(cbo科主任).ForeColor = vbBlue
        
        '签名按钮可操作状态
        If rsTmp!信息值 = UserInfo.姓名 Then
            cmdSign(cmd科主任).Enabled = False: cmdUnSign(cmd科主任).Enabled = True
        Else '非自已的签名不能取消
            cmdSign(cmd科主任).Enabled = False: cmdUnSign(cmd科主任).Enabled = False
        End If
        
        '低级别签名不能变更
        cmdSign(cmd住院医师).Enabled = False: cmdUnSign(cmd住院医师).Enabled = False
        cmdSign(cmd主治医师).Enabled = False: cmdUnSign(cmd主治医师).Enabled = False
        cmdSign(cmd主任医师).Enabled = False: cmdUnSign(cmd主任医师).Enabled = False
    End If
    If intHave > 0 Then
        '涉及签名的项都不允许再更改,不然权限混乱
        cboinfo(cbo科主任).Locked = True: cboinfo(cbo科主任).BackColor = vbButtonFace
        cboinfo(cbo主任医师).Locked = True: cboinfo(cbo主任医师).BackColor = vbButtonFace
        cboinfo(cbo主治医师).Locked = True: cboinfo(cbo主治医师).BackColor = vbButtonFace
        cboinfo(cbo住院医师).Locked = True: cboinfo(cbo住院医师).BackColor = vbButtonFace
    End If
    
    '如果当前人员签名级别不高于已签名级别，则不可编辑
    If intCurr <= intHave And intHave > 0 Then
        blnReadOnly = True
    End If
    
    SetSignature = blnReadOnly
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str性别 As String, int性别 As Integer
    Dim strInput As String, vPoint As POINTAPI
    
    On Error GoTo errH
    With vsOPS
        If Col = col手术编码 Or Col = col手术名称 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf Col = col手术名称 And .TextMatrix(Row, col手术编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '判断加了前缀后的名称是否存在其他的诊断编码
                strInput = UCase(.EditText)
                StrSQL = GetSQL(2, strInput, str性别)
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str性别, int性别)
                If rsTmp.RecordCount <> 1 Then
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, col手术名称) = .EditText
                Else
                    Call OPSSetInput(Row, Col, rsTmp)
                    .EditText = .Text
                End If
'                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
                .Tag = ""
                mblnChange = True
            Else
                strInput = UCase(.EditText)
                StrSQL = GetSQL(2, strInput, str性别)
                If str性别 = "男" Then
                    int性别 = 1
                ElseIf str性别 = "女" Then
                    int性别 = 2
                End If
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(optInput(4).Value, "手术项目", "手术编码"), False, "", "", False, True, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, int性别)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        If chkInfo(chk手术自由录入).Value = 0 Or Col = col手术编码 Then
                            MsgBox "没有找到您查找的手术项目。", vbInformation, Me.Caption
                            Cancel = True
                        Else
                            .TextMatrix(Row, col手术编码) = ""
                            .Cell(flexcpData, Row, col手术编码) = ""
                            .TextMatrix(Row, col诊疗项目ID) = ""
                            .TextMatrix(Row, col手术操作ID) = ""
                            .Tag = ""
                            '输入后始终保持一新行
                            If Row = .Rows - 1 Then .AddItem ""
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col麻醉方式 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            Else
                strInput = UCase(.EditText)
                StrSQL = _
                    " Select A.ID,A.编码,A.名称,A.操作类型 as 麻醉类型" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.类别='G' And A.ID=B.诊疗项目ID" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or B.简码 Like [2] Or B.名称 Like [2])" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
                
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "麻醉项目", False, "", "", False, True, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的麻醉项目！", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col主刀医师 Or Col = col助手1 Or Col = col助手2 Or Col = col麻醉医师 Then
            If (Col = col助手1 Or Col = col助手2) And .EditText = "" Then
                .TextMatrix(Row, Col) = "": .Cell(flexcpData, Row, Col) = ""
                If Col = col助手1 Then
                    .TextMatrix(Row, col助手2) = "": .Cell(flexcpData, Row, col助手2) = ""
                End If
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            Else
                strInput = UCase(.EditText)
                StrSQL = "Select A.ID,A.编号,A.姓名,A.简码" & _
                    " From 人员表 A,人员性质说明 B" & _
                    " Where A.ID=B.人员ID And B.人员性质='医生'" & _
                    " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                    " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2])" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编号"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "医生", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        If (Col = col主刀医师 Or Col = col助手1 Or Col = col助手2) And zlCommFun.IsCharChinese(.EditText) Then
                            If MsgBox("没有找到匹配的本院医生，是否录入未在本院建档的医生？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                                .Tag = ""
                                mblnChange = True
                                If mblnReturn Then Call OPSEnterNextCell
                                Exit Sub
                            End If
                        Else
                            MsgBox "没有找到匹配的医生！", vbInformation, gstrSysName
                        End If
                    End If
                    Cancel = True
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col助产护士 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            Else
                strInput = UCase(.EditText)
                StrSQL = "Select A.ID,A.编号,A.姓名,A.简码" & _
                    " From 人员表 A,人员性质说明 B" & _
                    " Where A.ID=B.人员ID And B.人员性质='护士'" & _
                    " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                    " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2])" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编号"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "护士", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的护士！", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = COL手术情况.COL手术情况 Or Col = colASA分级 Or Col = colNNIS分级 Then
            If .TextMatrix(Row, Col) <> .EditText Then
                If .Tag = "未修改" Then .Tag = "": mblnChange = True
            End If
        ElseIf Col = col再次手术 Or Col = col预防用抗菌药 Or Col = col非预期的二次手术 Or Col = col麻醉并发症 Or Col = col术中异物遗留 Or Col = col手术并发症 _
                Or Col = col术后出血或血肿 Or Col = col手术伤口裂开 Or Col = col术后深静脉血栓 Or Col = col术后生理代谢紊乱 Or Col = col术后呼吸衰竭 _
                Or Col = col术后肺栓塞 Or Col = col术后败血症 Or Col = col术后髋关节骨折 Then
            If .TextMatrix(Row, Col) <> IIf(.EditText = 2, 0, -1) Then
                If .Tag = "未修改" Then .Tag = "": mblnChange = True
            End If
        ElseIf Col = col抗菌药天数 Then
            If Len(Trim(.EditText)) > 5 Then
                MsgBox "抗菌用药天数不能超过5位数。", vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            Else
                If .EditText <> .TextMatrix(Row, Col) And .Tag = "未修改" Then .Tag = "": mblnChange = True
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsTSJC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsTSJC_AfterRowColChange(-1, -1, vsTSJC.Row, vsTSJC.Col)
End Sub

Private Sub vsTSJC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsTSJC.ComboList = "..."
End Sub

Private Sub vsTSJC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    With vsTSJC
        strSQLItem = _
            " From 诊疗项目目录 A" & _
            " Where A.类别='D' And A.服务对象 IN(2,3) And A.单独应用=1" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
        StrSQL = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select A.分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
            " Group by ID,上级ID,编码,名称"
        StrSQL = StrSQL & " Union ALL" & _
            " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位" & _
            strSQLItem & " Order By 末级,级ID Desc,编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, "特殊检查", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有检查项目数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call TSJCSetDiagInput(Row, rsTmp)
            Call TSJCEnterNextCell
        End If
    End With
End Sub

Private Sub vsTSJC_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    If KeyCode = vbKeyF4 Then
        Call zlCommFun.PressKey(vbKeySpace)
    ElseIf KeyCode = vbKeyDelete Then
        If MsgBox("确实要删除该行内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsTSJC
                .TextMatrix(.Row, 1) = ""
            End With
            mblnChange = True
        End If
    ElseIf KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsTSJC_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsTSJC_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    
    With vsTSJC
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call TSJCEnterNextCell
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsTSJC_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsTSJC_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsTSJC_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsTSJC.EditSelStart = 0
    vsTSJC.EditSelLength = zlCommFun.ActualLen(vsTSJC.EditText)
End Sub

Private Sub vsTSJC_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsTSJC
        If .EditText = "" Then
            .EditText = .Cell(flexcpData, Row, Col)
            If mblnReturn Then Call TSJCEnterNextCell
        ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
            If mblnReturn Then Call TSJCEnterNextCell
        Else
            strInput = UCase(.EditText)
            If LenB(StrConv(strInput, vbFromUnicode)) > 100 Then
                MsgBox "您输入的内容不能超过50个汉字。"
                Cancel = True
                Exit Sub
            End If
            If zlCommFun.IsCharChinese(strInput) Then
                StrSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
            Else
                StrSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
            End If
            StrSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='D' And A.服务对象 IN(2,3)" & _
                " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And A.单独应用=1 And B.码类=[3] And (" & StrSQL & ")" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
            If zlCommFun.IsCharChinese(strInput) Then
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint简码 + 1)
                If rsTmp.EOF Then
                    Set rsTmp = Nothing
                ElseIf rsTmp.RecordCount > 1 Then
                    Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                End If
                Call TSJCSetDiagInput(Row, rsTmp)
                .EditText = .Text
                If mblnReturn Then Call TSJCEnterNextCell
            Else
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "特殊检查", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint简码 + 1)
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    Call TSJCSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then Call TSJCEnterNextCell
                End If
            End If
        End If
        mblnReturn = False
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDateRange(ByVal strDate As String, Optional ByVal blnCheckData As Boolean) As Boolean
    '检查录入日期是否在入出院日期范围内，
    'blnCheckData true:只检查日期范围，不检查时间范围，false:检查具体时间范围
    ' 入院日期为空，返回false,出院日期为空则处理为3000-01-01
    
    Dim DateStart As Date, dateEnd As Date
    
    On Error GoTo errH
    CheckDateRange = False
    If Not IsDate(strDate) Then Exit Function
    
    If Trim("" & txtInfo(txt入院时间).Text) = "" Then
        DateStart = CDate(0)
    Else
        DateStart = CDate(Trim("" & txtInfo(txt入院时间).Text))
    End If
    If Trim("" & txtInfo(txt出院时间).Text) = "" Then
        dateEnd = CDate(0)
    Else
        dateEnd = CDate(Trim("" & txtInfo(txt出院时间).Text))
    End If
    
    If DateStart = CDate(0) Then Exit Function
    If dateEnd = CDate(0) Then dateEnd = zlDatabase.Currentdate
    
    If blnCheckData Then
        If Between(Format(strDate, "yyyy-MM-dd"), Format(DateStart, "yyyy-MM-dd"), Format(dateEnd, "yyyy-MM-dd")) Then
            CheckDateRange = True
        End If
    Else
        If CDate(strDate) >= DateStart And CDate(strDate) <= dateEnd Then
            CheckDateRange = True
        End If
    End If
    Exit Function
errH:
    CheckDateRange = False
End Function

Private Sub vs放疗_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '设置相关的格式
    '刘兴宏:2007/09/17
    '--------------------------------------------------------------------------------
    With vs放疗
        Select Case Col
        Case .ColIndex("放射治疗编码")
'            .ColComboList(Col) = "..."
            If .ComboIndex < 0 Then Exit Sub
            .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
        Case .ColIndex("设野部位")
           ' .ColComboList(Col) = "..."
        End Select
    End With
End Sub
 

Private Sub vs放疗_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
  Call zl_VsGridRowChange(vs放疗, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs放疗_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vs放疗
        Select Case Col
        Case .ColIndex("放射治疗编码"), .ColIndex("设野部位")
        Case .ColIndex("开始日期"), .ColIndex("结束日期")
        Case .ColIndex("放射剂量"), .ColIndex("累计量")
        Case .ColIndex("放疗效果")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vs放疗_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '功能:按钮选择
    '参数:
    '--------------------------------------------------------------------------
    With vs放疗
        Select Case Col
        Case .ColIndex("放射治疗编码")
'            If Select化疗与放疗("", False) = False Then
'                Exit Sub
'            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs放疗_GotFocus()
    Call zl_VsGridGotFocus(vs放疗)
End Sub

Private Sub vs放疗_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lngCol As Long, lngRow As Long, strKEY As String
   If mbln护士站 Or mblnReadOnly Then Exit Sub
    With vs放疗
        If (.Col = .ColIndex("设野部位")) And KeyCode <> vbKeyReturn Then
          '  .ColComboList(.Col) = ""
        End If
        If KeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行的放疗信息吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                    .RowData(.Row) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            zlCtlSetFocus vs放疗, True
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs放疗
        If Val(.Cell(flexcpData, .Row, .ColIndex("放射治疗编码"))) = 0 Or Trim(.TextMatrix(.Row, .ColIndex("开始日期"))) = "" Then
            Err = 0: On Error Resume Next
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call vsKSS.SetFocus
            Exit Sub
        End If
        Select Case .Col
        Case .Cols - 1
            If Not .Row >= .Rows - 1 Then
                .Col = .ColIndex("放射治疗编码")
                .Row = .Row + 1
            Else
                Call vs放疗_KeyDownEdit(.Row, .Col, KeyCode, Shift)
            End If
            .SetFocus
        Case Else
            zlCommFun.PressKey vbKeyRight
        End Select
    End With
End Sub

Private Sub vs放疗_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer, lngRow As Long, strKEY As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs放疗
        Select Case Col
        Case .ColIndex("放射治疗编码")
'            strKey = Trim(.EditText)
'            strKey = Replace(strKey, Chr(vbKeyReturn), "")
'            strKey = Replace(strKey, Chr(10), "")
'            If strKey = "" Then Exit Sub
'            If Select化疗与放疗(strKey, True) = False Then
'                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
'                For intCol = 0 To .Cols - 1
'                    If intCol <> Col Then
'                        .TextMatrix(Row, intCol) = ""
'                        .Cell(flexcpData, Row, intCol) = ""
'                    End If
'                Next
'                Exit Sub
'            End If
'            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vs放疗, .ColIndex("开始日期"), .Cols - 1, True, lngRow)
        If lngRow > 0 Then
            '表示新增加了一行,需要设置相关的缺省值
            strKEY = .ColData(.ColIndex("放射治疗编码"))
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngRow, .ColIndex("放射治疗编码")) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngRow, .ColIndex("放射治疗编码")) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
            End If
        End If
    End With
     
End Sub

Private Sub vs放疗_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs放疗_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vs放疗
        Select Case Col
        Case .ColIndex("放射治疗编码"), .ColIndex("设野部位")
            Call VsFlxGridCheckKeyPress(vs放疗, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("开始日期"), .ColIndex("结束日期")
            Call VsFlxGridCheckKeyPress(vs放疗, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("放射剂量"), .ColIndex("累计量")
            Call VsFlxGridCheckKeyPress(vs放疗, Row, Col, KeyAscii, m金额式)
        Case .ColIndex("放疗效果")
        Case Else
        End Select
    End With
End Sub

Private Sub vs放疗_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs放疗)
End Sub

Private Sub vs放疗_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKEY As String
    Dim intCol As Integer
    Dim strTemp As String
    
    With vs放疗
        strKEY = Trim(.EditText): strKEY = Replace(strKEY, Chr(vbKeyReturn), ""): strKEY = Replace(strKEY, Chr(10), "")
        Select Case Col
        Case .ColIndex("放射治疗编码")
        Case .ColIndex("开始日期")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "开始日期")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check日期有效性(strKEY, "开始日期") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("结束日期"))) <> "" Then
                If strKEY > Trim(.TextMatrix(Row, .ColIndex("结束日期"))) Then
                    MsgBox "开始日期不能大于结束日期,请检查!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            .EditText = strKEY
        Case .ColIndex("结束日期")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "结束日期")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check日期有效性(strKEY, "结束日期") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("开始日期"))) <> "" Then
                If strKEY < Trim(.TextMatrix(Row, .ColIndex("开始日期"))) Then
                    MsgBox "结束日期不能小于开始日期,请检查!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            .EditText = strKEY
        Case .ColIndex("设野部位")
            If strKEY = "" Then Exit Sub
            If zlCommFun.StrIsValid(strKEY, 50, 0, "设野部位") = False Then
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("放射剂量"), .ColIndex("累计量")
            If strKEY = "" Then Exit Sub
            If DblIsValid(strKEY, 10, True, False, 0, .ColKey(Col)) = False Then Cancel = True: Exit Sub
            If strKEY = "" Then Cancel = True: Exit Sub
            .EditText = strKEY
        Case .ColIndex("放疗效果")
        End Select
        mblnChange = True
        vs放疗.Tag = ""
    End With
End Sub
Private Function Load化疗与放疗(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载放疗与化疗信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-21 15:55:27
    '问题:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
    StrSQL = " " & _
    "   Select A.病人id, A.主页id, A.序号, A.疾病id, A.开始日期, A.结束日期, A.疗程数, A.总量, A.化疗方案, A.化疗效果, " & _
    "          B.编码 || '-' || B.名称 As 疾病信息 " & _
    "   From 病案化疗记录 A, 疾病编码目录 B " & _
    "   Where A.疾病id = B.Id And a.病人id=[1] And a.主页id=[2] " & _
    "   Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng病人ID, lng主页ID)
    With vs化疗
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("化学治疗编码")) = Nvl(rsTemp!疾病信息)
            .Cell(flexcpData, lngRow, .ColIndex("化学治疗编码")) = Nvl(rsTemp!疾病id)
            .TextMatrix(lngRow, .ColIndex("开始日期")) = Format(rsTemp!开始日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("结束日期")) = Format(rsTemp!结束日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("疗程数")) = Format(Val(Nvl(rsTemp!疗程数)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("总量")) = Format(Val(Nvl(rsTemp!总量)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("化疗方案")) = Trim(Nvl(rsTemp!化疗方案))
            .TextMatrix(lngRow, .ColIndex("化疗效果")) = Trim(Nvl(rsTemp!化疗效果))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    StrSQL = " " & _
    "   Select A.病人id, A.主页id, A.序号, A.疾病id, A.开始日期, A.结束日期,A.设野部位, A.放射剂量, A.累计量, A.放疗效果, " & _
    "          B.编码 || '-' || B.名称 As 疾病信息 " & _
    "   From 病案放疗记录 A, 疾病编码目录 B " & _
    "   Where A.疾病id = B.Id And a.病人id=[1] And a.主页id=[2] " & _
    "   Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng病人ID, lng主页ID)
    With vs放疗
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("放射治疗编码")) = Nvl(rsTemp!疾病信息)
            .Cell(flexcpData, lngRow, .ColIndex("放射治疗编码")) = Nvl(rsTemp!疾病id)
            .TextMatrix(lngRow, .ColIndex("开始日期")) = Format(rsTemp!开始日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("结束日期")) = Format(rsTemp!结束日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("放射剂量")) = Format(Val(Nvl(rsTemp!放射剂量)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("累计量")) = Format(Val(Nvl(rsTemp!累计量)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("设野部位")) = Trim(Nvl(rsTemp!设野部位))
            .TextMatrix(lngRow, .ColIndex("放疗效果")) = Trim(Nvl(rsTemp!放疗效果))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Load化疗与放疗 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vs化疗_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '设置相关的格式
    '刘兴宏:2007/09/17
    '--------------------------------------------------------------------------------
    With vs化疗
        Select Case Col
        Case .ColIndex("化学治疗编码")
            '.ColComboList(Col) = "..."
             If .ComboIndex < 0 Then Exit Sub
            .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
        End Select
    End With
End Sub
 

Private Sub vs化疗_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
  Call zl_VsGridRowChange(vs化疗, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs化疗_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs化疗
        Select Case Col
        Case .ColIndex("化学治疗编码"), .ColIndex("化疗方案")
        Case .ColIndex("开始日期"), .ColIndex("结束日期")
        Case .ColIndex("疗程数"), .ColIndex("总量")
        Case .ColIndex("化疗效果")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vs化疗_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '功能:按钮选择
    '参数:
    '--------------------------------------------------------------------------
    With vs化疗
        Select Case Col
'        Case .ColIndex("化学治疗编码")
'            If Select化疗与放疗("", True) = False Then
'                Exit Sub
'            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs化疗_GotFocus()
    Call zl_VsGridGotFocus(vs化疗)
End Sub

Private Sub vs化疗_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lngCol As Long, lngRow As Long, strKEY As String
   If mbln护士站 Or mblnReadOnly Then Exit Sub
    With vs化疗
        If (.Col = .ColIndex("化学治疗编码")) And KeyCode <> vbKeyReturn Then
           ' .ColComboList(.Col) = ""
        End If
        If KeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行的化疗信息吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                    .RowData(.Row) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            zlCtlSetFocus vs化疗, True
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs化疗
        If Val(.Cell(flexcpData, .Row, .ColIndex("化学治疗编码"))) = 0 Or Trim(.TextMatrix(.Row, .ColIndex("开始日期"))) = "" Then
            zlCtlSetFocus vs放疗, True
            Exit Sub
        End If
        Select Case .Col
        Case .Cols - 1
            If Not .Row >= .Rows - 1 Then
                .Col = 0
                .Row = .Row + 1
            Else
                Call vs化疗_KeyDownEdit(.Row, .Col, KeyCode, Shift)
            End If
            .SetFocus
        Case Else
            zlCommFun.PressKey vbKeyRight
        End Select
    End With
End Sub

Private Sub vs化疗_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer, lngRow As Long
    Dim strKEY As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs化疗
        Select Case Col
        Case .ColIndex("化学治疗编码")
'            strKey = Trim(.EditText)
'            strKey = Replace(strKey, Chr(vbKeyReturn), "")
'            strKey = Replace(strKey, Chr(10), "")
'            If strKey = "" Then Exit Sub
'            If Select化疗与放疗(strKey, True) = False Then
'                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
'                For intCol = 0 To .Cols - 1
'                    If intCol <> Col Then
'                        .TextMatrix(Row, intCol) = ""
'                        .Cell(flexcpData, Row, intCol) = ""
'                    End If
'                Next
'                Exit Sub
'            End If
'            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vs化疗, .ColIndex("开始日期"), .Cols - 1, True, lngRow)
        If lngRow > 0 Then
            '表示新增加了一行,需要设置相关的缺省值
            strKEY = .ColData(.ColIndex("化学治疗编码"))
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngRow, .ColIndex("化学治疗编码")) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngRow, .ColIndex("化学治疗编码")) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
                .TextMatrix(lngRow, .ColIndex("疗程数")) = 1
            End If
        End If
    End With
     
End Sub

Private Sub vs化疗_KeyPress(KeyAscii As Integer)
    If mbln护士站 Or mblnReadOnly Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs化疗_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vs化疗
        Select Case Col
        Case .ColIndex("化学治疗编码"), .ColIndex("化疗方案")
            Call VsFlxGridCheckKeyPress(vs化疗, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("开始日期"), .ColIndex("结束日期")
            Call VsFlxGridCheckKeyPress(vs化疗, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("疗程数"), .ColIndex("总量")
            Call VsFlxGridCheckKeyPress(vs化疗, Row, Col, KeyAscii, m金额式)
        Case .ColIndex("化疗效果")
        Case Else
        End Select
    End With
End Sub

Private Sub vs化疗_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs化疗)
End Sub

Private Sub vs化疗_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKEY As String
    Dim intCol As Integer
    Dim strTemp As String
    
    With vs化疗
        strKEY = Trim(.EditText): strKEY = Replace(strKEY, Chr(vbKeyReturn), ""): strKEY = Replace(strKEY, Chr(10), "")
        Select Case Col
        Case .ColIndex("化学治疗编码")
        Case .ColIndex("开始日期")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "开始日期")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check日期有效性(strKEY, "开始日期") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("结束日期"))) <> "" Then
                If strKEY > Trim(.TextMatrix(Row, .ColIndex("结束日期"))) Then
                    MsgBox "开始日期不能大于结束日期,请检查!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            .EditText = strKEY
        Case .ColIndex("结束日期")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "结束日期")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check日期有效性(strKEY, "结束日期") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("开始日期"))) <> "" Then
                If strKEY < Trim(.TextMatrix(Row, .ColIndex("开始日期"))) Then
                    MsgBox "结束日期不能小于开始日期,请检查!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            .EditText = strKEY
        Case .ColIndex("化疗方案")
            If strKEY = "" Then Exit Sub
            If zlCommFun.StrIsValid(strKEY, 50, 0, "化疗方案") = False Then
                
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("疗程数")
            If strKEY = "" Then Exit Sub
            If DblIsValid(strKEY, 3, True, False, 0, .ColKey(Col)) = False Then Cancel = True: Exit Sub
            If strKEY = "" Then Cancel = True: Exit Sub
            .EditText = strKEY
        Case .ColIndex("总量")
            If strKEY = "" Then Exit Sub
            If DblIsValid(strKEY, 10, True, False, 0, .ColKey(Col)) = False Then Cancel = True: Exit Sub
            If strKEY = "" Then Cancel = True: Exit Sub
            .EditText = strKEY
        End Select
        mblnChange = True
        vs化疗.Tag = ""
    End With
End Sub

Private Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '功能：离开网格控件时选择的颜色
    '入参：CustomColor-是否用自定义颜色来设置(BackColor)的方式来进行)
    '编制：刘兴洪
    '日期：2010-03-23 11:03:05
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
             If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
    End With
End Sub

Private Sub zl_VsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngoldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：行列改变时,设置相关的颜色
    '入参：CustomColor-自定义颜色
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-23 11:22:38
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    '行改变时
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

Private Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：进入网格控件时选择的颜色
    '入参：CustomColor-自定颜色
    '编制：刘兴洪
    '日期：2010-03-23 10:52:23
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '进入控件
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '清除选择颜色
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zl_VsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub

Private Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub

Private Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng主例 As Long = -1, Optional lng尾列 As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '功能:移动单元格的列
    '入参:blnEdit-当前正处于编辑状态,允许新增行
    '     lng主例-主列,如果<0,则主列为0列,否则为指定的列
    '     lng尾列-尾列,如果<0,则主列为.cols-1,否则为指定的列
    '出参:lngRow-如果存在插入行,则返回被插入的行号,否则返回-1
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    
    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
    If lng主例 <> -1 Then
        lngCol = lng主例
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lng尾列 < 0, vsGrid.Cols - 1, lng尾列)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
Errhand:
End Sub

Private Sub VsFlxGridCheckKeyPress(ByVal objCtl As Object, Row As Long, Col As Long, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '功能:只能输入数字和回车及退格
    '参数:
    '   objctl:Vsgrid8.0控件
    '   Keyascii:
    '           Keyascii:8 (退格)
    '   Row-当前行
    '   Col-当前列
    '   TextType:(0-文本式;1-数字式;2-金额式)
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    
    If TextType = m文本式 Then
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        Select Case KeyAscii
        Case vbKeyReturn       '回车
        Case 8                 '退格
        Case Asc(".")
            If TextType = m金额式 Or TextType = m负金额式 Then
                If InStr(objCtl.EditText, ".") <> 0 Then     '只能存在一个小数点
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        Case Asc("-")          '负数
            Dim iRow As Long
            Dim icol As Long
            If Trim(objCtl.EditText) = "" Then Exit Sub
            If TextType <> m负金额式 Then KeyAscii = 0: Exit Sub
            If objCtl.EditSelStart <> 0 Then KeyAscii = 0: Exit Sub      '光标不存第一位,不能输入负数
            If InStr(1, objCtl.EditText, "-") <> 0 Then   '只能存在一个负数
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    KeyAscii = 0
End Sub

Private Function CheckIsDate(ByVal strKEY As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '功能:检查是否合法的日期型,可以为:(20070101或2007-01-01)或则(01-01或0101)或则(01<01-31>)
    '参数:strKey-需要检查的关建字
    '返回:合法的日期,返回标准格式(yyyy-mm-dd),否则返回""
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKEY) = 4 And InStr(1, strKEY, "-") = 0 Then
        '0101,需要再前面加年
        strKEY = Year(Now) & strKEY
    ElseIf Len(Replace(strKEY, "-", "")) = 4 And InStr(1, strKEY, "-") > 0 Then
        '01-01形式,需要补零
        strKEY = Year(Now) & Replace(strKEY, "-", "")
    ElseIf Len(strKEY) <= 2 And IsNumeric(strKEY) Then
        '指是日
        strKEY = Format(Now, "YYYYMM") & IIf(Len(strKEY) = 2, strKEY, "0" & strKEY)
    End If
    If Len(strKEY) = 8 And InStr(1, strKEY, "-") = 0 Then
        strKEY = TranNumToDate(strKEY)
        If strKEY = "" Then
            MsgBox strTittle & "必须为日期型,请检查！", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    If Not IsDate(strKEY) Then
        MsgBox strTittle & "必须为日期型如(2000-10-10) 或（20001010）,请检查！", vbInformation, Me.Caption
        Exit Function
    End If
    CheckIsDate = strKEY
End Function

Private Function Check日期有效性(ByVal strDate As String, ByVal strTittle As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查日期的有效性
    '入参:strDate-当前日期
    '     strTittle-标题:如:放疗在第几行
    '出参:
    '返回:有效或strDate="",返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-21 17:03:30
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strCurDate As String
    Dim str入院时间 As String, str出院时间 As String
    
    If strDate = "" Then Check日期有效性 = True: Exit Function
    '检查日期是否合法
    If IsDate(strDate) = False Or IsNumeric(strDate) Then
        MsgBox strTittle & "不是一个有效的日期范围,请检查!", vbInformation, Me.Caption
        Exit Function
    End If
    str入院时间 = Format(txtInfo(txt入院时间).Text, "yyyy-mm-dd")
    If txtInfo(txt出院时间).Text <> "" Then str出院时间 = Format(txtInfo(txt出院时间).Text, "yyyy-mm-dd")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    If strDate > strCurDate Then
        MsgBox strTittle & "比当前日期还要大,请检查!", vbInformation, Me.Caption
        Exit Function
    End If
    
    If strDate < str入院时间 Then
        MsgBox strTittle & "比入院日期还要小,请检查!", vbInformation, Me.Caption
        Exit Function
    End If
    If str出院时间 <> "" Then
        If str出院时间 < strDate Then
            MsgBox strTittle & "比出院日期还要大,请检查!", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    Check日期有效性 = True
End Function

Private Function DblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     blnNegative     是否进行负数检查
    '     blnZero         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
   
    Dim dblValue As Double
    If blnZero = True Then
        If strInput = "" Then
            MsgBox str项目 & "未输入，请检查!", vbInformation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If strInput = "" Then DblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    
    If blnZero = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    DblIsValid = True
End Function

Private Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '功能:判断控件是否可
    '返回:初如成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:插入行
    '参数:vsGrid-插入行的网格格件
    '     lngRow-当前行
    '     blnBefor-在lngrow之间或之后.true:之间,false-之后
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
        Else
            .AddItem "", lngRow + 1
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'转换数值为日期
Private Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    If blnDec Then strDate = DateAdd("d", -1, Format(strDate, "yyyy-mm-dd"))
    TranNumToDate = strDate
End Function

Private Sub FillVsf()
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = "select 名称,内容 from 病案项目 order by 编码"
    vsfMain.Clear
    
    Call zlDatabase.OpenRecordset(rsTemp, StrSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then vsfMain.Rows = 1: vsfMain.Cols = 1: Exit Sub
    If (rsTemp.RecordCount Mod 2) <> 0 Then
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 2
    Else
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 1
    End If
    With vsfMain
        .Cols = 6
        For lngRow = 0 To 3 Step 3
            .TextMatrix(0, lngRow) = "项目"
            .TextMatrix(0, lngRow + 1) = "内容"
            .TextMatrix(0, lngRow + 2) = "说明"
            .ColWidth(0 + lngRow) = 1500
            .ColWidth(1 + lngRow) = 1200
            .ColHidden(2 + lngRow) = True
        Next lngRow
        .Cell(flexcpAlignment, 0, 0, 0, vsfMain.Cols - 1) = 4
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &HFCE7D8
        .Cell(flexcpBackColor, 1, 3, .Rows - 1, 3) = &HFCE7D8
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
    End With
    lngRow = 1
    lngCol = 0
    While Not rsTemp.EOF
        If lngCol < 4 Then
            With vsfMain
                .TextMatrix(lngRow, lngCol + 0) = rsTemp!名称
                .TextMatrix(lngRow, lngCol + 2) = rsTemp!内容 & ""
                If InStr(rsTemp!内容, "是否") > 0 Then
                    vsfMain.TextMatrix(lngRow, lngCol + 1) = "是"
                    vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = 2
                End If
            End With
            lngCol = lngCol + 3
            rsTemp.MoveNext
        Else
            lngCol = 0
            lngRow = lngRow + 1
        End If
    Wend
    vsfMain.Editable = flexEDKbdMouse
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SavePageDataUnit(ByRef blnDiagnose As Boolean, ByVal blnBeforSign As Boolean) As Boolean

'功能：检查保存于一体的首页保存方法
'参数：blnBeforSign-是否签名时保存前调用
'返回：blnDiagnose=是否填写了诊断
'返回：SavePageDataUnit=保存成功

    If Not CheckPageData(blnDiagnose, blnBeforSign) Then Exit Function
    
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("要求的诊断信息还没有输入，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
    End If
    
    If Not SavePageData(blnBeforSign) Then Exit Function
    '设置界面可用性
    Call SetFaceEditable(mblnReadOnly)
    
    SavePageDataUnit = True
    
    
End Function
