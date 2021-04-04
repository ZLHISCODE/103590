VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClinicItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊疗项目编辑"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   Icon            =   "frmClinicItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.CheckBox chkGoOn 
      Caption         =   "连续增加诊疗项目"
      Height          =   180
      Left            =   5880
      TabIndex        =   106
      Top             =   7785
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   2040
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   7440
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
   Begin TabDlg.SSTab stbInfo 
      Height          =   7020
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   555
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   12383
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "项目属性(&B)"
      TabPicture(0)   =   "frmClinicItem.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblComment"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgComment"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl项目编码"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl项目名称"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl执行频率"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl适用性别"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl计算方式"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl名称简码"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl操作类型"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl其他别名"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl别名简码"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl计算单位"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl英文"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl分类说明"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl执行分类"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl计算规则"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl病理类别"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblML"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbllel"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl输液类型"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblZLPL"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl试管编码"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cboZLPL"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cbo病理类别"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "fra标本部位"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "fra录入量"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "fra检查部位"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "fra标准编码"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chk检验组合"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt参考"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cbo操作类型"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txt项目编码"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txt项目名称"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cbo适用性别"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cbo计算方式"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt名称拼音"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt名称五笔"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt其他别名"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "chk单独应用"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "chk执行安排"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt别名拼音"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt别名五笔"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt计算单位"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt录入限量"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cbo录入限量范围"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt英文"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cbo分类说明"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cbo执行频率"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "picFound"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "chk加收"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Frame3"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtML"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cbo输液类型"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cbo计算规则"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "vsfBloodLis"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmd参考"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmdDel参考"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "picTestTubeCode"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cbo执行分类"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cboBloodType"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "chkNoTMSY"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "chkYYPS"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).ControlCount=   67
      TabCaption(1)   =   "执行科室(&E)"
      TabPicture(1)   =   "frmClinicItem.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDeptFind"
      Tab(1).Control(1)=   "picDept"
      Tab(1).Control(2)=   "fra执行部门"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "检查部位(&L)"
      TabPicture(2)   =   "frmClinicItem.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optList(0)"
      Tab(2).Control(1)=   "optList(1)"
      Tab(2).Control(2)=   "vfgList"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "皮试结果(&P)"
      TabPicture(3)   =   "frmClinicItem.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra皮试结果"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "频率设置(&R)"
      TabPicture(4)   =   "frmClinicItem.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblFreq"
      Tab(4).Control(1)=   "vsfFreq"
      Tab(4).ControlCount=   2
      Begin VB.CheckBox chkYYPS 
         Caption         =   "源液皮试"
         Height          =   270
         Left            =   8775
         TabIndex        =   155
         Top             =   3075
         Width           =   1100
      End
      Begin VB.CheckBox chkNoTMSY 
         Caption         =   "不允许脱敏使用"
         Height          =   180
         Left            =   7080
         TabIndex        =   154
         Top             =   3075
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   300
         Left            =   5115
         Style           =   2  'Dropdown List
         TabIndex        =   147
         Top             =   2295
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.ComboBox cbo执行分类 
         Height          =   300
         Left            =   5130
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   2290
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picTestTubeCode 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5115
         ScaleHeight     =   300
         ScaleWidth      =   1785
         TabIndex        =   150
         Top             =   2280
         Width           =   1785
         Begin VB.PictureBox picTubeColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1515
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   152
            Top             =   15
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.ComboBox cboTestTubeCode 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   15
            Width           =   1500
         End
      End
      Begin VB.CommandButton cmdDel参考 
         Height          =   285
         Left            =   6650
         Picture         =   "frmClinicItem.frx":0616
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   3360
         Width           =   285
      End
      Begin VB.CommandButton cmd参考 
         Caption         =   "…"
         Height          =   285
         Left            =   6360
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   3360
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBloodLis 
         Height          =   1000
         Left            =   7080
         TabIndex        =   144
         Top             =   3090
         Visible         =   0   'False
         Width           =   2760
         _cx             =   4860
         _cy             =   1764
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicItem.frx":09D9
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VB.ComboBox cbo计算规则 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   3030
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cbo输液类型 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   2295
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtML 
         Height          =   300
         Left            =   7080
         TabIndex        =   138
         Top             =   2660
         Width           =   675
      End
      Begin VB.Frame Frame3 
         Caption         =   "适用范围"
         Height          =   2800
         Left            =   120
         TabIndex        =   119
         Top             =   4100
         Width           =   9735
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1080
            TabIndex        =   128
            Top             =   2400
            Width           =   8385
            Begin VB.OptionButton OptAppUse 
               Caption         =   "应用于本项"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   132
               Top             =   60
               Value           =   -1  'True
               Width           =   1725
            End
            Begin VB.OptionButton OptAppUse 
               Caption         =   "应用于同级"
               Height          =   225
               Index           =   1
               Left            =   1920
               TabIndex        =   131
               Top             =   60
               Width           =   1605
            End
            Begin VB.OptionButton OptAppUse 
               Caption         =   "应用于分类下所有"
               Height          =   225
               Index           =   2
               Left            =   3720
               TabIndex        =   130
               Top             =   60
               Width           =   2235
            End
            Begin VB.OptionButton OptAppUse 
               Caption         =   "应用于当前类别"
               Height          =   225
               Index           =   3
               Left            =   6120
               TabIndex        =   129
               Top             =   60
               Width           =   2055
            End
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   5340
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   240
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox chk服务对象 
            Caption         =   "体检(&P)"
            Height          =   225
            Index           =   2
            Left            =   3060
            TabIndex        =   122
            Top             =   285
            Width           =   930
         End
         Begin VB.CheckBox chk服务对象 
            Caption         =   "住院(&I)"
            Height          =   225
            Index           =   1
            Left            =   2100
            TabIndex        =   121
            Top             =   285
            Value           =   1  'Checked
            Width           =   930
         End
         Begin VB.CheckBox chk服务对象 
            Caption         =   "门诊(&W)"
            Height          =   225
            Index           =   0
            Left            =   1140
            TabIndex        =   120
            Top             =   285
            Value           =   1  'Checked
            Width           =   930
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   1800
            Left            =   1080
            TabIndex        =   124
            Top             =   600
            Width           =   8565
            _cx             =   15108
            _cy             =   3175
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClinicItem.frx":0A31
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
         Begin VB.Label Label9 
            Caption         =   "使用科室"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   2460
            Width           =   855
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "院区编号(&Z)"
            Height          =   180
            Left            =   4320
            TabIndex        =   127
            Top             =   300
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label7 
            Caption         =   "服务范围"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "使用科室"
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Top             =   660
            Width           =   855
         End
      End
      Begin VB.CheckBox chk加收 
         Caption         =   "允许床旁或术中执行(&J)"
         Height          =   240
         Left            =   7200
         TabIndex        =   118
         Top             =   3825
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Frame fraDeptFind 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   -72970
         TabIndex        =   103
         Top             =   1750
         Width           =   5475
         Begin VB.TextBox txtLocate 
            Height          =   320
            Left            =   4380
            TabIndex        =   111
            ToolTipText     =   "查找下一个F3或回车，定位输入框F4"
            Top             =   57
            Width           =   1000
         End
         Begin VB.OptionButton optDeptKind 
            Caption         =   "执行科室(&E)"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   110
            Top             =   30
            Value           =   -1  'True
            Width           =   1300
         End
         Begin VB.OptionButton optDeptKind 
            Caption         =   "病人科室(&B)"
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   109
            Top             =   30
            Width           =   1300
         End
         Begin VB.Label lblX 
            Caption         =   "/"
            Height          =   195
            Left            =   2040
            TabIndex        =   137
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblAn 
            Caption         =   "按"
            Height          =   195
            Left            =   480
            TabIndex        =   136
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblLocate 
            Caption         =   "查找(&F)"
            Height          =   195
            Left            =   3540
            TabIndex        =   112
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.PictureBox picFound 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6360
         ScaleHeight     =   210
         ScaleWidth      =   3570
         TabIndex        =   101
         Top             =   0
         Width           =   3570
         Begin VB.Label lblFound 
            AutoSize        =   -1  'True
            Caption         =   "该项目于2002-12-20建立"
            Height          =   180
            Left            =   1560
            TabIndex        =   102
            Top             =   0
            Width           =   1980
         End
      End
      Begin VB.PictureBox picDept 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   -69840
         ScaleHeight     =   3105
         ScaleWidth      =   4890
         TabIndex        =   96
         Top             =   3120
         Visible         =   0   'False
         Width           =   4920
         Begin VB.CommandButton cmdFindCancle 
            Caption         =   "取消"
            Height          =   270
            Left            =   4200
            TabIndex        =   117
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdFindOk 
            Caption         =   "确定"
            Height          =   270
            Left            =   3480
            TabIndex        =   116
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "查找"
            Height          =   270
            Left            =   1740
            TabIndex        =   108
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFind 
            Height          =   270
            Left            =   50
            TabIndex        =   107
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox ChkSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "全选"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2115
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   88
            Width           =   675
         End
         Begin VB.ComboBox cboProperty 
            Height          =   300
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   45
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwItems 
            Height          =   2280
            Left            =   75
            TabIndex        =   99
            Top             =   795
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   4022
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
            BackColor       =   &H00E0E0E0&
            Caption         =   "工作性质"
            Height          =   180
            Left            =   50
            TabIndex        =   100
            Top             =   110
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   6135
         Left            =   -74910
         TabIndex        =   72
         Top             =   720
         Width           =   9795
         _cx             =   17277
         _cy             =   10821
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.ComboBox cbo执行频率 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   2290
         Width           =   2115
      End
      Begin VB.ComboBox cbo分类说明 
         Height          =   300
         Left            =   8160
         TabIndex        =   87
         Text            =   "cbo分类说明"
         Top             =   3780
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fra皮试结果 
         Caption         =   "皮试结果"
         Height          =   6375
         Left            =   -74760
         TabIndex        =   77
         Top             =   480
         Width           =   9495
         Begin VB.CommandButton cmdTestDel 
            Caption         =   "删除(&D)"
            Height          =   350
            Left            =   6480
            TabIndex        =   84
            Top             =   2280
            Width           =   1100
         End
         Begin VB.CommandButton cmdTestAdd 
            Caption         =   "增加(&A)"
            Height          =   350
            Left            =   5280
            TabIndex        =   83
            Top             =   2280
            Width           =   1100
         End
         Begin VB.CheckBox chk皮试过敏 
            Caption         =   "过敏"
            Height          =   180
            Left            =   6000
            TabIndex        =   82
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txt皮试文字 
            Height          =   300
            Left            =   6000
            MaxLength       =   13
            TabIndex        =   81
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txt皮试标注 
            Height          =   300
            Left            =   6000
            MaxLength       =   8
            TabIndex        =   79
            Top             =   540
            Width           =   1515
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfTest 
            Height          =   5775
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   4755
            _cx             =   8387
            _cy             =   10186
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
            AllowUserFreezing=   1
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   8
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "文字(&W)"
            Height          =   180
            Left            =   5280
            TabIndex        =   80
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lbl标注 
            AutoSize        =   -1  'True
            Caption         =   "标注(&B)"
            Height          =   180
            Left            =   5280
            TabIndex        =   78
            Top             =   600
            Width           =   630
         End
      End
      Begin VB.TextBox txt英文 
         Height          =   300
         Left            =   7845
         MaxLength       =   12
         TabIndex        =   75
         Top             =   3405
         Width           =   1980
      End
      Begin VB.OptionButton optList 
         Caption         =   "仅显示已选择部位"
         Height          =   375
         Index           =   1
         Left            =   -66960
         TabIndex        =   74
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optList 
         Caption         =   "显示所有部位"
         Height          =   375
         Index           =   0
         Left            =   -68760
         TabIndex        =   73
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox cbo录入限量范围 
         Height          =   300
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3765
         Width           =   1650
      End
      Begin VB.TextBox txt录入限量 
         Height          =   300
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   31
         Top             =   3765
         Width           =   2115
      End
      Begin VB.TextBox txt计算单位 
         Height          =   300
         Left            =   5115
         TabIndex        =   25
         Top             =   2660
         Width           =   1785
      End
      Begin VB.TextBox txt别名五笔 
         Height          =   300
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   19
         Top             =   1940
         Width           =   2115
      End
      Begin VB.TextBox txt别名拼音 
         Height          =   300
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   18
         Top             =   1940
         Width           =   2115
      End
      Begin VB.CheckBox chk执行安排 
         Caption         =   "需要执行安排(&W)"
         Height          =   210
         Left            =   3465
         TabIndex        =   28
         Top             =   3075
         Width           =   1740
      End
      Begin VB.CheckBox chk单独应用 
         Caption         =   "允许单独应用(&Y)"
         Height          =   210
         Left            =   5250
         TabIndex        =   21
         Top             =   3075
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.TextBox txt其他别名 
         Height          =   300
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1560
         Width           =   5715
      End
      Begin VB.TextBox txt名称五笔 
         Height          =   300
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1215
         Width           =   2115
      End
      Begin VB.TextBox txt名称拼音 
         Height          =   300
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1215
         Width           =   2115
      End
      Begin VB.ComboBox cbo计算方式 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2660
         Width           =   2115
      End
      Begin VB.ComboBox cbo适用性别 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3015
         Width           =   2115
      End
      Begin VB.TextBox txt项目名称 
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   820
         Width           =   5715
      End
      Begin VB.TextBox txt项目编码 
         Height          =   300
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   7
         Top             =   465
         Width           =   2115
      End
      Begin VB.ComboBox cbo操作类型 
         Height          =   300
         ItemData        =   "frmClinicItem.frx":0BDD
         Left            =   5085
         List            =   "frmClinicItem.frx":0BDF
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txt参考 
         Height          =   300
         Left            =   1200
         TabIndex        =   30
         Top             =   3377
         Width           =   5460
      End
      Begin VB.CheckBox chk检验组合 
         Caption         =   "组合检验项目"
         Height          =   210
         Left            =   7860
         TabIndex        =   60
         Top             =   3825
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame fra执行部门 
         BorderStyle     =   0  'None
         Caption         =   "执行科室"
         Height          =   6405
         Left            =   -74760
         TabIndex        =   41
         Top             =   480
         Width           =   9795
         Begin VSFlex8Ctl.VSFlexGrid msf定向执行 
            Height          =   4455
            Left            =   180
            TabIndex        =   115
            Top             =   1680
            Width           =   9405
            _cx             =   16589
            _cy             =   7858
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
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
         Begin VB.OptionButton opt执行部门 
            Caption         =   "开单人所在科室(&6)"
            Height          =   195
            Index           =   6
            Left            =   4500
            TabIndex        =   66
            Top             =   555
            Width           =   1860
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   120
            TabIndex        =   61
            Top             =   6135
            Width           =   8985
            Begin VB.OptionButton OptApp 
               Caption         =   "应用于当前类别"
               Height          =   225
               Index           =   3
               Left            =   6240
               TabIndex        =   65
               Top             =   60
               Width           =   2415
            End
            Begin VB.OptionButton OptApp 
               Caption         =   "应用于分类下所有"
               Height          =   225
               Index           =   2
               Left            =   3840
               TabIndex        =   64
               Top             =   60
               Width           =   2235
            End
            Begin VB.OptionButton OptApp 
               Caption         =   "应用于同级"
               Height          =   225
               Index           =   1
               Left            =   1920
               TabIndex        =   63
               Top             =   60
               Width           =   1605
            End
            Begin VB.OptionButton OptApp 
               Caption         =   "应用于本项"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   62
               Top             =   60
               Value           =   -1  'True
               Width           =   1605
            End
         End
         Begin VB.Frame Frame1 
            Height          =   120
            Left            =   120
            TabIndex        =   59
            Top             =   780
            Width           =   9550
         End
         Begin VB.TextBox txt住院执行 
            Enabled         =   0   'False
            Height          =   300
            Left            =   7725
            MaxLength       =   30
            TabIndex        =   50
            Top             =   940
            Width           =   1860
         End
         Begin VB.TextBox txt门诊执行 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3750
            MaxLength       =   30
            TabIndex        =   49
            Top             =   940
            Width           =   1890
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "医院外执行(&5)"
            Height          =   180
            Index           =   5
            Left            =   6660
            TabIndex        =   47
            Top             =   280
            Width           =   1485
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "指定科室执行(&4)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   4
            Left            =   2385
            TabIndex        =   46
            Top             =   555
            Width           =   1665
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "操作员所在科室(&3)"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   45
            Top             =   555
            Width           =   2025
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "由病人病区执行(&2)"
            Height          =   180
            Index           =   2
            Left            =   4500
            TabIndex        =   44
            Top             =   280
            Width           =   1845
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "由病人科室执行(&1)"
            Height          =   180
            Index           =   1
            Left            =   2385
            TabIndex        =   43
            Top             =   280
            Width           =   1845
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "不跟踪执行的叮嘱(&0)"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   280
            Width           =   2025
         End
         Begin VB.Label lbl住院执行 
            AutoSize        =   -1  'True
            Caption         =   "住院病人执行科室"
            Height          =   180
            Left            =   6210
            TabIndex        =   58
            Top             =   1005
            Width           =   1440
         End
         Begin VB.Label lbl一般情况 
            AutoSize        =   -1  'True
            Caption         =   "1、除指定病人科室外："
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   90
            TabIndex        =   57
            Top             =   1000
            Width           =   1890
         End
         Begin VB.Label lbl定向执行 
            AutoSize        =   -1  'True
            Caption         =   "2、指定病人科室："
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   90
            TabIndex        =   51
            Top             =   1380
            Width           =   1530
         End
         Begin VB.Label lbl门诊执行 
            AutoSize        =   -1  'True
            Caption         =   "门诊病人执行科室"
            Height          =   180
            Left            =   2250
            TabIndex        =   48
            Top             =   1005
            Width           =   1440
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFreq 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   91
         Top             =   720
         Width           =   9735
         _cx             =   17171
         _cy             =   10821
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Frame fra标准编码 
         Caption         =   "手术标准编码"
         Height          =   1400
         Left            =   7395
         TabIndex        =   33
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.CommandButton cmd选择 
            Caption         =   "…"
            Height          =   285
            Left            =   1720
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox txt标准编码 
            Height          =   300
            Left            =   135
            TabIndex        =   34
            Top             =   255
            Width           =   1875
         End
         Begin VB.Label lbl标准编码 
            Height          =   600
            Left            =   180
            TabIndex        =   35
            Top             =   615
            Width           =   1800
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fra检查部位 
         Height          =   1380
         Left            =   7395
         TabIndex        =   36
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.TextBox txt检查部位 
            Height          =   300
            Left            =   390
            MaxLength       =   40
            TabIndex        =   38
            Top             =   435
            Width           =   1665
         End
         Begin VB.OptionButton opt检查部位 
            Caption         =   "可选多部位检查(&X)"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   795
            Width           =   1980
         End
         Begin VB.OptionButton opt检查部位 
            Caption         =   "固定单部位检查(&G)"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   165
            Value           =   -1  'True
            Width           =   1980
         End
      End
      Begin VB.Frame fra录入量 
         Caption         =   "记录入出量"
         Height          =   1395
         Left            =   7380
         TabIndex        =   93
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.CheckBox chk尿量 
            Caption         =   "尿量(&E)"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fra标本部位 
         Caption         =   "默认标本部位"
         Height          =   1350
         Left            =   7380
         TabIndex        =   69
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.CommandButton cmd标本 
            Caption         =   "…"
            Height          =   285
            Left            =   1740
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   263
            Width           =   285
         End
         Begin VB.TextBox txt标本部位 
            Height          =   300
            Left            =   135
            TabIndex        =   70
            Top             =   255
            Width           =   1575
         End
      End
      Begin VB.ComboBox cbo病理类别 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   450
         Width           =   1575
      End
      Begin VB.ComboBox cboZLPL 
         Height          =   300
         Left            =   5115
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   2280
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lbl试管编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "试管编码(&C)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   153
         Top             =   2340
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblZLPL 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗频率(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   146
         Top             =   2340
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl输液类型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "输液类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7200
         TabIndex        =   141
         Top             =   2355
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbllel 
         Caption         =   "="
         Height          =   180
         Left            =   6960
         TabIndex        =   140
         Top             =   2720
         Width           =   135
      End
      Begin VB.Label lblML 
         Caption         =   "毫升"
         Height          =   180
         Left            =   7800
         TabIndex        =   139
         Top             =   2720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "(五笔)"
         Height          =   255
         Left            =   6195
         TabIndex        =   114
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "(五笔)"
         Height          =   255
         Left            =   6195
         TabIndex        =   113
         Top             =   1963
         Width           =   615
      End
      Begin VB.Label lbl病理类别 
         Caption         =   "号别名称"
         Height          =   180
         Left            =   7200
         TabIndex        =   104
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lbl计算规则 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "计算规则(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7080
         TabIndex        =   95
         Top             =   3090
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblFreq 
         AutoSize        =   -1  'True
         Caption         =   "请选择该诊疗项目的执行频率，在第一列打勾。"
         Height          =   180
         Left            =   -74880
         TabIndex        =   92
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label lbl执行分类 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "执行分类(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4065
         TabIndex        =   89
         Top             =   2350
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl分类说明 
         AutoSize        =   -1  'True
         Caption         =   "分类说明"
         Height          =   180
         Left            =   7200
         TabIndex        =   86
         Top             =   3840
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl英文 
         AutoSize        =   -1  'True
         Caption         =   "英文缩写"
         Height          =   255
         Left            =   7080
         TabIndex        =   76
         Top             =   3450
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "录入限量应用于"
         Height          =   180
         Left            =   3480
         TabIndex        =   68
         Top             =   3825
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "录入限量(&X)"
         Height          =   180
         Left            =   180
         TabIndex        =   67
         Top             =   3825
         Width           =   990
      End
      Begin VB.Label lbl计算单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "计算单位(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4080
         TabIndex        =   24
         Top             =   2720
         Width           =   990
      End
      Begin VB.Label lbl别名简码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "别名简码(&N)                        (拼音)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   2000
         Width           =   3690
      End
      Begin VB.Label lbl其他别名 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其他别名(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   15
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label lbl操作类型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "操作类型(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4080
         TabIndex        =   8
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lbl名称简码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称简码(&S)                        (拼音)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   12
         Top             =   1275
         Width           =   3690
      End
      Begin VB.Label lbl计算方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "计算方式(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   22
         Top             =   2720
         Width           =   990
      End
      Begin VB.Label lbl适用性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "适用性别(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   26
         Top             =   3075
         Width           =   990
      End
      Begin VB.Label lbl执行频率 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "执行频率(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   20
         Top             =   2350
         Width           =   990
      End
      Begin VB.Label lbl项目名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   880
         Width           =   990
      End
      Begin VB.Label lbl项目编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "项目编码(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   6
         Top             =   525
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "参考项目(&F)"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   3400
         Width           =   1095
      End
      Begin VB.Image imgComment 
         Height          =   240
         Left            =   7305
         Picture         =   "frmClinicItem.frx":0BE1
         Top             =   1845
         Width           =   240
      End
      Begin VB.Label lblComment 
         Caption         =   $"frmClinicItem.frx":1A23
         Height          =   1545
         Left            =   7305
         TabIndex        =   40
         Top             =   1905
         Width           =   2430
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9025
      TabIndex        =   53
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   105
      Picture         =   "frmClinicItem.frx":1AC7
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   7695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7920
      TabIndex        =   52
      Top             =   7680
      Width           =   1100
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1185
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   75
      Width           =   5010
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   6225
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   285
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      ItemData        =   "frmClinicItem.frx":1C11
      Left            =   8310
      List            =   "frmClinicItem.frx":1C13
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   60
      Width           =   1710
   End
   Begin VB.Frame fraLine 
      Height          =   120
      Left            =   -240
      TabIndex        =   56
      Top             =   360
      Width           =   10290
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8160
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":1C15
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":21AF
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":2749
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":2CE3
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":327D
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2280
      Left            =   1320
      TabIndex        =   134
      Top             =   7680
      Visible         =   0   'False
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4022
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
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "项目分类(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
   Begin VB.Label lbl类别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "诊疗类别(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7230
      TabIndex        =   3
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmClinicItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、权限、编辑项目的分类ID、ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"增加"、"修改"、"查阅"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private lngClassId As Long       '被编辑的分类ID，上级程序通过ShowMe传递进入
Private lngItemID As Long        '被编辑的项目ID，修改、查阅时由上级程序通过ShowMe传递进入,增加时为0，
Private lngVItemID As Long       '检验项目相关的指标ID（存放在诊治所见项中）
Private mlngOldId As Long

Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer
Dim mstrMatch As String, strRefer As String '参考名称
Dim mbln组合部位项目 As Boolean             '是否是组合项目的部位项目
Dim mbln组合项目 As Boolean                 '是否是组合项目（含子项目）
Private mlng简码长度 As Long
Private mLast操作类型 As String
Private mFromLoad As Boolean                '是否第一次调用
Private mbln连续增加 As Boolean             '是否连续增加
Private mblnIniTest As Boolean
Private mstr已选执行科室 As String
Private mstr已选使用科室 As String
Private mblnRefresh As Boolean
Private mrs性质分类 As ADODB.Recordset
Private mstrPageCaption     '用来记录上次的页中的标题
Private mblnOK As Boolean
Private mlngFind As Long
Private mstrFindStyle As String '匹配方式
Private mstrOldBlood As String  '记录修改前输血检验记录列表中的值
Private mblnPACSInterface As Boolean        '启用影像信息系统接口
Private mstr应用范围 As String
Private Enum 执行科室COL
    col病人科室ID = 0
    col病人科室 = 1
    col执行科室ID = 2
    col执行科室 = 3
End Enum

Private Sub Ini性质分类()
    '取部门性质分类，如果已经提取了则退出
    On Error GoTo ErrHandle
    If Not mrs性质分类 Is Nothing Then
        mrs性质分类.Filter = ""
        If Not mrs性质分类.EOF Then
            Exit Sub
        End If
    End If
    
    gstrSql = "Select 名称,服务病人 From 部门性质分类"
    Set mrs性质分类 = zlDatabase.OpenSQLRecord(gstrSql, "取部门性质分类")
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init诊疗频率(Optional ByVal strNo As String)
    Dim rsTemp As Recordset
    Dim strTemp As String
    Dim intIndex As Integer
    Dim i As Integer
    
    '取诊疗频率项目中间隔单位为小时或者分的项目
    On Error GoTo ErrHandle
    
    gstrSql = "select 编码,名称 from 诊疗频率项目 where 间隔单位='小时' or 间隔单位='分钟'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取诊疗频率项目")
    With Me.cboZLPL
        .Clear
        .AddItem ""
        strTemp = "|"
        
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            strTemp = strTemp & rsTemp!名称 & "-" & rsTemp!编码 & "|"
            i = i + 1
            If strNo = rsTemp!编码 Then
                intIndex = i
            End If
            rsTemp.MoveNext
        Loop
        
        .ListIndex = intIndex
        Me.lblZLPL.Tag = strTemp
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub
Private Sub load性质分类(ByVal intType As Integer)
    'intType:0-执行科室（所有性质，服务于病人）；1-病人科室（临床性质）;使用科室(根据服务范围和站点：临床、检查、检验、手术、治疗、体检性质的科室)
    
    mblnRefresh = True
    
    With cboProperty
        .Clear
        
        If mrs性质分类 Is Nothing Then Exit Sub
        
        If intType = 0 Then
            mrs性质分类.Filter = "服务病人=1 Or 服务病人=2 Or 服务病人=3"
        ElseIf intType = 1 Then
            mrs性质分类.Filter = "名称='临床'"
        ElseIf intType = 2 Then
            mrs性质分类.Filter = "名称='临床' Or 名称='检查' Or 名称='检验' Or 名称='手术' Or 名称='治疗'" & IIf(chk服务对象(2).Value = 1, " Or 名称='体检'", "")
        End If
        
        If mrs性质分类.RecordCount = 0 Then Exit Sub
        
        If intType = 0 Or intType = 2 Then
            .AddItem "所有性质"
            
            Do While Not mrs性质分类.EOF
                .AddItem mrs性质分类!名称
                
                mrs性质分类.MoveNext
            Loop
        ElseIf intType = 1 Then
            .AddItem "临床"
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    DoEvents
    
    mblnRefresh = False
End Sub
Private Sub Load部门(ByVal intType As Integer, ByVal str工作性质 As String)
    'intType:0-执行科室（所有性质，服务于病人）；1-病人科室（临床性质）;2-使用科室
    Dim rsData As ADODB.Recordset
    Dim objItem As ListItem
    Dim strTmp As String
    Dim str站点 As String
    
    On Error GoTo ErrHandle
    If intType = 1 Then
        gstrSql = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and 工作性质=[1] " & _
                "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " order by 编码"
    ElseIf intType = 0 Then
        gstrSql = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and T.服务对象 in (1,2,3) " & _
                " and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
                
        If str工作性质 <> "所有性质" Then
            gstrSql = gstrSql & " and 工作性质=[1] "
        End If
                
        gstrSql = gstrSql & " order by 编码"
    ElseIf intType = 2 Then
        If chk服务对象(1).Value = 1 Then strTmp = " T.服务对象=2"
        If chk服务对象(2).Value = 1 Or chk服务对象(0).Value = 1 Then strTmp = strTmp & IIf(strTmp = "", "", " Or") & " T.服务对象=1"
        If strTmp <> "" Then strTmp = " And (" & strTmp & " Or T.服务对象=3)"
        gstrSql = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID " & _
                " and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) And t.服务对象<>0 " & strTmp
        If cmbStationNo.Text <> "" Then
            str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
            gstrSql = gstrSql & " And (D.站点=[2] Or D.站点 is Null)"
        End If
                
        If str工作性质 <> "所有性质" Then
            gstrSql = gstrSql & " and 工作性质=[1] "
        Else
            gstrSql = gstrSql & " and 工作性质 In('临床','检查','检验','手术','治疗'" & IIf(chk服务对象(2).Value = 1, ",'体检'", "") & ") "
        End If
                
        gstrSql = gstrSql & " order by 编码"
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str工作性质, str站点)
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsData.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsData!ID, rsData!名称)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsData!编码
        objItem.Checked = False
        If Me.lvwItems.Tag = "开单" Then
            If InStr(Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) & ",", rsData!ID & ",") > 0 Then
                objItem.Checked = True
            End If
        End If
        
        If Me.lvwItems.Tag = "执行" Then
            If InStr(mstr已选执行科室, rsData!ID & "," & rsData!名称) > 0 Then
                objItem.Checked = True
            End If
        End If
        
        If Me.lvwItems.Tag = "使用" Then
            If InStr(mstr已选使用科室, rsData!ID & "," & rsData!名称) > 0 Then
                objItem.Checked = True
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

Private Sub Load试管编码(Optional ByVal strCode As String = "")
    Dim rsTemp As Recordset
    Dim strTmp As String
    Dim i As Integer, intIndex As Integer
    '取部门性质分类，如果已经提取了则退出
    On Error GoTo ErrHandle
    With Me.cboTestTubeCode
        If strCode = "" Then
            If .ListCount > 0 Then strTmp = .Text
        Else
            strTmp = strCode
        End If
        If .ListCount <= 1 Then
            gstrSql = "Select 编码 || '-' || 名称 名称,颜色  From 采血管类型"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "采血管类型")
            .Clear
            .AddItem "<未设置>": .ItemData(0) = 0
            Do While Not rsTemp.EOF
                i = i + 1
                .AddItem rsTemp!名称
                .ItemData(.NewIndex) = Val(rsTemp!颜色)
                If Split(rsTemp!名称, "-")(0) = strTmp Or rsTemp!名称 = strTmp Then
                    intIndex = i
                End If
                rsTemp.MoveNext
            Loop
            .ListIndex = intIndex
        Else
            For i = 1 To .ListCount - 1
                If Split(.List(i), "-")(0) = strTmp Or .List(i) = strTmp Then
                    intIndex = i
                    Exit For
                End If
            Next
            .ListIndex = intIndex
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboProperty_Click()
    If picDept.Tag = "2" Then
        Load部门 2, cboProperty.Text
    Else
        If Me.msf定向执行.Col = col病人科室 Then
            Load部门 1, cboProperty.Text
        Else
            Load部门 0, cboProperty.Text
        End If
    End If
    
    ChkSelect.Value = 0
End Sub

Private Sub cboTestTubeCode_Click()
    If cboTestTubeCode.ListIndex > 0 And cboTestTubeCode.ListIndex < cboTestTubeCode.ListCount - 1 Then
        picTubeColor.Visible = True
        picTubeColor.BackColor = Val(cboTestTubeCode.ItemData(cboTestTubeCode.ListIndex))
    Else
        picTubeColor.Visible = False
        picTubeColor.BackColor = picTestTubeCode.BackColor
    End If
End Sub

Private Sub cboTestTubeCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo录入限量范围_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo执行分类_Click()
    Me.lbl输液类型.Visible = False
    Me.cbo输液类型.Visible = False
    
    If Left(Me.cbo执行分类.Text, 1) = "1" Then      '输液
        Me.lbl输液类型.Visible = True
        Me.cbo输液类型.Visible = True
    End If
End Sub


Private Sub ChkSelect_Click()
    Dim i As Integer
    Dim str名称 As String
    
    On Error GoTo ErrHandle
    If mblnRefresh = True Then Exit Sub
    
    If ChkSelect.Value = 2 Then Exit Sub
    Call SetSelect(lvwItems, ChkSelect.Value)
    
    If cboProperty.Text = "所有性质" Then
        If lvwItems.Tag = "执行" Then
            mstr已选执行科室 = ""
        ElseIf lvwItems.Tag = "使用" Then
            mstr已选使用科室 = ""
        End If
    End If
    
    If ChkSelect.Value = 1 Then
        '当前性质全选
        For i = 1 To lvwItems.ListItems.Count
            str名称 = Mid(lvwItems.ListItems(i).Key, 2) & "," & lvwItems.ListItems(i).Text
            
            If InStr(mstr已选执行科室, str名称) = 0 Or cboProperty.Text = "所有性质" Then
                If lvwItems.Tag = "执行" Then
                    mstr已选执行科室 = IIf(mstr已选执行科室 = "", "", mstr已选执行科室 & ";") & str名称
                ElseIf lvwItems.Tag = "使用" Then
                    mstr已选使用科室 = IIf(mstr已选使用科室 = "", "", mstr已选使用科室 & ";") & str名称
                End If
            End If
        Next
    ElseIf cboProperty.Text <> "所有性质" Then
        '当前性质全清

        For i = 1 To lvwItems.ListItems.Count
            str名称 = Mid(lvwItems.ListItems(i).Key, 2) & "," & "[" & lvwItems.ListItems(i).SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) & "]" & lvwItems.ListItems(i).Text
            If lvwItems.Tag = "执行" Then
                If InStr(mstr已选执行科室, str名称) > 0 Then
                    mstr已选执行科室 = Replace(mstr已选执行科室, str名称, "")
                End If
            ElseIf lvwItems.Tag = "使用" Then
                If InStr(mstr已选使用科室, str名称) > 0 Then
                    mstr已选使用科室 = Replace(mstr已选使用科室, str名称, "")
                End If
            End If
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdDel参考_Click()
    Me.txt参考.Text = ""
    Me.txt参考.Tag = ""
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To lvwItems.ListItems.Count
        If zlCommFun.SpellCode(Mid(lvwItems.ListItems(i).Text, InStr(lvwItems.ListItems(i).Text, "-") + 1)) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(lvwItems.ListItems(i).Text) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Then
            lvwItems.ListItems(i).Selected = True
            lvwItems.ListItems(i).EnsureVisible
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

Private Sub cmdFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picDept.Visible = False
        txtFind.Text = ""
    End If
End Sub

Private Sub cmdFindCancle_Click()
    Call lvwItems_KeyPress(vbKeyEscape)
End Sub

Private Sub cmdFindOk_Click()
    Call lvwItems_DblClick
End Sub

Private Sub cmd选择_Click()
    Dim rsTemp As Recordset

    On Error GoTo ErrHand
    gstrSql = "select A.ID,A.编码,A.手术类型 手术类型,A.名称,A.简码" & _
            " from 疾病编码目录 A" & _
            " where A.类别='S' and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)

    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到指定手术标准编码", vbExclamation, gstrSysName
            Me.txt标准编码.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt标准编码.Tag = !ID
            Me.txt标准编码.Text = IIf(IsNull(!编码), "", !编码)
            Me.lbl标准编码.Caption = IIf(IsNull(!手术类型), "", "【" & NVL(!手术类型) & "】") & IIf(IsNull(!名称), "", !名称)
            Me.stbInfo.Tab = 1: Me.chk服务对象(0).SetFocus
            Exit Sub
        End If

        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !名称, "expend", "expend")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItem.ColumnHeaders("类别").Index - 1) = NVL(!手术类型)
            .MoveNext
        Loop
        With Me.lvwItem
            .ListItems(1).Selected = True
            .Tag = "手术"
            .Left = Me.stbInfo.Left + Me.fra标准编码.Left + Me.fra标准编码.Width - .Width
            .Top = Me.stbInfo.Top + Me.fra标准编码.Top + Me.txt标准编码.Top + Me.txt标准编码.Height
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim str名称 As String
    
    If Me.lvwItems.Tag = "执行" Then
        str名称 = Mid(Item.Key, 2) & "," & Item.Text
        
        If Item.Checked = True Then
            If InStr(mstr已选执行科室, str名称) = 0 Then
                mstr已选执行科室 = IIf(mstr已选执行科室 = "", "", mstr已选执行科室 & ";") & str名称
            End If
        Else
            If InStr(mstr已选执行科室, str名称) > 0 Then
                mstr已选执行科室 = Replace(mstr已选执行科室, str名称, "")
            End If
        End If
    ElseIf Me.lvwItems.Tag = "使用" Then
        str名称 = Mid(Item.Key, 2) & "," & Item.Text
        
        If Item.Checked = True Then
            If InStr(mstr已选使用科室, str名称) = 0 Then
                mstr已选使用科室 = IIf(mstr已选使用科室 = "", "", mstr已选使用科室 & ";") & str名称
            End If
        Else
            If InStr(mstr已选使用科室, str名称) > 0 Then
                mstr已选使用科室 = Replace(mstr已选使用科室, str名称, "")
            End If
        End If
    End If
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.lvwItems.Tag = "开单" Or Me.lvwItems.Tag = "执行" Or Me.lvwItems.Tag = "使用" Then
        If KeyCode = vbKeyA And Shift = vbCtrlMask Then '全选 Ctrl+A
            If Me.lvwItems.Tag = "执行" Or Me.lvwItems.Tag = "使用" Then
                If Me.ChkSelect.Value = 0 Then
                    Me.ChkSelect.Value = 1
                    Call SetSelect(lvwItems, True)
                End If
            Else
                Call SetSelect(lvwItems, True)
            End If
        End If
        
        If KeyCode = vbKeyR And Shift = vbCtrlMask Then     '全消 Ctrl+R
            If Me.lvwItems.Tag = "执行" Or Me.lvwItems.Tag = "使用" Then
                If Me.ChkSelect.Value = 1 Then
                    Me.ChkSelect.Value = 0
                    Call SetSelect(lvwItems, False)
                End If
            Else
                Call SetSelect(lvwItems, False)
            End If
        End If
    End If
End Sub
Private Sub cboProperty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
         txtFind.Text = ""
    End Select
End Sub

Private Sub cboProperty_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub msf定向执行_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If msf定向执行.Editable = flexEDNone Then
        msf定向执行.FocusRect = flexFocusLight
        msf定向执行.ComboList = ""
    Else
        msf定向执行.FocusRect = flexFocusSolid
        msf定向执行.ComboList = "..."
    End If
End Sub

Private Sub msf定向执行_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    msf定向执行.AutoSize msf定向执行.FixedCols, msf定向执行.Cols - 1
End Sub

Private Sub msf定向执行_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If msf定向执行.TextMatrix(NewRow, OldCol) <> msf定向执行.Cell(flexcpData, NewRow, OldCol) Then
        msf定向执行.TextMatrix(NewRow, OldCol) = msf定向执行.Cell(flexcpData, NewRow, OldCol)
    End If
End Sub

Private Sub msf定向执行_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim i As Integer
    
    mstr已选执行科室 = ""
    If Me.msf定向执行.Col = col执行科室 Then
        With Me.msf定向执行
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col执行科室) <> "" Then
                    mstr已选执行科室 = IIf(mstr已选执行科室 = "", "", mstr已选执行科室 & ";") & .TextMatrix(i, col执行科室ID) & "," & .TextMatrix(i, col执行科室)
                End If
            Next
        End With
    End If
    
    With Me.picDept
        If Me.msf定向执行.Col = col病人科室 Then
            .Tag = ""
            Me.lvwItems.Tag = "开单"
            .Left = Me.fra执行部门.Left + Me.msf定向执行.Left + Me.msf定向执行.ColWidth(col执行科室ID)
            .Width = IIf(Me.msf定向执行.ColWidth(col病人科室) < 3000, 3000, Me.msf定向执行.ColWidth(col病人科室))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "执行"
            .Left = Me.fra执行部门.Left + Me.msf定向执行.Left + Me.msf定向执行.ColWidth(col执行科室ID) + Me.msf定向执行.ColWidth(col病人科室) + Me.msf定向执行.ColWidth(col病人科室ID)
            .Width = IIf(Me.msf定向执行.ColWidth(col病人科室ID) < 5000, 5000, Me.msf定向执行.ColWidth(col病人科室ID))
            If .Left > Me.Width - .Width - stbInfo.Left - Me.fra执行部门.Left - Me.msf定向执行.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.fra执行部门.Left - Me.msf定向执行.Left
        End If
        
        .Top = 50
        .Height = Me.fra执行部门.Top + Me.msf定向执行.Top + (IIf(Me.msf定向执行.Row > 14, 14, Me.msf定向执行.Row) - Me.msf定向执行.FixedRows + 1) * Me.msf定向执行.RowHeight(col病人科室)
        
        lbl工作性质.Visible = (Me.msf定向执行.Col = col执行科室)
        cboProperty.Visible = lbl工作性质.Visible
        ChkSelect.Visible = lbl工作性质.Visible
        
        If Me.lvwItems.Tag = "执行" Then
            lbl工作性质.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        If .Tag = "执行" Then
            .Left = lbl工作性质.Left
            .Top = cboProperty.Top + cboProperty.Height + 50 + txtFind.Height + 50
            .Width = Me.picDept.Width - .Left - 50
            .Height = Me.picDept.Height - .Top - 10
            txtFind.Top = cboProperty.Top + cboProperty.Height + 50
            cmdFind.Top = cboProperty.Top + cboProperty.Height + 50
            cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
            cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
            cmdFindOk.Top = cmdFind.Top
            cmdFindCancle.Top = cmdFind.Top
        Else
            .Left = 0
            .Top = txtFind.Height + 100
            .Width = Me.picDept.Width
            .Height = Me.picDept.Height - txtFind.Height - 50 - 50
            txtFind.Top = 50
            cmdFind.Top = 50
            cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
            cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
            cmdFindOk.Top = cmdFind.Top
            cmdFindCancle.Top = cmdFind.Top
        End If
        
        .SetFocus
        .Refresh
    End With
    
    If Me.msf定向执行.Col = col病人科室 Then
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

Private Sub msf定向执行_EnterCell()
    strInputed = Me.msf定向执行.TextMatrix(msf定向执行.Row, msf定向执行.Col)
End Sub

Private Sub msf定向执行_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call msf定向执行_KeyPress(KeyCode)
    ElseIf KeyCode = vbKeyDelete Then
        If msf定向执行.TextMatrix(msf定向执行.Row, msf定向执行.Col) <> "" Then
            If MsgBox("您确定要删除这一行数据？", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption) = vbYes Then
                If msf定向执行.Rows <= 2 Then
                    msf定向执行.Cell(flexcpText, 1, col病人科室ID, 1, col执行科室) = ""
                    msf定向执行.Cell(flexcpData, 1, col病人科室ID, 1, col执行科室) = ""
                Else
                    msf定向执行.RemoveItem msf定向执行.Row
                End If
            End If
        End If
    End If
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf定向执行
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = col执行科室 And .TextMatrix(.Row, col执行科室) = "" Then
            If .Row = 1 Then Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If .Col = col病人科室 And .TextMatrix(.Row, col病人科室) = "" Then
            .TextMatrix(.Row, col病人科室) = "（所有部门）"
            .TextMatrix(.Row, col病人科室ID) = "（所有部门）"
            Exit Sub
        End If
    End With
End Sub

Private Sub msf定向执行_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf定向执行
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = col执行科室 And Trim(.EditText) = "" Then
            If .Row = 1 Then .SetFocus: Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strTemp = UCase(Trim(.EditText))
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    If strTemp = "" Then Exit Sub
    
    err = 0: On Error GoTo ErrHand

    If Me.msf定向执行.Col = col病人科室 Then
        gstrSql = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and 工作性质='临床'" & _
                "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
                " order by 编码"
    Else
        gstrSql = "select distinct ID,编码,名称" & _
                " from 部门表 D,部门性质说明 T" & _
                " where D.ID=T.部门ID and T.服务对象 in (1,2,3)" & _
                "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
                " order by 编码"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%")

    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName
            Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, msf定向执行.Col) = msf定向执行.Cell(flexcpData, msf定向执行.Row, msf定向执行.Col)
            msf定向执行.EditText = msf定向执行.Cell(flexcpData, msf定向执行.Row, msf定向执行.Col)
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msf定向执行.Text = !名称
            If Me.msf定向执行.Col = col执行科室 Then
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col执行科室ID) = !ID
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col执行科室) = Me.msf定向执行.Text
                msf定向执行.EditText = Me.msf定向执行.Text
                msf定向执行.Cell(flexcpData, msf定向执行.Row, msf定向执行.Col) = Me.msf定向执行.Text
            Else
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) = !ID
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室) = Me.msf定向执行.Text
                msf定向执行.EditText = Me.msf定向执行.Text
                msf定向执行.Cell(flexcpData, msf定向执行.Row, msf定向执行.Col) = Me.msf定向执行.Text
            End If
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = (Me.msf定向执行.Col = col病人科室)
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码

            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        If Me.msf定向执行.Col = col病人科室 Then
            .Tag = ""
            Me.lvwItems.Tag = "开单"
            .Left = Me.fra执行部门.Left + Me.msf定向执行.Left
            .Width = IIf(Me.msf定向执行.ColWidth(col病人科室) < 3000, 3000, Me.msf定向执行.ColWidth(col病人科室))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "执行"
            .Left = Me.fra执行部门.Left + Me.msf定向执行.Left + Me.msf定向执行.ColWidth(col执行科室ID) + Me.msf定向执行.ColWidth(col病人科室) + Me.msf定向执行.ColWidth(col病人科室ID)
            .Width = IIf(Me.msf定向执行.ColWidth(col病人科室ID) < 3000, 3000, Me.msf定向执行.ColWidth(col病人科室ID))
            If .Left > Me.Width - .Width - stbInfo.Left - Me.fra执行部门.Left - Me.msf定向执行.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.fra执行部门.Left - Me.msf定向执行.Left
        End If
        
        .Top = 50
        .Height = Me.fra执行部门.Top + Me.msf定向执行.Top + (IIf(Me.msf定向执行.Row > 14, 14, Me.msf定向执行.Row) - Me.msf定向执行.FixedRows + 1) * Me.msf定向执行.RowHeight(col病人科室)
        
        lbl工作性质.Visible = False
        cboProperty.Visible = lbl工作性质.Visible
        ChkSelect.Visible = lbl工作性质.Visible
        
        If Me.msf定向执行.Col = col执行科室 Then
            lbl工作性质.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        txtFind.Visible = False
        cmdFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
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
    KeyCode = 0
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msf定向执行_KeyPress(KeyAscii As Integer)
    If msf定向执行.Editable = flexEDNone Then Exit Sub

    With msf定向执行
        If KeyAscii = 13 Then
            KeyAscii = 0
            If .Col = col执行科室 Then
                If .Row = .Rows - 1 Then
                    If (.TextMatrix(.Row, col执行科室) <> "" Or .TextMatrix(.Row, col病人科室) <> "") Then
                        .Rows = .Rows + 1
                    Else
                        zlCommFun.PressKey (vbKeyTab)
                        Exit Sub
                    End If
                End If
                .Row = .Row + 1
                .Col = col病人科室
            ElseIf .Col = col病人科室 Then
                .Col = col执行科室
            End If
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call msf定向执行_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub msf定向执行_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If msf定向执行.Editable = flexEDNone Then
        msf定向执行.FocusRect = flexFocusLight
        msf定向执行.ComboList = ""
    Else
        msf定向执行.FocusRect = flexFocusSolid
        msf定向执行.ComboList = "..."
    End If
End Sub

Private Sub msf定向执行_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call msf定向执行_KeyDownEdit(Row, Col, vbKeyReturn, 0)
End Sub

Private Sub OptApp_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To OptApp.UBound
        If i = Index Then
            OptApp(i).FontBold = True
        Else
            OptApp(i).FontBold = False
        End If
    Next
End Sub

Private Sub OptAppUse_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To OptAppUse.UBound
        If i = Index Then
            OptAppUse(i).FontBold = True
        Else
            OptAppUse(i).FontBold = False
        End If
    Next
End Sub

Private Sub optDeptKind_Click(Index As Integer)
    lblLocate.Tag = ""
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDOKDEPT,CMDCANCELDEPT,LVWITEMS,CBOPROPERTY,PICDEPT,CHKSELECT,TXTFIND,CMDFIND,CMDFINDOK", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
    If Me.lvwItems.Tag = "使用" Then
        vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col) = vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col)
        vsUseDept.AutoSize vsUseDept.FixedCols, vsUseDept.Cols - 1
        Call vsUseDept.SetFocus
    Else
        msf定向执行.TextMatrix(msf定向执行.Row, msf定向执行.Col) = msf定向执行.Cell(flexcpData, msf定向执行.Row, msf定向执行.Col)
        msf定向执行.AutoSize msf定向执行.FixedCols, msf定向执行.Cols - 1
        Call msf定向执行.SetFocus
    End If
    txtFind.Text = ""
    mlngFind = 1
End Sub
Private Sub ChkSelect_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
         txtFind.Text = ""
    End Select
End Sub
Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select A.编码, A.计算单位, B.简码, B.名称 其他别名 From 诊疗项目目录 A, 诊疗项目别名 B " & _
            " Where A.ID = B.诊疗项目id And A.ID = 0 And B.码类 = 1"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)

    mlng简码长度 = rsTmp.Fields("简码").DefinedSize

    txt项目编码.MaxLength = rsTmp.Fields("编码").DefinedSize
    txt计算单位.MaxLength = rsTmp.Fields("计算单位").DefinedSize
    txt名称拼音.MaxLength = mlng简码长度
    txt名称五笔.MaxLength = mlng简码长度
    txt别名拼音.MaxLength = mlng简码长度
    txt别名五笔.MaxLength = mlng简码长度
    txt其他别名.MaxLength = rsTmp.Fields("其他别名").DefinedSize
    txt项目名称.MaxLength = rsTmp.Fields("其他别名").DefinedSize
    txt英文.MaxLength = 40
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsfFreq()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With vsfFreq
        '初始化表格
        .Clear
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = 7
        
        .RowHeightMin = 300
        
        .TextMatrix(0, 0) = "选择"
        .TextMatrix(0, 1) = "编码"
        .TextMatrix(0, 2) = "名称"
        .TextMatrix(0, 3) = "英文名称"
        .TextMatrix(0, 4) = "频率次数"
        .TextMatrix(0, 5) = "频率间隔"
        .TextMatrix(0, 6) = "间隔单位"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 2000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        
        .Editable = flexEDNone
        
        '提取数据,填充表格
        gstrSql = "Select A.项目id, B.编码, B.名称, B.英文名称, B.频率次数, B.频率间隔, B.间隔单位 " & _
            " From (Select 项目id, 频次 From 诊疗用法用量 Where 项目id = [1]) A, 诊疗频率项目 B " & _
            " Where A.频次(+) = B.编码 And B.适用范围 = 1 " & _
            " Order By A.项目id, B.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption & " 诊疗项目频率", lngItemID)
        
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = IIf(IsNull(rsTmp!项目id), "", "√")
            .TextMatrix(.Rows - 1, 1) = rsTmp!编码
            .TextMatrix(.Rows - 1, 2) = rsTmp!名称
            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(rsTmp!英文名称), "", rsTmp!英文名称)
            .TextMatrix(.Rows - 1, 4) = rsTmp!频率次数
            .TextMatrix(.Rows - 1, 5) = rsTmp!频率间隔
            .TextMatrix(.Rows - 1, 6) = IIf(IsNull(rsTmp!间隔单位), "", rsTmp!间隔单位)
            rsTmp.MoveNext
        Loop
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InivsfTest()
    Dim rsTemp As ADODB.Recordset
    Dim strTest As String
    Dim strTemp As String
    Dim strArr
    Dim n As Integer
    Const strDefault As String = "阳性(+);阴性(-)"
    
    On Error GoTo ErrHandle
    With vsfTest
        .Clear
        .Cols = 3
        .Rows = 1
        
        .TextMatrix(0, 0) = "标注"
        .TextMatrix(0, 1) = "文字"
        .TextMatrix(0, 2) = "过敏"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 1500
        .ColWidth(2) = 800
        
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        
        strTest = strDefault
        
        If lngItemID > 0 Then
            gstrSql = "Select 标本部位 From 诊疗项目目录 Where ID = [1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取皮试结果", lngItemID)

            If rsTemp.RecordCount > 0 Then
                strTemp = IIf(IsNull(rsTemp!标本部位), "", rsTemp!标本部位)
                
                If strTemp <> "" And InStrB(strTemp, ";") > 0 Then
                    strTest = strTemp
                End If
            End If
        End If
        
        '阴性结果
        strArr = Split(Split(strTest, ";")(1), ",")
        For n = 0 To UBound(strArr)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = MidB(strArr(n), InStrB(strArr(n), "("))
            .TextMatrix(.Rows - 1, 1) = MidB(strArr(n), 1, InStrB(strArr(n), "(") - 1)
        Next
        
        '阳性结果
        strArr = Split(Split(strTest, ";")(0), ",")
        For n = 0 To UBound(strArr)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = MidB(strArr(n), InStrB(strArr(n), "("))
            .TextMatrix(.Rows - 1, 1) = MidB(strArr(n), 1, InStrB(strArr(n), "(") - 1)
            .TextMatrix(.Rows - 1, 2) = "√"
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowMe(ByVal frmParent As Object, ByVal byt状态 As Byte, ByVal lng分类id As Long, Optional ByVal lng项目id As Long) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.Tag = Switch(byt状态 = 0, "增加", byt状态 = 1, "修改", byt状态 = 2, "查阅", byt状态 = 3, "复制增加")
    lngClassId = lng分类id: lngItemID = lng项目id: lngVItemID = 0
    mlngOldId = lng项目id
    
    '填写需要选择的数据
    aryTemp = Split("0-可选频率;1-一次性;2-持续性", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo执行频率.AddItem aryTemp(intCount)
    Next
    Me.cbo执行频率.ListIndex = 0

    aryTemp = Split("0-不明确;1-计量;2-计时;3-计次", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo计算方式.AddItem aryTemp(intCount)
    Next
    Me.cbo计算方式.ListIndex = 0

    aryTemp = Split("0-无性别区分;1-男性;2-女性", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo适用性别.AddItem aryTemp(intCount)
    Next
    Me.cbo适用性别.ListIndex = 0
    
    aryTemp = Split("0-正常计算;1-取整计算", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo计算规则.AddItem aryTemp(intCount)
    Next
    Me.cbo计算规则.ListIndex = 0

    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗分类目录" & _
            " Where 类型 = 5" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then MsgBox "请首先建立诊疗分类项目之后增加项目", vbExclamation, gstrSysName: Unload Me: Exit Function
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Nodes("_" & lng分类id).Selected = True
        Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With
    gstrSql = "select 编码||'-'||名称 from 诊疗项目类别 where 编码>'9' order by 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then MsgBox "诊疗项目类别数据丢失(请系统管理员处理)", vbExclamation, gstrSysName: Unload Me: Exit Function
        Me.cbo类别.Clear
        Do While Not .EOF
            Me.cbo类别.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo类别.ListCount > 0 Then Me.cbo类别.ListIndex = 0
    End With
    'Me.cbo项目类型.ListIndex = 0: Me.cbo结果类型.ListIndex = 0
    
    '取给药途径的分类说明
    With Me.cbo分类说明
        .Clear
        gstrSql = "Select Distinct 标本部位 From 诊疗项目目录 Where 类别 = 'E' And 操作类型 = '2' And 标本部位 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取给药途径的分类说明")
        
        If rsTemp.RecordCount > 0 Then
            Do While Not rsTemp.EOF
                .AddItem rsTemp.Fields(0).Value
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    '显示窗体
    Me.Show 1, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo操作类型_Click()
    Dim i As Long
    
    stbInfo.TabVisible(3) = False
    Me.lbl分类说明.Visible = False
    Me.cbo分类说明.Visible = False
    Me.fra录入量.Visible = False
    cbo执行分类.Visible = False
    cboBloodType.Visible = False
    lbl执行分类.Visible = False
    Me.lbl输液类型.Visible = False
    Me.cbo输液类型.Visible = False
    Me.picTestTubeCode.Visible = False
    Me.lbl试管编码.Visible = False
    Me.chkNoTMSY.Visible = False
    Me.chkYYPS.Visible = False
    
    imgComment.Top = lbl别名简码.Top + 50
    lblComment.Top = lbl别名简码.Top + 50
                
    If Left(Me.cbo类别.Text, 1) = "E" Then      '治疗
        Select Case Val(Left(Me.cbo操作类型.Text, 1))
        Case 0, 5     '0-普通;5-特殊治疗
            Me.chk单独应用.Enabled = True
            Me.cbo执行频率.Enabled = True
            Me.cbo计算方式.Enabled = True
            Me.opt执行部门(5).Enabled = True
        Case 1      '1-过敏试验
            Me.chk单独应用.Enabled = True
            Me.chkNoTMSY.Visible = True
            Me.chkYYPS.Visible = True
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
            If Me.opt执行部门(5).Value = True Then Me.opt执行部门(5).Value = False: Me.opt执行部门(2).Value = True
            Me.opt执行部门(5).Enabled = False
            stbInfo.TabVisible(3) = True
            Call InivsfTest
        Case 2, 3, 4, 6, 9  '2-给药方法(西药);3-中药煎法;4-中药用(服)法;6-标本采集;9-输血采集
            Me.chk单独应用.Value = 0: Me.chk单独应用.Enabled = False
            Me.cbo执行频率.ListIndex = 0: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
            If Me.opt执行部门(5).Value = True Then Me.opt执行部门(5).Value = False: Me.opt执行部门(2).Value = True
            Me.opt执行部门(5).Enabled = False
            Me.lbl分类说明.Visible = True
            Me.cbo分类说明.Visible = True
            
            If Val(Left(Me.cbo操作类型.Text, 1)) = 2 Then
                cbo执行分类.Visible = True
                lbl执行分类.Visible = True
                imgComment.Top = lbl名称简码.Top + 50
                lblComment.Top = lbl名称简码.Top + 50
            ElseIf Val(Left(Me.cbo操作类型.Text, 1)) = 9 Then
                Call Load试管编码
                Me.lbl试管编码.Visible = True
                Me.picTestTubeCode.Visible = True
                Me.cboTestTubeCode.Visible = True
            End If
            
            If cbo执行分类.Visible = True And Val(Left(Me.cbo执行分类.Text, 1)) = 1 Then
                Me.lbl输液类型.Visible = True
                Me.cbo输液类型.Visible = True
            End If
        Case 7 '7-配血方法
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
            Me.chk单独应用.Value = 0: Me.chk单独应用.Enabled = False
        Case 8 '8-输血途径
            Me.cbo执行频率.ListIndex = 0: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
            Me.chk单独应用.Enabled = True: Me.chk单独应用.Value = IIf(Val(Me.chk单独应用.Tag) = 1, 1, 0)
            lbl执行分类.Visible = True
            cboBloodType.Visible = True
        End Select
    End If
    
    If Left(Me.cbo类别.Text, 1) = "H" Then
        Select Case Val(Left(Me.cbo操作类型.Text, 1))
        Case 0
            Me.cbo执行频率.Enabled = True
        Case 1
            Me.cbo执行频率.ListIndex = 1
            Me.cbo执行频率.Enabled = False
        End Select
    End If

'    If mLast操作类型 <> cbo操作类型.Text Then
        If Left(Me.cbo类别.Text, 1) = "D" Then
            If cbo类别.Text = "D-检查" And cbo操作类型.Text = "18-病理" Then
                lbl病理类别.Visible = True
                cbo病理类别.Visible = True
                stbInfo.TabCaption(2) = "病理标本"
                optList(0).Caption = "显示所有标本"
                optList(1).Caption = "仅显示已选择标本"
            Else
                lbl病理类别.Visible = False
                cbo病理类别.Visible = False
                
                If cbo操作类型.Text <> "18-病理" Then
                    stbInfo.TabCaption(2) = "检查部位(&L)"
                    optList(0).Caption = "显示所有部位"
                    optList(1).Caption = "仅显示已选择部位"
                End If
            End If
            
            '检查项目 显示部位方法
            Call initVfgList
            
        End If
'    End If
    
    If Left(Me.cbo类别.Text, 1) = "Z" Then      '其他
        Me.cbo操作类型.Width = Me.fra标准编码.Left + Me.fra标准编码.Width - Me.cbo操作类型.Left
        Me.txt项目名称.Width = Me.fra标准编码.Left + Me.fra标准编码.Width - Me.txt项目名称.Left
    
        Select Case Val(Mid(Me.cbo操作类型.Text, 1, InStr(1, Me.cbo操作类型.Text, "-") - 1))
        Case 0      '0-普通
            Me.cbo执行频率.Enabled = True: Me.cbo计算方式.Enabled = True
            Me.chk服务对象(0).Enabled = True: Me.chk服务对象(1).Enabled = True: Me.chk服务对象(2).Enabled = True
            Me.vsUseDept.Editable = flexEDKbdMouse
            For i = 0 To OptAppUse.Count - 1
                If i = 0 Then
                    OptAppUse(i).Enabled = True
                Else
                    '根据参数来确定是否可用
                    OptAppUse(i).Enabled = (Val(Mid(mstr应用范围, i, 1)) = 1)
                End If
            Next
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = True
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                If intCount = 0 Then
                    Me.OptApp(intCount).Enabled = True
                Else
                    Me.OptApp(intCount).Enabled = (Val(Mid(mstr应用范围, intCount, 1)) = 1)
                End If
            Next
        Case 1, 2     '1-留观,2-住院
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 1: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 0: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.opt执行部门(1).Value = True
            vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 3, 5    '3-转科,5-出院
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.opt执行部门(1).Value = True
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 4, 14     '4-术后; 14-术前
            Me.cbo执行频率.ListIndex = 2: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.opt执行部门(1).Value = True
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 6   '6-转院
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 1: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.opt执行部门(1).Value = True
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 7   '7-会诊
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.opt执行部门(1).Value = True
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 8, 11
            Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.opt执行部门(1).Value = True
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 9, 10
            Me.cbo执行频率.ListIndex = 2: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.opt执行部门(1).Value = True
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 12
            Me.cbo操作类型.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.cbo操作类型.Left
            Me.txt项目名称.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.txt项目名称.Left

            Me.cbo执行频率.ListIndex = 2: Me.cbo执行频率.Enabled = False
            Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = ""
            Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
            Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
            Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.opt执行部门(2).Value = True
            Me.fra录入量.Visible = True
            For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
                Me.opt执行部门(intCount).Enabled = False
            Next
            Me.OptApp(0).Value = True
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        End Select
    End If
End Sub

Private Sub cbo操作类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo计算方式_Click()
    If cbo执行频率.ListIndex = 0 And cbo计算方式.ListIndex > 0 Then
        lbl计算规则.Visible = True
        cbo计算规则.Visible = True
    Else
        lbl计算规则.Visible = False
        cbo计算规则.Visible = False
    End If
End Sub

Private Sub cbo计算方式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo类别_Click()
    Dim i As Long
    
    Me.fra检查部位.Visible = False: Me.fra标准编码.Visible = False
    Me.fra标本部位.Visible = False: lbl英文.Visible = False: Me.txt英文.Visible = False
    Me.chk单独应用.Value = 1: Me.chk单独应用.Enabled = True
    Me.chk服务对象(0).Enabled = True: Me.chk服务对象(1).Enabled = True: Me.chk服务对象(2).Enabled = True
    Me.lbl分类说明.Visible = False
    Me.cbo分类说明.Visible = False
    Me.cbo执行分类.Visible = False
    Me.lbl执行分类.Visible = False
    Me.lbl输液类型.Visible = False
    Me.cbo输液类型.Visible = False
    Me.cboBloodType.Visible = False
    vsfBloodLis.Visible = False
    
    Me.vsUseDept.Editable = flexEDKbdMouse
    For i = 0 To OptAppUse.Count - 1
        If i = 0 Then
            OptAppUse(i).Enabled = True
        Else
            '根据参数来确定是否可用
            OptAppUse(i).Enabled = (Val(Mid(mstr应用范围, i, 1)) = 1)
        End If
    Next
    
    Me.imgComment.Top = Me.txt其他别名.Top + 250
    Me.lblComment.Top = Me.imgComment.Top + 50
    
    On Error GoTo ErrHand
    For intCount = Me.opt执行部门.LBound To Me.opt执行部门.UBound
        Me.opt执行部门(intCount).Enabled = True
    Next
    For intCount = Me.OptApp.LBound To Me.OptApp.UBound
        If intCount = 0 Then
            Me.OptApp(intCount).Enabled = True
        Else
            Me.OptApp(intCount).Enabled = (Val(Mid(mstr应用范围, intCount, 1)) = 1)
        End If
    Next
    Me.cbo操作类型.Width = Me.fra标准编码.Left + Me.fra标准编码.Width - Me.cbo操作类型.Left
    Me.txt项目名称.Width = Me.fra标准编码.Left + Me.fra标准编码.Width - Me.txt项目名称.Left
    Me.chk检验组合.Visible = False
    
    Me.stbInfo.TabVisible(2) = False '检查部位
    chk加收.Visible = False
    
    If cbo类别.Text = "D-检查" And cbo操作类型.Text = "18-病理" Then
        lbl病理类别.Visible = True
        cbo病理类别.Visible = True
        stbInfo.TabCaption(2) = "病理标本"
    Else
        lbl病理类别.Visible = False
        cbo病理类别.Visible = False
    End If
    
    Me.cbo执行频率.Clear
    Select Case Left(Me.cbo类别.Text, 1)
    Case "C", "D"
        aryTemp = Split("0-可选频率;1-一次性", ";")
    Case "H"
        aryTemp = Split("0-可选频率;2-持续性", ";")
    Case Else
        aryTemp = Split("0-可选频率;1-一次性;2-持续性", ";")
    End Select
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo执行频率.AddItem aryTemp(intCount)
    Next

    Select Case Left(Me.cbo类别.Text, 1)
    Case "C"        '检验
        Me.cbo操作类型.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.cbo操作类型.Left
        Me.txt项目名称.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.txt项目名称.Left
        Me.fra标本部位.Visible = True

        Me.chk检验组合.Visible = True
        Me.txt英文.Visible = True: lbl英文.Visible = True
        
        Me.lbl操作类型.Caption = "检验类型(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        err = 0: On Error GoTo ErrHand
        
        gstrSql = "select 编码||'-'||名称 from 诊疗检验类型 order by 编码"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo类别_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo操作类型.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo操作类型.ListCount > 0 Then Me.cbo操作类型.ListIndex = 0
        End With
        Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = True
        Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
        Me.imgComment.Top = Me.fra标本部位.Top + Me.fra标本部位.Height + 50
        Me.lblComment.Top = Me.imgComment.Top + 50
        Me.lblComment.Caption = Space(4) & "检验类项目只能是一次性项目，只允许在门诊和住院临时医嘱中使用；为有效执行，需要在检验项目管理中指定其标本和参考取值，对于组合项目还应设置其对应的基本指标项。"
        Me.chk服务对象(0).Value = 1: Me.chk服务对象(1).Value = 1
        Me.opt执行部门(0).Value = False: Me.opt执行部门(0).Enabled = False
        Me.opt执行部门(1).Value = True
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "D"        '检查
        Me.cbo操作类型.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.cbo操作类型.Left
        Me.txt项目名称.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.txt项目名称.Left
        Me.lbl操作类型.Caption = "检查类型(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        err = 0: On Error GoTo ErrHand
        gstrSql = "select 编码||'-'||名称 from 诊疗检查类型 order by 编码"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo类别_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo操作类型.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo操作类型.ListCount > 0 Then Me.cbo操作类型.ListIndex = 0
        End With
        Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = True
        Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
        Me.lblComment.Caption = Space(4) & "检查类项目只能是一次性项目，固定部位检查必须指明部位，可选部位项目需要通过部位构成程序指定其可选单部位项目才能使用。"
        'Me.fra检查部位.Visible = True
        Me.stbInfo.TabVisible(2) = True '检查部位
        stbInfo.TabCaption(2) = "检查部位"
        Me.opt执行部门(0).Value = False: Me.opt执行部门(0).Enabled = False
        Me.opt执行部门(1).Value = True
        chk加收.Visible = True
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "E"        '处置
        Me.lbl操作类型.Caption = "处置性质(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        aryTemp = Split("0-普通;1-过敏试验;2-给药方法(西药);3-中药煎法;4-中药用(服)法;5-特殊治疗;6-标本采集;7-配血方法;8-输血途径;9-输血采集", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo操作类型.AddItem aryTemp(intCount)
        Next
        Me.cbo操作类型.ListIndex = 0
        Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = True
        Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = True: Me.txt计算单位.Text = "次"
        Me.lblComment.Caption = Space(4) & "治疗类项目包括普通治疗和过敏试验、给药途径、中药煎法等，对特殊项目请准确标明其性质，以便医嘱和其执行时利用。"
        If Me.cbo操作类型.ListIndex = 1 Then
            stbInfo.TabVisible(3) = True
            Call InivsfTest
        Else
            stbInfo.TabVisible(3) = False
        End If

        If Me.cbo操作类型.ListIndex = 2 Then
            Me.lbl分类说明.Visible = True
            Me.cbo分类说明.Visible = True
            
            Me.cbo执行分类.Visible = True
            Me.lbl执行分类.Visible = True
        End If
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "F"        '手术
        Me.cbo操作类型.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.cbo操作类型.Left
        Me.txt项目名称.Width = Me.txt其他别名.Left + Me.txt其他别名.Width - Me.txt项目名称.Left

        Me.lbl操作类型.Caption = "手术规模(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        err = 0: On Error GoTo ErrHand
        gstrSql = "select 编码||'-'||名称 from 诊疗手术规模 order by 编码"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo类别_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo操作类型.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo操作类型.ListCount > 0 Then Me.cbo操作类型.ListIndex = 0
        End With
        Me.cbo操作类型.ListIndex = 0
        Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
        Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
        Me.lblComment.Caption = Space(4) & "手术类项目只能是一次性项目，划分为不同规模的手术，目前系统只允许在门诊和住院临时医嘱中使用。"
        Me.fra标准编码.Visible = True
        Me.opt执行部门(0).Value = False: Me.opt执行部门(0).Enabled = False
        Me.opt执行部门(1).Value = True
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "G"        '麻醉
        Me.lbl操作类型.Caption = "麻醉类型(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        err = 0: On Error GoTo ErrHand
        gstrSql = "select 编码||'-'||名称 from 诊疗麻醉类型 order by 编码"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo类别_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo操作类型.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo操作类型.ListCount > 0 Then Me.cbo操作类型.ListIndex = 0
        End With
        Me.cbo操作类型.ListIndex = 0
        Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
        Me.cbo计算方式.ListIndex = 3: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = "次"
        Me.chk单独应用.Value = 0: Me.chk单独应用.Enabled = False
        Me.lblComment.Caption = Space(4) & "麻醉类项目只能是一次性项目，只能在手术医嘱中根据需要指定，不允许麻醉类项目单独使用。"
        Me.opt执行部门(0).Value = False: Me.opt执行部门(0).Enabled = False
        Me.opt执行部门(1).Value = True
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "H"        '护理
        Me.lbl操作类型.Caption = "项目类型(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        aryTemp = Split("0-护理常规;1-护理等级", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo操作类型.AddItem aryTemp(intCount)
        Next
        Me.cbo操作类型.ListIndex = 0
        Me.cbo执行频率.ListIndex = 1
        Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = False: Me.txt计算单位.Text = " "
        Me.chk服务对象(0).Value = 0: Me.chk服务对象(0).Enabled = False
        Me.chk服务对象(1).Value = 1: Me.chk服务对象(1).Enabled = False
        Me.chk服务对象(2).Value = 0: Me.chk服务对象(2).Enabled = False
        Me.vsUseDept.Editable = flexEDNone
        For i = 0 To OptAppUse.Count - 1
            OptAppUse(i).Enabled = False
        Next
        Me.lblComment.Caption = Space(4) & "护理类项目包括护理常规和护理等级，为持续性的项目，只在住院长期医嘱中使用。"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "I"        '膳食
        Me.lbl操作类型.Visible = False
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = False
        Me.cbo执行频率.ListIndex = 2: Me.cbo执行频率.Enabled = False
        Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = True: Me.txt计算单位.Text = " "
        Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 1
        Me.lblComment.Caption = Space(4) & "膳食类项目是医生嘱咐病人配合医疗的饮食要求，为持续性的项目，通常只在住院长期医嘱中使用。"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "K"        '输血
        Me.opt执行部门(0).Enabled = False
        Me.lbl操作类型.Visible = False
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = False
        Me.cbo执行频率.ListIndex = 1: Me.cbo执行频率.Enabled = False
        Me.cbo计算方式.ListIndex = 1: Me.cbo计算方式.Enabled = True: Me.txt计算单位.Text = " "
        Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 1
        Me.lblComment.Caption = Space(4) & "输血通常作为病人辅助治疗措施一次性应用，根据实际输血量执行。"
        Me.lbllel.Visible = True
        Me.txtML.Visible = True: Me.txtML.Text = ""
        Me.lblML.Visible = True
        vsfBloodLis.Visible = True
    Case "L"        '输氧
        Me.lbl操作类型.Visible = False
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = False
        Me.cbo执行频率.ListIndex = 2: Me.cbo执行频率.Enabled = True
        Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = True: Me.txt计算单位.Text = " "
        Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 1
        Me.lblComment.Caption = Space(4) & "输氧通常作为病人常用的辅助治疗措施持续应用，根据实际用量或时间计算执行。"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "M"        '材料
        Me.lbl操作类型.Visible = False
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = False
        Me.cbo执行频率.ListIndex = 0: Me.cbo执行频率.Enabled = True
        Me.cbo计算方式.ListIndex = 1: Me.cbo计算方式.Enabled = True: Me.txt计算单位.Text = " "
        Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 1
        Me.lblComment.Caption = Space(4) & "在病人诊疗过程中，可以根据实际情况应用某些材料，执行频率和计算方式多样化，根据实际项目设置。"
        stbInfo.TabCaption(2) = "频率设置"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "Z"        '其他
        Me.lbl操作类型.Caption = "特殊标志(&T)": Me.lbl操作类型.Visible = True
        Me.cbo操作类型.Clear: Me.cbo操作类型.Visible = True
        aryTemp = Split("0-普通;1-留观;2-住院;3-转科;4-术后;5-出院;6-转院;7-会诊;8-抢救;9-病重;10-病危;11-死亡;12-记录入出量;14-术前", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo操作类型.AddItem aryTemp(intCount)
        Next
        Me.cbo操作类型.ListIndex = 0
        Me.cbo执行频率.ListIndex = 0: Me.cbo执行频率.Enabled = True
        Me.cbo计算方式.ListIndex = 0: Me.cbo计算方式.Enabled = True: Me.txt计算单位.Text = " "
        Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 1
        Me.lblComment.Caption = Space(4) & "无明确特性的诊疗项目可列入其他，但根据具体项目特点正确设置其属性，直接影响医嘱的下达和有效执行。"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    End Select
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo适用性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo执行频率_Click()
    If Left(Me.cbo类别.Text, 1) = "H" Then
        If cbo执行频率.ListIndex = 0 Then
            cbo计算方式.ListIndex = 2
        ElseIf cbo执行频率.ListIndex = 1 Then
            cbo计算方式.ListIndex = 0
        End If
    End If
    
    '可选频率及非检验项目时，允许设置频率
    If cbo执行频率.ListIndex = 0 And Left(Me.cbo类别.Text, 1) <> "C" Then
        stbInfo.TabVisible(4) = True
        Call InitVsfFreq
    Else
        stbInfo.TabVisible(4) = False
    End If
    
    If cbo执行频率.ListIndex = 0 And cbo计算方式.ListIndex > 0 Then
        lbl计算规则.Visible = True
        cbo计算规则.Visible = True
    Else
        lbl计算规则.Visible = False
        cbo计算规则.Visible = False
    End If
    
    If InStr(1, cbo执行频率.Text, "持续性") > 0 Then
        Me.lblZLPL.Visible = True
        Me.cboZLPL.Visible = True
    Else
        Me.lblZLPL.Visible = False
        Me.cboZLPL.Visible = False
    End If
End Sub

Private Sub cbo执行频率_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub chk单独应用_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk服务对象_Click(Index As Integer)
    Dim i As Long, j As Long
    Dim blnIsUse As Boolean
    
    If Me.chk服务对象(0).Enabled Then
        If Me.chk服务对象(0).Value = 1 Then
            Me.txt门诊执行.Enabled = True
            txt门诊执行.BackColor = vbWindowBackground
        Else
            Me.txt门诊执行.Enabled = False
            txt门诊执行.BackColor = vbButtonFace
        End If
        If Me.chk服务对象(1).Value = 1 Then
            Me.txt住院执行.Enabled = True
            txt住院执行.BackColor = vbWindowBackground
            Me.chk加收.Enabled = True
        Else
            Me.txt住院执行.Enabled = False
            Me.chk加收.Enabled = False
            txt住院执行.BackColor = vbButtonFace
            Me.chk加收.Value = 0
        End If
    Else
        
    End If
    If Me.chk服务对象(0).Value = 0 And Me.chk服务对象(1).Value = 0 And Me.chk服务对象(2).Value = 0 Then
        If chk服务对象(0).Enabled = True Then
            For i = 0 To vsUseDept.Rows - 1
                For j = 0 To vsUseDept.Cols - 1
                    If vsUseDept.ColHidden(j) = False Then
                        If vsUseDept.TextMatrix(i, j) <> "" Then
                            blnIsUse = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If blnIsUse Then
                chk服务对象(Index).Value = 1
            Else
                vsUseDept.Editable = flexEDNone
                For i = 0 To OptAppUse.Count - 1
                    OptAppUse(i).Enabled = False
                Next
            End If
        Else
            vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
        End If
    Else
        If Me.chk服务对象(0).Enabled Then
            vsUseDept.Editable = flexEDKbdMouse
            For i = 0 To OptAppUse.Count - 1
                If i = 0 Then
                    OptAppUse(i).Enabled = True
                Else
                    '根据参数来确定是否可用
                    OptAppUse(i).Enabled = (Val(Mid(mstr应用范围, i, 1)) = 1)
                End If
            Next
        End If
    End If
    '体检项目和门诊、住院互斥
    If Index = 0 Or Index = 1 Then
        If chk服务对象(Index).Value = 1 And chk服务对象(2).Value = 1 Then
            chk服务对象(2).Value = 0
        End If
    ElseIf chk服务对象(Index).Value = 1 Then
        If chk服务对象(0).Value = 1 Then
            chk服务对象(0).Value = 0
        End If
        If chk服务对象(1).Value = 1 Then
            chk服务对象(1).Value = 0
        End If
    End If
End Sub

Private Sub chk服务对象_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk检验组合_Click()
    If Me.chk检验组合.Value = 1 Then
        Me.chk单独应用.Enabled = False: Me.chk单独应用.Value = 1
    Else
        Me.chk单独应用.Enabled = True
    End If
'    Me.stbInfo.TabVisible(2) = Not Me.chk检验组合.Value = 1
End Sub

Private Sub chk执行安排_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Left(Me.cbo类别.Text, 1) = "D" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Me.fra标准编码.Visible Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        Me.stbInfo.Tab = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Function CheckUseDept(ByVal strDept As String) As String
'检查使用科室的站点和服务对象是否和界面上的吻合
'返回：如果有不吻合的，返回提示信息
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim strMsg As String
    Dim str站点 As String
    
    On Error GoTo errH
    If strDept = "" Then Exit Function
    If cmbStationNo.Text <> "" Then
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        
        strSql = "Select a.ID,a.名称,a.站点 From 部门表 A Where ID In(" & strDept & ") And (a.站点<>[1] And a.站点 Is Not Null)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str站点)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strMsg = strMsg & "," & rsTmp!名称
                rsTmp.MoveNext
            Loop
            strMsg = Mid(strMsg, 2)
            CheckUseDept = strMsg & " 不是 " & str站点 & " 站点的科室，请检查。"
        End If
    End If
    strSql = ""
    If chk服务对象(0).Value = 0 And chk服务对象(2).Value = 0 Then
        '没有勾选门诊，检查是否有只服务于门诊的科室。
        strSql = "Select ID,名称 From (Select a.ID,a.名称,Decode(Max(服务对象), 3, 3, 2, Decode(Min(服务对象), 1, 3, 2, 2), 1, 1) As 服务对象 " & vbNewLine & _
                " From 部门表 A, 部门性质说明 B" & vbNewLine & _
                " Where a.Id = b.部门id And b.服务对象 <> 0" & vbNewLine & _
                " And ID In(" & strDept & ") " & _
                " Group By a.Id,a.名称, 站点 ) Where 服务对象=1"

    ElseIf chk服务对象(1).Value = 0 Then
        '没有勾选住院，检查是否有只服务于住院的科室。
        strSql = "Select ID,名称 From (Select a.ID,a.名称,Decode(Max(服务对象), 3, 3, 2, Decode(Min(服务对象), 1, 3, 2, 2), 1, 1) As 服务对象 " & vbNewLine & _
                " From 部门表 A, 部门性质说明 B" & vbNewLine & _
                " Where a.Id = b.部门id And b.服务对象 <> 0" & vbNewLine & _
                " And ID In(" & strDept & ")" & _
                " Group By a.Id,a.名称, 站点) Where 服务对象=2"
    End If
    If strSql <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strMsg = strMsg & "," & rsTmp!名称
                rsTmp.MoveNext
            Loop
            strMsg = Mid(strMsg, 2)
            If chk服务对象(0).Value = 0 And chk服务对象(2).Value = 0 Then
                CheckUseDept = strMsg & " 是只服务于门诊的科室，请检查。"
            Else
                CheckUseDept = strMsg & " 是只服务于住院的科室，请检查。"
            End If
            
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim strFormula As String
    Dim strErrorMsg As String, iErrorPos As Integer
    Dim strMid As Variant
    Dim i As Integer, lngVItemID0 As Long
    Dim mAppType As Integer                 '应用类型 =0应用于本项;=1应用于同级;=2应用于本级下所有;=3应用于所有类别
    Dim strTmp As String, blnBegin As Boolean
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strFreq As String
    Dim str站点 As String
    Dim j As Long
    Dim strDeptId As String
    Dim strMsg As String
    Dim str编码 As String
    Dim int循环 As Integer
    Dim lngNum As Long
    Dim lngSum As Long
    Dim intFirst As Integer
    Dim strLast As String
    Dim intRow As Integer
    Dim str输血检验对照 As String
    Dim str试管编码 As String
    Dim blnRisTrans As Boolean
    Dim varTemp As Variant
    Dim strItem As String
    
    '皮试结果
    Dim strTest As String
    Dim str阳性 As String
    Dim str阴性 As String
    
    lngNum = 1
    int循环 = 1
    '重新检查名称，并去掉特殊字符
    strTmp = MoveSpecialChar(txt项目名称.Text)
    If txt项目名称.Text <> strTmp Then
        txt项目名称.Text = strTmp
        Me.txt名称拼音.Text = zlStr.GetCodeByORCL(strTmp, False, mlng简码长度)
        Me.txt名称五笔.Text = zlStr.GetCodeByORCL(strTmp, True, mlng简码长度)
    End If
    
    '检查使用科室
    With Me.vsUseDept
        For i = 0 To .Rows - 1
            For j = 0 To 4
                If .TextMatrix(i, j) <> "" Then
                    strDeptId = strDeptId & "," & .TextMatrix(i, j + 5)
                End If
            Next
        Next
        strDeptId = Mid(strDeptId, 2)
        strMsg = CheckUseDept(strDeptId)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Me.stbInfo.Tab = 0: Me.vsUseDept.SetFocus: Exit Sub
        End If
    End With
    strTmp = MoveSpecialChar(txt其他别名.Text)
    If txt其他别名.Text <> strTmp Then
        txt其他别名.Text = strTmp
        Me.txt别名拼音.Text = zlStr.GetCodeByORCL(strTmp, False, mlng简码长度)
        Me.txt别名五笔.Text = zlStr.GetCodeByORCL(strTmp, True, mlng简码长度)
    End If

    '一般特性检查
    If Trim(Me.txt项目编码.Text) = "" Then
        MsgBox "请输入项目编码！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt项目编码.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt项目编码.Text), vbFromUnicode)) > Me.txt项目编码.MaxLength Then
        MsgBox "项目编码的长度超长（最多" & Me.txt项目编码.MaxLength & " 个字符）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt项目编码.SetFocus: Exit Sub
    End If
    If Trim(Me.txt项目名称.Text) = "" Then
        MsgBox "请输入项目名称！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt项目名称.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt项目名称.Text), vbFromUnicode)) > Me.txt项目名称.MaxLength Then
        MsgBox "项目名称超长（最多" & Me.txt项目名称.MaxLength & "个字符或" & Me.txt项目名称.MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt项目名称.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt名称拼音.Text), vbFromUnicode)) > Me.txt名称拼音.MaxLength Then
        MsgBox "项目名称超长（最多" & Me.txt名称拼音.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt名称拼音.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt名称五笔.Text), vbFromUnicode)) > Me.txt名称五笔.MaxLength Then
        MsgBox "项目名称超长（最多" & Me.txt名称五笔.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt名称五笔.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt别名拼音.Text), vbFromUnicode)) > Me.txt别名拼音.MaxLength Then
        MsgBox "项目名称超长（最多" & Me.txt别名拼音.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt别名拼音.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt别名五笔.Text), vbFromUnicode)) > Me.txt别名五笔.MaxLength Then
        MsgBox "项目名称超长（最多" & Me.txt别名五笔.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt别名五笔.SetFocus: Exit Sub
    End If
    If Me.cbo计算方式.ListIndex = 1 And Trim(Me.txt计算单位.Text) = "" Then
        MsgBox "计量类项目，请输入计算单位！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt计算单位.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt计算单位.Text), vbFromUnicode)) > Me.txt计算单位.MaxLength Then
        MsgBox "计算单位超长（最多" & Me.txt计算单位.MaxLength & "个字符或" & Me.txt计算单位.MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt计算单位.SetFocus: Exit Sub
    End If
    If Left(Me.cbo类别.Text, 1) = "D" And Me.opt检查部位(0).Value = True Then
'        If LenB(StrConv(Trim(Me.txt检查部位.Text), vbFromUnicode)) > 60 Then
'            MsgBox "检查部位超长（最多40个字符或20个汉字）！", vbInformation, gstrSysName
'            Me.stbInfo.Tab = 0: Me.txt检查部位.SetFocus: Exit Sub
'        End If
'        If mbln组合部位项目 And Trim(Me.txt检查部位.Text) = "" Then
'            MsgBox "检查部位不能为空！", vbInformation, gstrSysName
'            Me.stbInfo.Tab = 0: Me.txt检查部位.SetFocus: Exit Sub
'        End If
    End If
    If LenB(StrConv(Trim(Me.txt其他别名.Text), vbFromUnicode)) > Me.txt其他别名.MaxLength Then
        MsgBox "其他别名（最多" & Me.txt其他别名.MaxLength & "个字符或" & Int(Me.txt其他别名.MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt其他别名.SetFocus: Exit Sub
    End If
    If Len(Trim(Me.txt录入限量.Text)) > 0 Then
        If CDbl(Me.txt录入限量.Text) > CDbl("99999999999.99999") Then
            MsgBox "录入限量不能大于（99999999999.99999）值！", vbInformation, gstrSysName
            Me.stbInfo.Tab = 0: Me.txt录入限量.SetFocus: Exit Sub
        End If
    End If

    If Left(Me.cbo类别.Text, 1) = "C" And Trim(Me.txt标本部位.Text) = "" Then
        MsgBox "检验项目必须设置默认标本部位！", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt标本部位.SetFocus: Exit Sub
    End If
    '10804 并发操作时，检查检验类型是否被删除
    If Left(Me.cbo类别.Text, 1) = "C" Then
        If Not zlExistItem("诊疗检验类型", "名称", Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型, "-") + 1), "操作类型：" & Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型, "-") + 1)) Then
            Me.cbo操作类型.SetFocus:  Exit Sub
        End If
    End If

    '检查部位
    If Left(Me.cbo类别.Text, 1) = "D" Then
        '检查项目 需完成数据正确性验证
        
    End If

    If Me.opt执行部门(4).Value = True Then
        '定向执行检查
        strTemp = ""
        With Me.msf定向执行
            For intCount = 1 To .Rows - 1
                If Val(.TextMatrix(intCount, 0)) <> 0 Then
                    '不再检查是否重复 By 赵彤宇
                    'If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 0)) & "-" & .TextMatrix(intCount, 1) & ";") > 0 Then
                    If InStr(1, strTemp & ";", ";" & .TextMatrix(intCount, col执行科室) & ";") > 0 Then
                        MsgBox "重复指定了执行科室“" & .TextMatrix(intCount, col执行科室) & "”！", vbInformation, gstrSysName
                        Me.stbInfo.Tab = 1: .SetFocus: Exit Sub
                    Else
                        strTemp = strTemp & ";" & .TextMatrix(intCount, col执行科室)
                    End If
'                    If Val(.TextMatrix(intCount, 2)) = 0 Then
'                        MsgBox "“" & .TextMatrix(intCount, 1) & "”未指定执行科室！", vbInformation, gstrSysName
'                        Me.stbInfo.Tab = 1: .SetFocus: Exit Sub
'                    End If
                End If
            Next
        End With
    End If
    
    '执行科室应用为大范围时提示
    If OptApp(0).Enabled = True And OptApp(0).Value = False Then
        For i = 0 To Me.OptApp.Count - 1
            If OptApp(i).Enabled = True And OptApp(i).Value = True Then
                If MsgBox("该项目设置的执行科室将" & OptApp(i).Caption & "项目，是否保存？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbInfo.Tab = 1:  Exit Sub
                End If
            End If
        Next
    End If
    
    '使用科室应用为大范围时提示
    If OptAppUse(0).Enabled = True And OptAppUse(0).Value = False Then
        For i = 0 To Me.OptAppUse.Count - 1
            If OptAppUse(i).Enabled = True And OptAppUse(i).Value = True Then
                If MsgBox("该项目设置的使用科室将" & OptAppUse(i).Caption & "项目，是否保存？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbInfo.Tab = 0:  Exit Sub
                End If
            End If
        Next
    End If
    
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '新增项目时，保证不出现重复编码，如果有重复自动在原编码基础上加1，直到不重复
    str编码 = Trim(txt项目编码.Text)
    If Me.Tag = "增加" Or Me.Tag = "复制增加" Then
        Do While True
            gstrSql = "select a.编码 from 诊疗项目目录 a,诊疗项目类别 b where a.编码=[1] and a.类别=b.编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "编码是否重复", str编码)
            If rsTemp.RecordCount <> 0 Then
                str编码 = zlCommFun.IncStr(str编码)
            Else
                Exit Do
            End If
        Loop
    End If
        
    '数据保存
    If Me.Tag = "增加" Or Me.Tag = "复制增加" Then
        lngItemID = zlDatabase.GetNextId("诊疗项目目录")
'        If zlClinicCodeRepeat(Trim(Me.txt项目编码.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(str编码, lngItemID) = True Then Exit Sub
        If zlExistItem("诊疗项目目录", "ID", lngItemID, Trim(Me.txt项目名称.Text)) = False Then Exit Sub
    End If

    gcnOracle.BeginTrans
    blnBegin = True
    Do While int循环 <> 0
        intFirst = intFirst + 1
        gstrSql = "'" & Left(Me.cbo类别.Text, 1) & "'," & Me.txt分类.Tag & "," & lngItemID & ",'" & str编码 & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt项目名称.Text) & "','" & Trim(Me.txt名称拼音.Text) & "','" & Trim(Me.txt名称五笔.Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt其他别名.Text) & "','" & Trim(Me.txt别名拼音.Text) & "','" & Trim(Me.txt别名五笔.Text) & "'"
        Select Case Left(Me.cbo类别.Text, 1)
        Case "C", "D", "F", "G"     '"C-检验", "D-检查", "F-手术", "G-麻醉"
            gstrSql = gstrSql & ",'" & Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型.Text, "-") + 1) & "'"
        Case "E", "H", "Z"              '"E-治疗", "H-护理"
            gstrSql = gstrSql & ",'" & Mid(Me.cbo操作类型.Text, 1, InStr(Me.cbo操作类型.Text, "-") - 1) & "'"
        Case Else
            gstrSql = gstrSql & ",''"
        End Select
        gstrSql = gstrSql & "," & Mid(Me.cbo执行频率.Text, 1, 1) & "," & Me.chk单独应用.Value
        gstrSql = gstrSql & "," & Me.cbo计算方式.ListIndex & ",'" & Trim(Me.txt计算单位.Text) & "'"
        gstrSql = gstrSql & "," & Me.cbo适用性别.ListIndex & "," & Me.chk执行安排.Value
        gstrSql = gstrSql & "," & IIf(Me.chk服务对象(0).Value = 0, 0, 1) + IIf(Me.chk服务对象(1).Value = 0, 0, 2) + IIf(Me.chk服务对象(2).Value = 0, 0, 4)
    
        Select Case Left(Me.cbo类别.Text, 1)
        Case "D"
            
            ' strSql & ",0,'" & Trim(Me.txt检查部位.Text) & "',null"
            '            是否组合项目0-不是( 1-是,没有检查部位)
            '            计算方式固定为Null
                
            '新的方式下,检查项目没有组合项目
                Dim str检查部位 As String, lngRow As Long
                Dim str方法 As String
                Dim strModusSQL() As String, arrItem() As String, lngCount As Long, lngItem As Long
                
                With vfgList
                lngCount = 0
                For lngRow = .FixedRows To .Rows - 1
                    If .RowData(lngRow) = 1 Then
                        str检查部位 = str检查部位 & Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & ","
                        '检查项目 生成保存诊疗项目部位的SQL
                        str方法 = .Cell(flexcpText, lngRow, .ColIndex("方法"))
                        arrItem = Split(.Cell(flexcpText, lngRow, .ColIndex("方法")), "  ")
                        For i = 0 To UBound(arrItem)
                            If Trim(arrItem(i)) <> "" Then
                                If InStr(arrItem(i), "〈") > 0 Then
                                    strTemp = Mid(arrItem(i), 1, InStr(arrItem(i), "〈") - 1)
                                    strItem = Mid(arrItem(i), InStr(arrItem(i), "〈") + 1, InStr(arrItem(i), "〉") - InStr(arrItem(i), "〈") - 1)
                                    If InStr(strTemp, "●") > 0 Or InStr(strTemp, "■") > 0 Then
                                        strTemp = Replace(strTemp, "●", "")
                                        strTemp = Trim(Replace(strTemp, "■", ""))
                                        lngCount = lngCount + 1
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_诊疗项目部位_insert(" & lngItemID & ",'" & Mid(cbo操作类型.Text, InStr(cbo操作类型, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & "','" & strTemp & "',1,'')"
                                    Else
                                        strTemp = Replace(strTemp, "○", "")
                                        strTemp = Trim(Replace(strTemp, "□", ""))
                                        lngCount = lngCount + 1
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_诊疗项目部位_insert(" & lngItemID & ",'" & Mid(cbo操作类型.Text, InStr(cbo操作类型, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & "','" & strTemp & "','','')"
                                    End If
                                    varTemp = Split(strItem, " ")
                                    For j = 0 To UBound(varTemp)
                                        If InStr(varTemp(j), "■") > 0 Then
                                            strTmp = Trim(Replace(varTemp(j), "■", ""))
                                            lngCount = lngCount + 1
                                            ReDim Preserve strModusSQL(lngCount) As String
                                            strModusSQL(lngCount) = "zl_诊疗项目部位_insert(" & lngItemID & ",'" & Mid(cbo操作类型.Text, InStr(cbo操作类型, "-") + 1) & "','" & _
                                                            Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & "','" & strTmp & "',1,'" & strTemp & "')"
                                        Else
                                            strTmp = Trim(Replace(varTemp(j), "□", ""))
                                            lngCount = lngCount + 1
                                            ReDim Preserve strModusSQL(lngCount) As String
                                            strModusSQL(lngCount) = "zl_诊疗项目部位_insert(" & lngItemID & ",'" & Mid(cbo操作类型.Text, InStr(cbo操作类型, "-") + 1) & "','" & _
                                                            Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & "','" & strTmp & "','','" & strTemp & "')"
                                        End If
                                    Next
                                Else
                                    strTemp = arrItem(i)
                                    If InStr(strTemp, "●") > 0 Or InStr(strTemp, "■") > 0 Then
                                        strTemp = Replace(strTemp, "●", "")
                                        strTemp = Trim(Replace(strTemp, "■", ""))
                                        lngCount = lngCount + 1
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_诊疗项目部位_insert(" & lngItemID & ",'" & Mid(cbo操作类型.Text, InStr(cbo操作类型, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & "','" & strTemp & "',1,'')"
                                    Else
                                        lngCount = lngCount + 1
                                        strTemp = Replace(strTemp, "○", "")
                                        strTemp = Trim(Replace(strTemp, "□", ""))
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_诊疗项目部位_insert(" & lngItemID & ",'" & Mid(cbo操作类型.Text, InStr(cbo操作类型, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("名称"))) & "','" & strTemp & "','','')"
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
                End With
                str检查部位 = zlCommFun.ToVarchar(str检查部位, 60)
                
                gstrSql = gstrSql & ",0,'" & str检查部位 & "',Null"
                
                
                '        If Me.opt检查部位(0).Value = True Then
                '            gstrSql = gstrSql & ",0,'" & Trim(Me.txt检查部位.Text) & "',null"
                '        Else
                '            gstrSql = gstrSql & ",1,'',null"
                '        End If
        Case "F"
            If Val(Me.txt标准编码.Tag) <> 0 Then
                gstrSql = gstrSql & ",0,''," & Val(Me.txt标准编码.Tag)
            Else
                gstrSql = gstrSql & ",0,'',null"
            End If
        Case "C"
            gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'" & Trim(Me.txt标本部位.Text) & "',null"
        Case "E"
            If cbo操作类型.ListIndex = 1 Then
                With vsfTest
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 2) = "√" Then
                            str阳性 = IIf(str阳性 = "", "", str阳性 & ",") & .TextMatrix(i, 1) & .TextMatrix(i, 0)
                        Else
                            str阴性 = IIf(str阴性 = "", "", str阴性 & ",") & .TextMatrix(i, 1) & .TextMatrix(i, 0)
                        End If
                    Next
                End With
                strTest = str阳性 & ";" & str阴性
                If LenB(strTest) > 60 Then
                    MsgBox "皮试结果项目总字符数太多，请减少项目或减少字符数！", vbInformation, gstrSysName
                    Me.stbInfo.Tab = 3
                    Exit Sub
                End If
                gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'" & strTest & "',null"
            ElseIf cbo操作类型.ListIndex = 2 Then
                gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'" & Trim(Me.cbo分类说明.Text) & "',null"
            Else
                gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'',null"
            End If
        Case "Z"
            If cbo操作类型.ListIndex = 12 Then
                gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'" & IIf(chk尿量.Value = 1, "尿量", "") & "',null"
            Else
                gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'',null"
            End If
        Case Else
            gstrSql = gstrSql & "," & Me.chk检验组合.Value & ",'',null"
        End Select
    
        If Me.opt执行部门(6).Value Then
            gstrSql = gstrSql & ",6"
        ElseIf Me.opt执行部门(5).Value Then
            gstrSql = gstrSql & ",5"
        ElseIf Me.opt执行部门(4).Value Then
            gstrSql = gstrSql & ",4"
        ElseIf Me.opt执行部门(3).Value Then
            gstrSql = gstrSql & ",3"
        ElseIf Me.opt执行部门(2).Value Then
            gstrSql = gstrSql & ",2"
        ElseIf Me.opt执行部门(1).Value Then
            gstrSql = gstrSql & ",1"
        Else
            gstrSql = gstrSql & ",0"
        End If
    
        If Me.opt执行部门(4).Value Then
            If Me.txt门诊执行.Enabled Then
                gstrSql = gstrSql & "," & IIf(Val(Me.txt门诊执行.Tag) = 0 Or Me.txt门诊执行.Text = "", "null", Val(Me.txt门诊执行.Tag))
            Else
                gstrSql = gstrSql & ",null"
            End If
            If Me.txt住院执行.Enabled Then
                gstrSql = gstrSql & "," & IIf(Val(Me.txt住院执行.Tag) = 0 Or Me.txt住院执行.Text = "", "null", Val(Me.txt住院执行.Tag))
            Else
                gstrSql = gstrSql & ",null"
            End If
            strTemp = ""
            strLast = ""
            With Me.msf定向执行
                For intCount = lngNum To .Rows - 1
                    lngSum = lngSum + 1
                    If Val(.TextMatrix(intCount, col执行科室ID)) <> 0 Then
                        strMid = Split(.TextMatrix(intCount, col病人科室ID), ",")
                        If UBound(strMid) <> -1 Then
                            For i = LBound(strMid) To UBound(strMid)
                                strTemp = strTemp & "|" & Trim(IIf(strMid(i) = "（所有部门）", 0, strMid(i))) & "^" & Trim(.TextMatrix(intCount, col执行科室ID))
                            Next
                        Else
                            strTemp = strTemp & "|" & 0 & "^" & Trim(.TextMatrix(intCount, col执行科室ID))
                        End If
                    End If
                    
                    If intCount < .Rows - 1 Then
                        If Len(strTemp) > 4000 Then
                            lngNum = lngNum - 1
                            strTemp = strLast
                            Exit For
                        ElseIf Len(strTemp & .TextMatrix(intCount + 1, col病人科室ID)) > 4000 Then
                            Exit For
                        Else
                            strLast = strTemp
                        End If
                    End If
                Next
                
                lngNum = lngSum + 1
                If intCount = .Rows Or lngSum = 0 Then int循环 = 0
            End With
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            gstrSql = gstrSql & ",'" & strTemp & "'"
        Else
            int循环 = 0
            gstrSql = gstrSql & ",null,null,''"
        End If
    
        For i = 0 To Me.OptApp.Count - 1
            If Me.OptApp(i).Enabled = True And Me.OptApp(i).Value = True Then
                mAppType = i
                Exit For
            End If
        Next
        If Me.OptApp(0).Enabled = False Then
            mAppType = 0
        End If
        If Len(Me.txt参考.Tag) = 0 Then
            If Me.Tag = "增加" Or Me.Tag = "复制增加" Then
                gstrSql = gstrSql & ",Null" & "," & mAppType
            Else
                gstrSql = gstrSql & ",Null" & ",0," & mAppType
            End If
        Else
            If Me.Tag = "增加" Or Me.Tag = "复制增加" Then
                gstrSql = gstrSql & "," & Me.txt参考.Tag & "," & mAppType
            Else
                gstrSql = gstrSql & "," & Me.txt参考.Tag & ",0," & mAppType
            End If
        End If
    
        '录入限量及应用范围
        gstrSql = gstrSql & "," & Val(Me.txt录入限量.Text) & "," & cbo录入限量范围.ListIndex
        '执行标记 (床旁加收，给药途径为输液时为输液类型)
        If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 2 And Me.cbo执行分类.ListIndex = 1 Then
            gstrSql = gstrSql & "," & IIf(Val(Me.cbo输液类型.ListIndex) = 0, 0, 2)
        ElseIf Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 1 Then
            gstrSql = gstrSql & "," & IIf(1 = Val(Me.chkNoTMSY.Value), 2, 0)
        Else
            gstrSql = gstrSql & "," & Val(Me.chk加收.Value)
        End If
        
        '执行分类
        If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 2 Then
            '-输液，注射，其他，口服
            gstrSql = gstrSql & "," & Val(cbo执行分类.List(cbo执行分类.ListIndex))
        ElseIf Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 1 Then
            '皮试
            gstrSql = gstrSql & IIf(1 = Val(Me.chkYYPS.Value), ",5", ",3")
        ElseIf Left(Me.cbo类别.Text, 1) = "D" And Me.cbo操作类型.Text = "18-病理" Then
            '病理
            If UBound(Split(cbo病理类别.Text, "-")) > 0 Then
                gstrSql = gstrSql & "," & Val(Split(cbo病理类别.Text, "-")(0))
            End If
        ElseIf Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 8 Then
            gstrSql = gstrSql & "," & Val(cboBloodType.ListIndex)
        Else
            '其他
            gstrSql = gstrSql & ",0"
        End If
        
        
        '站点
        gstrSql = gstrSql & "," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", str站点)
        
        '频率设置
        If Left(Me.cbo类别.Text, 1) <> "C" And stbInfo.TabVisible(4) = True Then
            With vsfFreq
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) = "√" Then
                        strFreq = IIf(strFreq = "", "", strFreq & "|") & .TextMatrix(i, 1)
                    End If
                Next
            End With
        End If
        gstrSql = gstrSql & "," & IIf(strFreq = "", "Null", "'" & strFreq & "'")
        
        '计算规则
        If Me.cbo计算规则.Visible = True Then
            gstrSql = gstrSql & "," & Val(cbo计算规则.List(cbo计算规则.ListIndex))
        Else
            gstrSql = gstrSql & ",Null"
        End If
        
        '使用科室
        gstrSql = gstrSql & ",'" & strDeptId & "'"
        '使用科室范围
        For i = 0 To Me.OptAppUse.Count - 1
            If Me.OptAppUse(i).Enabled = True And Me.OptAppUse(i).Value = True Then
                mAppType = i
                Exit For
            End If
        Next
        If Me.OptAppUse(0).Enabled = False Then
            mAppType = 0
        End If
        gstrSql = gstrSql & "," & mAppType
        
        '59964-是否第一次执行
        gstrSql = gstrSql & "," & IIf(intFirst = 1, 1, 0)
        
        gstrSql = gstrSql & "," & IIf(Me.txtML.Text = "", "NULL", Val(Me.txtML.Text))
        '输血检验对照
        With vsfBloodLis
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    str输血检验对照 = str输血检验对照 & "|" & .TextMatrix(intRow, 0)
                End If
            Next
            str输血检验对照 = Mid(str输血检验对照, 2)
        End With
        gstrSql = gstrSql & "," & IIf(str输血检验对照 = "", "NULL", "'" & str输血检验对照 & "'")
        
        '保存诊疗频率
        gstrSql = gstrSql & "," & IIf(Me.cboZLPL.Text = "", "NULL", "'" & Mid(Mid(Me.lblZLPL.Tag, InStr(1, Me.lblZLPL.Tag, "|" & Me.cboZLPL.Text & "-") + Len(Me.cboZLPL.Text) + 2), 1, InStr(1, Mid(Me.lblZLPL.Tag, InStr(1, Me.lblZLPL.Tag, "|" & Me.cboZLPL.Text & "-") + Len(Me.cboZLPL.Text) + 2), "|") - 1) & "'")
        
        '治疗类的输血采集方式，保存对应的试管编码
        str试管编码 = ""
        If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 9 Then
            If cboTestTubeCode.ListIndex > 0 Then
                str试管编码 = Split(cboTestTubeCode.List(cboTestTubeCode.ListIndex), "-")(0)
            End If
        End If
        
        If Me.Tag = "增加" Then
            gstrSql = "zl_诊疗项目_Insert(" & gstrSql & IIf(str试管编码 = "", "", ",0,'" & str试管编码 & "'") & ")"
        ElseIf Me.Tag = "复制增加" Then
            gstrSql = "zl_诊疗项目_Insert(" & gstrSql & "," & mlngOldId & IIf(str试管编码 = "", "", ",'" & str试管编码 & "'") & ")"
        Else
            gstrSql = "zl_诊疗项目_Update(" & gstrSql & IIf(str试管编码 = "", "", ",'" & str试管编码 & "'") & ")"
        End If
    
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Loop
    
    '检查项目 诊疗项目部位保存
    If Left(Me.cbo类别.Text, 1) = "D" Then
        '新网RIS接口，新增/修改诊疗项目
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItem, IIf(Me.Tag = "增加", RISBaseItemOper.AddNew, RISBaseItemOper.Modify), lngItemID) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '出错时提示接口错误信息
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                    End If
                    
                    Exit Sub
                End If
                
                '新网RIS接口，删除诊疗项目部位；放到HIS删除过程之前
                If Me.Tag = "修改" Then
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.Delete, lngItemID) <> 1 Then
                        gcnOracle.RollbackTrans
                        
                        '出错时提示接口错误信息
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                        End If
                        
                        Exit Sub
                    End If
                End If
                
                blnRisTrans = True
            Else
                '接口部件无效时禁止并提示
                gcnOracle.RollbackTrans
                
                MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                
                Exit Sub
            End If
        End If
        
        'HIS删除/修改项目部位
        gstrSql = "zl_诊疗项目部位_Delete(" & lngItemID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
        If str检查部位 <> "" Then
            For lngCount = LBound(strModusSQL) To UBound(strModusSQL) - 1
                If strModusSQL(lngCount + 1) <> "" Then
                    Call zlDatabase.ExecuteProcedure(strModusSQL(lngCount + 1), Me.Caption)
                    
                End If
            Next
            
            '新网RIS接口，新增诊疗项目部位
            '放到HIS新增过程之后
            If mblnPACSInterface = True Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.AddNew, lngItemID) <> 1 Then
                        gcnOracle.RollbackTrans
                        
                        '出错时提示接口错误信息
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                        End If
                        
                        Exit Sub
                    End If
                    
                    blnRisTrans = True
                Else
                    gcnOracle.RollbackTrans
                    
                    '接口部件无效时禁止并提示
                    MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                    
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '检验数据保存
    lngVItemID0 = lngVItemID
    If Left(Me.cbo类别.Text, 1) = "C" Then
        
        gstrSql = "Null,'" & str编码 & "','" & Me.txt项目名称 & "','" & _
            Me.txt名称拼音 & "'," & "0,10,0,Null," & _
            "Null,0,Null,Null,Null,Null,Null,0"

        If lngVItemID0 = 0 And (Me.Tag = "增加" Or Me.Tag = "复制增加") Then
            lngVItemID0 = zlDatabase.GetNextId("诊治所见项目")
            gstrSql = "ZL_所见项目_INSERT(" & lngVItemID0 & "," & gstrSql & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        Else
            strSql = "Select id,分类id,编码,中文名,英文名,类型,长度,小数,单位,临床意义,表示法,性别域,数值域,初始值,文字表述,空值文字 From 诊治所见项目 Where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngVItemID0)
            Do Until rsTmp.EOF
                gstrSql = "'" & rsTmp!分类id & "','" & str编码 & "','" & Me.txt项目名称 & "','" & _
                        Me.txt英文 & "'," & rsTmp!类型 & "," & rsTmp!长度 & "," & rsTmp!小数 & ",'" & Me.txt计算单位 & "','" & _
                        rsTmp!临床意义 & "'," & rsTmp!表示法 & "," & rsTmp!性别域 & ",'" & rsTmp!数值域 & "','" & rsTmp!初始值 & "','" & _
                        rsTmp!文字表述 & "','" & rsTmp!空值文字 & "'"
                rsTmp.MoveNext
                gstrSql = "ZL_所见项目_UPDATE(" & lngVItemID0 & "," & gstrSql & ")"
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Loop
        End If
        
        If Not (lngVItemID0 = 0 And Me.Tag <> "增加") Then '导的演示数据有问题,加上判断,这种错误数据不处理.
            gstrSql = "Select 报告代号,项目类别,结果类型,单位,打印类型,打印序号,计算公式,检验方法,合并后代码,结果异常条件 From 检验项目 Where 诊治项目id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngVItemID0)
            Do Until rsTmp.EOF
                gstrSql = "'" & Me.txt英文 & "','" & rsTmp!报告代号 & "','" & rsTmp!项目类别 & "','" & rsTmp!结果类型 & "','" & Me.txt计算单位 & "','" & _
                          rsTmp!打印类型 & "','" & rsTmp!打印序号 & "','" & rsTmp!计算公式 & "','" & rsTmp!检验方法 & "','" & rsTmp!合并后代码 & "','" & _
                          rsTmp!结果异常条件 & "'"
                          
                gstrSql = "ZL_检验项目_UPDATE(" & lngVItemID0 & "," & gstrSql & ")"
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                rsTmp.MoveNext
            Loop
            gstrSql = "'^" & lngVItemID0 & "'"
            gstrSql = "ZL_检验报告项目_UPDATE(" & lngItemID & "," & gstrSql & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        End If
    ElseIf lngVItemID0 > 0 Then
        '删除原来基础项目的报告项目
        gstrSql = "ZL_检验报告项目_UPDATE(" & lngItemID & ",'')"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        '如果将基础项目改为组合项目，则删除诊治所见项目
        gstrSql = "ZL_所见项目_DELETE(" & lngVItemID0 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If

    lngVItemID = lngVItemID0
    
    gcnOracle.CommitTrans
    
    blnBegin = False
    blnRisTrans = False

    mblnOK = True
    '连续增加处理
    If Me.Tag = "增加" Or Me.Tag = "复制增加" Then
        If chkGoOn.Value Then
            
            mbln连续增加 = True
            lngItemID = 0
            mLast操作类型 = ""
            Call Form_Activate
            Me.stbInfo.Tab = 0
            Me.txt分类.SetFocus
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub

ErrHand:
    If blnBegin Then gcnOracle.RollbackTrans
    
    'Ris接口和HIS不同步时，写错误日志
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS" & IIf(Me.Tag = "增加", "增加", "修改") & "诊疗项目错误，RIS接口和HIS数据不同步，请与系统管理员联系。", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmClinicItem：cmdOK_Click", "HIS" & IIf(Me.Tag = "增加", "增加", "修改") & "诊疗项目错误，RIS接口和HIS数据不同步", "诊疗项目ID=" & lngItemID, 0)
    End If
        
    'If ErrCenter() = 1 Then Resume
    Call ErrCenter
    'Resume
    Call SaveErrLog
End Sub


Private Sub IniStationNo()
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
    On Error GoTo ErrHandle
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSql = "select 编号,名称 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "站点查询")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!编号 & "-" & rsRecord!名称
                rsRecord.MoveNext
            Loop
        End With
        
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
Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub
Private Sub cmdTestAdd_Click()
    Dim n As Integer
    Dim str标注 As String
    Dim str文字 As String
    Dim str过敏 As String
    
    If Trim(txt皮试标注.Text) = "" Or Trim(txt皮试文字.Text) = "" Then Exit Sub
    
    str标注 = "(" & Trim(txt皮试标注.Text) & ")"
    str文字 = Trim(txt皮试文字.Text)
    str过敏 = IIf(chk皮试过敏.Value = 1, "√", "")
    
    With vsfTest
        '检查是否重复
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If .TextMatrix(n, 0) = str标注 And .TextMatrix(n, 1) = str文字 Then
                    MsgBox "有重复项目，请重新输入！", vbExclamation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
    
        '增加新项目
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str标注
        .TextMatrix(.Rows - 1, 1) = str文字
        .TextMatrix(.Rows - 1, 2) = str过敏
    End With
End Sub

Private Sub cmdTestDel_Click()
    Dim int阳性数量 As Integer
    Dim int阴性数量 As Integer
    Dim n As Integer
    Dim intCurr As Integer
    
    With vsfTest
        If .Row > 0 Then
            '当前的阳性和阴性数量
            For n = 1 To .Rows - 1
                If .TextMatrix(n, 2) = "√" Then
                    int阳性数量 = int阳性数量 + 1
                Else
                    int阴性数量 = int阴性数量 + 1
                End If
            Next
            
            '至少要保证阳性和阴性项目各剩余1个
            If (.TextMatrix(.Row, 2) = "√" And int阳性数量 <= 1) Or (.TextMatrix(.Row, 2) = "" And int阴性数量 <= 1) Then
                MsgBox "不能删除该项目，至少要保证阳性和阴性项目各剩余1个", vbInformation, gstrSysName
                Exit Sub
            End If
            
            .RemoveItem .Row
        End If
    End With
End Sub
Private Sub cmd标本_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim vPoint As POINTAPI

    On Error GoTo ErrHandle
    strSql = "Select Rownum As ID, 编码, 名称, 简码 From 诊疗检验标本 Order By 编码"

    vPoint = zlControl.GetCoordPos(txt标本部位.hWnd, txt标本部位.Left - 165, txt标本部位.Top - 30)

    Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "标本部位", , , , , True, True, vPoint.x, vPoint.y)

    If rsTmp.State = 0 Then
        Me.txt标本部位.Text = ""
        If Trim(Me.txt标本部位.Text) = "" And Trim(Me.txt标本部位.Tag) <> "" Then
            Me.txt标本部位.Text = Trim(Me.txt标本部位.Tag)
        End If
        Exit Sub
    End If
    If Not rsTmp Is Nothing Then
        Me.txt标本部位.Text = rsTmp("名称")
        Me.txt标本部位.Tag = rsTmp("名称")
    Else
        If Trim(Me.txt标本部位.Text) = "" And Trim(Me.txt标本部位.Tag) <> "" Then
            Me.txt标本部位.Text = Trim(Me.txt标本部位.Tag)
        Else
            Me.txt标本部位.Text = ""
            Me.txt标本部位.SetFocus
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd参考_Click()
    Dim rsTmp As ADODB.Recordset

    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
    Else
        MsgBox "没有找到可参考的项目。", vbInformation, Me.Caption
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer

    strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngClassId)

    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = "Select 0 As 末级,ID,上级ID,编码,名称,'' As 说明 From 诊疗参考分类 a" & _
            " Where 类型=" & iAttr & _
            " Start With a.上级id Is Null Connect By Prior a.id=a.上级id " & _
            " Union All" & _
            " Select 1,ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=" & iAttr & " Order By 编码"
    Else
        strSQLItem = " From 诊疗参考目录 A,诊疗参考别名 B" & _
            " Where A.ID=B.参考目录ID And A.类型=" & iAttr & _
            " And (Upper(A.编码) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.简码) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = "Select Distinct 0 As 末级,ID,上级ID,编码,名称,'' As 说明 From 诊疗参考分类 a" & _
            " Where 类型=" & iAttr & _
            " Start With ID In (Select 分类ID " & strSQLItem & ") Connect By Prior a.上级id=a.id " & _
            " Union All" & _
            " Select Distinct 1,A.ID,A.分类ID,A.编码,A.名称,A.说明 " & strSQLItem & " Order By 编码"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 2, "参考", , , , , True)
End Function

Private Sub cmd分类_Click()
    With Me.tvwClass
        .Left = Me.txt分类.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub InitVsf()
    '初始化输血检查对照表
    With vsfBloodLis
        .Cols = 2
        .ColHidden(0) = True '隐藏第一列
        .ExtendLastCol = True '最后一列填充满表格
        .ColComboList(1) = "|..."
        .Editable = flexEDKbdMouse
        .AllowSelection = False '不能多选单元格
    End With
End Sub

Private Sub Form_Activate()
    Dim aTmp() As String
    Dim strTmp As String
    Dim i As Integer
    Dim n As Integer
        
    If mFromLoad And Not mbln连续增加 Then Exit Sub
    If Me.Tag = "增加" Or Me.Tag = "复制增加" Then chkGoOn.Visible = True
    mFromLoad = True
    
    stbInfo.TabVisible(3) = False
    
    Call GetDefineSize
    Call IniStationNo

    '提取执行项目的信息
    err = 0: On Error GoTo ErrHand

    '因类别装入可能导致其他特性改变，所以首先装入
    gstrSql = "select 类别 from 诊疗项目目录 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

    With rsTemp
        If .RecordCount > 0 Then
            For intCount = 0 To Me.cbo类别.ListCount - 1
                If Left(Me.cbo类别.List(intCount), 1) = IIf(IsNull(!类别), "", !类别) Then
                    Me.cbo类别.ListIndex = intCount: Exit For
                End If
            Next
        End If
    End With

    '装入类别外的其他性质
    gstrSql = "select A.编码,A.名称,执行频率,单独应用,计算方式,计算单位,适用性别,执行安排,服务对象,执行科室,操作类型,组合项目,标本部位,A.试管编码," & _
            "        建档时间,nvl(撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,参考目录ID,B.名称 As 参考名称,A.录入限量,A.执行标记,A.执行分类,A.计算规则,A.站点,A.计算系数,A.诊疗频率编码 " & _
            " from 诊疗项目目录 A,诊疗参考目录 B" & _
            " where A.参考目录ID=B.ID(+) And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    cbo执行分类.Visible = False
    lbl执行分类.Visible = False
    cbo执行分类.ListIndex = 0
    
    With rsTemp
        Me.txt项目编码.MaxLength = .Fields("编码").DefinedSize
        If .RecordCount > 0 Then
            Me.txt项目编码.Text = !编码
            Me.txt项目名称.Text = !名称
            Me.txt项目名称.Tag = !名称
            '护理类的执行频率可选为"0-可选频率;2-持续性"，需要单独处理
            If Left(Me.cbo类别.Text, 1) = "H" Then
                If Val(IIf(IsNull(!执行频率), 0, !执行频率)) = 2 Then
                    Me.cbo执行频率.ListIndex = 1
                Else
                    Me.cbo执行频率.ListIndex = 0
                End If
            Else
                Me.cbo执行频率.ListIndex = IIf(IsNull(!执行频率), 0, !执行频率)
            End If
            
            Call Init诊疗频率(NVL(!诊疗频率编码))
            
            Me.chk单独应用.Value = IIf(IsNull(!单独应用), 0, !单独应用)
            Me.chk单独应用.Tag = IIf(IsNull(!单独应用), 0, !单独应用)
            Me.cbo计算方式.ListIndex = IIf(IsNull(!计算方式), 0, !计算方式)
            Me.txt计算单位.Text = IIf(IsNull(!计算单位), "", !计算单位)
            Me.cbo计算规则.ListIndex = IIf(IsNull(!计算规则), 0, !计算规则)
            Me.cbo适用性别.ListIndex = IIf(IsNull(!适用性别), 0, !适用性别)
            Me.chk执行安排.Value = IIf(IsNull(!执行安排), 0, !执行安排)
            Me.chk检验组合.Value = IIf(IsNull(!组合项目), 0, !组合项目)
            Me.txt录入限量.Text = IIf(IsNull(!录入限量), "", !录入限量)
            Me.txtML.Text = IIf(IsNull(!计算系数), "", !计算系数)
            SetStationNo IIf(IsNull(!站点), "", !站点)
            Select Case !服务对象
            Case 4
                Me.chk服务对象(2).Value = 1:
                Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 0
            Case 3
                Me.chk服务对象(0).Value = 1: Me.chk服务对象(1).Value = 1
            Case 2
                Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 1
            Case 1
                Me.chk服务对象(0).Value = 1: Me.chk服务对象(1).Value = 0
            Case Else
                Me.chk服务对象(0).Value = 0: Me.chk服务对象(1).Value = 0
            End Select
            Me.opt执行部门(IIf(IsNull(!执行科室), 0, !执行科室)).Value = True
            If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "该项目于" & Format(!建档时间, "YYYY-MM-DD") & "建立。"
            Else
                Me.lblFound.Caption = ""
            End If
            If Left(Me.cbo类别.Text, 1) = "D" Then
                If IIf(IsNull(!组合项目), 0, !组合项目) = 0 Then
                    Me.opt检查部位(0).Value = True: Me.txt检查部位.Text = IIf(IsNull(!标本部位), "", !标本部位): Me.txt检查部位.Enabled = True
                Else
                    Me.opt检查部位(1).Value = True: Me.txt检查部位.Text = "": Me.txt检查部位.Enabled = False
                End If
            End If
            If Left(Me.cbo类别.Text, 1) = "C" Then
                Me.txt标本部位.Text = IIf(IsNull(!标本部位), "", !标本部位)
                Me.txt标本部位.Tag = Me.txt标本部位.Text
            End If
            Select Case Left(Me.cbo类别.Text, 1)
            Case "C", "D", "F", "G"     'C-检验, D-检查, F-手术, G-麻醉
                For intCount = 0 To Me.cbo操作类型.ListCount - 1
                    If Mid(Me.cbo操作类型.List(intCount), InStr(1, Me.cbo操作类型.List(intCount), "-") + 1) = IIf(IsNull(!操作类型), "", !操作类型) Then
                        If mLast操作类型 = "" Then
                            Me.cbo操作类型.ListIndex = intCount
                            mLast操作类型 = !操作类型
                            Exit For
                        End If
                    End If
                Next
                Me.chk检验组合 = NVL(!组合项目, 0)
                If Me.chk检验组合.Value = 1 Then
                    Me.chk单独应用.Enabled = False
                Else
                    Me.chk单独应用.Enabled = True
                End If
                If Me.chk服务对象(1).Value = 1 Then
                    Me.chk加收.Enabled = True
                    Me.chk加收.Value = NVL(!执行标记, 0)
                Else
                    Me.chk加收.Enabled = False
                    Me.chk加收.Value = 0
                End If
                
                '判断以前设置的是什么值
                If !执行分类 = "" Then
                    cbo病理类别.ListIndex = 0
                Else
                    For n = 1 To cbo病理类别.ListCount - 1
                        If Val(Mid(cbo病理类别.List(n), 1, InStr(1, cbo病理类别.List(n), "-") - 1)) = !执行分类 Then
                            cbo病理类别.ListIndex = n
                        End If
                    Next
                End If
                
            Case "E", "H"         'E-治疗, H-护理
                For intCount = 0 To Me.cbo操作类型.ListCount - 1
                    If Val(Left(Me.cbo操作类型.List(intCount), 1)) = Val(IIf(IsNull(!操作类型), "", !操作类型)) Then
                        Me.cbo操作类型.ListIndex = intCount: Exit For
                    End If
                Next
                If Me.cbo操作类型.ListIndex = 2 Then
                    Me.cbo分类说明.Text = IIf(IsNull(!标本部位), "", !标本部位)
                End If
                
                If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 2 Then
                    For intCount = 0 To Me.cbo执行分类.ListCount - 1
                        If Val(Me.cbo执行分类.List(intCount)) = Val(IIf(IsNull(!执行分类), "", !执行分类)) Then
                            Me.cbo执行分类.ListIndex = intCount: Exit For
                        End If
                    Next
                    
                    If Me.cbo执行分类.ListIndex = 1 Then
                        If Val(IIf(IsNull(!执行标记), "", !执行标记)) = 2 Then
                            Me.cbo输液类型.ListIndex = 1
                        Else
                            Me.cbo输液类型.ListIndex = 0
                        End If
                    End If
                    
                    cbo执行分类.Visible = True
                    lbl执行分类.Visible = True
                End If
                If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 8 Then
                    cboBloodType.ListIndex = Val(!执行分类 & "")
                    cboBloodType.Visible = True
                    lbl执行分类.Visible = True
                End If
                If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 9 Then
                    Call Load试管编码(NVL(!试管编码 & ""))
                    Me.lbl试管编码.Visible = True
                    Me.picTestTubeCode.Visible = True
                    Me.cboTestTubeCode.Visible = True
                End If
                If Left(Me.cbo类别.Text, 1) = "E" And Me.cbo操作类型.ListIndex = 1 Then
                    Me.chkNoTMSY.Value = IIf(2 = Val(!执行标记 & ""), 1, 0)
                    Me.chkYYPS.Value = IIf(5 = Val(!执行分类 & ""), 1, 0)
                End If
            Case "Z"            'Z-其他
                For intCount = 0 To Me.cbo操作类型.ListCount - 1
                    If Mid(Me.cbo操作类型.List(intCount), 1, InStr(1, Me.cbo操作类型.List(intCount), "-") - 1) = IIf(IsNull(!操作类型), "", !操作类型) Then
                        Me.cbo操作类型.ListIndex = intCount: Exit For
                    End If
                Next
                If Me.cbo操作类型.ListIndex = 12 Then
                    If IIf(IsNull(!标本部位), "", !标本部位) = "尿量" Then
                        chk尿量.Value = 1
                    End If
                End If
            Case "K"
                '将输血检验对照数据查询出来
                If Me.Tag = "修改" Or Me.Tag = "查阅" Or Me.Tag = "复制增加" Then
                    gstrSql = "Select ID, 名称 From 诊疗项目目录 Where ID In (Select 检验项目id From 输血检验对照 Where 项目id = [1])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
                    If rsTemp.RecordCount > 0 Then
                        With vsfBloodLis
                            .Rows = rsTemp.RecordCount + 1
                            For n = 1 To rsTemp.RecordCount
                                .TextMatrix(n, 0) = rsTemp!ID
                                .TextMatrix(n, 1) = rsTemp!名称
                                rsTemp.MoveNext
                            Next
                        End With
                    End If
                End If
            Case Else
            End Select

            Me.txt参考 = NVL(!参考名称): Me.txt参考.Tag = NVL(!参考目录ID): strRefer = Me.txt参考
        End If
    End With

    If Left(Me.cbo类别.Text, 1) = "F" Then
        gstrSql = "select I.ID,I.编码,I.手术类型,I.名称" & _
                " from 疾病编码目录 I,疾病诊断对照 R" & _
                " where I.ID=R.疾病ID and R.手术ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

        If rsTemp.RecordCount > 0 Then
            Me.txt标准编码.Tag = rsTemp!ID
            Me.txt标准编码.Text = IIf(IsNull(rsTemp!编码), "", rsTemp!编码)
            Me.lbl标准编码.Caption = IIf(IsNull(rsTemp!手术类型), "", "【" & rsTemp!手术类型 & "】") & IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
        End If
    End If

    gstrSql = "select 名称,性质,简码,码类 from 诊疗项目别名 where 诊疗项目ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

    With rsTemp
        Do While Not .EOF
            If !性质 = 1 And !码类 = 1 Then Me.txt名称拼音.Text = !简码
            If !性质 = 1 And !码类 = 2 Then Me.txt名称五笔.Text = !简码
            If !性质 = 9 And !码类 = 1 Then Me.txt其他别名.Text = !名称: Me.txt别名拼音.Text = !简码
            If !性质 = 9 And !码类 = 2 Then Me.txt其他别名.Text = !名称: Me.txt别名五笔.Text = !简码
            .MoveNext
        Loop
    End With
    
    '使用科室
    If Me.Tag <> "增加" Then
        gstrSql = "Select b.id,b.名称 from 诊疗适用科室 A,部门表 B Where a.项目ID=[1] And a.科室id=b.id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
        With rsTemp
            i = 0: n = 0
            Do While Not .EOF
                vsUseDept.TextMatrix(i, n) = rsTemp!名称 & ""
                vsUseDept.Cell(flexcpData, i, n) = rsTemp!名称 & ""
                vsUseDept.TextMatrix(i, n + 5) = rsTemp!ID
                vsUseDept.Cell(flexcpData, i, n + 5) = rsTemp!ID & ""
                mstr已选使用科室 = IIf(mstr已选使用科室 = "", "", mstr已选使用科室 & ";") & rsTemp!ID & "," & rsTemp!名称
                If i = vsUseDept.Rows - 1 And n = 4 Then
                    vsUseDept.AddItem ""
                End If
                If n = 4 Then
                    n = 0
                    i = i + 1
                Else
                    n = n + 1
                End If
                .MoveNext
            Loop
        End With
    End If

    gstrSql = "select R.病人来源,E.ID,E.名称" & _
            " from 诊疗执行科室 R,部门表 E" & _
            " where R.执行科室ID=E.ID and R.病人来源 in (1,2) and R.开单科室id is null and R.诊疗项目ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

    With rsTemp
        Do While Not .EOF
            If !病人来源 = 1 Then Me.txt门诊执行.Text = !名称: Me.txt门诊执行.Tag = !ID
            If !病人来源 = 2 Then Me.txt住院执行.Text = !名称: Me.txt住院执行.Tag = !ID
            .MoveNext
        Loop
    End With

    gstrSql = "select K.ID as 开单部门ID,K.编码 as 开单科室编码,K.名称 as 开单部门名称," & _
            "         E.ID as 执行部门ID,E.编码 as 执行科室编码,E.名称 as 执行部门名称" & _
            " from 诊疗执行科室 R,部门表 K,部门表 E" & _
            " where R.开单科室ID=K.ID(+) and R.执行科室ID=E.ID and nvl(R.病人来源,0)=0 and R.诊疗项目ID=[1] " & _
            " order by e.名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    

    With rsTemp
        Me.msf定向执行.Clear (1)

        Do While Not .EOF
            'If Me.msf定向执行.Rows - 1 < .AbsolutePosition Then Me.msf定向执行.Rows = Me.msf定向执行.Rows + 1

            If strTmp <> !执行部门名称 Then
                i = i + 1
                Me.msf定向执行.Rows = i + 1
                Me.msf定向执行.TextMatrix(i, col病人科室ID) = IIf(IsNull(!开单部门ID), "（所有部门）", !开单部门ID)
                msf定向执行.Cell(flexcpData, i, col病人科室ID) = msf定向执行.TextMatrix(i, col病人科室ID)
                Me.msf定向执行.TextMatrix(i, col病人科室) = IIf(IsNull(!开单部门ID), "（所有部门）", !开单部门名称)
                msf定向执行.Cell(flexcpData, i, col病人科室) = msf定向执行.TextMatrix(i, col病人科室)
                Me.msf定向执行.TextMatrix(i, col执行科室ID) = !执行部门ID
                msf定向执行.Cell(flexcpData, i, col执行科室ID) = msf定向执行.TextMatrix(i, col执行科室ID)
                Me.msf定向执行.TextMatrix(i, col执行科室) = !执行部门名称
                msf定向执行.Cell(flexcpData, i, col执行科室) = msf定向执行.TextMatrix(i, col执行科室)
            Else
                Me.msf定向执行.TextMatrix(i, col病人科室ID) = Me.msf定向执行.TextMatrix(i, col病人科室ID) & "," & !开单部门ID
                msf定向执行.Cell(flexcpData, i, col病人科室ID) = msf定向执行.TextMatrix(i, col病人科室ID)
                Me.msf定向执行.TextMatrix(i, col病人科室) = Me.msf定向执行.TextMatrix(i, col病人科室) & "," & !开单部门名称
                msf定向执行.Cell(flexcpData, i, col病人科室) = msf定向执行.TextMatrix(i, col病人科室)
            End If

            strTmp = !执行部门名称
            .MoveNext
        Loop
    End With

    '查询基础检验项目对应的检验指标
    If Left(Me.cbo类别.Text, 1) = "C" And Me.chk检验组合.Value = 0 And Me.Tag <> "增加" Then
        gstrSql = "Select A.*,B.ID,B.临床意义,B.中文名 " & _
            "From 检验项目 A,诊治所见项目 B,检验报告项目 C " & _
            "Where A.诊治项目ID=B.ID And B.ID=C.报告项目ID And C.诊疗项目ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

        With rsTemp
            If Not .EOF Then
                lngVItemID = !ID
                Me.txt英文.Text = "" & !缩写
'                Me.txt检验(0) = Nvl(!缩写)
'                Me.txt检验(1) = Nvl(!单位)
'                Me.cbo项目类型.ListIndex = Nvl(!项目类别, 1) - 1
'                Me.cbo结果类型.ListIndex = Nvl(!结果类型, 1) - 1
'                Me.txt检验(2) = TransFormula1(Nvl(!计算公式))
'                If Len(Nvl(!结果异常条件)) > 0 Then
'                    aTmp = Split(!结果异常条件, ";")
'                    Me.txt检验(3) = aTmp(0)
'                    If UBound(aTmp) > 0 Then Me.txt检验(4) = aTmp(1)
'                End If
'                Me.txt检验(5) = Nvl(!临床意义)
            End If
        End With
    End If

    If Me.Tag = "增加" Or Me.Tag = "复制增加" Then
        lngItemID = 0: lngVItemID = 0
        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then '诊疗项目编码递增模式
            gstrSql = "select nvl(max(编码),'0000000') as 编码" & _
                    " From 诊疗项目目录" & _
                    " Where 类别 >= 'A'"
'            If rsTemp.State = adStateOpen Then rsTemp.Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
'            Call SQLTest
            Me.txt项目编码.Text = zlCommFun.IncStr(rsTemp!编码)
        Else
            strTemp = Mid(Me.txt分类.Text, 2, InStr(1, Me.txt分类.Text, "]") - 2)
            
            gstrSql = "select nvl(max(编码),'0000000') as 编码" & _
                    " From 诊疗项目目录" & _
                    " Where 类别 >= 'A' and 编码 like [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%")
    
            err = 0: On Error Resume Next
            If rsTemp!编码 = "0000000" Then
                Me.txt项目编码.Text = zlCommFun.IncStr(strTemp & "000")
            Else
                Me.txt项目编码.Text = zlCommFun.IncStr(rsTemp!编码)
            End If
        End If

        '清除命名信息
        Me.txt项目名称.Text = "": Me.txt名称拼音.Text = "": Me.txt名称五笔.Text = ""
        Me.txt其他别名.Text = "": Me.txt别名拼音.Text = "": Me.txt别名五笔.Text = "": Me.txt英文.Text = ""
        Me.lblFound.Visible = False
        If Me.Tag = "增加" Then Me.txt参考 = "": Me.txt参考.Tag = "": strRefer = ""
    End If

'        If Me.Tag = "修改" Then
'            Me.chk检验组合.Enabled = False
'        End If

    If Me.Tag = "查阅" Then
        Me.cmdOK.Visible = False
        Me.cmdCancel.Caption = "关闭(&C)"
        Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False: Me.cbo类别.Enabled = False
        Me.txt项目编码.Enabled = False: Me.cbo操作类型.Enabled = False
        Me.txt项目名称.Enabled = False: Me.txt名称拼音.Enabled = False: Me.txt名称五笔.Enabled = False
        Me.txt其他别名.Enabled = False: Me.txt别名拼音.Enabled = False: Me.txt别名五笔.Enabled = False

        Me.cbo执行频率.Enabled = False: Me.chk单独应用.Enabled = False
        Me.cbo计算方式.Enabled = False: Me.txt计算单位.Enabled = False
        Me.cbo适用性别.Enabled = False: Me.chk执行安排.Enabled = False
        Me.fra检查部位.Enabled = False: Me.txtML.Enabled = False
        Me.txt参考.Enabled = False: Me.cmd参考.Enabled = False
        Me.fra标本部位.Enabled = False: Me.cmd标本.Enabled = False
        Me.cboZLPL.Enabled = False


        Me.chk服务对象(0).Enabled = False: Me.chk服务对象(1).Enabled = False: Me.chk服务对象(2).Enabled = False
        Me.fra执行部门.Enabled = False
        Me.fra标准编码.Enabled = False

        Me.chk检验组合.Enabled = False
        Me.txt英文.Enabled = False
'        Me.txt检验(0).Enabled = False: Me.txt检验(1).Enabled = False
'        Me.txt检验(2).Enabled = False: Me.txt检验(3).Enabled = False
'        Me.txt检验(4).Enabled = False: Me.txt检验(5).Enabled = False
'        Me.cbo结果类型.Enabled = False: Me.cbo项目类型.Enabled = False
        Me.txt录入限量.Enabled = False
        Me.cbo录入限量范围.Enabled = False
        Me.cbo分类说明.Enabled = False
        Me.cmbStationNo.Enabled = False
        Me.chk加收.Enabled = False
        For i = 0 To OptAppUse.Count - 1
            OptAppUse(i).Enabled = False
        Next
    End If

    '判断是否是检查组合中的部位项目
    mbln组合部位项目 = False
    If lngItemID <> 0 And Left(Me.cbo类别.Text, 1) = "D" Then
        '组合项目
        gstrSql = "Select 1 From 诊疗项目组合 Where 诊疗组合id = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

        mbln组合项目 = (rsTemp.RecordCount > 0)

        If Not mbln组合项目 Then
            '子项目
            gstrSql = "Select 1 From 诊疗项目目录 Where ID In (Select 诊疗组合id From 诊疗项目组合 Where 诊疗项目id = [1]) And 类别 = 'D'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

            mbln组合部位项目 = (rsTemp.RecordCount > 0)
            Me.opt检查部位(1).Enabled = Not mbln组合部位项目
        End If

    End If

    '对已经设置为检查组合中的子项目，以及包含子项目的主项目，不允许更改诊疗类别。
    If mbln组合项目 = True Or mbln组合部位项目 = True Then
        Me.cbo类别.Enabled = False
    End If
    
    If Me.Tag = "增加" Then
        Call cbo类别_Click
        Call cbo操作类型_Click
    End If
    
    If Me.Tag = "修改" And Left(Me.cbo类别.Text, 1) = "G" Then
        Me.chk单独应用.Value = 0: Me.chk单独应用.Enabled = False
    End If
    
    '是否连续增加
    chkGoOn.Value = Val(zlDatabase.GetPara("诊疗项目连续增加", glngSys, 1054, 0, Array(Me.chkGoOn), True))
    msf定向执行.AutoSize msf定向执行.FixedCols, msf定向执行.Cols - 1
    Call chk服务对象_Click(0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Me.picDept.Visible = True Then
            Call picDept_LostFocus
            Exit Sub
        End If
        If Me.tvwClass.Visible Then
            Me.tvwClass.Visible = False: Me.txt分类.SetFocus: Exit Sub
        End If
        Call cmdCancel_Click
        
    ElseIf KeyCode = vbKeyF3 Then
        If txtLocate.Enabled And txtLocate.Visible Then Call txtLocate_KeyPress(vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    mstrFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    mstr应用范围 = zlDatabase.GetPara("项目应用范围", glngSys, 1054, "000")
    
    With Me.msf定向执行
        .Editable = flexEDKbdMouse
        .FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, col执行科室ID) = "执行科室ID": .TextMatrix(0, col执行科室) = "执行科室"
        .TextMatrix(0, col病人科室ID) = "病人科室ID": .TextMatrix(0, col病人科室) = "病人科室"
        .colData(col执行科室ID) = 5: .colData(col执行科室) = 1: .colData(col病人科室ID) = 5: .colData(col病人科室) = 1
        .ColWidth(col执行科室ID) = 0: .ColWidth(col执行科室) = 1700: .ColWidth(col病人科室ID) = 0: .ColWidth(col病人科室) = 7300
        .Row = 1: .Col = 1
        .ColHidden(col执行科室ID) = True: .ColHidden(col病人科室ID) = True
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeBothUniform
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .RowHeightMin = 250
    End With
    vsUseDept.Editable = flexEDKbdMouse
    For i = 0 To OptAppUse.Count - 1
        If i = 0 Then
            OptAppUse(i).Enabled = True
        Else
            '根据参数来确定是否可用
            OptAppUse(i).Enabled = (Val(Mid(mstr应用范围, i, 1)) = 1)
        End If
    Next
    
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
    
    
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1500
        .Add , "编码", "编码", 900
        .Add , "类别", "类别", 0
    End With
    
    With Me.lvwItem
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    With cbo病理类别
        .Clear
        .AddItem "0-常规"
        .AddItem "1-冰冻"
        .AddItem "2-细胞"
        .AddItem "3-会诊"
        .AddItem "4-尸检"
        .AddItem "5-快速石蜡"
        .ListIndex = 0
    End With
        
        strSql = "select ID,名称 from 病理号码规则"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    If rsData.RecordCount > 0 Then
        With cbo病理类别
        .Clear
            rsData.MoveFirst
            Do While Not rsData.EOF
                If NVL(rsData!名称, "  ") <> "  " Then
                    .AddItem NVL(rsData!ID, 0) & "-" & rsData!名称
                End If
                rsData.MoveNext
            Loop
        .ListIndex = 0
        End With
    End If
    
    With Me.cbo录入限量范围
        .Clear
        .AddItem "本项目"
        .AddItem "本级"
        .AddItem "本分类"
        .AddItem "本类别"
        .AddItem "所有"
        .ListIndex = 0
    End With
    
    
    With Me.cbo执行分类
        .Clear
        .AddItem "0-其他"
        .AddItem "1-输液"
        .AddItem "2-注射"
        .AddItem "4-口服"
        .ListIndex = 0
    End With
    
    With Me.cbo输液类型
        .Clear
        .AddItem "0-常规"
        .AddItem "2-静脉营养"
        .ListIndex = 0
    End With
    
    With Me.cboBloodType '输血途径分类 0－备血，1－用血
        .Clear
        .AddItem "0-备血"
        .AddItem "1-用血"
        .ListIndex = 0
    End With
    
    Call InitVsf '初始化表格
    
    mstrMatch = gstrMatch
    strRefer = ""
    mLast操作类型 = ""
    mblnOK = False
    mlngFind = 1
    Ini性质分类
    Call Init诊疗频率
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr已选使用科室 = ""
    Call zlDatabase.SetPara("诊疗项目连续增加", chkGoOn.Value, glngSys, 1054)
    mFromLoad = False
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str病人科室ID As String
    Dim str病人科室名称 As String
    Dim strTmp As String
    Dim strArr
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
        
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Select Case .Tag
        Case "手术"
            Me.txt标准编码.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt标准编码.Text = .SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1)
            Me.lbl标准编码.Caption = .SelectedItem.Text
            Me.stbInfo.Tab = 1: Me.chk服务对象(0).SetFocus
        Case "门诊"
            Me.txt门诊执行.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt门诊执行.Text = .SelectedItem.Text
            Me.txt门诊执行.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "住院"
            Me.txt住院执行.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt住院执行.Text = .SelectedItem.Text
            Me.txt住院执行.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "开单"
            With Me.lvwItems
                If Me.msf定向执行.Col = col病人科室 And Me.lvwItems.Checkboxes = True Then
                    Me.msf定向执行.Text = ""
                    For i = 1 To .ListItems.Count
                        If .ListItems(i).Checked = True Then
                            If Me.msf定向执行.Text = "" Then
                                Me.msf定向执行.Text = .ListItems(i).Text
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) = Mid(.ListItems(i).Key, 2)
                                msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室ID) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID)
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室) = Me.msf定向执行.Text
                                msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室)
                            Else
                                Me.msf定向执行.Text = Me.msf定向执行.Text & "," & .ListItems(i).Text
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) = Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) & "," & Mid(.ListItems(i).Key, 2)
                                msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室ID) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID)
                                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室) = Me.msf定向执行.Text
                                msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室)
                            End If
                            m = m + 1
                        End If
                    Next
                    If m = 0 Then
                        Me.msf定向执行.Text = ""
                        Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) = "（所有部门）"
                        msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室ID) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID)
                        Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室) = "（所有部门）"
                        msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室)
                    End If
                Else
                    Me.msf定向执行.Text = .SelectedItem.Text
                    Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID) = Mid(.SelectedItem.Key, col病人科室ID)
                    msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室ID) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID)
                    Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室) = Me.msf定向执行.Text
                    msf定向执行.Cell(flexcpData, Me.msf定向执行.Row, col病人科室) = msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室)
                End If
            End With
            
            '如果有其他未填的行，询问是否按同一方案增加
            For i = 1 To Me.msf定向执行.Rows - 1
                If Me.msf定向执行.TextMatrix(i, col执行科室ID) <> "" And Me.msf定向执行.TextMatrix(i, col病人科室) = "" Then
                    blnBatch = True
                    Exit For
                End If
            Next
            
            If blnBatch = True Then
                If MsgBox("是否应用与其他未设置的列？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    str病人科室ID = Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室ID)
                    str病人科室名称 = Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col病人科室)
                    For i = 1 To Me.msf定向执行.Rows - 1
                        If Me.msf定向执行.TextMatrix(i, col病人科室) = "" Then
                            Me.msf定向执行.TextMatrix(i, col病人科室ID) = str病人科室ID
                            msf定向执行.Cell(flexcpData, i, col病人科室ID) = msf定向执行.TextMatrix(i, col病人科室ID)
                            Me.msf定向执行.TextMatrix(i, col病人科室) = str病人科室名称
                            msf定向执行.Cell(flexcpData, i, col病人科室) = msf定向执行.TextMatrix(i, col病人科室)
                        End If
                    Next
                End If
            End If
            
            Me.msf定向执行.SetFocus
            Call zlCommFun.PressKey(vbKeyReturn)
        Case "执行"
            
            If Val(Me.picDept.Tag) = 1 And lbl工作性质.Visible = True Then
                '删除不在已选择列表中的执行科室
                For i = msf定向执行.Rows - 1 To 1 Step -1
                    If InStr(mstr已选执行科室, msf定向执行.TextMatrix(i, col执行科室ID) & "," & msf定向执行.TextMatrix(i, col执行科室)) = 0 Then
                        If i > 1 Then
                            msf定向执行.RemoveItem i
                        Else
                            msf定向执行.TextMatrix(1, col执行科室ID) = ""
                            msf定向执行.Cell(flexcpData, 1, col执行科室ID) = ""
                            msf定向执行.TextMatrix(1, col执行科室) = ""
                            msf定向执行.Cell(flexcpData, 1, col执行科室) = ""
                            msf定向执行.TextMatrix(1, col病人科室ID) = ""
                            msf定向执行.Cell(flexcpData, 1, col病人科室ID) = ""
                            msf定向执行.TextMatrix(1, col病人科室) = ""
                            msf定向执行.Cell(flexcpData, 1, col病人科室) = ""
                            
                            If msf定向执行.Rows > 2 Then
                                msf定向执行.RemoveItem 1
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
                            If strArr(i) = msf定向执行.TextMatrix(n, col执行科室ID) & "," & msf定向执行.TextMatrix(n, col执行科室) Then
                                blnNew = False
                            End If
                        Next
                        If blnNew = True Then
                            strNew = IIf(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            If msf定向执行.TextMatrix(msf定向执行.Rows - 1, col执行科室) <> "" Then
                                msf定向执行.Rows = msf定向执行.Rows + 1
                            End If
                            msf定向执行.TextMatrix(msf定向执行.Rows - 1, col执行科室ID) = Split(strArr(i), ",")(0)
                            msf定向执行.Cell(flexcpData, msf定向执行.Rows - 1, col执行科室ID) = msf定向执行.TextMatrix(msf定向执行.Rows - 1, col执行科室ID)
                            msf定向执行.TextMatrix(msf定向执行.Rows - 1, col执行科室) = Split(strArr(i), ",")(1)
                            msf定向执行.Cell(flexcpData, msf定向执行.Rows - 1, col执行科室) = msf定向执行.TextMatrix(msf定向执行.Rows - 1, col执行科室)
                        End If
                    Next
                End If
                
                msf定向执行.Row = msf定向执行.Rows - 1
                Me.msf定向执行.SetFocus
                Call zlCommFun.PressKey(vbKeyRight)
            Else
                Me.msf定向执行.Text = .SelectedItem.Text
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col执行科室ID) = Mid(.SelectedItem.Key, 2)
                msf定向执行.Cell(flexcpData, msf定向执行.Row, col执行科室ID) = msf定向执行.TextMatrix(msf定向执行.Row, col执行科室ID)
                Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, col执行科室) = Me.msf定向执行.Text
                msf定向执行.Cell(flexcpData, msf定向执行.Row, col执行科室) = msf定向执行.TextMatrix(msf定向执行.Row, col执行科室)
                Me.msf定向执行.SetFocus
                Call zlCommFun.PressKey(vbKeyRight)
            End If
        Case "使用"
            Dim j As Long
            
            If Val(Me.picDept.Tag) = 2 And lbl工作性质.Visible = True Then
                '删除不在已选择列表中的使用科室
                For i = 0 To vsUseDept.Rows - 1
                    For j = 0 To 4
                        If InStr(mstr已选使用科室, vsUseDept.TextMatrix(i, j + 5) & "," & vsUseDept.TextMatrix(i, j)) = 0 Then
                            vsUseDept.TextMatrix(i, j) = ""
                            vsUseDept.Cell(flexcpData, i, j) = ""
                            vsUseDept.TextMatrix(i, j + 5) = ""
                            vsUseDept.Cell(flexcpData, i, j + 5) = ""
                        End If
                    Next
                Next
                
                '增加新执行科室
                mstr已选使用科室 = mstr已选使用科室 & ";"
                strArr = Split(mstr已选使用科室, ";")
                
                For i = 0 To UBound(strArr) - 1
                    blnNew = True
                    If strArr(i) <> "" Then
                        For n = 0 To vsUseDept.Rows - 1
                            For j = 0 To 4
                                If strArr(i) = vsUseDept.TextMatrix(n, j + 5) & "," & vsUseDept.TextMatrix(n, j) Then
                                    blnNew = False
                                End If
                            Next
                        Next
                        If blnNew = True Then
                            strNew = IIf(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            For n = 0 To vsUseDept.Rows - 1
                                For j = 0 To 4
                                    If n = vsUseDept.Rows - 1 And j = 4 Then vsUseDept.AddItem ""
                                    If vsUseDept.TextMatrix(n, j) = "" Then
                                        vsUseDept.TextMatrix(n, j) = Split(strArr(i), ",")(1)
                                        vsUseDept.Cell(flexcpData, n, j) = vsUseDept.TextMatrix(n, j)
                                        vsUseDept.TextMatrix(n, j + 5) = Split(strArr(i), ",")(0)
                                        vsUseDept.Cell(flexcpData, n, j + 5) = vsUseDept.TextMatrix(n, j + 5)
                                        n = vsUseDept.Rows - 1
                                        Exit For
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
                
                Me.vsUseDept.SetFocus
            Else
                If InStr(mstr已选使用科室, Mid(.SelectedItem.Key, 2) & "," & .SelectedItem.Text) > 0 Then
                    MsgBox "已经存在相同的使用科室了，请检查。", vbInformation, gstrSysName
                    Me.vsUseDept.TextMatrix(Me.vsUseDept.Row, vsUseDept.Col) = ""
                    vsUseDept.SetFocus
                Else
                    Me.vsUseDept.Text = .SelectedItem.Text
                    Me.vsUseDept.TextMatrix(Me.vsUseDept.Row, vsUseDept.Col + 5) = Mid(.SelectedItem.Key, 2)
                    vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col + 5) = vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5)
                    Me.vsUseDept.TextMatrix(Me.vsUseDept.Row, vsUseDept.Col) = Me.vsUseDept.Text
                    vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col) = vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col)
                    Me.vsUseDept.SetFocus
                    mstr已选使用科室 = IIf(mstr已选使用科室 = "", "", mstr已选使用科室 & ";") & Mid(.SelectedItem.Key, 2) & "," & Me.vsUseDept.Text
                    Call zlCommFun.PressKey(vbKeyRight)
                End If
            End If
        End Select
        
        DoEvents
        picDept.Visible = False
        txtFind.Text = ""
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.SelectedItem.Checked = False And KeyAscii = vbKeyReturn Then
            lvwItems.SelectedItem.Checked = Not lvwItems.SelectedItem.Checked
            Exit Sub
        End If
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        picDept.Visible = False
        txtFind.Text = ""
    End Select
End Sub


Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub lvwItems_GotFocus()
    If Me.lvwItems.Tag = "开单" Or Me.lvwItems.Tag = "执行" Or Me.lvwItems.Tag = "使用" Then
        Me.lvwItems.ToolTipText = "全选Ctrl+A；全清Ctrl+R"
    Else
        Me.lvwItems.ToolTipText = ""
    End If
End Sub


Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItem.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItem.SortOrder = IIf(Me.lvwItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItem.SortKey = ColumnHeader.Index - 1
        Me.lvwItem.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItem_DblClick()
    Dim i As Integer
    Dim m As Integer
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItem
        Select Case .Tag
        Case "手术"
            Me.txt标准编码.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt标准编码.Text = .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1)
'            Me.lbl标准编码.Caption = .SelectedItem.Text & vbCrLf & .SelectedItem.SubItems(.ColumnHeaders("类别").Index - 1)
            Me.lbl标准编码.Caption = IIf(.SelectedItem.SubItems(.ColumnHeaders("类别").Index - 1) = "", "", "【" & .SelectedItem.SubItems(.ColumnHeaders("类别").Index - 1) & "】") & .SelectedItem.Text
            
            Me.stbInfo.Tab = 0: Me.chk服务对象(0).SetFocus
        Case "门诊"
            Me.txt门诊执行.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt门诊执行.Text = .SelectedItem.Text
            Me.txt门诊执行.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "住院"
            Me.txt住院执行.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt住院执行.Text = .SelectedItem.Text
            Me.txt住院执行.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
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
                    Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 0) = Mid(.SelectedItem.Key, 2)
                    Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 1) = Me.msf定向执行.Text
                End If
            End With
            Me.msf定向执行.SetFocus
            Call zlCommFun.PressKey(vbKeyReturn)
        Case "执行"
            Me.msf定向执行.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 0) = Mid(.SelectedItem.Key, 2)
            Me.msf定向执行.TextMatrix(Me.msf定向执行.Row, 1) = Me.msf定向执行.Text
            Me.msf定向执行.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End Select
    End With
End Sub

Private Sub lvwItem_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
        If lvwItem.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItem_DblClick
    Case vbKeyEscape
        Call lvwItem_LostFocus
    End Select
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub msf定向执行_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub optList_Click(Index As Integer)
    Dim lngRow As Long
    For lngRow = vfgList.FixedRows To vfgList.Rows - 1
        If vfgList.RowData(lngRow) = 0 Then
            If Index = 1 Then
                vfgList.RowHidden(lngRow) = True
            Else
                vfgList.RowHidden(lngRow) = False
            End If
        End If
    Next
End Sub

Private Sub opt检查部位_Click(Index As Integer)
    If Me.opt检查部位(0).Value = True Then
        Me.txt检查部位.Enabled = True
    Else
        Me.txt检查部位.Enabled = False
    End If
End Sub
Private Sub opt检查部位_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        Me.stbInfo.Tab = 1: Me.chk服务对象(0).SetFocus
    End If
End Sub

Private Sub opt执行部门_Click(Index As Integer)
    If Me.opt执行部门(4).Value = True Then
        If Me.chk服务对象(0).Value = 1 Then
            Me.txt门诊执行.Enabled = True
            txt门诊执行.BackColor = vbWindowBackground
        Else
            Me.txt门诊执行.Enabled = False
            txt门诊执行.BackColor = vbButtonFace
        End If
        If Me.chk服务对象(1).Value = 1 Then
            Me.txt住院执行.Enabled = True
            txt住院执行.BackColor = vbWindowBackground
        Else
            Me.txt住院执行.Enabled = False
            txt住院执行.BackColor = vbButtonFace
        End If
        load性质分类 0
        txtLocate.Enabled = True
        fraDeptFind.Enabled = True
        txtLocate.BackColor = vbWindowBackground
        optDeptKind(0).Enabled = True
        optDeptKind(1).Enabled = True
        msf定向执行.Editable = flexEDKbdMouse
    Else
        Me.txt门诊执行.Enabled = False: Me.txt住院执行.Enabled = False
        txt门诊执行.BackColor = vbButtonFace: txt住院执行.BackColor = vbButtonFace
        txtLocate.Enabled = False
        fraDeptFind.Enabled = False
        txtLocate.BackColor = vbButtonFace
        optDeptKind(0).Enabled = False
        optDeptKind(1).Enabled = False
        msf定向执行.Editable = flexEDNone
    End If
End Sub

Private Sub opt执行部门_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub stbInfo_Click(PreviousTab As Integer)
    If Me.tvwClass.Visible Then Me.stbInfo.Tab = 0: Me.tvwClass.SetFocus: Exit Sub
    If Me.lvwItems.Visible Then Me.stbInfo.Tab = 1: Me.lvwItems.SetFocus: Exit Sub
    
    Select Case Me.stbInfo.Tab
    Case 0
        If Me.txt项目编码.Enabled Then Me.txt项目编码.SetFocus
    Case 1
        If Me.chk服务对象(1).Enabled Then Me.chk服务对象(1).SetFocus
        If Me.chk服务对象(0).Enabled Then Me.chk服务对象(0).SetFocus
    Case 2
        '检查项目 显示检查部位选择页
        If Me.vfgList.Enabled Then
            Me.vfgList.SetFocus
        End If
    Case 3
        txt皮试标注.SetFocus
    End Select
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub


Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
    
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picDept.Visible = False
        txtFind.Text = ""
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Call cmdFind_Click
End Sub

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long, lngCol As Long
    Dim strFind As String, BlnFind As Boolean
    Const col执行科室 = 3, col病人科室 = 1
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1
        If lngStart > msf定向执行.Rows - 1 Then
            MsgBox "已经查找到了最后一个科室。", vbInformation, Me.Caption
            lblLocate.Tag = 0
            Exit Sub
        End If
        strFind = IIf(gstrMatch <> "", "*", "") & UCase(txtLocate.Text) & "*"
        
        If optDeptKind(0).Value Then
            lngCol = col执行科室
        Else
            lngCol = col病人科室
        End If
        
        For i = lngStart To msf定向执行.Rows - 1
            If msf定向执行.TextMatrix(i, lngCol) Like strFind Or zlCommFun.SpellCode(msf定向执行.TextMatrix(i, lngCol)) Like strFind Then
                lblLocate.Tag = i
                msf定向执行.Select i, lngCol
                msf定向执行.ShowCell i, lngCol
                If msf定向执行.Visible Then msf定向执行.SetFocus
                BlnFind = True
                Exit For
            End If
        Next
        If Not BlnFind Then
            If Val(lblLocate.Tag & "") = 0 Then
                MsgBox "没有找到您查找的科室。", vbInformation, Me.Caption
            Else
                MsgBox "已经查找到了最后一个科室。", vbInformation, Me.Caption
                lblLocate.Tag = 0
            End If
        End If
    End If
End Sub

Private Sub txtML_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtML_LostFocus()
    Me.txtML.Text = FormatEx(Val(Me.txtML.Text), 5)
End Sub

Private Sub txt标本部位_GotFocus()
    Me.txt标本部位.SelStart = 0: Me.txt标本部位.SelLength = 100
End Sub

Private Sub txt标本部位_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim vPoint As POINTAPI
    Dim strName As String
    
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        strName = Trim(Me.txt标本部位.Text)
        If strName = "" Then Exit Sub

        strSql = "Select Rownum As ID, 编码, 名称, 简码 From 诊疗检验标本 " & _
               " Where (编码 Like '" & strName & "%'" & _
               " Or 名称 Like '" & mstrMatch & strName & "%'" & _
               " Or 简码 Like '" & mstrMatch & UCase(strName) & "%')"

        vPoint = zlControl.GetCoordPos(txt标本部位.hWnd, txt标本部位.Left - 165, txt标本部位.Top - 30)

        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "标本部位", , , , , True, True, vPoint.x, vPoint.y)

        If rsTmp.State = 0 Then
            Me.txt标本部位.Text = ""
            If Trim(Me.txt标本部位.Text) = "" And Trim(Me.txt标本部位.Tag) <> "" Then
                Me.txt标本部位.Text = Trim(Me.txt标本部位.Tag)
            End If
            Exit Sub
        End If
        If Not rsTmp Is Nothing Then
            Me.txt标本部位.Text = rsTmp("名称")
            Me.txt标本部位.Tag = rsTmp("名称")
        Else
            If Trim(Me.txt标本部位.Text) = "" And Trim(Me.txt标本部位.Tag) <> "" Then
                Me.txt标本部位.Text = Trim(Me.txt标本部位.Tag)
            Else
                Me.txt标本部位.Text = ""
                Me.txt标本部位.SetFocus
            End If
        End If

    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txt标本部位_Validate(Cancel As Boolean)
    txt标本部位.Text = txt标本部位.Tag
End Sub

Private Sub txt标准编码_GotFocus()
    Me.txt标准编码.SelStart = 0: Me.txt标准编码.SelLength = 100
End Sub

Private Sub txt标准编码_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Exit Sub
    End If
    If Trim(Me.txt标准编码.Text) = "" Then
        Me.txt标准编码.Tag = ""
        Me.txt标准编码.Text = ""
        Me.lbl标准编码.Caption = ""
        Me.stbInfo.Tab = 1: Me.chk服务对象(0).SetFocus
        Exit Sub
    End If

    err = 0: On Error GoTo ErrHand

    
    gstrSql = "select A.ID,A.编码,A.手术类型 手术类型,A.名称,A.简码" & _
            " from 疾病编码目录 A" & _
            " where A.类别='S'" & _
            "   and (A.编码 like [1] " & _
            "       OR A.简码 like [2] " & _
            "       OR A.名称 like [2]) and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txt标准编码.Text) & "%", gstrMatch & Trim(Me.txt标准编码.Text) & "%")

    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到指定手术标准编码", vbExclamation, gstrSysName
            Me.txt标准编码.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt标准编码.Tag = !ID
            Me.txt标准编码.Text = IIf(IsNull(!编码), "", !编码)
            Me.lbl标准编码.Caption = IIf(IsNull(!手术类型), "", "【" & NVL(!手术类型) & "】") & IIf(IsNull(!名称), "", !名称)
            Me.stbInfo.Tab = 1: Me.chk服务对象(0).SetFocus
            Exit Sub
        End If

        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !名称, "expend", "expend")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItem.ColumnHeaders("类别").Index - 1) = NVL(!手术类型)
            .MoveNext
        Loop
        With Me.lvwItem
            .ListItems(1).Selected = True
            .Tag = "手术"
            .Left = Me.stbInfo.Left + Me.fra标准编码.Left + Me.fra标准编码.Width - .Width
            .Top = Me.stbInfo.Top + Me.fra标准编码.Top + Me.txt标准编码.Top + Me.txt标准编码.Height
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt别名拼音_GotFocus()
    Me.txt别名拼音.SelStart = 0: Me.txt别名拼音.SelLength = 100
End Sub

Private Sub txt别名拼音_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt别名五笔_GotFocus()
    Me.txt别名五笔.SelStart = 0: Me.txt别名五笔.SelLength = 100
End Sub

Private Sub txt别名五笔_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
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
                MsgBox "没有找到可参考的项目。", vbInformation, Me.Caption
            Else
                Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
            End If
            If Left(Me.cbo类别.Text, 1) = "D" Then
                Call zlCommFun.PressKey(vbKeyTab)
            ElseIf Me.fra标准编码.Visible Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            If Left(Me.cbo类别.Text, 1) = "D" Then
                Call zlCommFun.PressKey(vbKeyTab)
            ElseIf Me.fra标准编码.Visible Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                Me.stbInfo.Tab = 1
            End If
        End If
        Exit Sub
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt参考_LostFocus()
    If Me.txt参考 <> strRefer And Me.txt参考.Text <> "" Then
        Me.txt参考 = strRefer
    End If
    
    If Me.txt参考.Text = "" Then
        Me.txt参考.Tag = ""
    End If
End Sub

Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt计算单位_GotFocus()
    Me.txt计算单位.SelStart = 0: Me.txt计算单位.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt计算单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt计算单位_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt检查部位_GotFocus()
    Me.txt检查部位.SelStart = 0: Me.txt检查部位.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt检查部位_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Me.stbInfo.Tab = 1: Me.chk服务对象(0).SetFocus
    End If
End Sub

Private Sub txt检查部位_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub


Private Sub txt录入限量_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt录入限量_LostFocus()
    Me.txt录入限量.Text = FormatEx(Val(Me.txt录入限量.Text), 5)
End Sub


Private Sub txt门诊执行_GotFocus()
    Me.txt门诊执行.SelStart = 0: Me.txt门诊执行.SelLength = 100
End Sub

Private Sub txt门诊执行_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txt门诊执行.Text) = "" Then Me.txt门诊执行.Tag = "": Me.txt门诊执行.Text = "": Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txt门诊执行.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct ID,编码,名称" & _
            " from 部门表 D,部门性质说明 T" & _
            " where D.ID=T.部门ID and T.服务对象 in (1,2,3)" & _
            "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
            " order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%")
    
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName:  Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt门诊执行.Tag = !ID: Me.txt门诊执行.Text = !名称: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.fra执行部门.Left + Me.txt门诊执行.Left
        .Top = Me.fra执行部门.Top + Me.txt门诊执行.Top + Me.txt门诊执行.Height
        
        lbl工作性质.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        txtFind.Visible = False
        cmdFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
        
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

Private Sub txt名称拼音_GotFocus()
    Me.txt名称拼音.SelStart = 0: Me.txt名称拼音.SelLength = 100
End Sub

Private Sub txt名称拼音_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称五笔_GotFocus()
    Me.txt名称五笔.SelStart = 0: Me.txt名称五笔.SelLength = 100
End Sub

Private Sub txt名称五笔_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt皮试标注_KeyPress(KeyAscii As Integer)
    If InStr("();,'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt皮试文字_KeyPress(KeyAscii As Integer)
    If InStr("();,'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt其他别名_GotFocus()
    Me.txt其他别名.SelStart = 0: Me.txt其他别名.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt其他别名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt其他别名.Text = MoveSpecialChar(txt其他别名.Text)
        Me.txt别名拼音.Text = zlStr.GetCodeByORCL(Me.txt其他别名.Text, False, mlng简码长度)
        Me.txt别名五笔.Text = zlStr.GetCodeByORCL(Me.txt其他别名.Text, True, mlng简码长度)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
'    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt其他别名_LostFocus()
    Me.txt别名拼音.Text = zlStr.GetCodeByORCL(Me.txt其他别名.Text, False, mlng简码长度)
    Me.txt别名五笔.Text = zlStr.GetCodeByORCL(Me.txt其他别名.Text, True, mlng简码长度)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt项目编码_GotFocus()
    Me.txt项目编码.SelStart = 0: Me.txt项目编码.SelLength = 100
End Sub

Private Sub txt项目编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt项目名称_GotFocus()
    Me.txt项目名称.SelStart = 0: Me.txt项目名称.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt项目名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt项目名称.Text = MoveSpecialChar(txt项目名称.Text)
        Me.txt名称拼音.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, False, mlng简码长度)
        Me.txt名称五笔.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, True, mlng简码长度)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
'    If InStr(" ~!@#$%^&*_+|=`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt项目名称_LostFocus()
    If mbln组合部位项目 And Me.txt项目名称.Text <> Me.txt项目名称.Tag Then
        Me.txt项目名称.Text = Me.txt项目名称.Tag
        Me.txt名称拼音.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, False, mlng简码长度)
        Me.txt名称五笔.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, True, mlng简码长度)
        MsgBox "该项目是检查组合中的部位项目，不能修改名称。"
        Exit Sub
    End If
    Me.txt名称拼音.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, False, mlng简码长度)
    Me.txt名称五笔.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, True, mlng简码长度)
    Call zlCommFun.OpenIme(False)
End Sub
Private Sub txt住院执行_GotFocus()
    Me.txt住院执行.SelStart = 0: Me.txt住院执行.SelLength = 100
End Sub

Private Sub txt住院执行_KeyPress(KeyAscii As Integer)
    Dim objItem As ListItem
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txt住院执行.Text) = "" Then Me.txt住院执行.Tag = "": Me.txt住院执行.Text = "": Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txt住院执行.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
   
    gstrSql = "select distinct ID,编码,名称" & _
            " from 部门表 D,部门性质说明 T" & _
            " where D.ID=T.部门ID and T.服务对象 in (1,2,3)" & _
            "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])" & _
            " order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%")
        
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt住院执行.Tag = !ID: Me.txt住院执行.Text = !名称: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.fra执行部门.Left + Me.txt住院执行.Left
        .Top = Me.fra执行部门.Top + Me.txt住院执行.Top + Me.txt住院执行.Height
        
        lbl工作性质.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        txtFind.Visible = False
        cmdFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
        
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

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function TransFormula(ByVal strFormula As String, strErrorMsg As String, Optional iErrorPos As Integer = 1) As String
'将检验项目计算公式转换为以ID标识的公式

    Dim i As Integer, strTmp As String, iElementStart As Integer, strCalcForm As String
    Dim strElement As String, strSql As String, rsTmp As New ADODB.Recordset

    On Error GoTo DBError
    strErrorMsg = "": iErrorPos = 1: TransFormula = ""
    strCalcForm = ""
    For i = 1 To Len(strFormula)
        strTmp = Mid(strFormula, i, 1)
        If iElementStart > 0 Then
            '已找到元素的开始位置
            If strTmp = "]" Then
                strElement = Trim(Mid(strFormula, iElementStart + 1, i - iElementStart - 1))
                strSql = "Select 诊治项目ID,nvl(项目类别,1),nvl(结果类型,1) From 检验项目" & _
                    " Where 缩写=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(strElement))

                If rsTmp.EOF Then
                    strErrorMsg = "计算公式未找到检验项目：" & strElement & "。"
                    iErrorPos = iElementStart + 1
                    TransFormula = ""
                    Exit Function
                End If
                If rsTmp(1) <> 1 Then
                    strErrorMsg = "检验项目：" & strElement & " 不是基础项目！"
                    iErrorPos = iElementStart + 1
                    TransFormula = ""
                    Exit Function
                End If
                If rsTmp(2) <> 1 Then
                    strErrorMsg = "检验项目：" & strElement & " 不是数字型！"
                    iErrorPos = iElementStart + 1
                    TransFormula = ""
                    Exit Function
                End If

                TransFormula = TransFormula & "[" & rsTmp(0) & "]"
                strCalcForm = strCalcForm & "1" '计算公式的模拟数为1
                iElementStart = 0
            End If
        Else
            If strTmp = "[" Then
                iElementStart = i
            Else
                TransFormula = TransFormula & strTmp
                strCalcForm = strCalcForm & strTmp
            End If
        End If
    Next
    '校验公式的语法是否正确
    strSql = "Select " & strCalcForm & " From Dual"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    On Error GoTo ValidError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    TransFormula = ""
    Call SaveErrLog
    Exit Function
ValidError:
    If gcnOracle.Errors(0).NativeError <> 1476 Then '忽略除数为0
        strErrorMsg = "计算公式语法错：" & Mid(err.Description, InStr(err.Description, ":") + 1)
        iErrorPos = 1
        TransFormula = ""
    End If
End Function

Private Function TransFormula1(ByVal strFormula As String) As String
'将以ID标识的公式转换为以缩写标识的公式

    Dim i As Integer, strTmp As String, iElementStart As Integer
    Dim strElement As String, strSql As String, rsTmp As New ADODB.Recordset

    On Error GoTo DBError
    TransFormula1 = ""
    For i = 1 To Len(strFormula)
        strTmp = Mid(strFormula, i, 1)
        If iElementStart > 0 Then
            '已找到元素的开始位置
            If strTmp = "]" Then
                strElement = Trim(Mid(strFormula, iElementStart + 1, i - iElementStart - 1))
                strSql = "Select 缩写 From 检验项目" & _
                    " Where 诊治项目ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strElement)

                If rsTmp.EOF Then
                    TransFormula1 = TransFormula1 & "[未知项目]"
                Else
                    TransFormula1 = TransFormula1 & "[" & UCase(NVL(rsTmp(0))) & "]"
                End If

                iElementStart = 0
            End If
        Else
            If strTmp = "[" Then
                iElementStart = i
            Else
                TransFormula1 = TransFormula1 & strTmp
            End If
        End If
    Next
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    TransFormula1 = ""
    Call SaveErrLog
End Function

Private Function CalcExist(ByVal lngItemID As Long) As String
'判断指定的项目是否被其他计算项目引用，返回引用的项目名称
    Dim strSql As String, rsTmp As New ADODB.Recordset

    On Error GoTo DBError
    CalcExist = ""
    strSql = "Select a.中文名,b.缩写 From 诊治所见项目 a,检验项目 b" & _
        " Where a.id=b.诊治项目id And b.项目类别=3 And b.计算公式 Like [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "%[" & lngItemID & "]%")

    If Not rsTmp.EOF Then CalcExist = rsTmp(0)
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initVfgList()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, rsZlxm As ADODB.Recordset
    Dim strTemp As String
    Dim strItem As String
    Dim strName As String
    Dim varTemp As Variant
    Dim i As Long
    
    On Error GoTo ErrHandle
    With vfgList
        '初始化表格
        .Clear
        .FixedCols = 0: .FixedRows = 1
        .Rows = 1: .Cols = 7
        
        .MergeRow(0) = True
        .MergeCellsFixed = flexMergeRestrictColumns
        
        .MergeCol(0) = True: .MergeCol(1) = True
        .MergeCells = flexMergeRestrictColumns
        
        .RowHeightMin = 300
        .ColWidthMin = 450
        If cbo类别.Text = "D-检查" And cbo操作类型.Text = "18-病理" Then
            .TextMatrix(0, 0) = "标本名称": .TextMatrix(0, 1) = "标本名称": .TextMatrix(0, 2) = "材料类别": .TextMatrix(0, 3) = "备注"
        Else
            .TextMatrix(0, 0) = "部位": .TextMatrix(0, 1) = "部位": .TextMatrix(0, 2) = "方法": .TextMatrix(0, 3) = "备注"
        End If
        .TextMatrix(0, 4) = "默认": .TextMatrix(0, 5) = "使用": .TextMatrix(0, 6) = "唯一项"
        .ColKey(0) = "分组": .ColKey(1) = "名称": .ColKey(2) = "方法": .ColKey(3) = "备注"
        .ColKey(4) = "默认": .ColKey(5) = "使用": .ColKey(6) = "唯一项"
        
        .ColHidden(.ColIndex("分组")) = False: .ColHidden(.ColIndex("名称")) = False
        .ColHidden(.ColIndex("方法")) = False: .ColHidden(.ColIndex("备注")) = False
        .ColHidden(.ColIndex("默认")) = True: .ColHidden(.ColIndex("使用")) = True
        .ColHidden(.ColIndex("唯一项")) = True
        
        .ColWidth(.ColIndex("分组")) = 950: .ColWidth(.ColIndex("名称")) = 450: .ColWidth(.ColIndex("方法")) = 5000
        .ColWidth(.ColIndex("备注")) = 1800: .ColWidth(.ColIndex("默认")) = 0: .ColWidth(.ColIndex("使用")) = 0
        .ColWidth(.ColIndex("唯一项")) = 0
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .WordWrap = True
        .AutoResize = True
        
        
        .Editable = flexEDKbdMouse
        
        '提取数据,填充表格
        '       2007-07-09 1.新增的部位排序; 2.无方法的部位不能使用。
        strSql = "Select a.编码, a.分组, a.名称, a.方法, a.备注" & vbNewLine & _
                "From 诊疗检查部位 a " & vbNewLine & _
                "Where a.方法 Is Not Null And a.类型=[1] Order by a.分组, a.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption & " 提取检查部位", Mid(cbo操作类型.Text, InStr(1, cbo操作类型.Text, "-") + 1))
        
        '更新已设置的方法
        strSql = "Select A.项目id, A.类型, A.部位, A.方法, A.默认, Decode(Nvl(B.收费项目id, 0), 0, 0, 1) As 使用,A.上级方法 " & vbNewLine & _
                "From 诊疗项目部位 A, 诊疗收费关系 B" & vbNewLine & _
                "Where A.部位 = B.检查部位(+) And A.方法 = B.检查方法(+) And A.项目id = B.诊疗项目id(+)" & vbNewLine & _
                " And instr([2],A.类型)>0 And A.项目ID=[1] order by id"

        Set rsZlxm = zlDatabase.OpenSQLRecord(strSql, Me.Caption & " 诊疗项目部位", lngItemID, cbo操作类型.Text)
        
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("分组")) = "" & rsTmp.Fields("分组")
            .TextMatrix(.Rows - 1, .ColIndex("名称")) = rsTmp.Fields("名称")
            .RowData(.Rows - 1) = 0
            Set .Cell(flexcpPicture, .Rows - 1, .ColIndex("名称")) = imgList.ListImages(4).Picture
            
            .Cell(flexcpData, .Rows - 1, .ColIndex("方法")) = "" & rsTmp.Fields("方法")
            .TextMatrix(.Rows - 1, .ColIndex("方法")) = Anslyze_MethodString("" & rsTmp.Fields("方法"))
            
            If InStr(2, rsTmp.Fields("方法"), vbTab) = 0 And InStr(2, rsTmp.Fields("方法"), ";") = 0 Then
                .TextMatrix(.Rows - 1, .ColIndex("唯一项")) = "1"
            End If
            
            strName = ""
            rsZlxm.Filter = ""
            rsZlxm.Filter = " 部位='" & "" & rsTmp.Fields("名称") & "' and 默认=1"
            .TextMatrix(.Rows - 1, .ColIndex("默认")) = ""
            varTemp = Split(.TextMatrix(.Rows - 1, .ColIndex("方法")), "  ")
            If rsZlxm.RecordCount > 0 Then
                For i = 0 To UBound(varTemp)
                    strItem = varTemp(i)
                    rsZlxm.MoveFirst
                    Do Until rsZlxm.EOF
                        If InStr(varTemp(i), rsZlxm.Fields("方法")) > 0 And varTemp(i) <> "" Then
                            If InStr(varTemp(i), "〈") > 0 Then
                                strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), "〈") - 1)
                                If "" & rsZlxm.Fields("上级方法") <> "" Then
                                    If InStr(strTemp, rsZlxm.Fields("上级方法")) > 0 Then
                                        strItem = Replace(strItem, "□" & rsZlxm.Fields("方法"), "■" & rsZlxm.Fields("方法"))
                                    End If
                                Else
                                    If InStr(varTemp(i), "" & rsZlxm.Fields("方法")) > 0 Then
                                        strItem = Replace(strItem, "□" & rsZlxm.Fields("方法"), "■" & rsZlxm.Fields("方法"))
                                        strItem = Replace(strItem, "○" & rsZlxm.Fields("方法"), "●" & rsZlxm.Fields("方法"))
                                    End If
                                End If
                            Else
                                If InStr(varTemp(i), "" & rsZlxm.Fields("方法")) > 0 Then
                                    strItem = Replace(strItem, "□" & rsZlxm.Fields("方法"), "■" & rsZlxm.Fields("方法"))
                                    strItem = Replace(strItem, "○" & rsZlxm.Fields("方法"), "●" & rsZlxm.Fields("方法"))
                                End If
                            End If
                        End If
                        rsZlxm.MoveNext
                    Loop
                    If strItem <> "" Then strName = strName & "  " & strItem
                Next
                .TextMatrix(.Rows - 1, .ColIndex("方法")) = strName
                rsZlxm.MoveFirst
                Do While Not rsZlxm.EOF
                    If rsZlxm.Fields("上级方法") <> "" Then
                        .TextMatrix(.Rows - 1, .ColIndex("默认")) = .TextMatrix(.Rows - 1, .ColIndex("默认")) & "" & rsZlxm.Fields("上级方法") & "〈□" & rsZlxm.Fields("方法") & ","
                    Else
                        .TextMatrix(.Rows - 1, .ColIndex("默认")) = .TextMatrix(.Rows - 1, .ColIndex("默认")) & "" & rsZlxm.Fields("方法") & ","
                    End If
                    rsZlxm.MoveNext
                Loop
            End If
            
            If InStr(.TextMatrix(.Rows - 1, .ColIndex("默认")), ",") > 0 Then
                .TextMatrix(.Rows - 1, .ColIndex("默认")) = Mid(.TextMatrix(.Rows - 1, .ColIndex("默认")), 1, Len(.TextMatrix(.Rows - 1, .ColIndex("默认"))) - 1)
            End If
            
            rsZlxm.Filter = ""
            rsZlxm.Filter = " 部位='" & "" & rsTmp.Fields("名称") & "'"
            If Not rsZlxm.EOF Then
                If .RowData(.Rows - 1) = 0 Then
                    .RowData(.Rows - 1) = 1
                    Set .Cell(flexcpPicture, .Rows - 1, .ColIndex("名称")) = imgList.ListImages(5).Picture
                End If
                .TextMatrix(.Rows - 1, .ColIndex("使用")) = rsZlxm.Fields("使用")
            Else
                If optList(1).Value Then
                    .RowHidden(.Rows - 1) = True
                End If
            End If
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = "" & rsTmp.Fields("备注")
            
            rsTmp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 1, .ColIndex("名称")
        .AutoSize 3, .ColIndex("备注")
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .ColIndex("分组")
        .AutoSize 2, .ColIndex("方法")
        
        .RowHeight(0) = 350
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function Anslyze_MethodString(ByVal strMethod As String) As String
    '分析检查方法串
    '   strMethod:方法串
    '   返回,格式化好的方法串
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long, lngCount As Long
    
    strItems = ""
    strMethod = Replace(strMethod, vbTab, ";" & vbTab)
    
    aryItem() = Split(strMethod, ";")
    For lngCount = 0 To UBound(aryItem)
        strTemp = aryItem(lngCount)
        If strTemp <> "" Then
            If InStr(strTemp, vbTab) >= 1 Then
                strTemp = Mid(aryItem(lngCount), 3)
                If InStr(1, strTemp, ",") > 0 Then
                    aryChild = Split(strTemp, ",")
                    strTemp = ""
                    For lngChild = 1 To UBound(aryChild)
                        strTemp = strTemp & " □" & Mid(aryChild(lngChild), 2)
                    Next
                    strTemp = aryChild(0) & "〈" & Trim(strTemp) & "〉"
                End If
                strItems = strItems & "  □" & strTemp '用两个空格，方便后面截取
            Else
                strTemp = Mid(aryItem(lngCount), 2)
                If InStr(1, strTemp, ",") > 0 Then
                    aryChild = Split(strTemp, ",")
                    strTemp = ""
                    For lngChild = 1 To UBound(aryChild)
                        strTemp = strTemp & " □" & Mid(aryChild(lngChild), 2)
                    Next
                    strTemp = aryChild(0) & "〈" & Trim(strTemp) & "〉"
                End If
                strItems = strItems & "  ○" & strTemp '用两个空格，方便后面截取
            End If
        End If
    Next
    
    Anslyze_MethodString = strItems
End Function

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vfgList.ColIndex("方法") Then
        Cancel = True
    Else
        If vfgList.RowData(Row) = 1 Then
            vfgList.ColComboList(vfgList.ColIndex("方法")) = "..."
        Else
            vfgList.ColComboList(vfgList.ColIndex("方法")) = ""
            Cancel = True
        End If
    End If
End Sub

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI, strDefault As String
    Dim arrItem() As String, lngCount As Long
    Dim strTemp As String, strItem As String
    Dim varTemp As Variant
    Dim i As Long
    
    pt.x = vfgList.ColPos(Col) \ Screen.TwipsPerPixelX
    pt.y = (vfgList.RowPos(Row) + vfgList.RowHeight(Row)) \ Screen.TwipsPerPixelY
    ClientToScreen vfgList.hWnd, pt
    
    If InStr(vfgList.Cell(flexcpText, 0, Col), "方法") > 0 Then
        If vfgList.RowData(Row) = 1 Then
            With frmClinicDefaultModus
                 strDefault = vfgList.TextMatrix(Row, vfgList.ColIndex("默认"))
                .Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
                Call .ShowModus(vfgList.Cell(flexcpData, Row, Col), strDefault)
            End With
            With vfgList
                If strDefault <> .TextMatrix(Row, .ColIndex("默认")) Then
                    '更新显示
                    If .TextMatrix(Row, .ColIndex("默认")) <> "" Then
                        .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")), "■", "□")
                        .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")), "●", "○")
                    End If
                    If strDefault <> "" Then
                        arrItem = Split(strDefault, ",")
                        For lngCount = LBound(arrItem) To UBound(arrItem)
                            If InStr(arrItem(lngCount), "〈□") > 0 Then
                                varTemp = Split(.TextMatrix(Row, .ColIndex("方法")), "  ")
                                For i = 0 To UBound(varTemp)
                                    strItem = ""
                                    If InStr(varTemp(i), Split(arrItem(lngCount), "〈□")(0)) > 0 And InStr(varTemp(i), Split(arrItem(lngCount), "〈□")(1)) > 0 Then
                                        strItem = Replace(varTemp(i), "□" & Split(arrItem(lngCount), "〈□")(1), "■" & Split(arrItem(lngCount), "〈□")(1))
                                    End If
                                    If strItem <> "" Then .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")), varTemp(i), strItem)
                                Next
                            Else
                                .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")) & " ", "□" & arrItem(lngCount) & " ", "■" & arrItem(lngCount) & " ")
                                .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")) & " ", "□" & arrItem(lngCount) & "〈", "■" & arrItem(lngCount) & "〈")
                                .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")) & " ", "□" & arrItem(lngCount) & "〉", "■" & arrItem(lngCount) & "〉")
                                .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")) & " ", "○" & arrItem(lngCount) & " ", "●" & arrItem(lngCount) & " ")
                                .TextMatrix(Row, .ColIndex("方法")) = Replace(.TextMatrix(Row, .ColIndex("方法")) & " ", "○" & arrItem(lngCount) & "〈", "●" & arrItem(lngCount) & "〈")
                            End If
                        Next
                    End If
                    .TextMatrix(Row, .ColIndex("默认")) = strDefault
                End If
            End With
        End If
    End If
End Sub

Private Sub vfgList_EnterCell()
    With vfgList
        If (.Col = .ColIndex("名称") Or .Col = .ColIndex("方法")) And .Row > 0 Then
            On Error Resume Next
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
End Sub

Private Sub vfgList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        
        If vfgList.Col = vfgList.ColIndex("名称") Then
            Call SwapPic(vfgList.Row, vfgList.Col)
        End If
    ElseIf KeyCode = vbKeyReturn Then
        With vfgList
            If .Col = .ColIndex("方法") Then
                If .Row >= .FixedCols And .Row + 1 <= .Rows - 1 Then
                    .Select .Row + 1, .ColIndex("名称")
                End If
            ElseIf .Col < .Cols - 1 Then
                
                .Select .Row, .Col + 1
            
            End If
        End With
        
    End If
End Sub


Private Sub SwapPic(ByVal lngRow As Long, ByVal lngCol As Long)
    
    If Not cbo操作类型.Enabled Then Exit Sub
    With vfgList
        If .Col = .ColIndex("名称") And .Row > 0 And .Row < .Rows Then
            lngRow = .Row
            lngCol = .Col
            If .RowData(lngRow) = 0 Then
                .RowData(lngRow) = 1
                Set .Cell(flexcpPicture, lngRow, lngCol) = imgList.ListImages(5).Picture
                If .TextMatrix(lngRow, .ColIndex("唯一项")) = "1" Then
                    .TextMatrix(lngRow, .ColIndex("方法")) = Replace(.TextMatrix(lngRow, .ColIndex("方法")), "○", "●")
                    .TextMatrix(lngRow, .ColIndex("方法")) = Replace(.TextMatrix(lngRow, .ColIndex("方法")), "□", "■")
                    .TextMatrix(lngRow, .ColIndex("默认")) = Replace(.TextMatrix(lngRow, .ColIndex("方法")), "●", "")
                    .TextMatrix(lngRow, .ColIndex("默认")) = Replace(.TextMatrix(lngRow, .ColIndex("默认")), "■", "")
                End If
            Else
                If Val(.TextMatrix(lngRow, .ColIndex("使用"))) = 0 Then
                    .RowData(lngRow) = 0
                    Set .Cell(flexcpPicture, lngRow, lngCol) = imgList.ListImages(4).Picture
                    If .TextMatrix(lngRow, .ColIndex("唯一项")) = "1" Then
                        .TextMatrix(lngRow, .ColIndex("方法")) = Replace(.TextMatrix(lngRow, .ColIndex("方法")), "●", "○")
                        .TextMatrix(lngRow, .ColIndex("方法")) = Replace(.TextMatrix(lngRow, .ColIndex("方法")), "■", "□")
                        .TextMatrix(lngRow, .ColIndex("默认")) = ""
                    End If
                Else
                    MsgBox "该部位已使用，不能取消！", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
End Sub

Private Sub vfgList_LeaveCell()
    With vfgList
        If (.Col = .ColIndex("名称") Or .Col = .ColIndex("方法")) And .Row > 0 Then
            On Error Resume Next
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vfgList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
        If vfgList.Col = vfgList.ColIndex("名称") And Button = 1 And x > vfgList.CellLeft And x < vfgList.CellLeft + 250 Then
            Call SwapPic(vfgList.Row, vfgList.Col)
        End If
    
End Sub

Private Sub txt英文_GotFocus()
    Me.txt英文.SelStart = 0: Me.txt英文.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt英文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub vsfBloodLis_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mstrOldBlood = vsfBloodLis.TextMatrix(vsfBloodLis.Row, vsfBloodLis.Col)
End Sub

Private Sub vsfBloodLis_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intRow As Integer
    Dim str检验id As String
    
    vRect = zlControl.GetControlRect(vsfBloodLis.hWnd) '获取位置
    dblLeft = vRect.Left + vsfBloodLis.CellLeft
    dblTop = vRect.Top + vsfBloodLis.CellTop + vsfBloodLis.CellHeight + 3200
    
    With vsfBloodLis
        gstrSql = "Select ID, 分类id, 编码, 名称 From 诊疗项目目录 Where 类别 = 'C'  Order By ID"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSql, 0, "检验项目", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfBloodLis.Height, blnCancel, False, True)
        
        If Not rsTemp Is Nothing Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    str检验id = str检验id & "," & .TextMatrix(intRow, 0)
                End If
            Next
            If InStr(1, "," & str检验id & ",", "," & rsTemp!ID & ",") > 0 Then
                MsgBox "已经有该检验项目了，不需要再添加！", vbInformation, Me.Caption
            Else
                .TextMatrix(Row, 0) = rsTemp!ID
                .TextMatrix(Row, 1) = rsTemp!名称
            End If
        Else
            MsgBox "没有找到可选择的检验项目！", vbInformation, Me.Caption
        End If
    End With
End Sub

Private Sub GetBloodLis(ByVal strInput As String)
    '手动输入时，获取输血对照表的检验对照项目
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intRow As Integer
    Dim intCol As Integer
    Dim str检验id As String
    
    vRect = zlControl.GetControlRect(vsfBloodLis.hWnd) '获取位置
    dblLeft = vRect.Left + vsfBloodLis.CellLeft
    dblTop = vRect.Top + vsfBloodLis.CellTop + vsfBloodLis.CellHeight + 3200
        
    strInput = UCase(mstrFindStyle & strInput & "%")
    With vsfBloodLis
        gstrSql = "Select distinct a.Id, a.分类id, a.编码, a.名称" & vbNewLine & _
            "From 诊疗项目目录 A, 诊疗项目别名 B" & vbNewLine & _
            "Where a.Id = b.诊疗项目id And a.类别 = 'C' And (b.名称 Like [1] Or b.简码 Like [1] or a.编码 Like [1])"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "检验项目", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfBloodLis.Height, blnCancel, False, True, strInput)
        
        If Not rsTemp Is Nothing Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    str检验id = str检验id & "," & .TextMatrix(intRow, 0)
                End If
            Next
            If InStr(1, "," & str检验id & ",", "," & rsTemp!ID & ",") > 0 Then
                MsgBox "已经有该检验项目了，不需要再添加！", vbInformation, Me.Caption
                If .TextMatrix(.Row, 0) = rsTemp!ID Then
                    .EditText = mstrOldBlood
                    .TextMatrix(.Row, .Col) = mstrOldBlood
                Else
                    If .TextMatrix(.Row, 0) <> "" Then
                        .EditText = mstrOldBlood
                        .TextMatrix(.Row, .Col) = mstrOldBlood
                    Else
                        .EditText = ""
                        .TextMatrix(.Row, .Col) = ""
                    End If
                End If
            Else
                .TextMatrix(.Row, 0) = rsTemp!ID
                .TextMatrix(.Row, 1) = rsTemp!名称
            End If
        Else
            MsgBox "没有找到可选择的检验项目！", vbInformation, Me.Caption
            If .TextMatrix(intRow, 0) = "" Then
                .EditText = ""
                .TextMatrix(.Row, .Col) = ""
            Else
                .EditText = mstrOldBlood
                .TextMatrix(.Row, .Col) = mstrOldBlood
            End If
        End If
    End With
End Sub

Private Sub vsfBloodLis_EnterCell()
    With vsfBloodLis
        .Editable = flexEDNone
    End With
End Sub

Private Sub vsfBloodLis_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBloodLis
        If KeyCode = vbKeyReturn Then
            If .Row = .Rows - 1 Then
                If Me.Tag <> "查阅" Then
                    If .TextMatrix(.Row, 0) <> "" Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    Else
                        KeyCode = 0
                    End If
                End If
            Else
                .Row = .Row + 1
            End If
        ElseIf KeyCode = vbKeyDelete And Me.Tag <> "查阅" Then
            If .Row = .Rows - 1 And .Row = 1 Then
                .TextMatrix(.Row, 0) = ""
                .TextMatrix(.Row, 1) = ""
            Else
                .RemoveItem .Row
            End If
        End If
    End With
End Sub

Private Sub vsfBloodLis_KeyPress(KeyAscii As Integer)
    With vsfBloodLis
        If KeyAscii <> vbKeyReturn And Me.Tag <> "查阅" Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsfBloodLis_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfBloodLis
        If .EditText <> "" And KeyAscii = vbKeyReturn Then
            Call GetBloodLis(.EditText)
        End If
    End With
End Sub


Private Sub vsfBloodLis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With vsfBloodLis
            .Editable = flexEDNone
            If Me.Tag <> "查阅" Then
                If .Col = 1 Then
                    .Editable = flexEDKbdMouse
                End If
            End If
        End With
    End If
End Sub


Private Sub vsfFreq_DblClick()
    With vsfFreq
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .TextMatrix(.Row, 0) = "" Then
            .TextMatrix(.Row, 0) = "√"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
    End With
End Sub


Private Sub vsfFreq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    With vsfFreq
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .TextMatrix(.Row, 0) = "" Then
            .TextMatrix(.Row, 0) = "√"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
    End With
End Sub

Private Sub vsUseDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUseDept.Editable = flexEDNone Then
        vsUseDept.FocusRect = flexFocusLight
        vsUseDept.ComboList = ""
    Else
        vsUseDept.FocusRect = flexFocusSolid
        vsUseDept.ComboList = "..."
    End If
End Sub

Private Sub vsUseDept_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsUseDept.AutoSize vsUseDept.FixedCols, vsUseDept.Cols - 1
End Sub

Private Sub vsUseDept_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vsUseDept.TextMatrix(OldRow, OldCol) <> vsUseDept.Cell(flexcpData, OldRow, OldCol) Then
        vsUseDept.TextMatrix(OldRow, OldCol) = vsUseDept.Cell(flexcpData, OldRow, OldCol)
    End If
End Sub

Private Sub vsUseDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim i As Integer, j As Long
    
    mstr已选使用科室 = ""
    With Me.vsUseDept
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If .TextMatrix(i, j) <> "" And .ColHidden(j) = False Then
                    If InStr(mstr已选使用科室, .TextMatrix(i, j + 5) & "," & .TextMatrix(i, j)) = 0 Then
                        mstr已选使用科室 = IIf(mstr已选使用科室 = "", "", mstr已选使用科室 & ";") & .TextMatrix(i, j + 5) & "," & .TextMatrix(i, j)
                    End If
                End If
            Next
        Next
    End With
    
    With Me.picDept
        .Tag = "2"
        Me.lvwItems.Tag = "使用"
        .Left = Me.stbInfo.Left + Me.vsUseDept.Left + Me.vsUseDept.ColWidth(0) * Col
        .Width = 4000
        If .Left > Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left
    
        .Top = 1050 + Row * vsUseDept.RowHeight(Row)
        .Height = 3850
        
        lbl工作性质.Visible = True
        cboProperty.Visible = lbl工作性质.Visible
        ChkSelect.Visible = lbl工作性质.Visible
        
        lbl工作性质.Left = 50
        ChkSelect.Left = .Width - ChkSelect.Width - 50
        cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        .Left = lbl工作性质.Left
        .Top = cboProperty.Top + cboProperty.Height + 50 + txtFind.Height + 50
        .Width = Me.picDept.Width - .Left - 50
        .Height = Me.picDept.Height - .Top - 10
        txtFind.Top = cboProperty.Top + cboProperty.Height + 50
        cmdFind.Top = cboProperty.Top + cboProperty.Height + 50
        cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
        cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
        cmdFindOk.Top = cmdFind.Top
        cmdFindCancle.Top = cmdFind.Top
        
        .SetFocus
        .Refresh
    End With
    
    load性质分类 2
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsUseDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsUseDept_KeyPress(KeyCode)
    ElseIf KeyCode = vbKeyDelete Then
        If InStr(mstr已选使用科室, vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5) & "," & vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col)) > 0 Then
            mstr已选使用科室 = Replace(mstr已选使用科室, vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5) & "," & vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col), "")
        End If
        vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col) = ""
        vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col) = ""
        vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5) = ""
        vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col + 5) = ""
    End If
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.vsUseDept
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = 4 And .Row = .Rows - 1 Then
            .AddItem "": .Row = .Rows - 1: .Col = 0
        ElseIf .Col = 4 And .Row < .Rows - 1 Then
            .Row = .Row + 1: .Col = 0
        Else
            .Col = .Col + 1
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsUseDept_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, str站点 As String
    
    With Me.vsUseDept
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If Trim(.EditText) = "" Then
            .EditText = .Cell(flexcpData, Row, Col)
            Exit Sub
        End If
        strTemp = UCase(Trim(.EditText))
    End With
    If vsUseDept.EditText = vsUseDept.Cell(flexcpData, Row, Col) Then Exit Sub
    
    If strTemp = "" Then Exit Sub
    
    err = 0: On Error GoTo ErrHand
    If chk服务对象(1).Value = 1 Then strTmp = " T.服务对象=2"
    If chk服务对象(2).Value = 1 Or chk服务对象(0).Value = 1 Then strTmp = strTmp & IIf(strTmp = "", "", " Or") & " T.服务对象=1"
    If strTmp <> "" Then strTmp = " And (" & strTmp & " Or T.服务对象=3)"
    gstrSql = "select distinct ID,编码,名称" & _
            " from 部门表 D,部门性质说明 T" & _
            " where D.ID=T.部门ID " & _
            " and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) And T.服务对象<>0" & strTmp
    If cmbStationNo.Text <> "" Then
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        gstrSql = gstrSql & " And (D.站点=[2] Or D.站点 is Null)"
    End If
            
    gstrSql = gstrSql & " and 工作性质 In('临床','检查','检验','手术','治疗'" & IIf(chk服务对象(2).Value = 1, ",'体检'", "") & ") "
            
    
        
    gstrSql = gstrSql & " and (D.编码 like [1] or D.名称 like [1] or D.简码 like [1])"
            
    gstrSql = gstrSql & " order by 编码"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%", str站点)

    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "未找到指定部门，请重新输入！", vbExclamation, gstrSysName
            vsUseDept.TextMatrix(Row, Col) = vsUseDept.Cell(flexcpData, Row, Col)
            vsUseDept.EditText = vsUseDept.Cell(flexcpData, Row, Col)
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.vsUseDept.Text = !名称
            vsUseDept.EditText = Me.vsUseDept.Text
            vsUseDept.Cell(flexcpData, Row, Col) = Me.vsUseDept.Text
            vsUseDept.TextMatrix(Row, Col) = Me.vsUseDept.Text
            vsUseDept.TextMatrix(Row, Col + 5) = !ID
            vsUseDept.Cell(flexcpData, Row, Col + 5) = !ID
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码

            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Tag = "2"
        Me.lvwItems.Tag = "使用"
        .Left = Me.stbInfo.Left + Me.vsUseDept.Left + Me.vsUseDept.ColWidth(0) * Col
        .Width = 4000
        If .Left > Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left
    
        .Top = 1050 + Row * vsUseDept.RowHeight(Row)
        .Height = 3850
        
        lbl工作性质.Visible = False
        cboProperty.Visible = lbl工作性质.Visible
        ChkSelect.Visible = lbl工作性质.Visible
        
        lbl工作性质.Left = 50
        ChkSelect.Left = .Width - ChkSelect.Width - 50
        cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        
        cmdFind.Visible = False
        txtFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
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
        KeyCode = 0
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsUseDept_KeyPress(KeyAscii As Integer)
    If vsUseDept.Editable = flexEDNone Then Exit Sub

    With vsUseDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUseDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsUseDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If vsUseDept.Editable = flexEDNone Then
        vsUseDept.FocusRect = flexFocusLight
        vsUseDept.ComboList = ""
    Else
        vsUseDept.FocusRect = flexFocusSolid
        vsUseDept.ComboList = "..."
    End If
End Sub



