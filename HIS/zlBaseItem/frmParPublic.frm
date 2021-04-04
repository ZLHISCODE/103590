VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmParPublic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "公共参数设置"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12690
   Icon            =   "frmParPublic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   12690
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   0
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   10245
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10275
      Begin VB.Frame fraDevSvr 
         Caption         =   "三方服务配置"
         Height          =   2730
         Left            =   345
         TabIndex        =   195
         Top             =   4755
         Width           =   8265
         Begin VB.CommandButton cmdSvrChk 
            Caption         =   "服务验证"
            Height          =   300
            Left            =   6900
            TabIndex        =   196
            Top             =   2295
            Width           =   1260
         End
         Begin VSFlex8Ctl.VSFlexGrid vsThirdSvr 
            Height          =   1935
            Left            =   105
            TabIndex        =   197
            Top             =   285
            Width           =   8055
            _cx             =   14208
            _cy             =   3413
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
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483637
            GridColorFixed  =   16777215
            TreeColor       =   16777215
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParPublic.frx":6852
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
      Begin VB.CheckBox chk 
         Caption         =   "医疗机构不允许自由录入"
         Height          =   255
         Index           =   25
         Left            =   390
         TabIndex        =   182
         Top             =   4470
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "外院医生必须先建档"
         Height          =   255
         Index           =   22
         Left            =   5895
         TabIndex        =   170
         Top             =   4455
         Width           =   2175
      End
      Begin VB.CheckBox chk 
         Caption         =   "病人地址结构化录入"
         Height          =   255
         Index           =   50
         Left            =   480
         TabIndex        =   168
         Top             =   3795
         Width           =   2055
      End
      Begin VB.Frame fraSTAddress 
         Height          =   645
         Left            =   375
         TabIndex        =   167
         Top             =   3795
         Width           =   3975
         Begin VB.CheckBox chk 
            Caption         =   "乡镇级地址结构化录入"
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   390
            TabIndex        =   169
            Top             =   300
            Width           =   2535
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   15
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1980
         Width           =   2445
      End
      Begin VB.Frame fra补充录入 
         Caption         =   "补充录入限制"
         Height          =   1020
         Left            =   360
         TabIndex        =   11
         Top             =   2700
         Width           =   3975
         Begin VB.TextBox txtUD 
            Height          =   300
            Index           =   0
            Left            =   1550
            MaxLength       =   4
            TabIndex        =   13
            Top             =   300
            Width           =   650
         End
         Begin VB.CheckBox chk 
            Caption         =   "转病区病人只允许补录临嘱"
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   720
            Value           =   1  'Checked
            Width           =   2520
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   0
            Left            =   2220
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   300
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196637
            BuddyIndex      =   0
            OrigLeft        =   2470
            OrigTop         =   315
            OrigRight       =   2725
            OrigBottom      =   585
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblInputHours 
            AutoSize        =   -1  'True
            Caption         =   "小时"
            Height          =   180
            Left            =   2520
            TabIndex        =   148
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl补充录入 
            AutoSize        =   -1  'True
            Caption         =   "时限(0-9999)"
            Height          =   180
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   10
         ItemData        =   "frmParPublic.frx":68E1
         Left            =   1890
         List            =   "frmParPublic.frx":68E3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2340
         Width           =   2445
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "12"
         Top             =   480
         Width           =   380
      End
      Begin VB.Frame Fra 
         Caption         =   " 诊断输入 "
         Height          =   1860
         Index           =   12
         Left            =   5895
         TabIndex        =   25
         Top             =   2400
         Width           =   4095
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   16
            ItemData        =   "frmParPublic.frx":68E5
            Left            =   960
            List            =   "frmParPublic.frx":68E7
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1260
            Width           =   2535
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   8
            ItemData        =   "frmParPublic.frx":68E9
            Left            =   960
            List            =   "frmParPublic.frx":68EB
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   795
            Width           =   2535
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            ItemData        =   "frmParPublic.frx":68ED
            Left            =   960
            List            =   "frmParPublic.frx":68EF
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   330
            Width           =   2535
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院"
            Height          =   180
            Index           =   51
            Left            =   405
            TabIndex        =   30
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊"
            Height          =   180
            Index           =   27
            Left            =   405
            TabIndex        =   28
            Top             =   855
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "来源"
            Height          =   180
            Index           =   39
            Left            =   405
            TabIndex        =   26
            Top             =   390
            Width           =   360
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "输入全是数字时只查找编码"
         Height          =   195
         Index           =   10
         Left            =   1860
         TabIndex        =   6
         Top             =   1245
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.CheckBox chk 
         Caption         =   $"frmParPublic.frx":68F1
         Height          =   195
         Index           =   11
         Left            =   1860
         TabIndex        =   7
         Top             =   1530
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.Frame Fra 
         Caption         =   " 对外上下班时间 "
         Height          =   1635
         Index           =   1
         Left            =   5895
         TabIndex        =   16
         Top             =   360
         Width           =   4095
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   1005
            TabIndex        =   18
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.3541666667
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   2475
            TabIndex        =   20
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.5
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   2
            Left            =   1005
            TabIndex        =   22
            Top             =   915
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.5625
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   3
            Left            =   2475
            TabIndex        =   24
            Top             =   915
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.75
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Index           =   5
            Left            =   2100
            TabIndex        =   23
            Top             =   990
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "下午"
            Height          =   195
            Index           =   3
            Left            =   435
            TabIndex        =   21
            Top             =   975
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Index           =   4
            Left            =   2100
            TabIndex        =   19
            Top             =   495
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "上午"
            Height          =   180
            Index           =   2
            Left            =   435
            TabIndex        =   17
            Top             =   480
            Width           =   360
         End
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   9
         Left            =   2265
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtUD(9)"
         BuddyDispid     =   196637
         BuddyIndex      =   9
         OrigLeft        =   2205
         OrigTop         =   1200
         OrigRight       =   2460
         OrigBottom      =   1470
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医保对码检查"
         Height          =   180
         Index           =   42
         Left            =   720
         TabIndex        =   1
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗项目编码规则"
         Height          =   180
         Index           =   36
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "儿童年龄界定上限         岁"
         Height          =   180
         Index           =   47
         Left            =   360
         TabIndex        =   8
         Top             =   525
         Width           =   2430
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费项目和诊疗项目输入匹配方式"
         Height          =   180
         Index           =   40
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   2700
      End
   End
   Begin VB.PictureBox picPar 
      Height          =   7575
      Index           =   7
      Left            =   2400
      ScaleHeight     =   7515
      ScaleWidth      =   9675
      TabIndex        =   172
      Top             =   0
      Width           =   9735
      Begin VB.CheckBox chk 
         Caption         =   "启用医学影像信息系统专业版"
         Height          =   255
         Index           =   52
         Left            =   120
         TabIndex        =   173
         Top             =   240
         Width           =   2895
      End
      Begin TabDlg.SSTab sstRIS 
         Height          =   6495
         Left            =   120
         TabIndex        =   185
         Top             =   960
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11456
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "RIS分场合启动"
         TabPicture(0)   =   "frmParPublic.frx":690F
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "chkShowSel"
         Tab(0).Control(1)=   "vsfRisDepts"
         Tab(0).Control(2)=   "vsfRISEnables"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "HIS医院设置"
         TabPicture(1)   =   "frmParPublic.frx":692B
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label9"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame3"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.CheckBox chkShowSel 
            Caption         =   "只显示已启用的场合"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -74880
            TabIndex        =   191
            Top             =   600
            Width           =   2775
         End
         Begin VB.Frame Frame1 
            Caption         =   "本院设置"
            Height          =   855
            Left            =   240
            TabIndex        =   188
            Top             =   2280
            Width           =   9135
            Begin VB.TextBox txtMainHosp 
               Height          =   375
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   189
               ToolTipText     =   "医院代码，最大20个字符"
               Top             =   300
               Width           =   2415
            End
            Begin VB.Label Label10 
               Caption         =   "医院代码"
               Height          =   255
               Left            =   360
               TabIndex        =   190
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "分院设置"
            Height          =   3015
            Left            =   240
            TabIndex        =   186
            Top             =   3240
            Width           =   9135
            Begin VSFlex8Ctl.VSFlexGrid vsfBranchHosp 
               Height          =   2640
               Left            =   120
               TabIndex        =   187
               Top             =   240
               Width           =   8895
               _cx             =   15690
               _cy             =   4657
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
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
               Rows            =   2
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
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
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfRisDepts 
            Height          =   5400
            Left            =   -68400
            TabIndex        =   192
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
            _cx             =   4895
            _cy             =   9525
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VSFlex8Ctl.VSFlexGrid vsfRISEnables 
            Height          =   5400
            Left            =   -74880
            TabIndex        =   193
            Top             =   960
            Width           =   6375
            _cx             =   11245
            _cy             =   9525
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label Label9 
            Caption         =   $"frmParPublic.frx":6947
            ForeColor       =   &H000000C0&
            Height          =   1695
            Left            =   240
            TabIndex        =   194
            Top             =   480
            Width           =   9135
         End
      End
      Begin VB.Label Label8 
         Caption         =   "如果不选择任何RIS和预约设置，表示全院启动“医学影像信息系统专业版”，不按照“检查类型+场合”控制。"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   174
         Top             =   600
         Width           =   9495
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   6
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CheckBox chk 
         Caption         =   "转病区时将未执行或部分执行的费用转到新病区"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   176
         Top             =   1560
         Width           =   4200
      End
      Begin VB.Frame fraBabyWristlet 
         Caption         =   "婴儿腕带"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   103
         Top             =   4005
         Width           =   4815
         Begin VB.OptionButton optBabyWristletPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2805
            TabIndex        =   106
            Top             =   285
            Width           =   1500
         End
         Begin VB.OptionButton optBabyWristletPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1425
            TabIndex        =   105
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optBabyWristletPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   104
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame fraPatiWristlet 
         Caption         =   "病人腕带"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   99
         Top             =   3165
         Width           =   4815
         Begin VB.OptionButton optPatiWristletPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   100
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optPatiWristletPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1425
            TabIndex        =   101
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPatiWristletPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2805
            TabIndex        =   102
            Top             =   285
            Width           =   1500
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "入院入住，允许调整入院科室"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   92
         Top             =   390
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "转科入住时护理等级默认为空"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   93
         Top             =   675
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "出院时，下达了死亡医嘱才允许死亡出院"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   95
         Top             =   1230
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "入住时必须指定医疗小组"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   91
         Top             =   120
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "出院时，提取入院诊断为默认的出院诊断"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   94
         Top             =   960
         Width           =   4200
      End
      Begin VB.Frame fraInDeptTime 
         Caption         =   "缺省入住时间"
         Height          =   615
         Left            =   240
         TabIndex        =   96
         Top             =   2310
         Width           =   4815
         Begin VB.OptionButton OptInDeptTime 
            Caption         =   "入院时间"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   97
            Top             =   285
            Width           =   1215
         End
         Begin VB.OptionButton OptInDeptTime 
            Caption         =   "系统时间"
            Height          =   180
            Index           =   1
            Left            =   1470
            TabIndex        =   98
            Top             =   285
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7650
      Left            =   0
      ScaleHeight     =   7650
      ScaleWidth      =   2415
      TabIndex        =   150
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   154
         Top             =   120
         Width           =   45
      End
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   151
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   152
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1920
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParPublic.frx":6AA2
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   153
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
         Height          =   6765
         Left            =   0
         TabIndex        =   155
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParPublic.frx":B758
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12690
      TabIndex        =   141
      Top             =   7650
      Width           =   12690
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   158
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   146
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   144
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11040
         TabIndex        =   143
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9885
         TabIndex        =   142
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   159
         Top             =   165
         Width           =   3855
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室查找(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   157
         Top             =   165
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
         TabIndex        =   145
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   2
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   156
      Top             =   0
      Width           =   9735
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   840
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   4
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   120
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   480
         Width           =   2475
      End
      Begin MSComctlLib.ListView lvwNo 
         Height          =   6120
         Left            =   240
         TabIndex        =   114
         Top             =   1500
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   10795
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "iltC32"
         SmallIcons      =   "imgC16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "单据类型"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "编码规则"
            Object.Width           =   2646
         EndProperty
      End
      Begin ZL9BillEdit.BillEdit Bill药品科室编号 
         Height          =   3960
         Left            =   4200
         TabIndex        =   116
         Top             =   360
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   6985
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
      Begin ZL9BillEdit.BillEdit Bill卫材科室编号 
         Height          =   2400
         Left            =   4200
         TabIndex        =   118
         Top             =   4740
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   4233
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "住院号规则"
         Height          =   180
         Index           =   10
         Left            =   240
         TabIndex        =   109
         Top             =   555
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "留观号规则"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   111
         Top             =   915
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "卫材科室对应的单据编号"
         Height          =   285
         Left            =   4200
         TabIndex        =   117
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "注意：科室编号可选范围A-Z、1-9，同组中科室编号不能重复。"
         Height          =   285
         Left            =   4200
         TabIndex        =   119
         Top             =   7215
         Width           =   5040
      End
      Begin VB.Label Label4 
         Caption         =   "药品科室对应的单据编号"
         Height          =   285
         Left            =   4200
         TabIndex        =   115
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "门诊号规则"
         Height          =   180
         Index           =   22
         Left            =   240
         TabIndex        =   107
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "单据号的编码规则（鼠标双击可改变设置）"
         Height          =   285
         Left            =   240
         TabIndex        =   113
         Top             =   1260
         Width           =   3675
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   1
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   147
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Frame fra门诊结帐 
         Caption         =   "门诊结帐"
         Height          =   1065
         Left            =   5160
         TabIndex        =   177
         Top             =   6180
         Width           =   4305
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   5
            ItemData        =   "frmParPublic.frx":14DCC
            Left            =   1980
            List            =   "frmParPublic.frx":14DCE
            Style           =   2  'Dropdown List
            TabIndex        =   181
            Top             =   660
            Width           =   2100
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            ItemData        =   "frmParPublic.frx":14DD0
            Left            =   1980
            List            =   "frmParPublic.frx":14DD2
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   300
            Width           =   2100
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "未执行诊疗项目检查"
            Height          =   180
            Index           =   53
            Left            =   270
            TabIndex        =   179
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "未发药品检查"
            Height          =   180
            Index           =   18
            Left            =   810
            TabIndex        =   178
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "病人每次住院使用新的住院号"
         Height          =   195
         Index           =   1
         Left            =   5160
         TabIndex        =   46
         Top             =   2760
         Width           =   2640
      End
      Begin VB.CommandButton cmd社区参数 
         Caption         =   "设置(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8400
         TabIndex        =   34
         Top             =   480
         Width           =   1100
      End
      Begin VB.Frame fra出院检查 
         Caption         =   "病人转科或出院(未发药品)"
         Height          =   1185
         Left            =   5160
         TabIndex        =   53
         Top             =   4815
         Width           =   4305
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   22
            ItemData        =   "frmParPublic.frx":14DD4
            Left            =   1080
            List            =   "frmParPublic.frx":14DD6
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   660
            Width           =   3015
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   23
            ItemData        =   "frmParPublic.frx":14DD8
            Left            =   1080
            List            =   "frmParPublic.frx":14DDA
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   300
            Width           =   3015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "出院时"
            Height          =   180
            Index           =   46
            Left            =   375
            TabIndex        =   56
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "转科时"
            Height          =   180
            Index           =   48
            Left            =   390
            TabIndex        =   54
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "入科时必须确定护理等级"
         Height          =   180
         Index           =   2
         Left            =   5160
         TabIndex        =   47
         Top             =   3120
         Width           =   2280
      End
      Begin VB.Frame FraChangeDept 
         Caption         =   "病人转科或出院"
         Height          =   1080
         Left            =   5160
         TabIndex        =   48
         Top             =   3480
         Width           =   4305
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   28
            ItemData        =   "frmParPublic.frx":14DDC
            Left            =   1905
            List            =   "frmParPublic.frx":14DDE
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   255
            Width           =   2205
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   29
            ItemData        =   "frmParPublic.frx":14DE0
            Left            =   1905
            List            =   "frmParPublic.frx":14DE2
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   630
            Width           =   2205
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(转科)未审销帐单据"
            Height          =   180
            Index           =   57
            Left            =   210
            TabIndex        =   49
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(出院)超期护理数据"
            Height          =   180
            Index           =   8
            Left            =   210
            TabIndex        =   51
            Top             =   690
            Width           =   1620
         End
      End
      Begin VB.Frame fra出院检查副 
         Caption         =   "病人转科或出院(未执行诊疗项目)"
         Height          =   4200
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   4425
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   6
            ItemData        =   "frmParPublic.frx":14DE4
            Left            =   1080
            List            =   "frmParPublic.frx":14DE6
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   675
            Width           =   3015
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   19
            ItemData        =   "frmParPublic.frx":14DE8
            Left            =   1080
            List            =   "frmParPublic.frx":14DEA
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   315
            Width           =   3015
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUnCheckItem 
            Height          =   2565
            Left            =   240
            TabIndex        =   41
            Top             =   1440
            Width           =   3900
            _cx             =   6879
            _cy             =   4524
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":14DEC
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
         Begin VB.Label Label17 
            Caption         =   "不检查以下未执行诊疗项目："
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "出院时"
            Height          =   180
            Index           =   50
            Left            =   255
            TabIndex        =   38
            Top             =   705
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "转科时"
            Height          =   180
            Index           =   17
            Left            =   255
            TabIndex        =   36
            Top             =   375
            Width           =   540
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 入院时允许 "
         Height          =   765
         Index           =   5
         Left            =   5160
         TabIndex        =   42
         Top             =   1800
         Width           =   4335
         Begin VB.CheckBox chk 
            Caption         =   "办理就诊卡"
            Height          =   195
            Index           =   5
            Left            =   2880
            TabIndex        =   45
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "收取预交款"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   43
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "分配床位号"
            Height          =   195
            Index           =   6
            Left            =   1560
            TabIndex        =   44
            Top             =   285
            Width           =   1200
         End
      End
      Begin MSComctlLib.ListView lvw社区 
         Height          =   1065
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1879
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "序号"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "说明"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "部件"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "启用"
            Object.Width           =   952
         EndProperty
      End
      Begin VB.Label lbl社区档案接口 
         AutoSize        =   -1  'True
         Caption         =   "社区档案接口(在挂号和医生站使用)"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2880
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   3
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   10245
      TabIndex        =   149
      Top             =   0
      Visible         =   0   'False
      Width           =   10275
      Begin VB.CheckBox chk 
         Caption         =   "血库"
         Enabled         =   0   'False
         Height          =   195
         Index           =   26
         Left            =   4320
         TabIndex        =   183
         Top             =   1080
         Width           =   660
      End
      Begin VB.CommandButton cmd 
         Caption         =   "设置"
         Height          =   350
         Index           =   0
         Left            =   7080
         TabIndex        =   171
         Top             =   102
         Width           =   1100
      End
      Begin VB.CheckBox chk 
         Caption         =   "新开医嘱签名时一组医嘱签名一次"
         Height          =   195
         Index           =   49
         Left            =   960
         TabIndex        =   122
         Top             =   540
         Width           =   3540
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   11
         ItemData        =   "frmParPublic.frx":14E2A
         Left            =   960
         List            =   "frmParPublic.frx":14E2C
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   127
         Width           =   5940
      End
      Begin VB.CheckBox chk 
         Caption         =   "护理记录,护理病历"
         Enabled         =   0   'False
         Height          =   195
         Index           =   47
         Left            =   6000
         TabIndex        =   127
         Top             =   825
         Width           =   1860
      End
      Begin VB.CheckBox chk 
         Caption         =   "医技医嘱,报告"
         Enabled         =   0   'False
         Height          =   195
         Index           =   46
         Left            =   4320
         TabIndex        =   126
         Top             =   833
         Width           =   1500
      End
      Begin VB.CheckBox chk 
         Caption         =   "住院医嘱,病历"
         Enabled         =   0   'False
         Height          =   195
         Index           =   45
         Left            =   2640
         TabIndex        =   125
         Top             =   833
         Width           =   1500
      End
      Begin VB.CheckBox chk 
         Caption         =   "门诊医嘱,病历"
         Enabled         =   0   'False
         Height          =   195
         Index           =   44
         Left            =   960
         TabIndex        =   124
         Top             =   833
         Width           =   1620
      End
      Begin VB.CheckBox chk 
         Caption         =   "药品发药"
         Enabled         =   0   'False
         Height          =   195
         Index           =   48
         Left            =   8040
         TabIndex        =   128
         Top             =   833
         Width           =   1020
      End
      Begin VB.CheckBox chk 
         Caption         =   "LIS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   43
         Left            =   960
         TabIndex        =   129
         Top             =   1080
         Width           =   660
      End
      Begin VB.CheckBox chk 
         Caption         =   "PACS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   42
         Left            =   2640
         TabIndex        =   130
         Top             =   1080
         Width           =   660
      End
      Begin TabDlg.SSTab sstSign 
         Height          =   5895
         Left            =   120
         TabIndex        =   132
         Top             =   1560
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   10398
         _Version        =   393216
         Style           =   1
         Tabs            =   9
         TabsPerRow      =   9
         TabHeight       =   520
         TabCaption(0)   =   "门诊医嘱,病历"
         TabPicture(0)   =   "frmParPublic.frx":14E2E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "vsDept(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "住院医生医嘱,病历"
         TabPicture(1)   =   "frmParPublic.frx":14E4A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vsDept(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "住院护士医嘱"
         TabPicture(2)   =   "frmParPublic.frx":14E66
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "vsDept(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "医技医嘱,报告"
         TabPicture(3)   =   "frmParPublic.frx":14E82
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "vsDept(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "护理记录,护理病历"
         TabPicture(4)   =   "frmParPublic.frx":14E9E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "vsDept(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "药品发药"
         TabPicture(5)   =   "frmParPublic.frx":14EBA
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "vsDept(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "LIS"
         TabPicture(6)   =   "frmParPublic.frx":14ED6
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "vsDept(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "PACS"
         TabPicture(7)   =   "frmParPublic.frx":14EF2
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "vsDept(7)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "血库"
         TabPicture(8)   =   "frmParPublic.frx":14F0E
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "vsDept(8)"
         Tab(8).ControlCount=   1
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   5
            Left            =   -74880
            TabIndex        =   138
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":14F2A
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   4
            Left            =   -74880
            TabIndex        =   137
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":14FBD
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   2
            Left            =   -74880
            TabIndex        =   135
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":15050
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   7
            Left            =   -74880
            TabIndex        =   140
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":150E3
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
            Begin ComctlLib.ImageList imgCheck 
               Left            =   0
               Top             =   720
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   327682
               BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
                  NumListImages   =   2
                  BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "frmParPublic.frx":15176
                     Key             =   "Checked"
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "frmParPublic.frx":15350
                     Key             =   "UnChecked"
                  EndProperty
               EndProperty
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   3
            Left            =   -74880
            TabIndex        =   136
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":1552A
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   1
            Left            =   -74880
            TabIndex        =   134
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":155BD
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5145
            Index           =   0
            Left            =   120
            TabIndex        =   133
            Top             =   495
            Width           =   9870
            _cx             =   17410
            _cy             =   9075
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":15650
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   6
            Left            =   -74880
            TabIndex        =   139
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":156E3
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   8
            Left            =   -74880
            TabIndex        =   184
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
            BorderStyle     =   0
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
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":15776
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
      Begin VB.Label Label6 
         Caption         =   "认证中心"
         Height          =   255
         Left            =   200
         TabIndex        =   120
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "场合"
         Height          =   180
         Left            =   575
         TabIndex        =   123
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label15 
         Caption         =   "启用某个场合后，未勾选任何科室，表示不按科室控制。"
         Height          =   255
         Left            =   200
         TabIndex        =   131
         Top             =   1320
         Width           =   4815
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   4
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VSFlex8Ctl.VSFlexGrid vsgInput 
         Height          =   4560
         Index           =   0
         Left            =   240
         TabIndex        =   163
         Top             =   2775
         Width           =   5940
         _cx             =   10477
         _cy             =   8043
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParPublic.frx":15809
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
      Begin VB.Frame fra票据格式 
         Caption         =   "预交票据格式"
         Height          =   1395
         Left            =   240
         TabIndex        =   61
         Top             =   975
         Width           =   5955
         Begin VSFlex8Ctl.VSFlexGrid vfgBillFormat 
            Height          =   1095
            Left            =   120
            TabIndex        =   62
            Top             =   225
            Width           =   5775
            _cx             =   10186
            _cy             =   1931
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
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
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
            FormatString    =   $"frmParPublic.frx":158AD
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
         Caption         =   "扫描身份证签约"
         Height          =   180
         Index           =   19
         Left            =   240
         TabIndex        =   58
         Top             =   120
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "建档同时必须发卡"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chk 
         Caption         =   "就诊卡费用以记账方式收取"
         Height          =   180
         Index           =   21
         Left            =   240
         TabIndex        =   60
         Top             =   660
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Label lblInput 
         AutoSize        =   -1  'True
         Caption         =   "输入项控制"
         Height          =   180
         Left            =   240
         TabIndex        =   164
         Top             =   2520
         Width           =   900
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   5
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   161
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CheckBox chk 
         Caption         =   "科室下无空床不能登记"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   175
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Frame fraDeptFirst 
         Caption         =   "科室、病区优先级"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   71
         Top             =   2400
         Width           =   4095
         Begin VB.OptionButton optDeptFirst 
            Caption         =   "先选病区"
            Height          =   255
            Index           =   1
            Left            =   1485
            MaskColor       =   &H00000000&
            TabIndex        =   73
            Top             =   285
            Width           =   1215
         End
         Begin VB.OptionButton optDeptFirst 
            Caption         =   "先选科室"
            Height          =   255
            Index           =   0
            Left            =   135
            MaskColor       =   &H00000000&
            TabIndex        =   72
            Top             =   285
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "允许通过输入姓名来模糊查找病人信息"
         Height          =   195
         Index           =   12
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   67
         Top             =   1185
         Width           =   3540
      End
      Begin VB.CheckBox chk 
         Caption         =   "扫描身份证签约"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   64
         Top             =   390
         Value           =   1  'Checked
         Width           =   2520
      End
      Begin VB.CheckBox chk 
         Caption         =   "入院时自动计算一次费用"
         Height          =   180
         Index           =   8
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   65
         Top             =   660
         Value           =   1  'Checked
         Width           =   2295
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
         Index           =   0
         Left            =   1875
         MaxLength       =   3
         TabIndex        =   69
         Text            =   "3"
         Top             =   1455
         Width           =   285
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1875
         TabIndex        =   70
         Top             =   1650
         Width           =   285
      End
      Begin VB.CheckBox chk 
         Caption         =   "门诊费用转住院费用后立即退费或销帐"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   86
         Top             =   5640
         Width           =   3615
      End
      Begin VB.Frame FraDepositMtoZ 
         Caption         =   "门诊转住院预交款票据"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   87
         Top             =   6000
         Width           =   4095
         Begin VB.OptionButton optDepositMtoZ 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   90
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optDepositMtoZ 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   89
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optDepositMtoZ 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   88
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame fraWristlet 
         Caption         =   "病人腕带"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   82
         Top             =   4560
         Width           =   4095
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   83
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   84
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   85
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.Frame fraPatientPage 
         Caption         =   "病案首页"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   78
         Top             =   3840
         Width           =   4095
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2655
            TabIndex        =   81
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   80
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   79
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "预交款票据"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   74
         Top             =   3120
         Width           =   4095
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   75
            Top             =   300
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   76
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   77
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "输入病人担保信息"
         Height          =   210
         Index           =   3
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   63
         Top             =   120
         Width           =   1740
      End
      Begin VB.CheckBox chk 
         Caption         =   "医疗卡费用以记账方式收取"
         Height          =   180
         Index           =   9
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   66
         Top             =   933
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsgInput 
         Height          =   6915
         Index           =   1
         Left            =   4560
         TabIndex        =   165
         Top             =   360
         Width           =   4860
         _cx             =   8572
         _cy             =   12197
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParPublic.frx":1593B
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "输入项控制"
         Height          =   180
         Index           =   1
         Left            =   4560
         TabIndex        =   166
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "预约接收时提取病人    天内的诊断信息"
         Height          =   180
         Left            =   240
         TabIndex        =   68
         Top             =   1455
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmParPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset   '参数与控件对应记录集（同一个参数可能对应一组多个控件）
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnOk As Boolean
Private mobjESign As Object                 '电子签名接口

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk '最大序号55
    chk只允许补录临嘱 = 0
    chk_每次住院使用新住院号 = 1
    
    chk_入科确定护理等级 = 2
    chk_收取预交款 = 4
    chk_时办理就诊卡 = 5
    chk_分配床位号 = 6
    
    chk_全数字只查编码 = 10
    chk_全字母只查简码 = 11
    
    chk_病人地址结构化录入 = 50
    chk_乡镇地址结构化录入 = 51
    chk_外院医生必须先建档 = 22
    chk_启用医学影像信息系统专业版接口 = 52
    chk_医疗机构不允许自由录入 = 25
    
    '病人入院管理
    chk_输入病人担保信息 = 3
    chk_入院时扫描身份证签约 = 7
    chk_入院时自动计算一次费用 = 8
    chk_入院时卡费记帐 = 9
    chk_入院时姓名模糊查找 = 12
    chk_费用转出立即退费 = 13
    chk_科室下无空床不能登记 = 23
    
    '病人入出管理
    chk_入住指定医疗小组 = 14
    chk_入住允许调整入院科室 = 15
    chk_转科入住时护理等级默认为空 = 16
    chk_出院默认诊断 = 17
    chk_下达死亡医嘱才允许死亡出院 = 18
    chk_转病区转费用 = 24
    
    '病人信息管理
    chk_扫描身份证签约 = 19
    chk_建档同时必须发卡 = 20
    chk_卡费记帐 = 21
    
    chk_Sign_pacs = 42
    chk_Sign_lis = 43
    chk_Sign_门诊 = 44
    chk_Sign_住院 = 45
    chk_Sign_医技 = 46
    chk_Sign_护理 = 47
    chk_Sign_药品 = 48
    chk_新开一组医嘱签名一次 = 49
    chk_sign_血库 = 26
End Enum

Private Enum constCbo
    cbo_门诊号规则 = 4
    cbo_留观号规则 = 3 '住院留观
    cbo_住院号规则 = 2
    
    cbo_诊断输入来源 = 1
    cbo_门诊诊断输入 = 8
    cbo_住院诊断输入 = 16
    
    cbo_诊疗编码模式 = 10
    cbo_医保对码检查 = 15
    cbo_电子签名认证中心 = 11
        
    cbo_转科时未执行项目检查 = 19
    cbo_出院时未执行项目检查 = 6
    cbo_出院时未发药项目检查 = 22
    cbo_转科时未发药项目检查 = 23
    cmd_转科时未审核销帐单据 = 28
    cmd_出院时超期护理数据 = 29
    
    cbo_结帐_门诊未发药品检查 = 0
    cbo_结帐_门诊未执行项目检查 = 5
    
End Enum

Private Enum constUpDown
    ud_补录时限 = 0
    ud_儿童年龄界定上限 = 9
End Enum

Private Enum const日期
    dtp_上午上班 = 0
    dtp_上午下班 = 1
    dtp_下午上班 = 2
    dtp_下午下班 = 3
End Enum

'不启用电子签名的部门
Private Enum constDeptCol
    col_选择 = 0
    col_站点 = 1
    col_编码 = 2
    col_名称 = 3
    col_简码 = 4
End Enum
'电子签名场合
Private Enum constSign
    sst_门诊 = 0
    sst_住院医生 = 1
    sst_住院护士 = 2
    sst_医技 = 3
    sst_护理 = 4
    sst_药品 = 5
    sst_lis = 6
    sst_Pacs = 7
    sst_血库 = 8
End Enum

'启用RIS设置
Private Enum constRisEnables
    col_RIS启用检查类型 = 0
    col_RIS启用场合 = 1
    col_RIS启用科室 = 2
    col_RIS启用科室ID = 3
    col_RIS启用预约科室全 = 4
    col_RIS启用预约科室 = 5
    col_RIS启用预约科室ID = 6
End Enum
'启用RIS的科室
Private Enum constRisDepts
    col_Ris科室选择 = 0
    col_Ris科室名称 = 1
    col_Ris科室编码 = 2
    col_Ris科室ID = 3
End Enum
'RIS选择
Private Const RIS_Checked = "Checked"

'RIS分院设置
Private Enum constRisBranchHosp
    col_RIS分院序号 = 0
    col_RIS分院名称 = 1
    col_ris分院代码 = 2
    col_ris分院用户名 = 3
    col_ris分院密码 = 4
    col_ris分院数据库服务名 = 5
End Enum

Private Enum constTxt
    txt_诊断查找天数 = 0
End Enum
'药房或发料部门的科室编号
Private Enum mGrdCol
    选择 = 0
    科室
    号码
End Enum

Private Enum constVSGInput
    VSGInput_病人信息输入项设置 = 0
    VSGInput_病人入院输入项设置 = 1
    
    COL_系统标识 = 1
    COL_服务名称 = 2
    COL_服务地址 = 3
End Enum

Private Enum constCmd
    cmd_电子签名设置 = 0
End Enum
'记录最后编辑的科室编号所在行、列和编号值
Private mintLastRow_Drug As Integer          '行
Private mintLastCol_Drug As Integer          '列
Private mstrLastCode_Drug As String          '编号

Private mintLastRow_Stuff As Integer          '行
Private mintLastCol_Stuff As Integer          '列
Private mstrLastCode_Stuff As String          '编号
Private mrsSvr As ADODB.Recordset   '三方服务配置目录


Private Sub chkShowSel_Click()
    Dim i As Integer
    
    With vsfRISEnables
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col_RIS启用检查类型) = 2 Then .RowHidden(i) = IIF(chkShowSel.value = 1, True, False)
        Next i
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = cmd_电子签名设置 Then
       Call mobjESign.Setup(Me, gcnOracle, glngSys)
    End If
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "初始成功" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    mblnOk = False
    strCategory = "参数设置,基础项目"
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "100,0,系统公共;102,1,病人管理公共;104,4,病人信息管理;105,5,病人入院管理;106,6,病人入出管理"
    marrFunc(1) = "103,2,单据编号设置;101,3,电子签名控制;107,7,影像信息系统"
    
    
    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    
    Me.Tag = "初始成功"
End Sub

Private Sub optBabyWristletPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optBabyWristletPrint, Index, mrsPar)
End Sub

Private Sub optBabyWristletPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBabyWristletPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBabyWristletPrint, Index, mrsPar)
End Sub

Private Sub optDepositMtoZ_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optDepositMtoZ, Index, mrsPar)
End Sub

Private Sub optDepositMtoZ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDepositMtoZ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDepositMtoZ, Index, mrsPar)
End Sub

Private Sub optDeptFirst_Click(Index As Integer)
     If Me.Visible Then Call SetParChange(optDeptFirst, Index, mrsPar)
End Sub

Private Sub optDeptFirst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDeptFirst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDeptFirst, Index, mrsPar)
End Sub

Private Sub optFpagePrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optFpagePrint, Index, mrsPar)
End Sub

Private Sub optFpagePrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optFpagePrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optFpagePrint, Index, mrsPar)
End Sub

Private Sub OptInDeptTime_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptInDeptTime, Index, mrsPar)
End Sub

Private Sub OptInDeptTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptInDeptTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptInDeptTime, Index, mrsPar)
End Sub

Private Sub optPatiWristletPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPatiWristletPrint, Index, mrsPar)
End Sub

Private Sub optPatiWristletPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPatiWristletPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPatiWristletPrint, Index, mrsPar)
End Sub

Private Sub optPrepayPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPrepayPrint, Index, mrsPar)
End Sub

Private Sub optPrepayPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrepayPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrepayPrint, Index, mrsPar)
End Sub

Private Sub optWristletPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optWristletPrint, Index, mrsPar)
End Sub

Private Sub optWristletPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optWristletPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optWristletPrint, Index, mrsPar)
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("电子签名控制", marrFunc) Or _
                                   Item.ID = GetFuncID("单据编号设置", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOk.Left - lblPrompt.Left - 120
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


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOk.Left = cmdCancel.Left - cmdOk.Width - 120
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
    Call InitSystemPara
    
    
    
    '2.初始化界面控件
    Call InitEnv
    Call Load社区接口
    
    Call Load单据编码规则
    Call Load药品卫材科室编号
    
    Call LoadThirdSvr
        
    '3.加载系统参数
    Call LoadPar
    
    
End Sub


Private Sub LoadPar()
'功能：读取并加载参数到界面控件
    Dim strValue As String
    Dim i As Long, arrTmp As Variant
    Dim blnFind As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String      '模块号1:参数号1:控件数组序号1,参数号2:控件数组序号2,......
    Dim arrObj As Variant  '数组对象：模块1,参数号1,控件对象1,模块2,参数号2,控件对象2,......
    Dim strBillFormat As String, strPrintMode As String '预交票据格式和打印方式
    Set rsTmp = GetPar(mrsPar, p病人信息管理 & "," & P病人入院管理 & "," & p病人入出管理)
        
     '1.设置CheckBox类参数
    strTmp = "0:10:" & chk_收取预交款 & _
            ",0:11:" & chk_时办理就诊卡 & _
            ",0:13:" & chk_分配床位号 & _
            ",0:99:" & chk_入科确定护理等级 & _
            ",0:191:" & chk只允许补录临嘱 & _
            ",0:145:" & chk_每次住院使用新住院号 & _
            ",0:239:" & chk_新开一组医嘱签名一次 & _
            ",0:251:" & chk_病人地址结构化录入 & _
            ",0:252:" & chk_乡镇地址结构化录入 & _
            ",0:253:" & chk_外院医生必须先建档 & _
            ",0:255:" & chk_启用医学影像信息系统专业版接口 & _
            ",0:287:" & chk_医疗机构不允许自由录入

    Call SetParToControl(strTmp, mrsPar, chk)
    
    '病人信息相关
    strTmp = "1101:扫描身份证签约:" & chk_扫描身份证签约 & ",1101:建档同时必须发卡:" & chk_建档同时必须发卡 & ",1101:卡费记帐:" & chk_卡费记帐
    Call SetParToControl(strTmp, mrsPar, chk)

    '入院管理相关
     strTmp = "1131:卡费记帐:" & chk_入院时卡费记帐 & ",1131:担保信息:" & chk_输入病人担保信息 & ",1131:姓名模糊查找:" & chk_入院时姓名模糊查找 & _
            ",1131:扫描身份证签约:" & chk_入院时扫描身份证签约 & ",1131:费用计算时机:" & chk_入院时自动计算一次费用 & ",1131:费用转出立即退费:" & chk_费用转出立即退费 & _
        ",1131:科室下无空床不能登记:" & chk_科室下无空床不能登记
     Call SetParToControl(strTmp, mrsPar, chk)
     '入出管理相关
     strTmp = "1132:默认诊断:" & chk_出院默认诊断 & ",1132:允许调整科室:" & chk_入住允许调整入院科室 & ",1132:护理等级默认为空:" & chk_转科入住时护理等级默认为空 & _
        ",1132:出院死亡:" & chk_下达死亡医嘱才允许死亡出院 & ",1132:入住指定医疗小组:" & chk_入住指定医疗小组 & ",1132:转病区转费用:" & chk_转病区转费用
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '2.设置ComboBox类参数
    strTmp = "0:22:" & cbo_出院时未执行项目检查 & _
            ",0:32:" & cbo_转科时未执行项目检查 & _
            ",0:59:" & cbo_医保对码检查 & _
            ",0:61:" & cbo_诊疗编码模式 & _
            ",0:154:" & cbo_出院时未发药项目检查 & _
            ",0:155:" & cbo_转科时未发药项目检查 & _
            ",0:227:" & cmd_转科时未审核销帐单据 & _
            ",0:265:" & cbo_结帐_门诊未发药品检查 & _
            ",0:266:" & cbo_结帐_门诊未执行项目检查 & _
            ",0:235:" & cmd_出院时超期护理数据

    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = "0:25:" & cbo_电子签名认证中心
    Call SetParToControl(strTmp, mrsPar, cbo, 2)
    
    '3.设置UpDown类参数
    strTmp = "0:147:" & ud_儿童年龄界定上限 & _
            ",0:158:" & ud_补录时限
                
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar存储的控件名是txtUD
            
    '4.设置TextBox类参数
    strTmp = "1131:诊断查找天数:" & txt_诊断查找天数
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '5.设置ListBox类参数
    strTmp = ""
    'Call SetParToControl(strTmp, mrsPar, lst)
    
    '6.设置OptionButton类参数
    '病人入院管理
    arrObj = Array(P病人入院管理, "先选病区", optDeptFirst, P病人入院管理, "预交款票据打印", optPrepayPrint, P病人入院管理, "病案首页打印", optFpagePrint, P病人入院管理, "病人腕带打印", optWristletPrint, P病人入院管理, "门诊转住院预交打印", optDepositMtoZ)
    Call SetParToControl("", mrsPar, arrObj)
    '病人入出管理
    arrObj = Array(p病人入出管理, "缺省入科时间", OptInDeptTime, p病人入出管理, "病人腕带打印", optPatiWristletPrint, p病人入出管理, "婴儿腕带打印", optBabyWristletPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    '7.其他系统参数
    rsTmp.Filter = "模块=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数号
        
        Case 1    '上午上下班时间
            i = InStr(UCase(strValue), "AND")
            strTmp = Mid(strValue, 1, i - 2)
            dtp(dtp_上午上班).value = CDate(strTmp)
            strTmp = Mid(strValue, i + 4)
            dtp(dtp_上午下班).value = CDate(strTmp)
            
            Call SetParRelation(dtp, dtp_上午上班, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(dtp, dtp_上午下班, mrsPar)
            
        Case 2    '下午上下班时间
            i = InStr(UCase(strValue), "AND")
            strTmp = Mid(strValue, 1, i - 2)
            dtp(dtp_下午上班).value = CDate(strTmp)
            strTmp = Mid(strValue, i + 4)
            dtp(dtp_下午下班).value = CDate(strTmp)
        
            Call SetParRelation(dtp, dtp_下午上班, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(dtp, dtp_下午下班, mrsPar)
    
        Case 26    '电子签名使用场合
            strTmp = chk_Sign_门诊 & "," & chk_Sign_住院 & "," & chk_Sign_医技 & "," & chk_Sign_护理 & "," & _
                    chk_Sign_药品 & "," & chk_Sign_lis & "," & chk_Sign_pacs & "," & chk_sign_血库
            arrTmp = Split(strTmp, ",")
            For i = 1 To 8
                chk(arrTmp(i - 1)).value = Val(Mid(strValue, i, 1))
                If i = 1 Then
                    Call SetParRelation(chk, arrTmp(i - 1), mrsPar, rsTmp!参数号)
                Else
                    Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                    Call SetParRelation(chk, arrTmp(i - 1), mrsPar)
                End If
            Next
            
        Case 44    '收费项目和诊疗项目的输入匹配方式
            chk(chk_全数字只查编码).value = IIF(Mid(NVL(strValue, "00"), 1, 1) = "1", 1, 0)
            chk(chk_全字母只查简码).value = IIF(Mid(NVL(strValue, "00"), 2, 1) = "1", 1, 0)
            
            Call SetParRelation(chk, chk_全数字只查编码, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_全字母只查简码, mrsPar)
        Case 55
            cbo(cbo_诊断输入来源).ListIndex = IIF(Val(strValue) > cbo(cbo_诊断输入来源).ListCount, 0, Val(strValue) - 1)
            
            Call SetParRelation(cbo, cbo_诊断输入来源, mrsPar, rsTmp!参数号)
        Case 65
            cbo(cbo_门诊诊断输入).ListIndex = Val(Mid(NVL(strValue, "11"), 1, 1)) - 1
            cbo(cbo_住院诊断输入).ListIndex = Val(Mid(NVL(strValue, "11"), 2, 1)) - 1
            
            Call SetParRelation(cbo, cbo_门诊诊断输入, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_住院诊断输入, mrsPar)
        Case 234
            Call Init转科出院不检查项目(strValue)
            
            Call SetParRelation(vsUnCheckItem, 0, mrsPar, rsTmp!参数号)
        End Select
        rsTmp.MoveNext
    Loop
    
    
    '8.其他模块参数
    strBillFormat = "": strPrintMode = ""
    rsTmp.Filter = "模块=" & p病人信息管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "预交发票格式" '预交发票格式
                strBillFormat = strValue
                Call SetParRelation(vfgBillFormat, 0, mrsPar, rsTmp!参数名, 1101, , vfgBillFormat.ColIndex("票据格式"))
            Case "预交发票打印方式" '预交发票打印方式
                strPrintMode = strValue
                Call SetParRelation(vfgBillFormat, 0, mrsPar, rsTmp!参数名, 1101, "", vfgBillFormat.ColIndex("预交打印方式"))
            Case "输入项控制" '病人信息管理 输入项控制
                If strValue = "" Then strValue = "国籍|民族|学历|婚姻状况|职业|身份|出生日期|其他证件|身份证号|出生地点|现住址|家庭地址邮编|家庭电话|联系人姓名|联系人关系|户口地址|户口地址邮编|区域|联系人地址|联系人电话|联系人身份证号|工作单位|单位电话|单位邮编|单位开户行|单位帐号|籍贯"
                Call LoadInputItem(VSGInput_病人信息输入项设置, strValue)
                Call SetParRelation(vsgInput, VSGInput_病人信息输入项设置, mrsPar, rsTmp!参数名, p病人信息管理)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & P病人入院管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "输入项控制" '病人信息管理 输入项控制
                If strValue = "" Then strValue = "国籍|民族|学历|婚姻状况|职业|身份|出生日期|其他证件|身份证号|出生地点|现住址|家庭地址邮编|家庭电话|联系人姓名|联系人关系|户口地址|户口地址邮编|区域|联系人地址|联系人电话|联系人身份证号|工作单位|单位电话|单位邮编|单位开户行|单位帐号|籍贯"
                Call LoadInputItem(VSGInput_病人入院输入项设置, strValue)
                Call SetParRelation(vsgInput, VSGInput_病人入院输入项设置, mrsPar, rsTmp!参数名, P病人入院管理)
        End Select
        rsTmp.MoveNext
    Loop
    
    
    '加载预交票据格式和打印方式信息
    Call LaodBillForamt(vfgBillFormat, strBillFormat, strPrintMode)
    
End Sub

Private Sub cmdOK_Click()
    If ValidateData() = False Then Exit Sub
    
    If cbo(cbo_电子签名认证中心).ListIndex > 0 Then Call Save电子签名
    
    Call Save社区接口
    
    Call Save单据编码规则
    Call Save科室编号
    
    Call SaveThirdSvr
    
    '保存“影像信息系统”启用控制
    If chk(52).value = 1 Then Call SaveRisEnable
    
    Call SaveRisBranchHosp
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub


Private Function ValidateData() As Boolean
'功能：验证数据的有效性
    Dim i As Long, strTmp As String
    
    '自动对科室编号最后一个编辑操作进行校验
    If mintLastRow_Drug > 0 And Len(Trim(mstrLastCode_Drug)) > 0 Then
        With Bill药品科室编号
            If .TextMatrix(mintLastRow_Drug, mintLastCol_Drug) <> UCase(mstrLastCode_Drug) Then
                .TextMatrix(mintLastRow_Drug, mintLastCol_Drug) = UCase(mstrLastCode_Drug)
            End If
        End With
    End If
    If mintLastRow_Stuff > 0 And Len(Trim(mstrLastCode_Stuff)) > 0 Then
        With Bill卫材科室编号
            If .TextMatrix(mintLastRow_Stuff, mintLastCol_Stuff) <> UCase(mstrLastCode_Stuff) Then
                .TextMatrix(mintLastRow_Stuff, mintLastCol_Stuff) = UCase(mstrLastCode_Stuff)
            End If
        End With
    End If
    
      
    If CheckNumberRule_Drug = True Then
        '同一个GRID里的科室编号不能重复
        With Bill药品科室编号
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "药品科室第" & i & "行编号重复！", vbQuestion, gstrSysName
                  
                        .Row = i
                        .Col = 2
                        .SetFocus
                        Exit Function
                    End If
                End If
                strTmp = strTmp & "," & .TextMatrix(i, 2)
            Next
        End With
        strTmp = ""
    Else
        With Bill药品科室编号
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = ""
            Next
            .Tag = "已修改"
        End With
    End If
    
    If CheckNumberRule_Stuff = True Then
        '同一个GRID里的科室编号不能重复
        With Bill卫材科室编号
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "卫材科室第" & i & "行编号重复！", vbQuestion, gstrSysName
                    
                        .Row = i
                        .Col = 2
                        .SetFocus
                        Exit Function
                    End If
                End If
                strTmp = strTmp & "," & .TextMatrix(i, 2)
            Next
        End With
    Else
        With Bill卫材科室编号
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = ""
            Next
            .Tag = "已修改"
        End With
    End If
    
    If ValidateRisBranchHosp = False Then Exit Function
    
    ValidateData = True
End Function


Private Sub InitEnv()
'功能：初始化界面控件，加载基础数据
    Dim strTmp As String
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim blnTmp As Boolean
    Dim arrTemp As Variant
    
    cbo(cbo_住院号规则).AddItem "0-顺序编号"
    cbo(cbo_住院号规则).AddItem "1-年月(YYMM)+顺序号(0000)"
    cbo(cbo_住院号规则).AddItem "2-年(YYYY)+顺序号(00000)"
    cbo(cbo_住院号规则).ListIndex = 0
    
    cbo(cbo_留观号规则).AddItem "0-顺序编号"
    cbo(cbo_留观号规则).AddItem "1-年月(YYMM)+顺序号(0000)"
    cbo(cbo_留观号规则).AddItem "2-年(YYYY)+顺序号(00000)"
    cbo(cbo_留观号规则).ListIndex = 0
    
    cbo(cbo_门诊号规则).AddItem "0-顺序编号"
    cbo(cbo_门诊号规则).AddItem "1-年月日(YYMMDD)+顺序号(0000)"
    cbo(cbo_门诊号规则).ListIndex = 0

    cbo(cbo_诊断输入来源).AddItem "1-可选择输入来源"
    cbo(cbo_诊断输入来源).AddItem "2-按诊断标准输入"
    cbo(cbo_诊断输入来源).AddItem "3-按疾病编码输入"
    cbo(cbo_诊断输入来源).ListIndex = 0
    
    cbo(cbo_门诊诊断输入).AddItem "1-允许自由输入"
    cbo(cbo_门诊诊断输入).AddItem "2-从数据库提取输入"
    cbo(cbo_门诊诊断输入).AddItem "3-仅医保病人从数据库输入"
    cbo(cbo_门诊诊断输入).ListIndex = 0
    
    cbo(cbo_住院诊断输入).AddItem "1-允许自由输入"
    cbo(cbo_住院诊断输入).AddItem "2-从数据库提取输入"
    cbo(cbo_住院诊断输入).AddItem "3-仅医保病人从数据库输入"
    cbo(cbo_住院诊断输入).ListIndex = 0
    
    
    cbo(cbo_诊疗编码模式).AddItem "顺序编号"
    cbo(cbo_诊疗编码模式).AddItem "种类+分类号+顺序编号"
        
    cbo(cbo_医保对码检查).AddItem "0-不进行检查"
    cbo(cbo_医保对码检查).AddItem "1-检查并提醒未对码项目"
    cbo(cbo_医保对码检查).AddItem "2-检查并禁止未对码项目"
    cbo(cbo_医保对码检查).ListIndex = 1
    
    
    cbo(cbo_出院时未执行项目检查).AddItem "0-不检查"
    cbo(cbo_出院时未执行项目检查).AddItem "1-检查并提示"
    cbo(cbo_出院时未执行项目检查).AddItem "2-检查并禁止"
    cbo(cbo_出院时未执行项目检查).ListIndex = 0
    
    cbo(cbo_转科时未执行项目检查).AddItem "0-不检查"
    cbo(cbo_转科时未执行项目检查).AddItem "1-检查并提示"
    cbo(cbo_转科时未执行项目检查).AddItem "2-检查并禁止"
    cbo(cbo_转科时未执行项目检查).ListIndex = 0
    
    cbo(cbo_出院时未发药项目检查).AddItem "0-不检查"
    cbo(cbo_出院时未发药项目检查).AddItem "1-检查并提示"
    cbo(cbo_出院时未发药项目检查).AddItem "2-检查并禁止"
    cbo(cbo_出院时未发药项目检查).ListIndex = 0
    
    cbo(cbo_转科时未发药项目检查).AddItem "0-不检查"
    cbo(cbo_转科时未发药项目检查).AddItem "1-检查并提示"
    cbo(cbo_转科时未发药项目检查).AddItem "2-检查并禁止"
    cbo(cbo_转科时未发药项目检查).ListIndex = 0
    
    cbo(cmd_转科时未审核销帐单据).AddItem "0-不检查"
    cbo(cmd_转科时未审核销帐单据).AddItem "1-检查并提示"
    cbo(cmd_转科时未审核销帐单据).AddItem "2-检查并禁止"
    cbo(cmd_转科时未审核销帐单据).ListIndex = 0
    
    cbo(cmd_出院时超期护理数据).AddItem "0-不检查"
    cbo(cmd_出院时超期护理数据).AddItem "1-检查并提示"
    cbo(cmd_出院时超期护理数据).AddItem "2-检查并禁止"
    cbo(cmd_出院时超期护理数据).ListIndex = 0
    
    cbo(cbo_结帐_门诊未发药品检查).AddItem "0-不检查"
    cbo(cbo_结帐_门诊未发药品检查).AddItem "1-检查并提示"
    cbo(cbo_结帐_门诊未发药品检查).AddItem "2-检查并禁止"
    cbo(cbo_结帐_门诊未发药品检查).ListIndex = 0
    
    cbo(cbo_结帐_门诊未执行项目检查).AddItem "0-不检查"
    cbo(cbo_结帐_门诊未执行项目检查).AddItem "1-检查并提示"
    cbo(cbo_结帐_门诊未执行项目检查).AddItem "2-检查并禁止"
    cbo(cbo_结帐_门诊未执行项目检查).ListIndex = 0
    
    
    vsUnCheckItem.ComboList = "..."
    
    '电子签名认证中心
    If mobjESign Is Nothing Then
        On Error Resume Next
        Set mobjESign = CreateObject("zl9ESign.clsESign")
        Err.Clear: On Error GoTo 0
    End If
    If Not mobjESign Is Nothing Then
        strTmp = mobjESign.GetESignType()
        arrTemp = Split(strTmp, ",")
        For i = LBound(arrTemp) To UBound(arrTemp)
            cbo(cbo_电子签名认证中心).AddItem arrTemp(i)
        Next
        If cbo(cbo_电子签名认证中心).ListCount > 0 Then cbo(cbo_电子签名认证中心).ListIndex = 0
    End If
    
    For i = 0 To sstSign.Tabs - 1
        sstSign.TabVisible(i) = False
        If i = sst_门诊 Then
            strTmp = " And t.服务对象 IN (1,3)  and T.工作性质 IN ('临床','手术','治疗')"
        ElseIf i = sst_住院医生 Then
            strTmp = " And t.服务对象 IN (2,3)  and T.工作性质 IN ('临床','手术','治疗')"
        ElseIf i = sst_住院护士 Then
            strTmp = " And t.服务对象 IN (2,3)  and T.工作性质='护理'"
        ElseIf i = sst_医技 Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质 IN('检查','检验','手术','治疗','营养')"
        ElseIf i = sst_护理 Then
            strTmp = " And t.服务对象 IN (2,3)  and T.工作性质='护理'"
        ElseIf i = sst_药品 Then
            strTmp = " and T.工作性质 in('西药房','中药房','成药房')"
        ElseIf i = sst_lis Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质='检验'"
        ElseIf i = sst_Pacs Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质='检查'"
        ElseIf i = sst_血库 Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质='血库'"
        End If
         '加载默认部门选择
        gstrSQL = "Select Distinct D.ID, d.站点,D.编码, D.名称,D.简码" & vbNewLine & _
                "From 部门表 D, 部门性质说明 T" & vbNewLine & _
                "Where d.Id = t.部门id And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) " & strTmp & vbNewLine & _
                "order by 站点,名称"
                
        On Error GoTo ErrHandle
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With vsDept(i)
            .Rows = 1
            .MergeCells = flexMergeFree
            .MergeCol(col_站点) = True
            .AllowUserResizing = flexResizeBoth
            .SelectionMode = flexSelectionByRow
            .Editable = flexEDKbdMouse
            .ExplorerBar = flexExSortShowAndMove
            .ColSort(col_选择) = flexSortNone
            .Cell(flexcpPicture, 0, col_选择) = imgCheck.ListImages("UnChecked").Picture
            .Cell(flexcpPictureAlignment, 0, col_选择) = flexAlignCenterCenter
            blnTmp = False
            Do While Not rsTmp.EOF
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ID & "")
                .TextMatrix(.Rows - 1, col_站点) = rsTmp!站点 & ""
                If rsTmp!站点 & "" <> "" Then
                    blnTmp = True
                Else
                    .TextMatrix(.Rows - 1, col_站点) = " "
                End If
                .TextMatrix(.Rows - 1, col_编码) = rsTmp!编码 & ""
                .TextMatrix(.Rows - 1, col_名称) = rsTmp!名称 & ""
                .TextMatrix(.Rows - 1, col_简码) = rsTmp!简码 & ""
                
                rsTmp.MoveNext
            Loop
            .ColHidden(col_站点) = Not blnTmp
        End With
    Next
    
    '电子签名控制
    Call cbo_Click(cbo_电子签名认证中心)
    Call LoadSign
    
        
    '初始化控件
    With Bill药品科室编号
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.选择) = "选择"
        .TextMatrix(0, mGrdCol.科室) = "科室"
        .TextMatrix(0, mGrdCol.号码) = "号码"

        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mGrdCol.选择) = 5
        .ColData(mGrdCol.科室) = 5
        .ColData(mGrdCol.号码) = 4


        .ColWidth(mGrdCol.选择) = 0
        .ColWidth(mGrdCol.科室) = 2000
        .ColWidth(mGrdCol.号码) = 1600
        
        .ColAlignment(mGrdCol.号码) = 1
        
        .PrimaryCol = mGrdCol.科室
        .LocateCol = mGrdCol.号码
        .AllowAddRow = False
        .Active = True
    End With
    
    With Bill卫材科室编号
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.选择) = "选择"
        .TextMatrix(0, mGrdCol.科室) = "科室"
        .TextMatrix(0, mGrdCol.号码) = "号码"

        .ColData(mGrdCol.选择) = 5
        .ColData(mGrdCol.科室) = 5
        .ColData(mGrdCol.号码) = 4


        .ColWidth(mGrdCol.选择) = 0
        .ColWidth(mGrdCol.科室) = 2000
        .ColWidth(mGrdCol.号码) = 1600
        
        .ColAlignment(mGrdCol.号码) = 1
        
        .PrimaryCol = mGrdCol.科室
        .LocateCol = mGrdCol.号码
        .AllowAddRow = False
        .Active = True
    End With
    
    '预交票据格式
    Call InitBillForamt(vfgBillFormat)
    
    'Ris启动控制
    LoadRisEnables
    
    'RIS 分院设置
    LoadRisBranchHosp
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnModi As Boolean, i As Long
    
    If Not mblnOk Then
        For i = 0 To vsDept.Count - 1
            With vsDept(i)
                If .Tag = "已修改" Then
                    blnModi = True
                    Exit Sub
                End If
            End With
        Next
        If blnModi = False Then
            blnModi = lvw社区.Tag = "已修改" Or lvwNo.Tag = "已修改" Or cbo(cbo_住院号规则).Tag = "已修改" _
                Or cbo(cbo_门诊号规则).Tag = "已修改" Or cbo(cbo_留观号规则).Tag = "已修改" _
                Or Bill药品科室编号.Tag = "已修改" Or Bill卫材科室编号.Tag = "已修改" Or vsfRISEnables.Tag = "已修改" _
                Or vsfBranchHosp.Tag = "已修改" Or txtMainHosp.Tag = "已修改"
        End If
        
        If Not blnModi Then
            blnModi = ThirdSvrChanged
        End If
        
        mrsPar.Filter = "(修改状态=1 ANd ErrType =Null) OR  (修改状态=1 And ErrType=" & PET_值超限 & ")"
        If mrsPar.RecordCount > 0 Or blnModi Then
            If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mobjESign = Nothing
    Set mrsPar = Nothing
    Set mrsSvr = Nothing
End Sub


Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub dtp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(dtp, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    If Me.Visible Then Call SetParChange(txt, Index, mrsPar)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Index = txt_诊断查找天数 Then
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If Index = txt_诊断查找天数 Then
        If Val(txt(Index).Text) <= 0 Then
            txt(Index).Text = 0
        ElseIf Val(txt(Index).Text) > 999 Then
            txt(Index).Text = 999
        End If
    End If
End Sub

Private Sub txtMainHosp_Change()
    If Me.Visible And txtMainHosp.Tag = "" Then txtMainHosp.Tag = "已修改"
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    If Index = cbo_留观号规则 Then
        Call zlCommFun.ShowTipInfo(cbo(cbo_留观号规则).hwnd, "控制说明" & vbCrLf & "决定住院留观登记时，留观号的生成规则。", True, True, 8800)
    Else
        Call SetParTip(cbo, Index, mrsPar)
    End If
End Sub


Private Sub Init转科出院不检查项目(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    If strIn = "" Then Exit Sub
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,名称 from 诊疗项目目录 where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnCheckItem
        .Row = 0: .Col = 0
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(.Row, .Col) = rsTmp!名称 & ""
            .Cell(flexcpData, .Row, .Col) = rsTmp!ID & ""
            Call EnterNextCell(vsUnCheckItem)
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get转科出院不检查项目() As String
    Dim i As Integer
    Dim j As Integer
    Dim strIds As String
    
    With vsUnCheckItem
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    strIds = strIds & "|" & Val(.Cell(flexcpData, i, j))
                End If
            Next
        Next
    End With
    Get转科出院不检查项目 = Mid(strIds, 2)
End Function


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
            
            If Bill药品科室编号.Visible Then
                Call LocateDept(strFind, Bill药品科室编号, 1) '第0列是隐藏列
            Else
                Call LocateDeptSign(strFind)
            End If
        End Select
    End If
End Sub

Private Sub LocateDeptSign(ByVal strFind As String)
'功能：查找启用电子签名的科室
    Dim i As Long
    
    With vsDept(sstSign.Tab)
        For i = mlngPreFind To .Rows - 1
            If .RowHidden(i) = False Then
                If .TextMatrix(i, col_名称) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or _
                    .TextMatrix(i, col_简码) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or .TextMatrix(i, col_编码) = strFind Then
                    .Row = i: .ShowCell i, col_名称
                    Exit For
                End If
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

Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object, ByVal lngCol As Long)
'功能：查找科室
'参数：lngCol-进行查找的列
    Dim i As Long, lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    With objTmp
        If TypeName(objTmp) = "ListView" Then 'lvw
            lngRows = .ListItems.Count
            For i = mlngPreFind To lngRows
                If .ListItems(i).ListSubItems(lngCol).Text Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    Call .ListItems(i).EnsureVisible
                    .ListItems(i).Selected = True
                    .SetFocus
                    Exit For
                End If
            Next
        ElseIf TypeName(objTmp) = "ListBox" Then 'lst
            With objTmp
                lngRows = .ListCount - 1
                
                lngStart = IIF(mlngPreFind = 1, 0, mlngPreFind)
                For i = lngStart To .ListCount - 1
                    strCode = Split(.List(i), "-")(0)
                    strName = Split(.List(i), "-")(1)
                    If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                        .ListIndex = i
                        .SetFocus
                        Exit For
                    End If
                Next
            End With
        Else
            lngRows = objTmp.Rows
            For i = mlngPreFind To .Rows - 1
                If InStr(.TextMatrix(i, lngCol), "-") > 0 Then
                    strCode = Split(.TextMatrix(i, lngCol), "-")(0)
                    strName = Split(.TextMatrix(i, lngCol), "-")(1)
                Else
                    strCode = ""
                    strName = .TextMatrix(i, lngCol)
                End If
                
                If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    objTmp.SetFocus
                    .Row = i: .Col = lngCol
                    .TopRow = i
                    Exit For
                End If
            Next
        End If
    End With
    If i < lngRows Then
        mlngPreFind = i + 1
    Else
        If mlngPreFind = 1 Then
            MsgBox "没有找到匹配的，请检查输入的内容。", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        Else
            MsgBox "全部找完了，后面没有了。", vbInformation, Me.Caption
            mlngPreFind = 1
        End If
    End If
End Sub

Private Sub vfgBillFormat_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer
    Dim strCol As String
    
    If Me.Visible Then
        With vfgBillFormat
            If Col = vfgBillFormat.ColIndex("票据格式") Or Col = vfgBillFormat.ColIndex("预交打印方式") Then
                For i = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                        If vfgBillFormat.ColIndex("票据格式") = Col Then
                            strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                        Else
                            strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("预交打印方式")), 1))
                        End If
                    End If
                Next
                If strValue <> "" Then strValue = Mid(strValue, 2)
                Call SetParChange(vfgBillFormat, 0, mrsPar, True, strValue, CStr(Col))
            End If
        End With
    End If
End Sub

Private Sub vfgBillFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub vfgBillFormat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vfgBillFormat
        If .MouseCol = .ColIndex("票据格式") Then
            Call SetParTip(vfgBillFormat, 0, mrsPar, , , CStr(.ColIndex("票据格式")))
        ElseIf .MouseCol = .ColIndex("预交打印方式") Then
            Call SetParTip(vfgBillFormat, 0, mrsPar, , , CStr(.ColIndex("预交打印方式")))
        End If
    End With
End Sub

Private Sub vsDept_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Me.Visible And vsDept(Index).Tag = "" Then vsDept(Index).Tag = "已修改"
End Sub

Private Sub vsfBranchHosp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsfBranchHosp.AutoSize(Row, Col)
End Sub

Private Sub vsfBranchHosp_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim lngLen As Long
    
    If Me.Visible And vsfBranchHosp.Tag = "" Then vsfBranchHosp.Tag = "已修改"
    
    If Col = col_RIS分院名称 Then
        lngLen = 100
    Else
        lngLen = 20
    End If
    
    If Len(vsfBranchHosp.TextMatrix(Row, Col)) > lngLen Then
        MsgBox "“影像信息系统---HIS医院设置---分院设置”中，" & vbCrLf & vbCrLf & "数据超出规定长度，只保留前" & lngLen & "位。", vbInformation, gstrSysName
        vsfBranchHosp.TextMatrix(Row, Col) = Left(vsfBranchHosp.TextMatrix(Row, Col), lngLen)
    End If
End Sub

Private Sub vsfBranchHosp_Click()
    If vsfBranchHosp.Rows = 1 Then
        vsfBranchHosp.Rows = 2
        vsfBranchHosp.TextMatrix(1, col_RIS分院序号) = 1
    End If
End Sub

Private Sub vsfBranchHosp_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBranchHosp
        If .Rows <= 1 Then Exit Sub
        
        If KeyCode = vbKeyReturn And .Col = col_ris分院数据库服务名 And .Row = .Rows - 1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, col_RIS分院序号) = .Rows - 1
        End If
    End With
End Sub

Private Sub vsfRisDepts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col_Ris科室选择 Then
        '向RIS控制表格回写科室名称和ID串
         Call WriteDeptsIntoVsfRisEnables
    End If
End Sub

Private Sub vsfRisDepts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_Ris科室选择 Then Cancel = True
End Sub

Private Sub vsfRisDepts_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    If Col = col_Ris科室选择 Then
        With vsfRisDepts
            If .MouseCol = col_Ris科室选择 And .MouseRow = 0 Then
                If .ColData(col_Ris科室选择) = RIS_Checked Then
                    .Cell(flexcpPicture, 0, col_Ris科室选择) = imgCheck.ListImages("UnChecked").Picture
                    .ColData(col_Ris科室选择) = ""
                Else
                    .Cell(flexcpPicture, 0, col_Ris科室选择) = imgCheck.ListImages("Checked").Picture
                    .ColData(col_Ris科室选择) = RIS_Checked
                End If
                
                For i = 1 To .Rows - 1
                    If .ColData(col_Ris科室选择) = RIS_Checked Then
                        .Cell(flexcpChecked, i, col_Ris科室选择) = 1
                    Else
                        .Cell(flexcpChecked, i, col_Ris科室选择) = 2
                    End If
                Next i
                
                Call WriteDeptsIntoVsfRisEnables
                
            End If
        End With
    End If
End Sub

Private Sub vsfRISEnables_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    
    If Me.Visible And vsfRISEnables.Tag = "" Then vsfRISEnables.Tag = "已修改"
    '如果选择了预约科室的全部，则清空预约科室
    If Col = col_RIS启用预约科室全 Then
        With vsfRISEnables
            If .Cell(flexcpChecked, Row, col_RIS启用预约科室全) = 1 Then
                .TextMatrix(Row, col_RIS启用预约科室ID) = ""
                .TextMatrix(Row, col_RIS启用预约科室) = ""
                Call .AutoSize(col_RIS启用科室, col_RIS启用预约科室)
            End If
        End With
    End If
    
    '如果取消了场合，则取消RIS科室，取消预约科室
    If Col = col_RIS启用场合 Then
        With vsfRISEnables
            If .Cell(flexcpChecked, Row, col_RIS启用场合) = 2 Then
                .TextMatrix(Row, col_RIS启用科室) = ""
                .TextMatrix(Row, col_RIS启用科室ID) = ""
                .Cell(flexcpChecked, Row, col_RIS启用预约科室全) = 2
                .TextMatrix(Row, col_RIS启用预约科室ID) = ""
                .TextMatrix(Row, col_RIS启用预约科室) = ""
                Call .AutoSize(col_RIS启用科室, col_RIS启用预约科室)
            End If
        End With
    End If
    '如果取消了检查类别，则取消下面所有选项
    If Col = col_RIS启用检查类型 Then
        With vsfRISEnables
            .Cell(flexcpChecked, Row, col_RIS启用场合) = 2
            .TextMatrix(Row, col_RIS启用科室) = ""
            .TextMatrix(Row, col_RIS启用科室ID) = ""
            .Cell(flexcpChecked, Row, col_RIS启用预约科室全) = 2
            .TextMatrix(Row, col_RIS启用预约科室ID) = ""
            .TextMatrix(Row, col_RIS启用预约科室) = ""
            
            If .TextMatrix(Row + 1, col_RIS启用检查类型) = .TextMatrix(Row, col_RIS启用检查类型) Then
                .Cell(flexcpChecked, Row + 1, col_RIS启用场合) = 2
                .TextMatrix(Row + 1, col_RIS启用科室) = ""
                .TextMatrix(Row + 1, col_RIS启用科室ID) = ""
                .Cell(flexcpChecked, Row + 1, col_RIS启用预约科室全) = 2
                .TextMatrix(Row + 1, col_RIS启用预约科室ID) = ""
                .TextMatrix(Row + 1, col_RIS启用预约科室) = ""
            End If
            
            If .TextMatrix(Row + 2, col_RIS启用检查类型) = .TextMatrix(Row, col_RIS启用检查类型) Then
                .Cell(flexcpChecked, Row + 2, col_RIS启用场合) = 2
                .TextMatrix(Row + 2, col_RIS启用科室) = ""
                .TextMatrix(Row + 2, col_RIS启用科室ID) = ""
                .Cell(flexcpChecked, Row + 2, col_RIS启用预约科室全) = 2
                .TextMatrix(Row + 2, col_RIS启用预约科室ID) = ""
                .TextMatrix(Row + 2, col_RIS启用预约科室) = ""
            End If
            Call .AutoSize(col_RIS启用科室, col_RIS启用预约科室)
        End With
    End If
End Sub

Private Sub vsfRISEnables_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col_RIS启用科室 Or Col = col_RIS启用预约科室 Then Cancel = True
    If Col = col_RIS启用场合 Then
        If vsfRISEnables.Cell(flexcpChecked, Row, col_RIS启用检查类型) = 2 Then Cancel = True
    End If
    If Col = col_RIS启用预约科室全 Then
        If vsfRISEnables.Cell(flexcpChecked, Row, col_RIS启用场合) = 2 Then Cancel = True
    End If
End Sub

Private Sub vsfRISEnables_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Me.Visible And vsfRISEnables.Tag = "" Then vsfRISEnables.Tag = "已修改"
End Sub

Private Sub vsfRISEnables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strDeptIDs As String
    Dim lngSource As Long
    
    With vsfRISEnables
        '从正面来控制科室的选择，先选择了检查类型和场合，才能选择RIS科室；先选择了RIS场合，才能选择预约科室
        lngSource = IIF(.TextMatrix(.MouseRow, col_RIS启用场合) = "门诊", 1, IIF(.TextMatrix(.MouseRow, col_RIS启用场合) = "住院", 2, 4))
        If .MouseRow >= 1 And (.MouseCol = col_RIS启用科室 Or .MouseCol = col_RIS启用预约科室) And lngSource <> 4 And .Cell(flexcpChecked, .MouseRow, col_RIS启用检查类型) = 1 And .Cell(flexcpChecked, .MouseRow, col_RIS启用场合) = 1 Then
            strDeptIDs = vsfRISEnables.TextMatrix(vsfRISEnables.MouseRow, IIF(vsfRISEnables.MouseCol = col_RIS启用科室, col_RIS启用科室ID, col_RIS启用预约科室ID))
            Call LoadRisDepts(strDeptIDs, lngSource)
            vsfRisDepts.Visible = True
        Else
            vsfRisDepts.Visible = False
        End If
    End With
End Sub

Private Sub vsgInput_DblClick(Index As Integer)
    Call SetInputItemValue(Index)
End Sub

Private Sub vsgInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsgInput(Index)
            Select Case Index
            Case VSGInput_病人入院输入项设置, VSGInput_病人信息输入项设置
                If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                   If cmdOk.Enabled And cmdOk.Visible Then cmdOk.SetFocus
                   Exit Sub
                End If
                
                zlVsMoveGridCell vsgInput(Index), 1, .Cols - 1
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call SetInputItemValue(Index)
End Sub

Private Sub vsgInput_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsgInput, Index, mrsPar)
End Sub

Private Sub vsUnCheckItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsUnCheckItem.ComboList = "..."
End Sub

Private Sub vsUnCheckItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSQL = "select A.ID,A.编码,A.名称 from 诊疗项目目录 A Where A.类别 not in('4','5','6','7') and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) Order By 编码"
    With vsUnCheckItem
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "诊疗项目", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            If SetItemInput(Row, Col, rsTmp) Then
                Call vsUnCheckItem_AfterRowColChange(-1, -1, Row, Col)
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有可用的诊疗项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsUnCheckItem_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = Get转科出院不检查项目
        Call SetParChange(vsUnCheckItem, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub vsUnCheckItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    With vsUnCheckItem
        If KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsUnCheckItem_KeyPress(KeyCode)
            
        ElseIf KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
             
        ElseIf KeyCode = vbKeyReturn Then
            Call EnterNextCell(vsUnCheckItem)
        End If
        
    End With
End Sub

Private Sub vsUnCheckItem_KeyPress(KeyAscii As Integer)
    With vsUnCheckItem
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUnCheckItem_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsUnCheckItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUnCheckItem, 0, mrsPar)
End Sub

Private Sub vsUnCheckItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsUnCheckItem
        If .EditText = CStr(.TextMatrix(Row, Col)) Then
            Call EnterNextCell(vsUnCheckItem)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select DISTINCT A.ID,A.编码,A.名称 from 诊疗项目目录 A, 诊疗项目别名 B where " & _
            " a.Id = b.诊疗项目id And B.码类=1 And B.性质=1 And A.类别 not in('4','5','6','7') and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(B.简码) Like [2])" & _
            " Order by A.编码"
        With vsUnCheckItem
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "诊疗项目", False, "", "", False, False, True, _
                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                If SetItemInput(Row, Col, rsTmp) Then
                    .EditText = .TextMatrix(Row, Col)
                     Call EnterNextCell(vsUnCheckItem)
                     Exit Sub
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsUnCheckItem_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
    End With
End Sub


Private Function SetItemInput(ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset) As Boolean
    '先检查下表格中是否存在
    Dim i As Long, j As Long
    
    With vsUnCheckItem
        For i = .FixedCols To .Cols - 1
            For j = .FixedRows To .Rows - 1
                If .Cell(flexcpData, j, i) = rsTmp!ID & "" Then
                    MsgBox "该诊疗项目已经加入列表中，请查看。", vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        Next
        
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!ID & ""
        .TextMatrix(lngRow, lngCol) = rsTmp!名称 & ""
        SetItemInput = True
        
    End With
End Function


Private Sub Save电子签名()
    Dim i As Integer, j As Long
    Dim strDept As String
    
    On Error GoTo ErrHandle
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            If .Tag = "已修改" Then
                strDept = ""
                For j = 1 To .Rows - 1
                    If .Cell(flexcpChecked, j, col_选择) = 1 Then
                        strDept = strDept & "," & .RowData(j)
                    End If
                Next
                gstrSQL = "Zl_电子签名启用部门_Update(" & i & ",'" & Mid(strDept, 2) & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                
                .Tag = ""
            End If
        End With
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Save社区接口()
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If lvw社区.Tag = "已修改" Then
        For i = 1 To lvw社区.ListItems.Count
            With lvw社区.ListItems(i)
                gstrSQL = "Zl_社区目录_启用(" & Mid(.Key, 2) & "," & IIF(.SubItems(4) <> "", 1, 0) & ")"
            End With
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
        lvw社区.Tag = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub lvw社区_DblClick()
    If Not lvw社区.SelectedItem Is Nothing Then
        If lvw社区.SelectedItem.SubItems(4) <> "" Then
            lvw社区.SelectedItem.SubItems(4) = ""
        Else
            lvw社区.SelectedItem.SubItems(4) = "√"
        End If
        lvw社区.Tag = "已修改"
        
        Call lvw社区_ItemClick(lvw社区.SelectedItem)
    End If
End Sub

Private Sub lvw社区_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmd社区参数.Enabled = Item.SubItems(4) <> ""
End Sub

Private Sub lvw社区_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Call lvw社区_DblClick
    End If
End Sub


Private Sub cmd社区参数_Click()
    Dim objCommunity As Object
    
    If lvw社区.SelectedItem Is Nothing Then Exit Sub
    If lvw社区.SelectedItem.SubItems(4) = "" Then
        MsgBox lvw社区.SelectedItem.SubItems(1) & "没有启用。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '先保存设置数据：因为接口初始化要判断是否启用
    Call Save社区接口
    
    '创建部件
    Err.Clear: On Error Resume Next
    Set objCommunity = CreateObject("zlCommunity.clsCommunity")
    Err.Clear: On Error GoTo 0
    
    '调用功能
    If Not objCommunity Is Nothing Then
        If objCommunity.Initialize(gcnOracle) Then
            Call objCommunity.Setup(Val(Mid(lvw社区.SelectedItem.Key, 2)))
        End If
    Else
        MsgBox "社区公共接口没有正确安装！", vbExclamation, gstrSysName
    End If
    
    Set objCommunity = Nothing
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function Load社区接口() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    strSQL = "Select 序号, 名称, 说明, 启用, 部件名 From 社区目录 Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw社区.ListItems.Add(, "_" & rsTmp!序号, rsTmp!序号)
        ObjItem.SubItems(1) = rsTmp!名称
        ObjItem.SubItems(2) = NVL(rsTmp!说明)
        ObjItem.SubItems(3) = rsTmp!部件名
        ObjItem.SubItems(4) = IIF(NVL(rsTmp!启用, 0) = 1, "√", "")
        rsTmp.MoveNext
    Loop
    
    If Not lvw社区.SelectedItem Is Nothing Then
        Call lvw社区_ItemClick(lvw社区.SelectedItem)
    End If
    Load社区接口 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Load单据编码规则() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lst As ListItem
    
    gstrSQL = "" & _
        "   Select 项目序号,项目名称,编号规则,decode(编号规则,2,'2-按执行科室分月编号',0,'0-按年顺序编号',1,'1-按日顺序编号','0-按年顺序编号') as 编号规则说明 " & _
        "   From 号码控制表 " & _
        "   where 项目序号 in ( 11,12,13,14,15,16,21,22,23,24,25,26,27,28,29,32,62,68,69,70,71,72,73,74,75,76,77) order by 项目序号 "
    
    Err = 0: On Error GoTo ErrHand:
    Load单据编码规则 = False
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    With rsTmp
        lvwNo.ListItems.Clear
        Do While Not rsTmp.EOF
            Set lst = lvwNo.ListItems.Add(, "K" & NVL(!项目序号, 0), NVL(!项目名称))
            lst.SubItems(1) = NVL(!编号规则说明)
            If NVL(!项目序号) >= 1 And NVL(!项目序号) <= 16 Then
                lst.ForeColor = &HC85422
                lvwNo.ListItems("K" & NVL(!项目序号, 0)).ListSubItems(1).ForeColor = &HC85422
            End If
            If NVL(!项目序号) >= 21 And NVL(!项目序号) <= 62 Then
                lst.ForeColor = &H68588
                lvwNo.ListItems("K" & NVL(!项目序号, 0)).ListSubItems(1).ForeColor = &H68588
            End If
            If NVL(!项目序号) >= 68 And NVL(!项目序号) <= 77 Then
                lst.ForeColor = &H856701
                lvwNo.ListItems("K" & NVL(!项目序号, 0)).ListSubItems(1).ForeColor = &H856701
            End If
            lst.Tag = NVL(!编号规则, 0)
            If lvwNo.SelectedItem Is Nothing Then
                lst.Selected = True
            End If
            .MoveNext
        Loop
    End With
    
    '2-住院号，3-门诊号，6-住院留关号
    Set rsTmp = New ADODB.Recordset
    gstrSQL = "Select 项目序号,编号规则 as 参数值 From 号码控制表 Where 项目序号 in (2,3,6)"
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    rsTmp.Filter = "项目序号=2"
    If rsTmp.RecordCount > 0 Then cbo(cbo_住院号规则).ListIndex = Val("" & rsTmp!参数值)
    rsTmp.Filter = "项目序号=3"
    If rsTmp.RecordCount > 0 Then cbo(cbo_门诊号规则).ListIndex = Val("" & rsTmp!参数值)
    rsTmp.Filter = "项目序号=6"
    If rsTmp.RecordCount > 0 Then cbo(cbo_留观号规则).ListIndex = Val("" & rsTmp!参数值)
    
    Load单据编码规则 = True

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub LoadSign()
'功能：加载电子签名启用部门
    Dim rsTmp As New Recordset
    Dim i As Long, lngTmp As Long
    
    gstrSQL = "select 部门ID,场合 from 电子签名启用部门"
    On Error GoTo ErrHandle
     Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            rsTmp.Filter = "场合=" & i
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                lngTmp = .FindRow(Val(rsTmp!部门ID & ""))
                If lngTmp <> -1 Then
                    .Cell(flexcpChecked, lngTmp, col_选择) = 1
                End If
                rsTmp.MoveNext
            Loop
            
        End With
    Next
    For i = 0 To sstSign.Tabs - 1
        If sstSign.TabVisible(i) = True Then sstSign.Tab = i: Exit For
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub lvwNo_DblClick()
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    Call Set单据编码规则
End Sub

Private Sub lvwNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        If lvwNo.SelectedItem Is Nothing Then Exit Sub
        Call Set单据编码规则
    End If
End Sub

Private Sub Set单据编码规则()
'改变编码规则
    Dim strNo As String
    
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    strNo = lvwNo.SelectedItem.SubItems(1) & "-"
    Select Case Split(strNo, "-")(0)
        Case 0
            If Mid(lvwNo.SelectedItem.Key, 2) >= 11 And Mid(lvwNo.SelectedItem.Key, 2) <= 16 Then
                strNo = "1-按日顺序编号"
                lvwNo.SelectedItem.Tag = "1"
            Else
                strNo = "2-按执行科室分月编号"
                lvwNo.SelectedItem.Tag = "2"
            End If
        Case 1
            strNo = "0-按年顺序编号"
            lvwNo.SelectedItem.Tag = "0"
        Case 2
            strNo = "0-按年顺序编号"
            lvwNo.SelectedItem.Tag = "0"
    End Select
    lvwNo.SelectedItem.SubItems(1) = strNo
    
    lvwNo.Tag = "已修改"
End Sub

Sub Save单据编码规则()
    Dim lst As ListItem
    
    On Error GoTo ErrHandle
    If lvwNo.Tag = "已修改" Then
        For Each lst In lvwNo.ListItems
            gstrSQL = "ZL_号码控制表_Rule(" & Mid(lst.Key, 2) & "," & Val(lst.Tag) & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        Next
        lvwNo.Tag = ""
    End If
    
    '2-住院号,3-门诊号
    If cbo(cbo_住院号规则).Tag = "已修改" Then
        gstrSQL = "ZL_号码控制表_Rule(2," & cbo(cbo_住院号规则).ListIndex & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        cbo(cbo_住院号规则).Tag = ""
    End If
    
    If cbo(cbo_门诊号规则).Tag = "已修改" Then
        gstrSQL = "ZL_号码控制表_Rule(3," & cbo(cbo_门诊号规则).ListIndex & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        cbo(cbo_门诊号规则).Tag = ""
    End If
    
    If cbo(cbo_留观号规则).Tag = "已修改" Then
        gstrSQL = "ZL_号码控制表_Rule(6," & cbo(cbo_留观号规则).ListIndex & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        cbo(cbo_留观号规则).Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub vsDept_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_选择 Then Cancel = True
End Sub

Private Sub vsDept_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    If Col = col_选择 Then
        Order = 0
        With vsDept(Index)
            If .MouseCol = col_选择 And .MouseRow = .FixedRows - 1 Then
                If sstSign.Enabled = False Then Exit Sub
                If .ColData(col_选择) = "Check" Then
                    .Cell(flexcpPicture, 0, col_选择) = imgCheck.ListImages("UnChecked").Picture
                    .ColData(col_选择) = ""
                Else
                    .Cell(flexcpPicture, 0, col_选择) = imgCheck.ListImages("Checked").Picture
                    .ColData(col_选择) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .ColData(col_选择) = "Check" Then
                        .Cell(flexcpChecked, i, col_选择) = 1
                    Else
                        .Cell(flexcpChecked, i, col_选择) = 2
                    End If
                    
                Next
            End If
        End With
    End If
End Sub

Private Sub vsDept_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If vsDept(Index).Row > 0 Then
            vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_选择) = IIF(vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_选择) = 1, 2, 1)
        End If
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    If Me.Visible Then Call SetParChange(txtUD, Index, mrsPar)
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

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub sstSign_Click(PreviousTab As Integer)
    mlngPreFind = 1
End Sub

Private Sub dtp_Change(Index As Integer)
    Dim intNext As Integer
    Dim blnValue As Boolean, strValue As String

    If Index < dtp_下午下班 Then
        intNext = Index + 1
        
        dtp(intNext).MinDate = dtp(Index).value
        If dtp(intNext).value < dtp(intNext).MinDate Then
            dtp(intNext).value = dtp(intNext).MinDate
            dtp_Change intNext
        End If
    End If
    
    If Me.Visible Then
        Select Case Index
        Case dtp_上午上班, dtp_上午下班
            blnValue = True
            strValue = Format(dtp(dtp_上午上班).value, "HH:mm") & " AND " & Format(dtp(dtp_上午下班).value, "HH:mm")
            If Index = dtp_上午上班 Then
                Call SetParChange(dtp, dtp_上午上班, mrsPar, blnValue, strValue)
            Else
                Call SetParChange(dtp, dtp_上午上班, mrsPar, blnValue, strValue)
            End If
            Exit Sub
        Case dtp_下午上班, dtp_下午下班
            blnValue = True
            strValue = Format(dtp(dtp_下午上班).value, "HH:mm") & " AND " & Format(dtp(dtp_下午下班).value, "HH:mm")
            If Index = dtp_下午上班 Then
                Call SetParChange(dtp, dtp_下午下班, mrsPar, blnValue, strValue)
            Else
                Call SetParChange(dtp, dtp_下午上班, mrsPar, blnValue, strValue)
            End If
            Exit Sub
        End Select
        
        Call SetParChange(dtp, Index, mrsPar, blnValue, strValue)
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo_Click(Index As Integer)
    Dim strTmp As String, i As Long
    Dim arrTmp As Variant
    Dim blnValue As Boolean, strValue As String
    
    Select Case Index
    Case cbo_电子签名认证中心
        strTmp = chk_Sign_门诊 & "," & chk_Sign_住院 & "," & chk_Sign_医技 & "," & chk_Sign_护理 & "," & _
                chk_Sign_药品 & "," & chk_Sign_lis & "," & chk_Sign_pacs & "," & chk_sign_血库
        arrTmp = Split(strTmp, ",")
        
        chk(chk_新开一组医嘱签名一次).Enabled = cbo(Index).ListIndex <> 0
        If cbo(Index).ListIndex = 0 Then chk(chk_新开一组医嘱签名一次).value = 0
        sstSign.Enabled = cbo(Index).ListIndex <> 0
        
        If cbo(Index).ListIndex = 0 Then
            For i = 1 To 8
                chk(arrTmp(i - 1)).value = 0
                chk(arrTmp(i - 1)).Enabled = False
            Next
            sstSign.TabVisible(sst_门诊) = True
        Else
            If Not chk(chk_Sign_门诊).Enabled Then
                chk(chk_Sign_门诊).value = 1
            End If
            
            For i = 1 To 8
                chk(arrTmp(i - 1)).Enabled = True
            Next
        End If
        
        blnValue = True
        strValue = Val(cbo(Index).List(cbo(Index).ListIndex))
        
        cmd(cmd_电子签名设置).Visible = False
        If cbo(Index).ListIndex <> 0 Then
            If Not mobjESign Is Nothing Then
                cmd(cmd_电子签名设置).Visible = mobjESign.SetEnabled(Val(strValue))
            End If
        End If
    Case cbo_诊断输入来源
        blnValue = True
        strValue = cbo(cbo_诊断输入来源).ListIndex + 1
    Case cbo_门诊诊断输入, cbo_住院诊断输入
        blnValue = True
        strValue = (cbo(cbo_门诊诊断输入).ListIndex + 1) & (cbo(cbo_住院诊断输入).ListIndex + 1)
        If Index = cbo_门诊诊断输入 Then
            If Me.Visible Then Call SetParChange(cbo, cbo_住院诊断输入, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_门诊诊断输入, mrsPar, blnValue, strValue)
        End If
    Case cbo_住院号规则, cbo_门诊号规则, cbo_留观号规则
        If Me.Visible Then cbo(Index).Tag = "已修改"
    
    End Select
    
    If Me.Visible Then Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
End Sub


Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    Select Case Index
    Case chk_Sign_门诊, chk_Sign_住院, chk_Sign_医技, chk_Sign_护理, chk_Sign_药品, chk_Sign_lis, chk_Sign_pacs, chk_sign_血库
            
        '在使用电子签名的情况下，至少有一个场合需要控制签名
        If cbo(cbo_电子签名认证中心).ListIndex <> 0 Then
            If chk(chk_Sign_门诊).value = 0 And chk(chk_Sign_住院).value = 0 _
                And chk(chk_Sign_医技).value = 0 And chk(chk_Sign_护理).value = 0 And chk(chk_Sign_药品).value = 0 _
                And chk(chk_Sign_lis).value = 0 And chk(chk_Sign_pacs).value = 0 And chk(chk_sign_血库).value = 0 Then
                    If Index = chk_Sign_护理 Then
                        chk(chk_Sign_药品).value = 1
                    ElseIf Index = chk_Sign_药品 Then
                         chk(chk_Sign_lis).value = 1
                    ElseIf Index = chk_Sign_lis Then
                         chk(chk_Sign_pacs).value = 1
                    ElseIf Index = chk_Sign_pacs Then
                        chk(chk_sign_血库).value = 1
                    ElseIf Index = chk_sign_血库 Then
                        chk(chk_Sign_门诊).value = 1
                    Else
                        chk(((Index - chk_Sign_门诊 + 1) Mod 4) + chk_Sign_门诊).value = 1
                    End If
            End If
        End If
        If Index = chk_Sign_护理 Then
            sstSign.TabVisible(sst_护理) = chk(chk_Sign_护理).value = 1
        ElseIf Index = chk_Sign_药品 Then
             sstSign.TabVisible(sst_药品) = chk(chk_Sign_药品).value = 1
        ElseIf Index = chk_Sign_lis Then
             sstSign.TabVisible(sst_lis) = chk(chk_Sign_lis).value = 1
        ElseIf Index = chk_Sign_pacs Then
             sstSign.TabVisible(sst_Pacs) = chk(chk_Sign_pacs).value = 1
        ElseIf Index = chk_Sign_门诊 Then
            sstSign.TabVisible(sst_门诊) = chk(chk_Sign_门诊).value = 1
        ElseIf Index = chk_Sign_住院 Then
            sstSign.TabVisible(sst_住院护士) = chk(chk_Sign_住院).value = 1
            sstSign.TabVisible(sst_住院医生) = chk(chk_Sign_住院).value = 1
        ElseIf Index = chk_Sign_医技 Then
            sstSign.TabVisible(sst_医技) = chk(chk_Sign_医技).value = 1
        ElseIf Index = chk_sign_血库 Then
            sstSign.TabVisible(sst_血库) = chk(chk_sign_血库).value = 1
        End If
        
        blnValue = True
        strValue = chk(chk_Sign_门诊).value & chk(chk_Sign_住院).value & chk(chk_Sign_医技).value & _
                   chk(chk_Sign_护理).value & chk(chk_Sign_药品).value & chk(chk_Sign_lis).value & chk(chk_Sign_pacs).value & chk(chk_sign_血库).value
                   
    Case chk_全数字只查编码, chk_全字母只查简码
        
        blnValue = True
        strValue = chk(chk_全数字只查编码).value & chk(chk_全字母只查简码).value
        If Index = chk_全数字只查编码 Then
            If Me.Visible Then Call SetParChange(chk, chk_全字母只查简码, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(chk, chk_全数字只查编码, mrsPar, blnValue, strValue)
        End If
    Case chk_病人地址结构化录入
        chk(chk_乡镇地址结构化录入).Enabled = chk(chk_病人地址结构化录入).value = 1
        If Not chk(chk_乡镇地址结构化录入).Enabled Then chk(chk_乡镇地址结构化录入).value = 0
    Case chk_启用医学影像信息系统专业版接口
        sstRIS.Enabled = chk(chk_启用医学影像信息系统专业版接口).value = 1
        vsfRISEnables.Enabled = sstRIS.Enabled
        chkShowSel.Enabled = vsfRISEnables.Enabled
    End Select
    
    If Me.Visible Then Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
End Sub



Private Sub Load药品卫材科室编号()
'功能：提取数据并显示出来
    Dim lng序号 As Long, str库房ID As String
    Dim rsTmp As New ADODB.Recordset
    Dim strType As String
    Dim strSequence As String
    
    '药品科室
    On Error GoTo ErrHandle
    strType = "('中药库','西药库','成药库','制剂室', '中药房', '西药房', '成药房')"
    strSequence = "(21,22,23,24,25,26,27,28,29,32,62)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.编码,a.名称,b.编号 " & _
        "   From 部门表 A,科室号码表 b" & _
        "   Where a.id=b.科室id and a.ID in (select distinct 部门id from 部门性质说明 where 工作性质 in " & strType & ")" & _
        "   And b.项目序号 In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.编码,a.名称,'' As 编号 " & _
        "   From 部门表 A " & _
        "   Where a.ID in (select distinct 部门id from 部门性质说明 " & _
        "   where 工作性质 in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct 科室id From 科室号码表 Where 科室id Is Not null " & _
        "   And 项目序号 In " & strSequence & ") " & _
        "   ORDER BY 编码 "
        
    zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取相关的科室"
    
    With rsTmp
        str库房ID = ""
        Do While Not .EOF
            Bill药品科室编号.TextMatrix(Bill药品科室编号.Rows - 1, mGrdCol.科室) = NVL(!名称)
            Bill药品科室编号.TextMatrix(Bill药品科室编号.Rows - 1, mGrdCol.号码) = NVL(!编号)
            Bill药品科室编号.RowData(Bill药品科室编号.Rows - 1) = !ID
            Bill药品科室编号.Rows = Bill药品科室编号.Rows + 1
            str库房ID = str库房ID & "," & rsTmp!ID
            .MoveNext
        Loop
    End With
    
    If str库房ID <> "" Then
        str库房ID = Mid(str库房ID, 2)
        Bill药品科室编号.Rows = Bill药品科室编号.Rows - 1
        Bill药品科室编号.Active = True
    Else
        Bill药品科室编号.Active = False
    End If
    
    rsTmp.Close
    
    '卫材科室
    strType = "('制剂室','卫材库','虚拟库房')"
    strSequence = "(68,69,70,71,72,73,74,75,76,77)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.编码,a.名称,b.编号 " & _
        "   From 部门表 A,科室号码表 b" & _
        "   Where a.id=b.科室id and a.ID in (select distinct 部门id from 部门性质说明 where 工作性质 in " & strType & ")" & _
        "   And b.项目序号 In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.编码,a.名称,'' As 编号 " & _
        "   From 部门表 A " & _
        "   Where a.ID in (select distinct 部门id from 部门性质说明 " & _
        "   where 工作性质 in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct 科室id From 科室号码表 Where 科室id Is Not null " & _
        "   And 项目序号 In " & strSequence & ") " & _
        "   ORDER BY 编码 "

    zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取相关的科室"
    
    With rsTmp
        str库房ID = ""
        Do While Not .EOF
            Bill卫材科室编号.TextMatrix(Bill卫材科室编号.Rows - 1, mGrdCol.科室) = NVL(!名称)
            Bill卫材科室编号.TextMatrix(Bill卫材科室编号.Rows - 1, mGrdCol.号码) = NVL(!编号)
            Bill卫材科室编号.RowData(Bill卫材科室编号.Rows - 1) = !ID
            Bill卫材科室编号.Rows = Bill卫材科室编号.Rows + 1
            str库房ID = str库房ID & "," & rsTmp!ID
            .MoveNext
        Loop
    End With
    
    If str库房ID <> "" Then
        str库房ID = Mid(str库房ID, 2)
        Bill卫材科室编号.Rows = Bill卫材科室编号.Rows - 1
        Bill卫材科室编号.Active = True
    Else
        Bill卫材科室编号.Active = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Save科室编号()
'功能：保存科室编号
    Dim i As Integer
    
    On Error GoTo ErrHandle
    With Bill药品科室编号
        If .Tag = "已修改" Then
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, mGrdCol.科室)) <> "" Then  'And Trim(.TextMatrix(i, mGrdCol.号码)) <> "" Then
                    '科室ID_IN   IN 科室编号.科室ID%TYPE,
                    '编号_IN     IN 科室编号.编号%TYPE
                    gstrSQL = "ZL_科室号码表_UPDATE("
                    gstrSQL = gstrSQL & .RowData(i) & ","
                    gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.号码)) & "',1)"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
            Next
            .Tag = ""
        End If
    End With

    With Bill卫材科室编号
        If .Tag = "已修改" Then
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, mGrdCol.科室)) <> "" Then 'And Trim(.TextMatrix(i, mGrdCol.号码)) <> "" Then
                    '科室ID_IN   IN 科室编号.科室ID%TYPE,
                    '编号_IN     IN 科室编号.编号%TYPE
                    gstrSQL = "ZL_科室号码表_UPDATE("
                    gstrSQL = gstrSQL & .RowData(i) & ","
                    gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.号码)) & "',2)"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
            Next
            .Tag = ""
        End If
    End With

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Function CheckNumberRule_Drug() As Boolean
 '功能       检查单据编码规则是否有"2"的
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 21 And Mid(.ListItems(i).Key, 2) <= 62 Then
                If .ListItems(i).SubItems(1) = "2-按执行科室分月编号" Then
                    CheckNumberRule_Drug = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Function CheckNumberRule_Stuff() As Boolean
'功能       检查单据编码规则是否有"2"的
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 68 And Mid(.ListItems(i).Key, 2) <= 77 Then
                If .ListItems(i).SubItems(1) = "2-按执行科室分月编号" Then
                    CheckNumberRule_Stuff = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Private Sub bill药品科室编号_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub bill药品科室编号_EditKeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub bill药品科室编号_EnterCell(Row As Long, Col As Long)
    With Bill药品科室编号
        Select Case .Col
            Case mGrdCol.号码
                .TxtCheck = True
                .MaxLength = 1
                .TextMask = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
                mintLastRow_Drug = Row
                mintLastCol_Drug = Col
            End Select
    End With
End Sub

Private Sub bill药品科室编号_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
     
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstrLastCode_Drug = ""
    
    
    With Bill药品科室编号
        .Tag = "已修改"
        .Text = Replace(UCase(Trim(.Text)), "'", "")
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case mGrdCol.号码
                If strKey <> "" Then
                    .Text = strKey
                End If
                If .Row = .Rows - 1 And .Col = 2 And .TextMatrix(.Row, .Col) <> "" Then
'                    zlCommFun.PressKey vbKeyTab
                    Bill卫材科室编号.SetFocus
                End If
            Case mGrdCol.科室
        End Select
    End With
End Sub

Private Sub bill药品科室编号_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789", Chr(KeyAscii)) > 0 Then
        mstrLastCode_Drug = Chr(KeyAscii)
    End If
End Sub

Private Sub bill卫材科室编号_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub bill卫材科室编号_EditKeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub bill卫材科室编号_EnterCell(Row As Long, Col As Long)
    With Bill卫材科室编号
        Select Case .Col
            Case mGrdCol.号码
                .TxtCheck = True
                .MaxLength = 1
                .TextMask = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
                mintLastRow_Stuff = Row
                mintLastCol_Stuff = Col
            End Select
    End With
End Sub

Private Sub bill卫材科室编号_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
     
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstrLastCode_Stuff = ""
    
    
    With Bill卫材科室编号
        .Tag = "已修改"
        .Text = Replace(UCase(Trim(.Text)), "'", "")
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case mGrdCol.号码
                
                If strKey <> "" Then
                    .Text = strKey
                End If
                If .Row = .Rows - 1 And .Col = 2 And .TextMatrix(.Row, .Col) <> "" Then
                    zlCommFun.PressKey vbKeyTab
                End If
            Case mGrdCol.科室
        End Select
    End With
End Sub

Private Sub bill卫材科室编号_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789", Chr(KeyAscii)) > 0 Then
        mstrLastCode_Stuff = Chr(KeyAscii)
    End If
End Sub

Private Sub InitBillForamt(ByVal vfgBill As VSFlexGrid)
    '初始化预交票据格式信息
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHand
    strSQL = "" & _
        "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
        "   Select B.说明,B.序号  " & _
        "   From zlReports A,zlRptFmts B" & _
        "   Where A.ID=B.报表ID And A.编号='" & "ZL" & glngSys \ 100 & "_BILL_1103" & "'  " & _
        "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "票据格式")
    
    With vfgBill
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("预交打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
        
        .TextMatrix(1, 0) = "门诊预交"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "住院预交"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("票据格式")) = "0"
        .ColData(.ColIndex("预交打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        .Editable = flexEDKbdMouse
    End With
        
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LaodBillForamt(ByVal vfgBill As VSFlexGrid, ByVal strBillFormat As String, ByVal strPrintMode As String)
    Dim varData As Variant, varType As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, i As Long
    
    varData = Split(strBillFormat, "|")
    varType = Split(strPrintMode, "|")
    With vfgBill
        .Tag = ""
        .Clear 1
        .Rows = 3
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
        If Val(.ColData(.ColIndex("预交打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("预交打印方式"), .Rows - 1, .ColIndex("预交打印方式")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("票据格式"), .Rows - 1, .ColIndex("票据格式")) = vbBlue
        End If
    End With
End Sub

Private Sub LoadInputItem(ByVal intIndex As Integer, ByVal strValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载输入项控制
    '入参:intIndex-索引值
    '     strValue-缺省参数值和输入项,格式:输入项目,禁止录入,光标是否跳过,必输项|....
    '编制:
    '日期:2015-06-11 17:32:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim intRow As Integer, i As Integer
    
    On Error GoTo ErrHandle
    varData = Split(strValue, "|")
    With vsgInput(intIndex)
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
'        .ColAlignment(.ColIndex("输入项目")) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandle:
    vsgInput(intIndex).redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetInputItemValue(ByVal intIndex As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置当前项目的相关值
    '入参:intIndex-网格控件数组的索引值
    '编制:
    '日期:2015-06-11 17:58:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
       
    On Error GoTo ErrHandle
    With vsgInput(intIndex)
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
    Call SetParChange(vsgInput, intIndex, mrsPar, True, GetInputItemSetValue(intIndex))
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
    '编制:
    '日期:2015-06-11 18:10:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTmp As String
    On Error GoTo ErrHandle
        
    With vsgInput(intIndex)
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

Private Sub LoadRisEnables()
'-----------------------------------------------------------
'功能:加载RIS启用控制列表
'入参:
'返回:无..
'-----------------------------------------------------------
   
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim strModality As String
    Dim lngSource As Long
    Dim j As Integer
    Dim strDeptIDs As String
    Dim strDeptNames As String
    
    On Error GoTo Err
    
    '查询所有的检查类型
    
    strSQL = "select 编码, 名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取影像检查类别")
    
    If rsTemp.EOF = True Then Exit Sub
    
    With vsfRISEnables
        .Rows = 1 + rsTemp.RecordCount * 3
        .Cols = 7
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True

        .TextMatrix(0, col_RIS启用检查类型) = "检查类型"
        .TextMatrix(0, col_RIS启用场合) = "场合"
        .TextMatrix(0, col_RIS启用科室) = "科室"
        .TextMatrix(0, col_RIS启用预约科室全) = "启用预约科室"
        .TextMatrix(0, col_RIS启用预约科室) = "启用预约科室"

        .ColWidth(col_RIS启用检查类型) = 850
        .ColWidth(col_RIS启用场合) = 650
        .ColWidth(col_RIS启用科室) = 2000
        .ColWidth(col_RIS启用预约科室全) = 650
        .ColWidth(col_RIS启用预约科室) = 1900
        
        .Cell(flexcpChecked, 1, col_RIS启用检查类型, .Rows - 1, col_RIS启用场合) = 2
        .Cell(flexcpChecked, 1, col_RIS启用预约科室全, .Rows - 1, col_RIS启用预约科室全) = 2
        
        For i = 0 To rsTemp.RecordCount - 1
            .TextMatrix(i * 3 + 1, col_RIS启用检查类型) = rsTemp!名称
            .TextMatrix(i * 3 + 2, col_RIS启用检查类型) = rsTemp!名称
            .TextMatrix(i * 3 + 3, col_RIS启用检查类型) = rsTemp!名称
            .RowData(i * 3 + 1) = NVL(rsTemp!编码)
            .RowData(i * 3 + 2) = NVL(rsTemp!编码)
            .RowData(i * 3 + 3) = NVL(rsTemp!编码)
            
            .TextMatrix(i * 3 + 1, col_RIS启用场合) = "门诊"
            .TextMatrix(i * 3 + 2, col_RIS启用场合) = "住院"
            .TextMatrix(i * 3 + 3, col_RIS启用场合) = "体检"
            
            .TextMatrix(i * 3 + 1, col_RIS启用预约科室全) = "全部"
            .TextMatrix(i * 3 + 2, col_RIS启用预约科室全) = "全部"
            .TextMatrix(i * 3 + 3, col_RIS启用预约科室全) = "全部"
            rsTemp.MoveNext
        Next i
        
        '设置单元格合并
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .MergeRow(0) = True
        
        
        '隐藏科室ID列
        .ColHidden(col_RIS启用科室ID) = True
        .ColHidden(col_RIS启用预约科室ID) = True
        
        '读取RIS启动控制参数，并显示
        strSQL = "select a.检查类型,a.场合,a.部门ID,a.是否启用RIS,a.是否启用预约,b.名称 from ris启用控制 a, 部门表 b where a.部门id = b.id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取RIS启用控制")
        If rsTemp.EOF = True Then Exit Sub
        
        '循环列表，填写数据,三行一个循环
        For i = 1 To .Rows - 1 Step 3
            Call loadOneModality(vsfRISEnables, rsTemp, i)
        Next i
        
        '重新调整行的高度，确保科室能完整显示
        Call .AutoSize(col_RIS启用科室, col_RIS启用预约科室)
        
        .Refresh
        
    End With
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadRisDepts(strDeptIDs As String, lngSource As Long)
'-----------------------------------------------------------
'功能:加载RIS启用控制中分科室控制的列表
'入参:  strDeptIDs -- 部门ID串
'       lngSource -- 病人来源 1 - 门诊；2 - 住院； 4 - 体检
'返回:无..
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo Err
    
    vsfRisDepts.Clear
    
    strSQL = "Select Distinct D.ID,D.编码, D.名称,D.简码,t.服务对象 From 部门表 D, 部门性质说明 T " & _
            " Where d.Id = t.部门id And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
    If lngSource = 1 Then   '门诊
        strSQL = strSQL & " And t.服务对象 IN (1,3)  and T.工作性质 IN ('临床','手术','治疗','护理','检查','检验','营养') order by 名称 "
    Else    '住院
        strSQL = strSQL & " And t.服务对象 IN (2,3)  and T.工作性质 IN ('临床','手术','治疗','护理','检查','检验','营养') order by 名称 "
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取部门")
    If rsTemp.EOF = True Then Exit Sub
    
    strDeptIDs = "," & strDeptIDs & ","
    
    With vsfRisDepts
        .Rows = rsTemp.RecordCount + 1
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSort
        .ColSort(col_Ris科室选择) = flexSortNone
        .ExtendLastCol = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpPictureAlignment, 0, col_Ris科室选择, .Rows - 1, col_Ris科室选择) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, col_Ris科室编码, .Rows - 1, col_Ris科室编码) = flexAlignLeftCenter
        
        .Cell(flexcpPicture, 0, col_Ris科室选择) = imgCheck.ListImages("UnChecked").Picture

        .TextMatrix(0, col_Ris科室选择) = ""
        .TextMatrix(0, col_Ris科室编码) = "编码"
        .TextMatrix(0, col_Ris科室名称) = "名称"
        
        .ColWidth(col_Ris科室选择) = 400
        .ColWidth(col_Ris科室编码) = 850
        .ColWidth(col_Ris科室名称) = 1200
        
        .Cell(flexcpChecked, 1, col_Ris科室选择, .Rows - 1, col_Ris科室选择) = 2
        
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, col_Ris科室编码) = rsTemp!编码
            .TextMatrix(i, col_Ris科室名称) = rsTemp!名称
            .TextMatrix(i, col_Ris科室ID) = rsTemp!ID
            If InStr(strDeptIDs, rsTemp!ID) > 0 Then
                .Cell(flexcpChecked, i, col_Ris科室选择) = 1
            End If
            rsTemp.MoveNext
        Next i
        
        .ColHidden(col_Ris科室ID) = True
        
        .Refresh
    End With
    
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function loadOneModality(vsfGrid As VSFlexGrid, rsData As ADODB.Recordset, iRow As Integer) As Boolean
'-----------------------------------------------------------
'功能:加载一个检查类别的数据
'入参:  vsfGrid -- vsflexGrid控件
'       rsData -- 数据源
'       iRow -- 要加载的行号
'返回:True -- 成功； False -- 失败
'-----------------------------------------------------------
    Dim lngSource As Long
    Dim strModality As String
    Dim strDeptIDs As String
    Dim strDeptNames As String
    Dim i As Integer
    
    On Error GoTo Err
    
    With vsfGrid
        '判断启用RIS和检查类型是否被选中
        strModality = .RowData(iRow)
        rsData.Filter = " 检查类型='" & strModality & "' and 是否启用RIS=1"
        If rsData.EOF = False Then
            '先勾选检查类型，再逐个判断门诊，住院，体检的选择情况
            .Cell(flexcpChecked, iRow, col_RIS启用检查类型) = 1
            .Cell(flexcpChecked, iRow + 1, col_RIS启用检查类型) = 1
            .Cell(flexcpChecked, iRow + 2, col_RIS启用检查类型) = 1
            
            For i = 0 To 2
                lngSource = IIF(.TextMatrix(iRow + i, col_RIS启用场合) = "门诊", 1, IIF(.TextMatrix(iRow + i, col_RIS启用场合) = "住院", 2, 4))
                If GetDeptString(rsData, lngSource, strDeptNames, strDeptIDs) = False Then
                    '场合没有被选中，不用处理
                Else
                    '先勾选场合，再判断是否有科室
                    .Cell(flexcpChecked, iRow + i, col_RIS启用场合) = 1
                    If strDeptIDs = "" Or lngSource = 4 Then
                        '没有选择科室，留空，不用处理
                        '体检作为一个单一科室，只区分场合，不区分科室
                    Else
                        '选择了科室，则填写科室
                        .TextMatrix(iRow + i, col_RIS启用科室) = strDeptNames
                        .TextMatrix(iRow + i, col_RIS启用科室ID) = strDeptIDs
                    End If
                End If
            Next i
            
            '最后判断是否选择了门诊、住院、体检其中之一,如果没有选择，则取消检查类型的勾选项
            If .Cell(flexcpChecked, iRow, col_RIS启用场合) = 2 And .Cell(flexcpChecked, iRow + 1, col_RIS启用场合) = 2 And .Cell(flexcpChecked, iRow + 2, col_RIS启用场合) = 2 Then
                .Cell(flexcpChecked, iRow, col_RIS启用检查类型) = 2
                .Cell(flexcpChecked, iRow + 1, col_RIS启用检查类型) = 2
                .Cell(flexcpChecked, iRow + 2, col_RIS启用检查类型) = 2
                '不用再处理预约数据，直接退出
                Exit Function
            End If
            
            '判断是否启用了预约
            rsData.Filter = " 检查类型='" & strModality & "' and 是否启用预约=1"
            If rsData.EOF = False Then
                '逐个判断门诊，住院，体检的选择情况
                For i = 0 To 2
                    lngSource = IIF(.TextMatrix(iRow + i, col_RIS启用场合) = "门诊", 1, IIF(.TextMatrix(iRow + i, col_RIS启用场合) = "住院", 2, 4))
                    '该场合启用了RIS，才填写预约
                    If .Cell(flexcpChecked, iRow + i, col_RIS启用场合) = 1 Then
                        If GetDeptString(rsData, lngSource, strDeptNames, strDeptIDs) = True Then
                            If strDeptIDs = "" Or lngSource = 4 Then
                                '按场合启用了预约,或者是体检
                                .Cell(flexcpChecked, iRow + i, col_RIS启用预约科室全) = 1
                            Else
                                '按科室启用了预约
                                .TextMatrix(iRow + i, col_RIS启用预约科室) = strDeptNames
                                .TextMatrix(iRow + i, col_RIS启用预约科室ID) = strDeptIDs
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        
    End With
    
    loadOneModality = True
    Exit Function
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDeptString(rsData As ADODB.Recordset, lngSource As Long, ByRef strDeptNames As String, ByRef strDeptIDs As String) As Boolean
'-----------------------------------------------------------
'功能:提取部门ID和部门名称串
'入参:  rsData -- 数据源
'       lngSource -- 病人来源，1=门诊；2=住院；4=体检
'       strDeptNames -- 返回值，部门名称串
'       strDeptIDs -- 返回值，部门ID串
'返回:True -- 成功； False -- 失败
'-----------------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFilter As String
    
    On Error GoTo Err
    
    strDeptNames = ""
    strDeptIDs = ""
    
    strFilter = rsData.Filter
    
    rsData.Filter = strFilter & " and 场合=" & lngSource
    
    If rsData.EOF = True Then
        rsData.Filter = strFilter
        GetDeptString = False
        Exit Function
    Else
        If rsData.RecordCount = 1 And IsNull(rsData!部门ID) Then
            '使用默认返回值
        Else
            '组合部门ID和名称串
            For i = 1 To rsData.RecordCount
                strDeptIDs = strDeptIDs & "," & NVL(rsData!部门ID)
                strDeptNames = strDeptNames & "," & NVL(rsData!名称)
                rsData.MoveNext
            Next i
            
            strDeptIDs = Mid(strDeptIDs, 2)
            strDeptNames = Mid(strDeptNames, 2)
        End If
    End If
    
    '恢复原来的Filter，这样就不需要复制一个数据集了
    rsData.Filter = strFilter
    GetDeptString = True
    Exit Function
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveRisEnable()
'-----------------------------------------------------------
'功能:保存RIS分科室启用设置,保存“影像信息系统”启用控制
'入参:
'返回:
'-----------------------------------------------------------

    Dim i As Integer
    Dim strRISDeptIDs As String
    Dim strSchDeptIDs As String
    Dim strModality As String
    Dim lngSource As Long
    Dim strSQL As String
    
    On Error GoTo Err
    
    With vsfRISEnables
        If .Tag = "已修改" Then
            strSQL = "b_Zlxwinterface.RIS启用控制_Delete()"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            For i = 1 To .Rows - 1
                '选择了的检查类型，才保存
                If .Cell(flexcpChecked, i, col_RIS启用检查类型) = 1 Then
                    strModality = .RowData(i)
                    '选择了场合，才保存
                    If .Cell(flexcpChecked, i, col_RIS启用场合) = 1 Then
                        strRISDeptIDs = .TextMatrix(i, col_RIS启用科室ID)
                        lngSource = IIF(.TextMatrix(i, col_RIS启用场合) = "门诊", 1, IIF(.TextMatrix(i, col_RIS启用场合) = "住院", 2, 4))
                        
                        strSQL = "b_Zlxwinterface.RIS启用控制_Update('" & strModality & "'," & lngSource & ",'" & strRISDeptIDs & "',1)"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        
                        '选择了预约，才保存
                        strSchDeptIDs = .TextMatrix(i, col_RIS启用预约科室ID)
                        If strSchDeptIDs <> "" Or .Cell(flexcpChecked, i, col_RIS启用预约科室全) = 1 Then
                            strSQL = "b_Zlxwinterface.RIS启用控制_Update('" & strModality & "'," & lngSource & ",'" & strSchDeptIDs & "',2)"
                            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        End If
                    End If
                End If
            Next i
            .Tag = ""
        End If
    End With
    
    Exit Sub
Err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub WriteDeptsIntoVsfRisEnables()
'-----------------------------------------------------------
'功能:将科室名称写回到“RIS启用控制列表”
'入参:
'返回:
'-----------------------------------------------------------
    Dim i As Integer
    Dim strDeptNames As String
    Dim strDeptIDs As String
    Dim iSelCount As Integer
    
    On Error GoTo Err
    
    iSelCount = 0
    If vsfRISEnables.ColSel = col_RIS启用科室 Or vsfRISEnables.ColSel = col_RIS启用预约科室 Then
        With vsfRisDepts
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col_Ris科室选择) = 1 Then
                    strDeptIDs = strDeptIDs & "," & .TextMatrix(i, col_Ris科室ID)
                    strDeptNames = strDeptNames & "," & .TextMatrix(i, col_Ris科室名称)
                    iSelCount = iSelCount + 1
                End If
            Next i
            
            '判断是否全选
            If iSelCount = vsfRisDepts.Rows - 1 Then
                '全选
                If vsfRISEnables.ColSel = col_RIS启用科室 Then
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用科室ID) = ""
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用科室) = ""
                Else
                    vsfRISEnables.Cell(flexcpChecked, vsfRISEnables.RowSel, col_RIS启用预约科室全) = 1
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用预约科室ID) = ""
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用预约科室) = ""
                End If
            Else
                '部分选择
                If strDeptIDs <> "" Then
                    strDeptIDs = Mid(strDeptIDs, 2)
                    strDeptNames = Mid(strDeptNames, 2)
                End If
                
                If vsfRISEnables.ColSel = col_RIS启用科室 Then
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用科室ID) = strDeptIDs
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用科室) = strDeptNames
                Else
                    '如果是预约科室，则取消预约全选
                    vsfRISEnables.Cell(flexcpChecked, vsfRISEnables.RowSel, col_RIS启用预约科室全) = 2
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用预约科室ID) = strDeptIDs
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS启用预约科室) = strDeptNames
                End If
            End If
            Call vsfRISEnables.AutoSize(col_RIS启用科室, col_RIS启用预约科室)
        End With
    End If
    Exit Sub
Err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRisBranchHosp()
'-----------------------------------------------------------
'功能:加载RIS的 分院设置
'入参:  无
'返回:无
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo Err
    
    vsfBranchHosp.Clear
    
    strSQL = "select a.id ,a.医院名称,a.医院代码,a.用户名,a.密码,a.数据库服务名 from ris分院设置 a"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取RIS分院设置")
    
    With vsfBranchHosp
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount)
        .Cols = 6
        .FixedRows = 1
        .FixedCols = 1
        .RowHeightMin = 400
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        
        .Cell(flexcpAlignment, 0, 0, 0, col_ris分院数据库服务名) = flexAlignCenterCenter
        
        .TextMatrix(0, col_RIS分院序号) = "序号"
        .TextMatrix(0, col_RIS分院名称) = "医院名称"
        .TextMatrix(0, col_ris分院代码) = "医院代码"
        .TextMatrix(0, col_ris分院用户名) = "用户名"
        .TextMatrix(0, col_ris分院密码) = "密码"
        .TextMatrix(0, col_ris分院数据库服务名) = "数据库服务名"
        
        .ColWidth(col_RIS分院序号) = 600
        .ColWidth(col_RIS分院名称) = 1600
        .ColWidth(col_ris分院代码) = 1600
        .ColWidth(col_ris分院用户名) = 1600
        .ColWidth(col_ris分院密码) = 1600
        .ColWidth(col_ris分院数据库服务名) = 1600
        
        i = 1
        While Not rsTemp.EOF
            '本院单独填写
            If rsTemp!医院名称 = "本院" Then
                txtMainHosp.Text = rsTemp!医院代码
            Else
                .TextMatrix(i, col_RIS分院序号) = i
                .TextMatrix(i, col_RIS分院名称) = rsTemp!医院名称
                .TextMatrix(i, col_ris分院代码) = rsTemp!医院代码
                .TextMatrix(i, col_ris分院用户名) = NVL(rsTemp!用户名)
                .TextMatrix(i, col_ris分院密码) = NVL(rsTemp!密码)
                .TextMatrix(i, col_ris分院数据库服务名) = rsTemp!数据库服务名
                i = i + 1
            End If
            
            rsTemp.MoveNext
        Wend
        .Refresh
    End With
    
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ValidateRisBranchHosp() As Boolean
'-----------------------------------------------------------
'功能:检查RIS分院信息的有效性
'入参:无
'返回: True -- 数据有效，可以保存；False -- 数据无效，不能保存
'-----------------------------------------------------------
    Dim i As Integer
    
    On Error GoTo Err
    
    If txtMainHosp.Tag = "已修改" Or vsfBranchHosp.Tag = "已修改" Then
        '先检查数据的完整性：
        '（1）没有设置任何信息
        '（2）必须由本院代码，且最少有一条分院数据；
        '（3）分院数据中用户名和密码可以同时为空，其他内容非空。
        
        With vsfBranchHosp
            '先把vsfBranchHosp中多余的行删除掉
            Do While .Rows > 1
                If (.TextMatrix(.Rows - 1, col_RIS分院名称) = "" And .TextMatrix(.Rows - 1, col_ris分院代码) = "" _
                    And .TextMatrix(.Rows - 1, col_ris分院用户名) = "" And .TextMatrix(.Rows - 1, col_ris分院用户名) = "" _
                    And .TextMatrix(.Rows - 1, col_ris分院数据库服务名) = "") Then
                    .Rows = .Rows - 1
                Else
                    Exit Do
                End If
            Loop
        
            If vsfBranchHosp.Rows > 1 Or txtMainHosp.Text <> "" Then
                If txtMainHosp.Text = "" Then
                    MsgBox "“影像信息系统---HIS医院设置---本院设置”中，医院代码不能为空。", vbInformation, gstrSysName
                    txtMainHosp.SetFocus
                    Exit Function
                End If
                
                If vsfBranchHosp.Rows <= 1 Then
                    MsgBox "“影像信息系统---HIS医院设置---分院设置”中，最少应该设置一个分院信息。", vbInformation, gstrSysName
                    Exit Function
                End If
                
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, col_RIS分院名称) = "" Then
                        MsgBox "“影像信息系统---HIS医院设置---分院设置”中，" & vbCrLf & vbCrLf & " 医院名称不能为空，请填写医院名称。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If .TextMatrix(i, col_ris分院代码) = "" Then
                        MsgBox "“影像信息系统---HIS医院设置---分院设置”中，" & vbCrLf & vbCrLf & " 医院代码不能为空，请填写医院代码。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If .TextMatrix(i, col_ris分院数据库服务名) = "" Then
                        MsgBox "“影像信息系统---HIS医院设置---分院设置”中，" & vbCrLf & vbCrLf & " 数据库服务名不能为空，请填写数据库服务名。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If (.TextMatrix(i, col_ris分院用户名) = "" And .TextMatrix(i, col_ris分院密码) = "") Or (.TextMatrix(i, col_ris分院用户名) <> "" And .TextMatrix(i, col_ris分院密码) <> "") Then
                        '正确的情况，不用处理
                    Else
                        MsgBox "“影像信息系统---HIS医院设置---分院设置”中，" & vbCrLf & vbCrLf & " 用户名和密码可以同时为空，或者同时不为空，请按照红色字体描述的规则，填写用户名和密码。", vbInformation, gstrSysName
                        Exit Function
                    End If
                Next i
            End If
        End With
    End If
    
    ValidateRisBranchHosp = True
    Exit Function
Err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveRisBranchHosp()
'-----------------------------------------------------------
'功能:保存RIS的分院信息
'入参:无
'返回:无
'-----------------------------------------------------------
    Dim blnInTrans As Boolean       '是否在事务处理之中
    Dim arrSQL() As Variant
    Dim strSQL As String
    Dim i As Integer

    On Error GoTo Err
    
    If txtMainHosp.Tag = "已修改" Or vsfBranchHosp.Tag = "已修改" Then
        arrSQL = Array()
        
        '先清空，再重新添加
        strSQL = "b_Zlxwinterface.Ris分院设置_Delete()"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        '添加本院设置
        strSQL = "b_Zlxwinterface.Ris分院设置_Update(1,'本院','" & txtMainHosp.Text & "',null,null,null)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
                
        '添加分院设置
        With vsfBranchHosp
            For i = 1 To .Rows - 1
            
                strSQL = "b_Zlxwinterface.Ris分院设置_Update(" & Val(.TextMatrix(i, col_RIS分院序号)) + 1 _
                     & ",'" & .TextMatrix(i, col_RIS分院名称) & "','" & .TextMatrix(i, col_ris分院代码) _
                    & "','" & .TextMatrix(i, col_ris分院用户名) & "','" & .TextMatrix(i, col_ris分院密码) _
                    & "','" & .TextMatrix(i, col_ris分院数据库服务名) & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Next i
        End With
        
        gcnOracle.BeginTrans        '开始保存参数
        blnInTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存Ris分院设置")
        Next i
        gcnOracle.CommitTrans
        blnInTrans = False
    
        '保存完成之后，设置成未修改
        txtMainHosp.Tag = ""
        vsfBranchHosp.Tag = ""
    End If
    
    Exit Sub
Err:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadThirdSvr()
'功能：初始化 三方服务配置目录
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    
    strTmp = ",200,1;系统标识,1000,1;服务名称,2100,1;服务地址,3000,1"
    Call Grid.Init(vsThirdSvr, strTmp, , 1)
    vsThirdSvr.Rows = 1
    Set mrsSvr = Nothing
    strSQL = "Select a.系统标识, a.服务名称, a.服务地址 From 三方服务配置目录 A Order By a.系统标识,a.服务名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取三方服务配置目录")
    If rsTmp.EOF Then Exit Sub
    Set mrsSvr = zlDatabase.CopyNewRec(rsTmp)
    With vsThirdSvr
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, COL_系统标识) = rsTmp!系统标识 & ""
            .TextMatrix(i, COL_服务名称) = rsTmp!服务名称 & ""
            .TextMatrix(i, COL_服务地址) = rsTmp!服务地址 & ""
            rsTmp.MoveNext
        Next
    End With
      
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsThirdSvr_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> COL_服务地址 Then
        Cancel = True
    End If
End Sub

Private Function CheckThirdSvr() As Boolean
'功能：检查 三方服务配置目录
'参数：blnSave true保存参数的时候调用
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim strMsg As String
    Dim strTmp1 As String
    
    On Error GoTo errH
    If mrsSvr Is Nothing Then Exit Function
    
    With vsThirdSvr
        mrsSvr.MoveFirst
        For i = 1 To mrsSvr.RecordCount
            .TextMatrix(i, COL_服务地址) = Trim(.TextMatrix(i, COL_服务地址))
            If mrsSvr!服务地址 & "" <> .TextMatrix(i, COL_服务地址) And .TextMatrix(i, COL_服务地址) <> "" Then
                '地址发生了变化的要检查
                strTmp = ""
                Call Sys.WebAPIByBasic(.TextMatrix(i, COL_服务地址), "", strTmp1, strTmp)
                If strTmp <> "" Then
                    strMsg = IIF("" = strMsg, "", strMsg & vbCrLf) & .TextMatrix(i, COL_系统标识) & ":" & .TextMatrix(i, COL_服务名称) & "  验证：" & strTmp
                End If
            End If
            mrsSvr.MoveNext
        Next
    End With
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, Me.Caption
    Else
        CheckThirdSvr = True
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ThirdSvrTest(ByVal lngRow As Long)
'功能：按行测试服务合法性
    Dim strTmp As String
    Dim strTmp1 As String
    Dim strMsg As String
    
    With vsThirdSvr
        If lngRow < 1 Then Exit Sub
        .TextMatrix(lngRow, COL_服务地址) = Trim(.TextMatrix(lngRow, COL_服务地址))
        If "" = .TextMatrix(lngRow, COL_服务地址) Then
            MsgBox "服务地址为空，请填写！", vbInformation, Me.Caption
            Exit Sub
        End If
        Call Sys.WebAPIByBasic(.TextMatrix(lngRow, COL_服务地址), "", strTmp1, strTmp)
        If strTmp <> "" Then
            strMsg = .TextMatrix(lngRow, COL_系统标识) & ":" & .TextMatrix(lngRow, COL_服务名称) & "  验证：" & strTmp
        End If
    End With
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, Me.Caption
    Else
        MsgBox "成功！", vbInformation, Me.Caption
    End If
End Sub

Private Function ThirdSvrChanged() As Boolean
'功能：判断三方服务地址是否发生变化
    Dim i As Long
    If mrsSvr Is Nothing Then Exit Function
    
    With vsThirdSvr
        mrsSvr.MoveFirst
        For i = 1 To mrsSvr.RecordCount
            .TextMatrix(i, COL_服务地址) = Trim(.TextMatrix(i, COL_服务地址))
            If mrsSvr!服务地址 & "" <> .TextMatrix(i, COL_服务地址) Then
                '地址发生了变化的要检查
                ThirdSvrChanged = True
                Exit Function
            End If
            mrsSvr.MoveNext
        Next
    End With
End Function

Private Sub SaveThirdSvr()
'功能：保存 三方服务配置目录
    Dim blnInTrans As Boolean
    Dim i As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not ThirdSvrChanged Then Exit Sub
    If Not CheckThirdSvr Then Exit Sub
    If mrsSvr Is Nothing Then Exit Sub
    
    arrSQL = Array()
    With vsThirdSvr
        mrsSvr.MoveFirst
        For i = 1 To mrsSvr.RecordCount
            If mrsSvr!服务地址 & "" <> .TextMatrix(i, COL_服务地址) Then
                '地址发生了变化的保存
                strSQL = "Zl_三方服务配置目录_Update('" & .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
            mrsSvr.MoveNext
        Next
    End With
    
    gcnOracle.BeginTrans        '开始保存参数
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存三方服务配置目录")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    Exit Sub
errH:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSvrChk_Click()
'功能：验证 三方服务
    Call ThirdSvrTest(vsThirdSvr.Row)
End Sub
