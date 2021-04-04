VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "综合查询"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstSelect 
      Height          =   1680
      Left            =   6360
      TabIndex        =   96
      Top             =   8550
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Frame fra统计 
      Height          =   900
      Left            =   60
      TabIndex        =   86
      Top             =   7785
      Visible         =   0   'False
      Width           =   5430
      Begin VB.CommandButton cmdCalc 
         Caption         =   "计算(&J)"
         Height          =   350
         Left            =   4200
         TabIndex        =   95
         Top             =   135
         Width           =   1100
      End
      Begin VB.TextBox txtCount 
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   165
         Width           =   800
      End
      Begin VB.CheckBox chk双击 
         Caption         =   "双击剔除数据"
         Height          =   255
         Left            =   3300
         TabIndex        =   93
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtSD 
         Height          =   300
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   165
         Width           =   800
      End
      Begin VB.TextBox txtDelSD 
         Height          =   300
         Left            =   615
         TabIndex        =   90
         Top             =   540
         Width           =   390
      End
      Begin VB.CommandButton cmd剔除 
         Caption         =   "剔除(&T)"
         Height          =   350
         Left            =   2145
         TabIndex        =   88
         Top             =   510
         Width           =   1100
      End
      Begin VB.TextBox txtAVG 
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   165
         Width           =   800
      End
      Begin VB.Label lbl均值SD 
         Caption         =   "统计数量          均值           SD"
         Height          =   240
         Left            =   120
         TabIndex        =   91
         Top             =   225
         Width           =   3660
      End
      Begin VB.Label lbl剔除 
         Caption         =   "剔除>      SD的数据"
         Height          =   240
         Left            =   105
         TabIndex        =   89
         Top             =   585
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "返回(&R)"
      Height          =   900
      Index           =   4
      Left            =   11625
      Picture         =   "frmMain.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7755
      Width           =   1500
   End
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   8685
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23715
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTMain 
      Height          =   7725
      Left            =   345
      TabIndex        =   4
      Top             =   45
      Width           =   16365
      _ExtentX        =   28866
      _ExtentY        =   13626
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "日常报表(&R)"
      TabPicture(0)   =   "frmMain.frx":29FE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pic(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "工作量统计(&G)"
      TabPicture(1)   =   "frmMain.frx":2A1A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pic(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "结果统计(&T)"
      TabPicture(2)   =   "frmMain.frx":2A36
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pic(2)"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   6540
         Index           =   0
         Left            =   225
         ScaleHeight     =   6540
         ScaleWidth      =   11955
         TabIndex        =   6
         Top             =   480
         Width           =   11955
         Begin VSFlex8Ctl.VSFlexGrid vfgData 
            Height          =   5700
            Index           =   0
            Left            =   30
            TabIndex        =   7
            Top             =   30
            Width           =   10905
            _cx             =   19235
            _cy             =   10054
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
            BackColorSel    =   16635590
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
            Cols            =   12
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
         Begin VB.Frame fraData 
            Caption         =   "搜索条件"
            Height          =   870
            Index           =   0
            Left            =   -30
            TabIndex        =   8
            Top             =   5880
            Width           =   11565
            Begin VB.ComboBox cbo类型 
               Height          =   300
               Index           =   0
               Left            =   4575
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   330
               Width           =   1500
            End
            Begin VB.ComboBox cbo小组 
               Height          =   300
               Index           =   0
               Left            =   6517
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   330
               Width           =   1800
            End
            Begin VB.ComboBox cbo仪器 
               Height          =   300
               Index           =   0
               Left            =   8820
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   330
               Width           =   2600
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Index           =   0
               Left            =   1035
               TabIndex        =   12
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Index           =   0
               Left            =   2640
               TabIndex        =   13
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin VB.Label lbl核收时间 
               Caption         =   "核收时间                －"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   17
               Top             =   390
               Width           =   2790
            End
            Begin VB.Label lbl类型 
               Caption         =   "类型"
               Height          =   210
               Index           =   0
               Left            =   4110
               TabIndex        =   16
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl小组 
               Caption         =   "小组"
               Height          =   210
               Index           =   0
               Left            =   6090
               TabIndex        =   15
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl仪器 
               Caption         =   "仪器"
               Height          =   210
               Index           =   0
               Left            =   8370
               TabIndex        =   14
               Top             =   390
               Width           =   435
            End
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   7290
         Index           =   1
         Left            =   -74910
         ScaleHeight     =   7290
         ScaleWidth      =   12630
         TabIndex        =   18
         Top             =   390
         Width           =   12630
         Begin VB.Frame fraData 
            Caption         =   "搜索条件"
            Height          =   1815
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   5430
            Width           =   12555
            Begin VB.Frame fra收费 
               Caption         =   "收费情况"
               Height          =   1395
               Left            =   11415
               TabIndex        =   70
               Top             =   195
               Width           =   1050
               Begin VB.OptionButton opt收费 
                  Caption         =   "未收费"
                  Height          =   180
                  Index           =   2
                  Left            =   90
                  TabIndex        =   73
                  Top             =   945
                  Width           =   855
               End
               Begin VB.OptionButton opt收费 
                  Caption         =   "已收费"
                  Height          =   180
                  Index           =   1
                  Left            =   90
                  TabIndex        =   72
                  Top             =   630
                  Width           =   855
               End
               Begin VB.OptionButton opt收费 
                  Caption         =   "所有"
                  Height          =   180
                  Index           =   0
                  Left            =   90
                  TabIndex        =   71
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   675
               End
            End
            Begin VB.ComboBox cbo仪器 
               Height          =   300
               Index           =   1
               Left            =   8805
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   300
               Width           =   2595
            End
            Begin VB.ComboBox cbo小组 
               Height          =   300
               Index           =   1
               Left            =   6495
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   300
               Width           =   1800
            End
            Begin VB.ComboBox cbo类型 
               Height          =   300
               Index           =   1
               Left            =   4560
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   300
               Width           =   1500
            End
            Begin VB.ComboBox cbo申请科室 
               Height          =   300
               Index           =   1
               Left            =   1020
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   675
               Width           =   1830
            End
            Begin VB.ComboBox cbo申请人 
               Height          =   300
               Index           =   1
               Left            =   3495
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   675
               Width           =   1425
            End
            Begin VB.ComboBox cbo检验人 
               Height          =   300
               Index           =   1
               Left            =   5565
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   675
               Width           =   1425
            End
            Begin VB.ComboBox cbo审核人 
               Height          =   300
               Index           =   1
               Left            =   7635
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   675
               Width           =   1425
            End
            Begin VB.Frame fra病人来源 
               Caption         =   "病人来源"
               Height          =   1005
               Left            =   9105
               TabIndex        =   31
               Top             =   585
               Width           =   2295
               Begin VB.OptionButton opt来源 
                  Caption         =   "所有"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   36
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   660
               End
               Begin VB.OptionButton opt来源 
                  Caption         =   "门诊"
                  Height          =   180
                  Index           =   1
                  Left            =   810
                  TabIndex        =   35
                  Top             =   390
                  Width           =   660
               End
               Begin VB.OptionButton opt来源 
                  Caption         =   "住院"
                  Height          =   180
                  Index           =   2
                  Left            =   1500
                  TabIndex        =   34
                  Top             =   390
                  Width           =   660
               End
               Begin VB.OptionButton opt来源 
                  Caption         =   "院外"
                  Height          =   180
                  Index           =   3
                  Left            =   120
                  TabIndex        =   33
                  Top             =   690
                  Width           =   660
               End
               Begin VB.OptionButton opt来源 
                  Caption         =   "体检"
                  Height          =   180
                  Index           =   4
                  Left            =   825
                  TabIndex        =   32
                  Top             =   705
                  Width           =   660
               End
            End
            Begin VB.Frame frm统计方式 
               Caption         =   "统计方式"
               Height          =   645
               Left            =   105
               TabIndex        =   22
               Top             =   945
               Width           =   8955
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "小组"
                  Height          =   180
                  Index           =   0
                  Left            =   180
                  TabIndex        =   30
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "仪器"
                  Height          =   180
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   29
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "申请项目"
                  Height          =   180
                  Index           =   2
                  Left            =   2325
                  TabIndex        =   28
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "申请科室"
                  Height          =   180
                  Index           =   3
                  Left            =   3405
                  TabIndex        =   27
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "申请人"
                  Height          =   180
                  Index           =   4
                  Left            =   4485
                  TabIndex        =   26
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "检验人"
                  Height          =   180
                  Index           =   5
                  Left            =   5565
                  TabIndex        =   25
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "审核人"
                  Height          =   180
                  Index           =   6
                  Left            =   6630
                  TabIndex        =   24
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton opt统计方式 
                  Caption         =   "病人来源"
                  Height          =   180
                  Index           =   7
                  Left            =   7710
                  TabIndex        =   23
                  Top             =   315
                  Width           =   1080
               End
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Index           =   1
               Left            =   1020
               TabIndex        =   44
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Index           =   1
               Left            =   2625
               TabIndex        =   45
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin VB.Label lbl仪器 
               Caption         =   "仪器"
               Height          =   210
               Index           =   1
               Left            =   8355
               TabIndex        =   53
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lbl小组 
               Caption         =   "小组"
               Height          =   210
               Index           =   1
               Left            =   6075
               TabIndex        =   52
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lbl类型 
               Caption         =   "类型"
               Height          =   210
               Index           =   1
               Left            =   4050
               TabIndex        =   51
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lbl核收时间 
               Caption         =   "核收时间                －"
               Height          =   255
               Index           =   1
               Left            =   225
               TabIndex        =   50
               Top             =   360
               Width           =   2790
            End
            Begin VB.Label lbl申请科室 
               Caption         =   "申请科室"
               Height          =   225
               Left            =   225
               TabIndex        =   49
               Top             =   690
               Width           =   780
            End
            Begin VB.Label Label1 
               Caption         =   "申请人"
               Height          =   225
               Left            =   2925
               TabIndex        =   48
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label2 
               Caption         =   "检验人"
               Height          =   225
               Left            =   4980
               TabIndex        =   47
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label3 
               Caption         =   "审核人"
               Height          =   225
               Left            =   7050
               TabIndex        =   46
               Top             =   720
               Width           =   780
            End
         End
         Begin VB.Frame fraLR 
            Height          =   1875
            Index           =   1
            Left            =   4605
            MousePointer    =   9  'Size W E
            TabIndex        =   19
            Top             =   15
            Width           =   45
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgData 
            Height          =   5250
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   90
            Width           =   4530
            _cx             =   7990
            _cy             =   9260
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
            BackColorSel    =   16635590
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
            Cols            =   12
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
         Begin VSFlex8Ctl.VSFlexGrid vfgItem 
            Height          =   5250
            Index           =   1
            Left            =   4755
            TabIndex        =   54
            Top             =   75
            Width           =   7470
            _cx             =   13176
            _cy             =   9260
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
            BackColorSel    =   16635590
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
            Cols            =   12
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
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   7260
         Index           =   2
         Left            =   -74490
         ScaleHeight     =   7260
         ScaleWidth      =   17070
         TabIndex        =   55
         Top             =   450
         Width           =   17070
         Begin VB.Frame fraData 
            Caption         =   "搜索条件"
            Height          =   1350
            Index           =   2
            Left            =   840
            TabIndex        =   57
            Top             =   5865
            Width           =   16065
            Begin VB.TextBox txt上限 
               Height          =   300
               Left            =   11670
               TabIndex        =   85
               Top             =   765
               Width           =   1200
            End
            Begin VB.ComboBox cbo符号 
               Height          =   300
               Left            =   9150
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   765
               Width           =   1185
            End
            Begin VB.TextBox txt下限 
               Height          =   300
               Left            =   10395
               TabIndex        =   83
               Top             =   765
               Width           =   1200
            End
            Begin VB.CommandButton cmd项目 
               Caption         =   "…"
               Height          =   300
               Left            =   12620
               TabIndex        =   81
               Top             =   330
               Width           =   250
            End
            Begin VB.TextBox txt项目 
               Height          =   300
               Left            =   9135
               TabIndex        =   79
               Top             =   330
               Width           =   3465
            End
            Begin VB.ComboBox cbo性别 
               Height          =   300
               Left            =   7530
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   765
               Width           =   765
            End
            Begin VB.ComboBox cbo年龄 
               Height          =   300
               Left            =   6045
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   765
               Width           =   765
            End
            Begin VB.TextBox txt年龄 
               Height          =   300
               Left            =   4665
               TabIndex        =   74
               ToolTipText     =   "支持输入20-30的方式指定范围"
               Top             =   765
               Width           =   1320
            End
            Begin VB.ComboBox cbo仪器 
               Height          =   300
               Index           =   2
               Left            =   675
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   765
               Width           =   2985
            End
            Begin VB.ComboBox cbo小组 
               Height          =   300
               Index           =   2
               Left            =   6517
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   330
               Width           =   1800
            End
            Begin VB.ComboBox cbo类型 
               Height          =   300
               Index           =   2
               Left            =   4575
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   330
               Width           =   1500
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Index           =   2
               Left            =   1035
               TabIndex        =   61
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Index           =   2
               Left            =   2640
               TabIndex        =   62
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin VB.Label lbl结果范围 
               Caption         =   "结果范围"
               Height          =   225
               Left            =   8385
               TabIndex        =   82
               Top             =   810
               Width           =   900
            End
            Begin VB.Label lbl检验项目 
               Caption         =   "检验项目"
               Height          =   225
               Left            =   8370
               TabIndex        =   80
               Top             =   390
               Width           =   900
            End
            Begin VB.Label lbl性别 
               Caption         =   "性别"
               Height          =   225
               Left            =   6990
               TabIndex        =   77
               Top             =   810
               Width           =   450
            End
            Begin VB.Label lbl年龄 
               Caption         =   "年龄范围"
               Height          =   225
               Left            =   3810
               TabIndex        =   75
               Top             =   810
               Width           =   795
            End
            Begin VB.Label lbl仪器 
               Caption         =   "仪器"
               Height          =   210
               Index           =   2
               Left            =   225
               TabIndex        =   66
               Top             =   810
               Width           =   435
            End
            Begin VB.Label lbl小组 
               Caption         =   "小组"
               Height          =   210
               Index           =   2
               Left            =   6090
               TabIndex        =   65
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl类型 
               Caption         =   "类型"
               Height          =   210
               Index           =   2
               Left            =   4110
               TabIndex        =   64
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl核收时间 
               Caption         =   "核收时间                －"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   63
               Top             =   390
               Width           =   2790
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgItem 
            Height          =   5250
            Index           =   2
            Left            =   5430
            TabIndex        =   68
            Top             =   195
            Width           =   6405
            _cx             =   11298
            _cy             =   9260
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
            BackColorSel    =   16635590
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
            Cols            =   12
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
         Begin VSFlex8Ctl.VSFlexGrid vfgData 
            Height          =   5250
            Index           =   2
            Left            =   0
            TabIndex        =   56
            Top             =   540
            Width           =   5070
            _cx             =   8943
            _cy             =   9260
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
            BackColorSel    =   16635590
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
            Cols            =   12
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
         Begin VB.Frame fraLR 
            Height          =   1875
            Index           =   2
            Left            =   4800
            MousePointer    =   9  'Size W E
            TabIndex        =   67
            Top             =   210
            Width           =   45
         End
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "输出到Excel(&E)"
      Height          =   900
      Index           =   3
      Left            =   10125
      Picture         =   "frmMain.frx":2A52
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7755
      Width           =   1500
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "打印设置(&S)"
      Height          =   900
      Index           =   2
      Left            =   8655
      Picture         =   "frmMain.frx":5444
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "页面设置"
      Top             =   7755
      Width           =   1500
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "打印(&P)"
      Height          =   900
      Index           =   1
      Left            =   7050
      Picture         =   "frmMain.frx":7E36
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7755
      Width           =   1500
   End
   Begin VB.CommandButton cmdRun 
      Appearance      =   0  'Flat
      Caption         =   "搜索(&F)"
      Height          =   900
      Index           =   0
      Left            =   5475
      Picture         =   "frmMain.frx":A828
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "搜索"
      Top             =   7755
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsHost As zl9LisQuery_Def.clsLisQueryHost
Private mlgIndex As Long
Private mintLastTab As Integer
Private Enum mCol_日常
    类型 = 1:  小组: 仪器: 无主: 已接收: 已核收: 审核: 未审
End Enum

Private mrs项目 As ADODB.Recordset

Private Sub cbo符号_Click()
    If cbo符号.List(cbo符号.ListIndex) = "在...之间" Then
        txt上限.Enabled = True
    Else
        txt上限.Enabled = False
    End If
End Sub

Private Sub cmdCalc_Click()
    Call CalcData
End Sub

Private Sub cmd剔除_Click()
    Dim lngRow As Long
    Dim strDeleteRow As String, varDelRow As Variant
    If txtDelSD.Text <> "" Then
        If Val(txtDelSD.Text) >= 1 And Val(txtDelSD.Text) <= 4 Then
            
            With vfgData(2)
                strDeleteRow = ""
                'lngCurrRow = .Row
                For lngRow = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(lngRow, .ColIndex("SD"))) = ">" & Val(txtDelSD.Text) & "S" Then
                        strDeleteRow = lngRow & "," & strDeleteRow
                    End If
                Next
                
                If strDeleteRow <> "" Then
                    varDelRow = Split(strDeleteRow, ",")
                    For lngRow = LBound(varDelRow) To UBound(varDelRow) - 1
                       .RemoveItem Val(varDelRow(lngRow))
                    Next
                End If
                
                'If lngCurrRow >= .FixedRows And lngCurrRow < .Rows Then
                Call vfgData_RowColChange(2)
                'End If
                
            End With
            Me.cmdCalc.Enabled = True
        Else
            MsgBox "输入的范围是1-4，请检查！", vbInformation, Me.Caption
        End If
    End If
End Sub

Private Sub cmd项目_Click()
    Call ShowSelect("")
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    '--- 当月开始日期，当月结束日期
    For i = Me.dtpBegin.LBound To Me.dtpBegin.UBound
        Me.dtpBegin(i).Value = Format(Now, "yyyy-MM-01")
    Next
    For i = Me.dtpEnd.LBound To Me.dtpEnd.UBound
        Me.dtpEnd(i).Value = Format(Now, "yyyy-MM-dd")
    Next
    
    '--- 初始化类型
    strSQL = "Select 名称 From 诊疗检验类型 Order By 编码"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo类型.LBound To cbo类型.UBound
        cbo类型(i).Clear
        cbo类型(i).AddItem ""
        
        Do Until rsTmp.EOF
            cbo类型(i).AddItem "" & rsTmp.Fields("名称")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo类型(i).ListIndex = 0
    Next
    
    '--- 初始化小组
    strSQL = "Select a.Id, a.编码, a.名称" & vbNewLine & _
        "From 检验小组 a, 检验小组成员 b, 上机人员表 c" & vbNewLine & _
        "Where a.Id = b.小组id And b.人员id = c.人员id And 用户名 = User" & vbNewLine & _
        "Order By a.编码"

    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo小组.LBound To cbo小组.UBound
        cbo小组(i).Clear
        cbo小组(i).AddItem ""
        Do Until rsTmp.EOF
            cbo小组(i).AddItem "" & rsTmp.Fields("编码") & "-" & rsTmp.Fields("名称")
            cbo小组(i).ItemData(cbo小组(i).NewIndex) = Val("" & rsTmp.Fields("ID"))
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo小组(i).ListIndex = 0
    Next
    
    '--- 初始化仪器
    strSQL = "Select e.Id, e.编码, e.名称" & vbNewLine & _
        "From 检验小组 a, 检验小组成员 b, 上机人员表 c, 检验小组仪器 d, 检验仪器 e" & vbNewLine & _
        "Where Nvl(e.微生物,0) = 0 and a.Id = b.小组id And b.人员id = c.人员id And 用户名 = User And a.Id = d.小组id And d.仪器id = e.Id" & vbNewLine & _
        "Order By e.编码"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo仪器.LBound To cbo仪器.UBound
        cbo仪器(i).Clear
        cbo仪器(i).AddItem ""
       
        Do Until rsTmp.EOF
            cbo仪器(i).AddItem "" & rsTmp.Fields("编码") & "-" & rsTmp.Fields("名称")
            cbo仪器(i).ItemData(cbo仪器(i).NewIndex) = Val("" & rsTmp.Fields("ID"))
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo仪器(i).ListIndex = 0
    Next
    '--- 初始化申请科室
    strSQL = "Select a.Id, a.编码, a.名称" & vbNewLine & _
            "From 部门性质说明 b, 部门表 a" & vbNewLine & _
            "Where a.Id = b.部门id And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & vbNewLine & _
            "           Instr(',临床,体检,', ',' || b.工作性质 || ',') > 0" & vbNewLine & _
            " Order by a.编码"

    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo申请科室.LBound To cbo申请科室.UBound
        cbo申请科室(i).Clear
        cbo申请科室(i).AddItem ""
        Do Until rsTmp.EOF
            cbo申请科室(i).AddItem "" & rsTmp.Fields("编码") & "-" & rsTmp.Fields("名称")
            cbo申请科室(i).ItemData(cbo申请科室(i).NewIndex) = Val("" & rsTmp.Fields("ID"))
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo申请科室(i).ListIndex = 0
    Next
    
    '--- 初始化申请人
    strSQL = "Select a.姓名" & vbNewLine & _
        "From 人员性质说明 b, 人员表 a" & vbNewLine & _
        "Where b.人员性质 = '医生' And a.Id = b.人员id And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd'))" & vbNewLine & _
        "Order By a.姓名"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo申请人.LBound To cbo申请人.UBound
        cbo申请人(i).Clear
        cbo申请人(i).AddItem ""
        Do Until rsTmp.EOF
            cbo申请人(i).AddItem "" & rsTmp.Fields("姓名")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo申请人(i).ListIndex = 0
    Next
    '--- 初始化审核人, 检验人
    strSQL = "Select  Distinct B.姓名" & vbNewLine & _
        "From 人员表 B,检验小组成员 A" & vbNewLine & _
        "Where A.人员id=B.ID" & vbNewLine & _
        "Order By B.姓名"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo审核人.LBound To cbo审核人.UBound
        cbo审核人(i).Clear
        cbo审核人(i).AddItem ""
        Do Until rsTmp.EOF
            cbo审核人(i).AddItem "" & rsTmp.Fields("姓名")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo审核人(i).ListIndex = 0
    Next
    
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo检验人.LBound To cbo检验人.UBound
        cbo检验人(i).Clear
        cbo检验人(i).AddItem ""
        Do Until rsTmp.EOF
            cbo检验人(i).AddItem "" & rsTmp.Fields("姓名")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo检验人(i).ListIndex = 0
    Next
    '年龄
    cbo年龄.Clear
    cbo年龄.AddItem "岁"
    cbo年龄.AddItem "月"
    cbo年龄.AddItem "天"
    cbo年龄.AddItem "小时"
    cbo年龄.AddItem "成人"
    cbo年龄.AddItem "婴儿"
    cbo年龄.ListIndex = 0
    
    '性别
    cbo性别.Clear
    cbo性别.AddItem ""
    cbo性别.AddItem "1-男"
    cbo性别.AddItem "2-女"
    cbo性别.AddItem "3-未知"
    cbo性别.AddItem "9-不明"
    cbo性别.ListIndex = 0
    
    '
    cbo符号.Clear
    cbo符号.AddItem "="
    cbo符号.AddItem "<>"
    cbo符号.AddItem ">"
    cbo符号.AddItem "<"
    cbo符号.AddItem ">="
    cbo符号.AddItem "<="
    cbo符号.AddItem "包含"
    cbo符号.AddItem "在...之间"
    cbo符号.ListIndex = 0
    '--- 初始化
    
    
    Call initvfgDataTitle(0)
    Call initvfgDataTitle(1): Call initvfgItemTitle(1)
    Call initvfgDataTitle(2): Call initvfgItemTitle(2)
    mintLastTab = 0
    ssTMain.Tab = 0
    Me.Show
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub cmdRun_Click(Index As Integer)
    Dim strFileName As String
    Select Case Index
        Case 0 '搜索
            Me.cmdRun(0).Enabled = False
            Call DoQuery
            Me.cmdRun(0).Enabled = True
        Case 1 '打印
            Me.cmdRun(1).Enabled = False
            If Not vsPrint Is Nothing Then Unload vsPrint
            Call vsPrint.vsPrint(Me.vfgData(Me.ssTMain.Tab).hWnd, Me.ssTMain.Tab)
            Me.cmdRun(1).Enabled = True
        Case 2 '打印设置
            Me.cmdRun(2).Enabled = False
            Call frmPrintSet.PageSetup(Me.ssTMain.Tab)
            Call DoQuery
            Me.cmdRun(2).Enabled = True
        Case 3 'Excel
            Me.cmdRun(3).Enabled = False
                strFileName = App.Path & "\Report" & Me.ssTMain.Tab & "_" & Format(Now, "yyyyMMddHHmmss") & ".xls"
                vfgData(Me.ssTMain.Tab).SaveGrid strFileName, flexFileExcel, flexXLSaveFixedCells
                MsgBox "已保存到" & strFileName, vbInformation, Me.Caption
            Me.cmdRun(3).Enabled = True
        Case 4 '返回
            Unload Me
    End Select
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim iTab As Integer
    On Error Resume Next
    With Me.ssTMain
        .Left = 45
        .Top = 45
        .Width = Me.ScaleWidth - 90
        .Height = Me.ScaleHeight - 90 - Me.cmdRun(0).Height - 45 - Me.stbBar.Height
    End With
    For i = pic.LBound To pic.UBound
        With Me.pic(i)
            .Left = Me.ssTMain.Left + 45
            .Top = Me.ssTMain.Top + 350
            .Width = Me.ssTMain.Width - 150
            .Height = Me.ssTMain.Height - 450
        End With
    Next
    
    With Me.cmdRun(0)
        .Left = Me.ScaleWidth - Me.cmdRun(0).Width * Me.cmdRun.Count - 90 * Me.cmdRun.Count
        .Top = Me.ssTMain.Top + Me.ssTMain.Height + 45
    End With
    
    For i = Me.cmdRun.LBound + 1 To Me.cmdRun.UBound
        Me.cmdRun(i).Left = Me.cmdRun(i - 1).Left + Me.cmdRun(i - 1).Width + 90
        Me.cmdRun(i).Top = Me.cmdRun(0).Top
    Next
    
    With fra统计
        .Left = Me.ssTMain.Left
        .Top = Me.cmdRun(0).Top
    End With
    
    If Me.lstSelect.Visible = True Then
        Call MoveSelect(txt项目)
    End If
    Me.Refresh
End Sub

Private Sub lstSelect_DblClick()

    If lstSelect.ListIndex >= 0 Then
        txt项目.Text = lstSelect.List(lstSelect.ListIndex)
        txt项目.Tag = lstSelect.ItemData(lstSelect.ListIndex)
    End If
    txt项目.SetFocus
    lstSelect.Visible = False
End Sub

Private Sub lstSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lstSelect.ListIndex >= 0 Then
            txt项目.Text = lstSelect.List(lstSelect.ListIndex)
            txt项目.Tag = lstSelect.ItemData(lstSelect.ListIndex)
        End If
        txt项目.SetFocus
        lstSelect.Visible = False
    ElseIf KeyAscii = vbKeyEscape Then
        txt项目.SetFocus
        lstSelect.Visible = False
    End If
End Sub

Private Sub lstSelect_LostFocus()
    Me.lstSelect.Visible = False
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error Resume Next
    '---- 日常

    With Me.vfgData(Index)
        .Left = Me.pic(Index).ScaleLeft
        .Top = Me.pic(Index).ScaleTop
        .Width = Me.pic(Index).ScaleWidth
        .Height = Me.pic(Index).Height - Me.fraData(Index).Height
    End With
    With Me.fraData(Index)
        .Left = Me.vfgData(Index).Left
        .Top = Me.vfgData(Index).Top + Me.vfgData(Index).Height
        .Width = Me.vfgData(Index).Width
    End With
    '--- 工作量
    If Index >= Me.vfgItem.LBound And Index <= Me.vfgItem.UBound Then
        With Me.vfgItem(Index)
             Me.vfgData(Index).Width = Me.vfgData(Index).Width - .Width - 45
            
            .Left = Me.vfgData(Index).Left + Me.vfgData(Index).Width + 45
            .Top = Me.vfgData(Index).Top
            .Height = Me.vfgData(Index).Height
            
            Me.fraLR(Index).Left = .Left - 45
            Me.fraLR(Index).Top = .Top
            Me.fraLR(Index).Height = .Height
        End With
    End If
End Sub

Private Sub ssTMain_Click(PreviousTab As Integer)
    Dim iTab As Integer
    iTab = PreviousTab
    If iTab >= pic.LBound And iTab <= pic.UBound Then
        pic(iTab).Visible = False
    End If
    
    iTab = ssTMain.Tab
    mintLastTab = iTab
    If iTab >= pic.LBound And iTab <= pic.UBound Then
        pic(iTab).Visible = True
    End If
    ssTMain.Tab = iTab
    fra统计.Visible = False
    If iTab = 2 Then fra统计.Visible = True
End Sub

Private Sub fraLR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     On Error Resume Next
    If Button = vbLeftButton Then
        If Index >= Me.vfgItem.LBound And Index <= Me.vfgItem.UBound Then
            Me.vfgData(Index).Width = Me.vfgData(Index).Width + X
            Me.vfgItem(Index).Width = Me.vfgItem(Index).Width - X
            Me.fraLR(Index).Left = Me.fraLR(Index).Left + X
            Me.vfgItem(Index).Left = Me.fraLR(Index).Left + Me.fraLR(Index).Width
        End If
    End If
End Sub

Private Sub txt项目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txt项目.Text) <> "" Then
            Call ShowSelect(Trim(txt项目.Text))
        End If
    End If
End Sub

Private Sub vfgData_DblClick(Index As Integer)
    Dim lngRow As Long
    
    If Index = 2 Then
        If chk双击.Value = 1 Then
            lngRow = vfgData(Index).Row
            If lngRow >= vfgData(Index).FixedRows And lngRow < vfgData(Index).Rows Then
                vfgData(Index).RemoveItem (vfgData(Index).Row)
                Call vfgData_RowColChange(2)
                Me.cmdCalc.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub vfgData_RowColChange(Index As Integer)
    Dim strSQL As String
    Dim strValue As String, strKey As String
    On Error GoTo errH
    Select Case Index
    Case 1
        '工作量统计
        With vfgData(Index)
             Call initvfgItemTitle(Index)
             If .ColIndex("ID") >= .FixedCols And .ColIndex("ID") <= .Cols - 1 Then
                strValue = Trim("" & .TextMatrix(.Row, .ColIndex("ID")))
                strKey = Trim("" & .TextMatrix(.FixedRows - 1, .ColIndex("小组")))
                If strValue <> "" And strKey <> "" Then
                    Call RefGrid_工作量Item(Index, strKey, strValue)
                End If
             End If
        End With
    Case 2  '结果统计
        With vfgData(Index)
            Call initvfgItemTitle(Index)
            If .ColIndex("标本ID") >= .FixedCols And .ColIndex("标本ID") <= .Cols - 1 Then
                strValue = Trim("" & .TextMatrix(.Row, .ColIndex("标本ID")))
                If strValue <> "" Then
                    Call RefGrid_结果Item(Index, strValue)
                End If
            End If
        End With
    End Select
    
    Exit Sub
errH:
    MsgBox Err.Description
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------
'外部调用过程

Public Function ShowMe(ByVal Index As Long, ShowMode As QueryShowMode, objHost As zl9LisQuery_Def.clsLisQueryHost) As Boolean
    mlgIndex = Index
    Set clsHost = objHost
    Me.Show ShowMode, objHost
End Function

'-----------------------------------------------------------------------------------------------------------------------------------
' 内部过程
Private Sub DoQuery()

    Dim curStart As Currency, curEnd As Currency
    On Error GoTo errH
    
    Me.MousePointer = vbHourglass
    Select Case Me.ssTMain.Tab
        Case 0
            Call RefGrid_日常(0)
        Case 1
            Call RefGrid_工作量(1)
        Case 2
            Call RefGrid_结果(2)
    End Select
    Me.MousePointer = vbDefault
    Exit Sub
errH:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbQuestion, Me.Caption
End Sub

Private Sub RefGrid_结果(ByVal lngIndex As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String, strEnd As String, lng仪器 As Long, str类型 As String, lng小组 As Long
    Dim strWhere As String, iCol As Integer, strTitle As String
    Dim str年龄单位 As String, str年龄 As String, lng年龄下限 As Long, lng年龄上限 As Long, str年龄符号 As String, strRecord年龄 As String
    Dim str性别 As String, str结果符号 As String, str结果下限 As String, str结果上限 As String
    Dim str检验结果 As String, curSD As Currency, curAVG As Currency
    
    Dim lng项目ID As Long ', str结果类型 As String, str取值序列 As String
    Dim blnAdd As Boolean, lng偏低 As Long, lng偏高 As Long, lng警示 As Long
    Dim lngColor  As Long, lngForeColor As Long, str标志 As String
    On Error GoTo errH
    
    lng偏低 = &H80FFFF: lng偏高 = &H80C0FF: lng警示 = &H40C0&
    Call initvfgDataTitle(lngIndex)
    
    lng项目ID = Val(txt项目.Tag)
    If lng项目ID = 0 Then
        MsgBox "请输入项目后再执行此功能!", vbInformation, Me.Caption
        Exit Sub
    End If
'    mrs项目.Filter = ""
'    mrs项目.Filter = "项目ID=" & lng项目ID
'    str结果类型 = "": str项目序列 = ""
'    Do Until mrs项目.EOF
'        str结果类型 = Trim("" & mrs项目!结果类型)
'        str取值序列 = Trim("" & mrs项目!取值序列)
'        mrs项目.MoveNext
'    Loop
    
    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    
    strTitle = "日期：" & strBegin & " 至 " & strEnd
    lng仪器 = Val(cbo仪器(lngIndex).ItemData(cbo仪器(lngIndex).ListIndex))
    lng小组 = Val(cbo小组(lngIndex).ItemData(cbo小组(lngIndex).ListIndex))
    str类型 = Trim(cbo类型(lngIndex).List(cbo类型(lngIndex).ListIndex))
    
    If str类型 <> "" Then strWhere = strWhere & " And D.仪器类型 =[6]"
    If lng小组 <> 0 Then strWhere = strWhere & " And C.ID=[5]"
    If lng仪器 <> 0 Then strWhere = strWhere & " And D.ID=[4]"
    
    str年龄单位 = cbo年龄.List(cbo年龄.ListIndex)
    
    str年龄 = Trim(txt年龄.Text)
    If str年龄 = "" Then str年龄单位 = ""
    
    str年龄符号 = "="
    If str年龄 Like "*-*" Then
        lng年龄下限 = Val(Split(str年龄, "-")(0))
        lng年龄上限 = Val(Split(str年龄, "-")(1))
         
        If lng年龄下限 >= lng年龄上限 Then
            MsgBox "年龄下限不能大于或等于年龄上限！", vbInformation, Me.Caption
            Exit Sub
        End If
        str年龄符号 = "Between"
    ElseIf str年龄 Like ">=*" Then
        lng年龄下限 = Val(Mid(str年龄, 3))
        str年龄符号 = ">="
    ElseIf str年龄 Like "<=*" Then
        lng年龄下限 = Val(Mid(str年龄, 3))
        str年龄符号 = "<="
    ElseIf str年龄 Like ">*" Then
        lng年龄下限 = Val(Mid(str年龄, 2))
        str年龄符号 = ">"
    ElseIf str年龄 Like "<*" Then
        lng年龄下限 = Val(Mid(str年龄, 2))
        str年龄符号 = "<"
    ElseIf str年龄 Like "<>*" Then
        lng年龄下限 = Val(Mid(str年龄, 3))
        str年龄符号 = "<>"
    Else
        If Not IsNumeric(str年龄) Then str年龄符号 = "NO"
    End If
    
    If str年龄单位 <> "" Then
        If InStr("成人,婴儿", str年龄单位) <= 0 Then
            If lng年龄下限 < 0 Or lng年龄上限 < 0 Then
                MsgBox "年龄下限或年龄上限不能小于0！", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    
    str性别 = Trim(cbo性别.List(cbo性别.ListIndex))
    If str性别 <> "" Then
        str性别 = Split(str性别, "-")(1)
        strWhere = strWhere & " And A.性别 = [7] "
    End If
    
    str结果符号 = cbo符号.List(cbo符号.ListIndex)
    str结果下限 = Trim(txt下限.Text)
    str结果上限 = Trim(txt上限.Text)
    
    strSQL = "Select /*+Rule */ g.检验标本id, a.核收时间, c.名称 As 小组, a.标本序号 As 样本号, h.中文名 As 项目, g.检验结果, a.姓名, a.性别, a.年龄, nvl(a.年龄单位,'岁') as 年龄单位," & vbNewLine & _
            " Decode(g.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') as 结果标志 " & vbNewLine & _
            "From 诊治所见项目 h, 检验普通结果 g, 上机人员表 f, 检验小组成员 e, 检验仪器 d, 检验小组 c, 检验小组仪器 b, 检验标本记录 a" & vbNewLine & _
            "Where a.核收时间 Between [1] And [2] And a.仪器id = b.仪器id And" & vbNewLine & _
            "      a.报告结果=g.记录类型 And b.小组id = c.Id And a.仪器id = d.Id And Nvl(a.微生物标本, 0) = 0 And c.Id = e.小组id And e.人员id = f.人员id And f.用户名 = User And" & vbNewLine & _
            "           a.审核人 Is Not Null And a.Id = g.检验标本id And g.检验项目id = h.Id And g.检验项目id+0 = [3] " & strWhere
    
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng项目ID, lng仪器, lng小组, str类型, str性别)
    
    Do Until rsTmp.EOF
        blnAdd = True
        If Not (InStr("成人,婴儿", str年龄单位) > 0) Then
            If str年龄单位 = Trim("" & rsTmp!年龄单位) And str年龄单位 <> "" Then
                Select Case str年龄符号
                    Case "="
                        If Not (Val(Trim("" & rsTmp!年龄)) = lng年龄下限) Then blnAdd = False
                    Case "Between"
                        If Not (Val(Trim("" & rsTmp!年龄)) >= lng年龄下限 And Val(Trim("" & rsTmp!年龄)) <= lng年龄上限) Then blnAdd = False
                    Case ">"
                        If Not (Val(Trim("" & rsTmp!年龄)) > lng年龄下限) Then blnAdd = False
                    Case ">="
                        If Not (Val(Trim("" & rsTmp!年龄)) >= lng年龄下限) Then blnAdd = False
                    Case "<"
                        If Not (Val(Trim("" & rsTmp!年龄)) < lng年龄下限) Then blnAdd = False
                    Case "<="
                        If Not (Val(Trim("" & rsTmp!年龄)) <= lng年龄下限) Then blnAdd = False
                    Case "<>"
                        If Not (Val(Trim("" & rsTmp!年龄)) <> lng年龄下限) Then blnAdd = False
                    Case "NO"
                        blnAdd = False
                End Select
            Else
                If str年龄单位 <> "" Then blnAdd = False
            End If
        End If
        
        If blnAdd Then
            str检验结果 = Trim("" & rsTmp!检验结果)
            
            Select Case str结果符号
                Case "="
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str结果下限) = Val(str检验结果)) Then blnAdd = False
                    Else
                        If Not (str结果下限 = str检验结果) Then blnAdd = False
                    End If
                Case "<>"
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str结果下限) <> Val(str检验结果)) Then blnAdd = False
                    Else
                        If Not (str结果下限 <> CStr(str检验结果)) Then blnAdd = False
                    End If
                Case ">"
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str检验结果) > Val(str结果下限)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case "<"
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str检验结果) < Val(str结果下限)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case ">="
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str检验结果) >= Val(str结果下限)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case "<="
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str检验结果) <= Val(str结果下限)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case "包含"
                    If Not (CStr(str检验结果) Like "*" & str结果下限 & "*") Then blnAdd = False
                Case "在...之间"
                    If IsNumeric(str检验结果) Then
                        If Not (Val(str检验结果) >= Val(str结果下限) And Val(str检验结果) <= Val(str结果上限)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
            End Select
            
        End If
        If blnAdd Then
            With vfgData(lngIndex)
                .TextMatrix(.Rows - 1, .ColIndex("标本ID")) = Val("" & rsTmp!检验标本ID)
                .TextMatrix(.Rows - 1, .ColIndex("核收日期")) = Format("" & rsTmp!核收时间, "yy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, .ColIndex("检验小组")) = Trim("" & rsTmp!小组)  'IIf(Val("" & rsTmp!数量) = 0, "", Val("" & rsTmp!数量))
                .TextMatrix(.Rows - 1, .ColIndex("样本号")) = Trim("" & rsTmp!样本号)
                .TextMatrix(.Rows - 1, .ColIndex("项目")) = Trim("" & rsTmp!项目)
                
                .TextMatrix(.Rows - 1, .ColIndex("项目结果")) = Trim("" & rsTmp!检验结果)
                If IsNumeric(.TextMatrix(.Rows - 1, .ColIndex("项目结果"))) Then
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("项目结果")) = flexAlignRightCenter
                Else
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("项目结果")) = flexAlignLeftCenter
                End If
                
                str标志 = Trim("" & rsTmp!结果标志)
                lngColor = .BackColor
                lngForeColor = .ForeColor
                If InStr("↓", str标志) > 0 And str标志 <> "" Then     '2
                    lngColor = lng偏低
                ElseIf InStr("↑,异常", str标志) > 0 And str标志 <> "" Then '3,异常
                    lngColor = lng偏高
                ElseIf InStr("↑↑,↓↓", str标志) > 0 And str标志 <> "" Then  '5,6
                    lngColor = lng警示
                End If
                .Cell(flexcpBackColor, .Rows - 1, .ColIndex("项目结果")) = lngColor
                .Cell(flexcpForeColor, .Rows - 1, .ColIndex("项目结果")) = lngForeColor
                
                
                .TextMatrix(.Rows - 1, .ColIndex("姓名")) = Trim("" & rsTmp!姓名)
                .TextMatrix(.Rows - 1, .ColIndex("性别")) = Trim("" & rsTmp!性别)
                .TextMatrix(.Rows - 1, .ColIndex("年龄")) = Trim("" & rsTmp!年龄)
                .Rows = .Rows + 1
            End With
        End If
        
        rsTmp.MoveNext
    Loop
    
    With vfgData(lngIndex)
        '加表格线
        If .Rows > 2 Then .Rows = .Rows - 1
        Call CalcData
        .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
         
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("检验小组"), .Rows - 1, .ColIndex("检验小组")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("项目"), .Rows - 1, .ColIndex("项目")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("姓名"), .Rows - 1, .ColIndex("姓名")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("样本号"), .Rows - 1, .ColIndex("样本号")) = flexAlignRightCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("年龄"), .Rows - 1, .ColIndex("年龄")) = flexAlignRightCenter
    End With
    Exit Sub
errH:
    MsgBox "未找到数据！", vbQuestion, Me.Caption
    Err.Clear
End Sub
Private Sub CalcData()
    Dim curSD As Currency, curAVG As Currency, curCount As Currency
    Dim lngRow As Long, str结果 As String
        '求均值，方差
    With vfgData(2)
'        .Subtotal flexSTClear
'        .OutlineCol = 1   '指定输出列
'        .SubtotalPosition = flexSTBelow '合计在底部
'
'        'SD
'        .Subtotal flexSTStd, -1, .ColIndex("项目结果"), , , , , ""
'        curSD = Val(.TextMatrix(.Rows - 1, .ColIndex("项目结果")))
'
'        'AVG
'        .Subtotal flexSTClear
'        .Subtotal flexSTAverage, -1, .ColIndex("项目结果"), , , , , ""
'        curAVG = Val(.TextMatrix(.Rows - 1, .ColIndex("项目结果")))
'
'        'Count
'        .Subtotal flexSTClear
'        .Subtotal flexSTCount, -1, .ColIndex("项目结果"), , , , , ""
'        curCount = Val(.TextMatrix(.Rows - 1, .ColIndex("项目结果")))
'
'        .Rows = .Rows - 1
        str结果 = ""
        For lngRow = .FixedRows To .Rows - 1
            str结果 = str结果 & "," & Val(.TextMatrix(lngRow, .ColIndex("项目结果")))
        Next
        If str结果 <> "" Then
            curAVG = CalcSVG(str结果)
            curSD = CalcSD(str结果)
            curCount = UBound(Split(str结果, ","))
        End If
        txtCount.Text = IIf(curCount = 0, "", CLng(curCount))
        txtAVG.Text = IIf(curAVG = 0, "", Format(curAVG, "0.00"))
        txtSD.Text = IIf(curSD = 0, "", Format(curSD, "0.000"))
        
        If curSD <> 0 Then
            For lngRow = .FixedRows To .Rows - 1
                str结果 = Trim(.TextMatrix(lngRow, .ColIndex("项目结果")))
                If IsNumeric(str结果) Then
                    If Val(str结果) > 4 * curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">4S"
                    ElseIf Val(str结果) > 3 * curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">3S"
                    ElseIf Val(str结果) > 2 * curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">2S"
                    ElseIf Val(str结果) > curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">1S"
                    End If
                End If
            Next
        End If
    End With
    Me.cmdCalc.Enabled = False
End Sub

Private Sub RefGrid_结果Item(ByVal Index As Long, ByVal strValue As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng标本ID As Long, lng偏低 As Long, lng偏高 As Long, lng警示 As Long
    Dim lngColor  As Long, lngForeColor As Long, str标志 As String
    
    lng偏低 = &H80FFFF: lng偏高 = &H80C0FF: lng警示 = &H40C0&
    
    lng标本ID = Val(strValue)
    If lng标本ID = 0 Then Exit Sub
    strSQL = "Select c.中文名 As 项目, d.缩写 as 英文名, b.检验结果, Decode(B.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') as 结果标志, b.结果参考" & vbNewLine & _
        "From 检验项目 d,诊治所见项目 c, 检验普通结果 b, 检验标本记录 a" & vbNewLine & _
        "Where b.检验项目id=d.诊治项目id And a.Id = b.检验标本id And a.报告结果 = b.记录类型 And b.检验项目id = c.Id And a.Id = [1]" & vbNewLine & _
        " Order by Nvl(to_number(d.排列序号),c.编码)"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, lng标本ID)
    Do Until rsTmp.EOF
        With vfgItem(Index)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = Trim("" & rsTmp!项目)
            .TextMatrix(.Rows - 1, .ColIndex("英文名")) = Trim("" & rsTmp!英文名)
            .TextMatrix(.Rows - 1, .ColIndex("项目值")) = Trim("" & rsTmp!检验结果)
            If IsNumeric(.TextMatrix(.Rows - 1, .ColIndex("项目值"))) Then
                .Cell(flexcpAlignment, .Rows - 1, .ColIndex("项目值")) = flexAlignRightCenter
            Else
                .Cell(flexcpAlignment, .Rows - 1, .ColIndex("项目值")) = flexAlignLeftCenter
            End If
            
            str标志 = Trim("" & rsTmp!结果标志)
            lngColor = .BackColor
            lngForeColor = .ForeColor
            If InStr("↓", str标志) > 0 And str标志 <> "" Then     '2
                lngColor = lng偏低
            ElseIf InStr("↑,异常", str标志) > 0 And str标志 <> "" Then '3,异常
                lngColor = lng偏高
            ElseIf InStr("↑↑,↓↓", str标志) > 0 And str标志 <> "" Then  '5,6
                lngColor = lng警示
            End If
            .Cell(flexcpBackColor, .Rows - 1, .ColIndex("项目值")) = lngColor
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("项目值")) = lngForeColor
            
            .TextMatrix(.Rows - 1, .ColIndex("状态")) = str标志
            .TextMatrix(.Rows - 1, .ColIndex("参考范围")) = Trim("" & rsTmp!结果参考)
            .Rows = .Rows + 1
        End With
        rsTmp.MoveNext
    Loop
    
    With vfgItem(Index)
        '加表格线
        If .Rows > 2 Then .Rows = .Rows - 1
        .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
        
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("项目"), .Rows - 1, .ColIndex("项目")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("英文名"), .Rows - 1, .ColIndex("英文名")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("参考范围"), .Rows - 1, .ColIndex("参考范围")) = flexAlignLeftCenter
    End With
End Sub

Private Sub RefGrid_工作量Item(ByVal lngIndex As Long, ByVal strKey As String, ByVal strValue As String)
    '
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String, strEnd As String, lng仪器 As Long, str类型 As String, lng小组 As Long
    Dim strWhere As String, iCol As Integer
    Dim lng申请科室 As Long, str申请人 As String, str检验人 As String, str审核人 As String
    Dim lng病人来源 As Long, str统计方式 As String
    Dim lng样本数 As Long, lng诊疗项目id As Long
    
    'Call initvfgItemTitle(lngIndex)
    If InStr("小组,仪器,项目,申请科室,申请人,检验人,审核人,病人来源", strKey) <= 0 Then Exit Sub
    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    
    lng仪器 = Val(cbo仪器(lngIndex).ItemData(cbo仪器(lngIndex).ListIndex))
    lng小组 = Val(cbo小组(lngIndex).ItemData(cbo小组(lngIndex).ListIndex))
    str类型 = Trim(cbo类型(lngIndex).List(cbo类型(lngIndex).ListIndex))
    lng申请科室 = Val(cbo申请科室(lngIndex).ItemData(cbo申请科室(lngIndex).ListIndex))
    str申请人 = Trim(cbo申请人(lngIndex).List(cbo申请人(lngIndex).ListIndex))
    str检验人 = Trim(cbo检验人(lngIndex).List(cbo检验人(lngIndex).ListIndex))
    str审核人 = Trim(cbo审核人(lngIndex).List(cbo审核人(lngIndex).ListIndex))

    
    lng病人来源 = 0
    If opt来源(0).Value = True Then
        lng病人来源 = 0
    ElseIf opt来源(1).Value = True Then
        lng病人来源 = 1
    ElseIf opt来源(2).Value = True Then
        lng病人来源 = 2
    ElseIf opt来源(3).Value = True Then
        lng病人来源 = 3
    ElseIf opt来源(4).Value = True Then
        lng病人来源 = 4
    End If
    
    Select Case strKey
        Case "小组"
            lng小组 = Val(strValue)
        Case "仪器"
            lng仪器 = Val(strValue)
        Case "项目"
            lng诊疗项目id = Val(strValue)
        Case "申请科室"
            lng申请科室 = Val(strValue)
        Case "申请人"
            str申请人 = Trim(strValue)
        Case "检验人"
            str检验人 = Trim(strValue)
        Case "审核人"
            str审核人 = Trim(strValue)
        Case "病人来源"
            lng病人来源 = Val(strValue)
    End Select
    
    If str类型 <> "" Then strWhere = strWhere & " And D.仪器类型 =[5]"
    If lng小组 <> 0 Then strWhere = strWhere & " And C.ID=[4]"
    If lng仪器 <> 0 Then strWhere = strWhere & " And N.仪器ID=[3]"
    If str审核人 <> "" Then strWhere = strWhere & " And N.审核人=[9]"
    If str检验人 <> "" Then strWhere = strWhere & " And N.检验人=[8]"
    If str申请人 <> "" Then strWhere = strWhere & " And N.申请人=[7]"
    If lng申请科室 <> 0 Then strWhere = strWhere & " And N.申请科室ID= [6]"
    If lng病人来源 <> 0 Then strWhere = strWhere & " And N.病人来源=[10]"
    If lng诊疗项目id <> 0 Then strWhere = strWhere & " And N.诊疗项目id=[11]"
    
    strSQL = "Select n.中文名, n.缩写, Count(n.Id) As 数量" & vbNewLine & _
            "From 上机人员表 f, 检验小组成员 e, 检验仪器 d, 检验小组 c, 检验小组仪器 b," & vbNewLine & _
            "        (Select b.诊疗项目id,a.病人来源, a.仪器id, a.申请科室id, a.申请人, a.审核人, a.检验人, a.Id," & vbNewLine & _
            "                           b.检验项目id, c.中文名, d.缩写" & vbNewLine & _
            "            From 检验项目 d, 诊治所见项目 c, 检验普通结果 b, 检验标本记录 a" & vbNewLine & _
            "            Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And" & vbNewLine & _
            "                  a.报告结果=b.记录类型 And a.核收时间 Between [1] And [2] And nvl(a.微生物标本,0) = 0 And" & vbNewLine & _
            "                        a.审核人 Is Not Null) n" & vbNewLine & _
            "Where n.仪器id = b.仪器id And b.小组id = c.Id And n.仪器id = d.Id And c.Id = e.小组id And e.人员id = f.人员id And f.用户名 = User" & vbNewLine & _
            strWhere & vbNewLine & _
            "Group By n.中文名, n.缩写"
        
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng仪器, lng小组, str类型, lng申请科室, str申请人, str检验人, str审核人, lng病人来源, lng诊疗项目id)
    With vfgItem(lngIndex)
        '填充数据
        lng样本数 = 0
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = Trim("" & rsTmp!中文名)
            .TextMatrix(.Rows - 1, .ColIndex("英文名")) = Trim("" & rsTmp!缩写)
            .TextMatrix(.Rows - 1, .ColIndex("数量")) = IIf(Val("" & rsTmp!数量) = 0, "", Val("" & rsTmp!数量))
            lng样本数 = lng样本数 + Val("" & rsTmp!数量)
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        .TextMatrix(.Rows - 1, .ColIndex("项目")) = "合计"
        .TextMatrix(.Rows - 1, .ColIndex("数量")) = IIf(lng样本数 = 0, "", lng样本数)
        
        '加表格线
        .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
         
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("英文名")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("英文名") + 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        
    End With

End Sub

Private Sub RefGrid_工作量(ByVal lngIndex As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String, strEnd As String, lng仪器 As Long, str类型 As String, lng小组 As Long
    Dim strWhere As String, iCol As Integer, strTitle As String
    Dim lng申请科室 As Long, str申请人 As String, str检验人 As String, str审核人 As String
    Dim lng病人来源 As Long, str统计方式 As String
    Dim cur金额 As Currency, lng样本数 As Long
    Call initvfgDataTitle(lngIndex)
    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    strTitle = "日期：" & strBegin & " 至 " & strEnd
    lng仪器 = Val(cbo仪器(lngIndex).ItemData(cbo仪器(lngIndex).ListIndex))
    lng小组 = Val(cbo小组(lngIndex).ItemData(cbo小组(lngIndex).ListIndex))
    str类型 = Trim(cbo类型(lngIndex).List(cbo类型(lngIndex).ListIndex))
    
    If str类型 <> "" Then
        strWhere = strWhere & " And D.仪器类型 =[5]"
        strTitle = strTitle & "  仪器类型:" & Trim(cbo类型(lngIndex).List(cbo类型(lngIndex).ListIndex))
    End If
    
    If lng小组 <> 0 Then
        strWhere = strWhere & " And C.ID=[4]"
        strTitle = strTitle & "  小组:" & Trim(cbo小组(lngIndex).List(cbo小组(lngIndex).ListIndex))
    End If

    If lng仪器 <> 0 Then
        strWhere = strWhere & " And N.仪器ID=[3]"
        strTitle = strTitle & "  仪器:" & Trim(cbo仪器(lngIndex).List(cbo仪器(lngIndex).ListIndex))
    End If
    
    If strWhere <> "" Then strWhere = strWhere & vbNewLine
    lng申请科室 = Val(cbo申请科室(lngIndex).ItemData(cbo申请科室(lngIndex).ListIndex))
    If lng申请科室 <> 0 Then
        strWhere = strWhere & " And N.申请科室ID= [6]"
        strTitle = strTitle & "  申请科室:" & Trim(cbo申请科室(lngIndex).List(cbo申请科室(lngIndex).ListIndex))
    End If
    
    str申请人 = Trim(cbo申请人(lngIndex).List(cbo申请人(lngIndex).ListIndex))
    If str申请人 <> "" Then
        strWhere = strWhere & " And N.申请人=[7]"
        strTitle = strTitle & "  申请人:" & str申请人
    End If
    
    str检验人 = Trim(cbo检验人(lngIndex).List(cbo检验人(lngIndex).ListIndex))
    If str检验人 <> "" Then
        strWhere = strWhere & " And N.检验人=[8]"
        strTitle = strTitle & "  检验人:" & str检验人
    End If
    
    str审核人 = Trim(cbo审核人(lngIndex).List(cbo审核人(lngIndex).ListIndex))
    If str审核人 <> "" Then
        strWhere = strWhere & " And N.审核人=[9]"
        strTitle = strTitle & "  审核人:" & str审核人
    End If
    
    lng病人来源 = 0
    If opt来源(0).Value = True Then
        lng病人来源 = 0
    ElseIf opt来源(1).Value = True Then
        lng病人来源 = 1
        strTitle = strTitle & "  病人来源:门诊"
    ElseIf opt来源(2).Value = True Then
        lng病人来源 = 2
        strTitle = strTitle & "  病人来源:住院"
    ElseIf opt来源(3).Value = True Then
        lng病人来源 = 3
        strTitle = strTitle & "  病人来源:院外"
    ElseIf opt来源(4).Value = True Then
        lng病人来源 = 4
        strTitle = strTitle & "  病人来源:体检"
    End If
    
    If lng病人来源 <> 0 Then
        strWhere = strWhere & " And N.病人来源=[10]"
    End If
    
    If opt收费(1).Value = True Then
        strWhere = strWhere & " And Nvl(N.记录状态,0) <> 0 "
        strTitle = strTitle & "  已收费"
    End If
    If opt收费(2).Value = True Then
        strWhere = strWhere & " And Nvl(N.记录状态,0) = 0 "
        strTitle = strTitle & "  未收费"
    End If
    
    str统计方式 = "小组"
    If opt统计方式(0).Value = True Then
        str统计方式 = "小组"
    ElseIf opt统计方式(1).Value = True Then
        str统计方式 = "仪器"
    ElseIf opt统计方式(2).Value = True Then
        str统计方式 = "项目"
    ElseIf opt统计方式(3).Value = True Then
        str统计方式 = "申请科室"
    ElseIf opt统计方式(4).Value = True Then
        str统计方式 = "申请人"
    ElseIf opt统计方式(5).Value = True Then
        str统计方式 = "检验人"
    ElseIf opt统计方式(6).Value = True Then
        str统计方式 = "审核人"
    ElseIf opt统计方式(7).Value = True Then
        str统计方式 = "病人来源"
    End If
    
    
    With vfgData(lngIndex)
        If .FixedRows >= 2 Then
            For iCol = .FixedCols To .Cols - 1
                .TextMatrix(.FixedRows - 2, iCol) = strTitle
            Next
        End If
    End With
    If lng病人来源 = 0 Then
        strSQL = "            From 部门表 g, 上机人员表 f, 检验小组成员 e, 检验仪器 d, 检验小组 c, 检验小组仪器 b," & vbNewLine & _
                "                       (Select Distinct c.诊疗项目id,a.病人来源,a.申请科室id, a.申请人, a.检验人, a.审核人, a.Id, a.医嘱id, a.仪器id, a.核收时间, a.标本序号, e.编码, e.名称 as 项目, D.No, D.序号, d.记录性质, d.记录状态, d.实收金额 " & vbNewLine & _
                "                           From 诊疗项目目录 e, 住院费用记录 d, 病人医嘱记录 c, 检验项目分布 b, 检验标本记录 a" & vbNewLine & _
                "                           Where a.Id = b.标本id And b.医嘱id = c.相关id And c.Id = d.医嘱序号(+) And c.诊疗项目id = e.Id And" & vbNewLine & _
                "                                       a.核收时间 Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.微生物标本, 0) = 0 And a.审核人 Is Not Null" & _
                "                         union all             " & _
                "                        Select Distinct c.诊疗项目id,a.病人来源,a.申请科室id, a.申请人, a.检验人, a.审核人, a.Id, a.医嘱id, a.仪器id, a.核收时间, a.标本序号, e.编码, e.名称 as 项目, D.No, D.序号, d.记录性质, d.记录状态, d.实收金额 " & vbNewLine & _
                "                           From 诊疗项目目录 e, 门诊费用记录 d, 病人医嘱记录 c, 检验项目分布 b, 检验标本记录 a" & vbNewLine & _
                "                           Where a.Id = b.标本id And b.医嘱id = c.相关id And c.Id = d.医嘱序号(+) And c.诊疗项目id = e.Id And" & vbNewLine & _
                "                                       a.核收时间 Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.微生物标本, 0) = 0 And a.审核人 Is Not Null" & _
                "                         ) n" & vbNewLine & _
                "            Where n.申请科室id = g.Id And n.仪器id = b.仪器id And b.小组id = c.Id And n.仪器id = d.Id And c.Id = e.小组id And" & vbNewLine & _
                "                        e.人员id = f.人员id And f.用户名 = User"
    
    ElseIf lng病人来源 = 2 Then
    
        strSQL = "            From 部门表 g, 上机人员表 f, 检验小组成员 e, 检验仪器 d, 检验小组 c, 检验小组仪器 b," & vbNewLine & _
                "                       (Select Distinct c.诊疗项目id,a.病人来源,a.申请科室id, a.申请人, a.检验人, a.审核人, a.Id, a.医嘱id, a.仪器id, a.核收时间, a.标本序号, e.编码, e.名称 as 项目, D.No, D.序号, d.记录性质, d.记录状态, d.实收金额 " & vbNewLine & _
                "                           From 诊疗项目目录 e, 住院费用记录 d, 病人医嘱记录 c, 检验项目分布 b, 检验标本记录 a" & vbNewLine & _
                "                           Where a.Id = b.标本id And b.医嘱id = c.相关id And c.Id = d.医嘱序号(+) And c.诊疗项目id = e.Id And" & vbNewLine & _
                "                                       a.核收时间 Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.微生物标本, 0) = 0 And a.审核人 Is Not Null) n" & vbNewLine & _
                "            Where n.申请科室id = g.Id And n.仪器id = b.仪器id And b.小组id = c.Id And n.仪器id = d.Id And c.Id = e.小组id And" & vbNewLine & _
                "                        e.人员id = f.人员id And f.用户名 = User"

    Else
        strSQL = "            From 部门表 g, 上机人员表 f, 检验小组成员 e, 检验仪器 d, 检验小组 c, 检验小组仪器 b," & vbNewLine & _
                "                       (Select Distinct c.诊疗项目id,a.病人来源,a.申请科室id, a.申请人, a.检验人, a.审核人, a.Id, a.医嘱id, a.仪器id, a.核收时间, a.标本序号, e.编码, e.名称 as 项目, D.No, D.序号, d.记录性质, d.记录状态, d.实收金额 " & vbNewLine & _
                "                           From 诊疗项目目录 e, 门诊费用记录 d, 病人医嘱记录 c, 检验项目分布 b, 检验标本记录 a" & vbNewLine & _
                "                           Where a.Id = b.标本id And b.医嘱id = c.相关id And c.Id = d.医嘱序号(+) And c.诊疗项目id = e.Id And" & vbNewLine & _
                "                                       a.核收时间 Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.微生物标本, 0) = 0 And a.审核人 Is Not Null) n" & vbNewLine & _
                "            Where n.申请科室id = g.Id And n.仪器id = b.仪器id And b.小组id = c.Id And n.仪器id = d.Id And c.Id = e.小组id And" & vbNewLine & _
                "                        e.人员id = f.人员id And f.用户名 = User"
    
    End If
    Select Case str统计方式
    Case "小组"
        strSQL = "Select 小组id As Id, 小组, Count(Id) As 样本数, Sum(实收金额) As 金额" & vbNewLine & _
                 "From (Select c.Id As 小组id, c.名称 As 小组,n.Id, Sum(Nvl(n.实收金额, 0)) As 实收金额" & vbNewLine & _
                 strSQL & strWhere & _
                "      Group By c.Id, c.名称, n.Id)" & vbNewLine & _
                "Group By 小组id, 小组"
    Case "仪器"
        strSQL = "Select 仪器id As Id, 仪器, Count(Id) As 样本数, Sum(实收金额) As 金额" & vbNewLine & _
                 "From (Select d.Id As 仪器id, d.名称 As 仪器,n.Id, Sum(Nvl(n.实收金额, 0)) As 实收金额" & vbNewLine & _
                 strSQL & strWhere & _
                "      Group By d.Id, d.名称, n.Id)" & vbNewLine & _
                "Group By 仪器id, 仪器"

    Case "申请科室"
        strSQL = "Select 申请科室id as id,申请科室, Count(Id) As 样本数, Sum(实收金额) As 金额" & vbNewLine & _
                "From (Select n.申请科室id,g.名称 As 申请科室, n.Id, Sum(Nvl(n.实收金额, 0)) As 实收金额" & vbNewLine & _
                 strSQL & strWhere & _
                "            Group By n.申请科室id,g.名称, n.Id)" & vbNewLine & _
                "Group By 申请科室id,申请科室"
    Case "项目"
        strSQL = "Select 诊疗项目id as id," & str统计方式 & ", Count(Id) As 样本数, Sum(实收金额) As 金额" & vbNewLine & _
                "From (Select n.诊疗项目id,n." & str统计方式 & ", n.Id, Sum(Nvl(n.实收金额, 0)) As 实收金额" & vbNewLine & _
                strSQL & strWhere & _
                "            Group By n.诊疗项目id, n." & str统计方式 & ", n.Id)" & vbNewLine & _
                "Group By 诊疗项目id," & str统计方式
    
    Case "病人来源"
        strSQL = "Select 病人来源 as id, decode(病人来源,1,'门诊',2,'住院',4,'体检','院外') as 病人来源, Count(Id) As 样本数, Sum(实收金额) As 金额" & vbNewLine & _
                "From (Select n." & str统计方式 & ", n.Id, Sum(Nvl(n.实收金额, 0)) As 实收金额" & vbNewLine & _
                strSQL & strWhere & _
                "            Group By n.病人来源,decode(n.病人来源,1,'门诊',2,'住院',4,'体检','院外'), n.Id)" & vbNewLine & _
                "Group By " & str统计方式
                
    Case "检验人", "审核人", "申请人"
        strSQL = "Select " & str统计方式 & " as id," & str统计方式 & ", Count(Id) As 样本数, Sum(实收金额) As 金额" & vbNewLine & _
                "From (Select n." & str统计方式 & ", n.Id, Sum(Nvl(n.实收金额, 0)) As 实收金额" & vbNewLine & _
                strSQL & strWhere & _
                "            Group By n." & str统计方式 & ", n.Id)" & vbNewLine & _
                "Group By " & str统计方式
    
    End Select
    
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng仪器, lng小组, str类型, lng申请科室, str申请人, str检验人, str审核人, lng病人来源)
    
    With vfgData(lngIndex)
        .TextMatrix(.FixedRows - 1, .ColIndex("小组")) = str统计方式
        '填充数据
        lng样本数 = 0: cur金额 = 0
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Trim("" & rsTmp!Id)
            .TextMatrix(.Rows - 1, .ColIndex("小组")) = Trim("" & rsTmp.Fields(1))
            .TextMatrix(.Rows - 1, .ColIndex("样本数")) = IIf(Val("" & rsTmp!样本数) = 0, "", Val("" & rsTmp!样本数))
            .TextMatrix(.Rows - 1, .ColIndex("金额")) = IIf(Val("" & rsTmp!金额) = 0, "", Format(Val("" & rsTmp!金额), "0.00"))
            lng样本数 = lng样本数 + Val("" & rsTmp!样本数)
            cur金额 = cur金额 + Val("" & rsTmp!金额)
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        .TextMatrix(.Rows - 1, .ColIndex("小组")) = "合计"
        .TextMatrix(.Rows - 1, .ColIndex("样本数")) = IIf(lng样本数 = 0, "", lng样本数)
        .TextMatrix(.Rows - 1, .ColIndex("金额")) = IIf(cur金额 = 0, "", Format(cur金额, "0.00"))
        
        '加表格线
        .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols + 1
         
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("小组")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("小组") + 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        
    End With
    
    
End Sub

Private Sub RefGrid_日常(ByVal lngIndex As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String
    Dim strEnd As String
    Dim lng仪器 As Long, str类型 As String, lng小组 As Long
    Dim strWhere As String, iCol As Integer, strTitle As String '附标题
    
    Call initvfgDataTitle(lngIndex)

    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    strTitle = "日期：" & strBegin & " 至 " & strEnd
    lng仪器 = Val(cbo仪器(lngIndex).ItemData(cbo仪器(lngIndex).ListIndex))
    lng小组 = Val(cbo小组(lngIndex).ItemData(cbo小组(lngIndex).ListIndex))
    str类型 = Trim(cbo类型(lngIndex).List(cbo类型(lngIndex).ListIndex))
    
    If str类型 <> "" Then
        strWhere = strWhere & " And D.仪器类型 =[5]"
        strTitle = strTitle & "  仪器类型:" & Trim(cbo类型(lngIndex).List(cbo类型(lngIndex).ListIndex))
    Else
        strTitle = strTitle & "  仪器类型:所有类型"
    End If
    
    If lng小组 <> 0 Then
        strWhere = strWhere & " And C.ID=[4]"
        strTitle = strTitle & "  小组:" & Trim(cbo小组(lngIndex).List(cbo小组(lngIndex).ListIndex))
    Else
        strTitle = strTitle & "  小组:所有小组"
    End If

    If lng仪器 <> 0 Then
        strWhere = strWhere & " And A.仪器ID = [3]"
        strTitle = strTitle & "  仪器:" & Trim(cbo仪器(lngIndex).List(cbo仪器(lngIndex).ListIndex))
    Else
        strTitle = strTitle & "  仪器:所有仪器"
    End If
        

    With vfgData(lngIndex)
        If .FixedRows >= 2 Then
            For iCol = .FixedCols To .Cols - 1
                .TextMatrix(.FixedRows - 2, iCol) = strTitle
            Next
        End If
    End With
    
    strSQL = "Select D.仪器类型 As 类型, c.名称 As 小组, d.名称 As 仪器, Sum(Decode(a.姓名, Null, 1, 0)) As 无主," & vbNewLine & _
            "            Sum(Decode(a.接收人, Null, 0, Decode(a.审核人, Null, 1, 0))) As 已接收, Count(a.Id) As 已核收," & vbNewLine & _
            "            Sum(Decode(a.审核人, Null, 0, 1)) As 审核, Sum(Decode(a.审核人, Null, 1, 0)) As 未审" & vbNewLine & _
            "From 上机人员表 f, 检验小组成员 e, 检验仪器 D, 检验小组 C, 检验小组仪器 b, 检验标本记录 a" & vbNewLine & _
            "Where a.核收时间 Between [1] And [2] And a.仪器id = b.仪器id And" & vbNewLine & _
            "           b.小组id = c.Id And a.仪器id = d.Id And Nvl(a.微生物标本,0)=0 And c.Id = e.小组id And e.人员id = f.人员id And f.用户名 = User" & vbNewLine & _
            strWhere & vbNewLine & _
            "Group By D.仪器类型, c.名称, d.名称"

    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng仪器, lng小组, str类型)
    
    With vfgData(lngIndex)
        '填充数据
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, mCol_日常.类型) = Trim("" & rsTmp!类型)
            .TextMatrix(.Rows - 1, mCol_日常.小组) = Trim("" & rsTmp!小组)
            .TextMatrix(.Rows - 1, mCol_日常.仪器) = Trim("" & rsTmp!仪器)
            .TextMatrix(.Rows - 1, mCol_日常.无主) = IIf(Val("" & rsTmp!无主) = 0, "", Val("" & rsTmp!无主))
            .TextMatrix(.Rows - 1, mCol_日常.已接收) = IIf(Val("" & rsTmp!已接收) = 0, "", Val("" & rsTmp!已接收))
            .TextMatrix(.Rows - 1, mCol_日常.已核收) = IIf(Val("" & rsTmp!已核收) = 0, "", Val("" & rsTmp!已核收))
            .TextMatrix(.Rows - 1, mCol_日常.审核) = IIf(Val("" & rsTmp!审核) = 0, "", Val("" & rsTmp!审核))
            .TextMatrix(.Rows - 1, mCol_日常.未审) = IIf(Val("" & rsTmp!未审) = 0, "", Val("" & rsTmp!未审))
            
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop

        If .Rows > 4 Then .Rows = .Rows - 1
        
        '求合计
        .Subtotal flexSTClear
        .OutlineCol = 1   '指定输出列
        .SubtotalPosition = flexSTBelow '合计在底部
        .Subtotal flexSTSum, -1, .ColIndex("无主"), , , , , "合计"
        .Subtotal flexSTSum, -1, .ColIndex("已核收"), , , , , "合计"
        .Subtotal flexSTSum, -1, .ColIndex("已接收"), , , , , "合计"
        .Subtotal flexSTSum, -1, .ColIndex("已核收"), , , , , "合计"
        .Subtotal flexSTSum, -1, .ColIndex("审核"), , , , , "合计"
        .Subtotal flexSTSum, -1, .ColIndex("未审"), , , , , "合计"
        
        For iCol = .ColIndex("无主") To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = Replace(.TextMatrix(.Rows - 1, iCol), ".00", "")
        Next
        
        '加表格线
        
        .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
         
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, mCol_日常.仪器) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, mCol_日常.无主, .Rows - 1, .Cols - 1) = flexAlignRightCenter
    End With

End Sub

Private Sub initvfgDataTitle(ByVal Index As Long)
    Dim strFiles As String, strTitle As String, strFont As String
    
    Select Case Index
    Case 0
        strFiles = ",100;类型,900;小组,900;仪器,1800;无主,900;已接收,900;已核收,900;审核,900;未审,900"
        strTitle = Trim(ReadIni("Report0", "标题", App.Path & "\PrintSetup.ini"))
        strFont = Trim(ReadIni("Report0", "标题字体", App.Path & "\PrintSetup.ini"))
        If strTitle = "" Then strTitle = "日常报表"
        
        vfgData(0).Rows = 4: vfgData(0).Cols = 9
        
        Call initVfg(vfgData(0), strFiles, strTitle, strFont)
        With vfgData(0)
            .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
        
    Case 1
        strFiles = ",100;ID,0;小组,2800;样本数,2000;金额,2000"
        strTitle = Trim(ReadIni("Report1", "标题", App.Path & "\zl9LisQuery_Base.ini"))
        strFont = Trim(ReadIni("Report1", "标题字体", App.Path & "\PrintSetup.ini"))
        If strTitle = "" Then strTitle = "工作量统计"
        vfgData(1).Rows = 4: vfgData(1).Cols = 5
        Call initVfg(vfgData(1), strFiles, strTitle, strFont)
        With vfgData(1)
            .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    Case 2
        strFiles = ",100;标本ID,0;核收日期,1800;检验小组,1200;样本号,900;项目,2200;项目结果,1200;姓名,900;性别,800;年龄,800;SD,1000"
        strTitle = Trim(ReadIni("Report2", "标题", App.Path & "\zl9LisQuery_Base.ini"))
        strFont = Trim(ReadIni("Report2", "标题字体", App.Path & "\PrintSetup.ini"))
        If strTitle = "" Then strTitle = "结果统计"
        
        vfgData(2).Rows = 4: vfgData(2).Cols = 11
        Call initVfg(vfgData(2), strFiles, strTitle, strFont)
        With vfgData(2)
            .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    End Select
End Sub

Private Sub initvfgItemTitle(ByVal Index As Long)
    Dim strFiles As String, strTitle As String
    Select Case Index
    Case 1
        strFiles = ",0;项目,2800;英文名,1000;数量,2000"
        strTitle = ""
        
        vfgItem(1).Rows = 2: vfgItem(1).Cols = 4
        Call initVfg(vfgItem(1), strFiles, strTitle, "")
        With vfgItem(1)
            .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    Case 2
        strFiles = ",0;项目,1800;英文名,1000;项目值,1000;状态,800;参考范围,2000;"
        strTitle = ""
        vfgItem(2).Rows = 2: vfgItem(2).Cols = 6
        Call initVfg(vfgItem(2), strFiles, strTitle, "")
        With vfgItem(2)
            .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    End Select
End Sub

Private Sub initVfg(objVfg As VSFlexGrid, ByVal str字段 As String, ByVal strTitle As String, ByVal strFont As String)
    Dim iCol As Integer
    Dim varTmp As Variant, varTmp1 As Variant
    On Error GoTo errH
    varTmp = Split(str字段, ";")
    
    If UBound(Split(strFont, "|")) <> 2 Then strFont = "宋体|18"
    
    With objVfg
        .Clear
        .Editable = flexEDNone
        .GridLines = flexGridNone
        
        .MergeCells = flexMergeRestrictRows
        .BackColorFixed = .BackColor
        .ForeColorFixed = .ForeColor
        .GridColorFixed = .GridColor
        .GridLinesFixed = flexGridNone
        
                
        If strTitle <> "" Then
            If .Rows < 4 Then Exit Sub
            .FixedCols = 1: .FixedRows = 3
            '-- 表头
            For iCol = 0 To 1
                .MergeRow(iCol) = True
            Next
            
            If strTitle <> "" Then
                For iCol = .FixedCols To .Cols - 1
                    .TextMatrix(0, iCol) = strTitle
                Next
            End If
            
            .Cell(flexcpFontName, 0, .FixedCols, 0, .Cols - 1) = Split(strFont, "|")(0)
            .Cell(flexcpFontSize, 0, .FixedCols, 0, .Cols - 1) = Split(strFont, "|")(1)
            .Cell(flexcpFontBold, 0, .FixedCols, 0, .Cols - 1) = True
            .RowHeight(0) = 600
            .RowHeight(1) = 500
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        Else
            If .Rows < 2 Then Exit Sub
            .FixedCols = 1: .FixedRows = 1
        End If
        
        For iCol = LBound(varTmp) To UBound(varTmp)
            If InStr(varTmp(iCol), ",") > 0 Then
                varTmp1 = Split(varTmp(iCol), ",")
                .TextMatrix(.FixedRows - 1, iCol) = Trim(varTmp1(0))
                If .TextMatrix(.FixedRows - 1, iCol) <> "" Then .ColKey(iCol) = .TextMatrix(.FixedRows - 1, iCol)
                .ColWidth(iCol) = Val(varTmp1(1))
                If .ColWidth(iCol) = 0 Then .ColHidden(iCol) = True
                .ColAlignment(iCol) = flexAlignCenterCenter
            End If
        Next
    End With
    Exit Sub
errH:
    MsgBox "initvfg" & vbCrLf & str字段 & vbCrLf & Err.Description, vbQuestion, Me.Caption
End Sub

Private Sub ShowSelect(ByVal strInput As String)
    Dim strSQL As String
    
    Dim strP1 As String
    Dim strWhere As String
    
    If strInput <> "" Then
        strWhere = " And (D.编码 Like '%" & UCase(strInput) & "%' Or Upper(D.中文名) Like '%" & UCase(strInput) & "%' Or Upper(C.缩写) Like '%" & UCase(strInput) & "%')"
    End If
    strSQL = "Select c.诊治项目id As 项目id, d.编码, d.中文名, c.缩写, d.单位, c.项目类别, c.结果类型, c.取值序列" & vbNewLine & _
            "From 诊治所见项目 d, 检验项目 c" & vbNewLine & _
            "Where c.诊治项目id = d.Id " & strWhere & " Order by D.编码"
            
    Set mrs项目 = clsHost.GetRecordSet(strSQL, Me.Caption)
    
    lstSelect.Clear
    Do Until mrs项目.EOF
        lstSelect.AddItem "" & mrs项目!编码 & "-" & mrs项目!中文名 & IIf(Trim("" & mrs项目!缩写) = "", "", "(" & mrs项目!缩写 & ")")
        lstSelect.ItemData(lstSelect.NewIndex) = Val("" & mrs项目!项目id)
        mrs项目.MoveNext
    Loop
    If lstSelect.ListCount > 0 Then
        Call MoveSelect(txt项目)
        lstSelect.ListIndex = 0
        lstSelect.Visible = True
        lstSelect.SetFocus
    End If
End Sub

Private Sub MoveSelect(ByVal ctrl As Control)
    
    Dim vRect As RECT
    Dim vRect1 As RECT
    
    vRect = GetControlRect(ctrl.hWnd)
    vRect1 = GetControlRect(lstSelect.hWnd)
    
    lstSelect.Top = lstSelect.Top + (vRect.Top - vRect1.Top) + ctrl.Height + 10
    lstSelect.Left = lstSelect.Left + (vRect.Left - vRect1.Left)
    lstSelect.Width = ctrl.Width

End Sub

Private Function CalcSVG(ByVal strVal As String) As Currency
'   均值
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + Val(varInData(i))
    Next
    If i > 0 Then
        CalcSVG = curX / i
    End If
End Function
Private Function CalcSD(ByVal strVal As String) As Currency
    '标准差
    Dim varInData As Variant, curX As Currency, i As Integer, cur均值 As Currency
    
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    cur均值 = CalcSVG(strVal)
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + (Val(varInData(i)) - cur均值) ^ 2
    Next
    If i - 1 > 0 Then
        CalcSD = Sqr(curX / (i - 1))
    End If
    'Sqr (∑(xn - x拨) ^ 2 / (N - 1))
End Function

