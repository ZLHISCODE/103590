VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTunning 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   11430
   ClientLeft      =   0
   ClientTop       =   360
   ClientWidth     =   18960
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TabStrip tabData 
      Height          =   310
      Left            =   0
      TabIndex        =   49
      Top             =   9480
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   556
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      TabFixedHeight  =   1411
      Placement       =   1
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "当前低性能"
            Object.ToolTipText     =   "当前共享池中，执行计划含有全表扫描、索引全扫描、索引快速全表扫描、索引跳跃扫描的SQL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "历史低性能"
            Object.ToolTipText     =   "SQL历史库中，执行计划含有全表扫描、索引全扫描、索引快速全表扫描、索引跳跃扫描的SQL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "高逻辑读"
            Object.ToolTipText     =   "当前共享池中，单次执行的逻辑读超过指定块数并且执行次数超过2次的SQL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "硬解析"
            Object.ToolTipText     =   "当前共享池中，没有使用绑定变量的SQL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "高频执行"
            Object.ToolTipText     =   "当前共享池中，累计执行次数超过指定次数的SQL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "含提示字"
            Object.ToolTipText     =   "当前共享池中，SQL文本中含有rule等hints的SQL"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pctLine 
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   6000
      MousePointer    =   9  'Size W E
      ScaleHeight     =   8055
      ScaleWidth      =   45
      TabIndex        =   9
      Top             =   480
      Width           =   45
   End
   Begin VB.PictureBox pctSqlList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9465
      Left            =   0
      ScaleHeight     =   9465
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.TextBox txtRate 
         Height          =   300
         Left            =   2760
         TabIndex        =   55
         Text            =   "1000"
         Top             =   8280
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4320
         TabIndex        =   51
         Top             =   8790
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         Format          =   253362179
         CurrentDate     =   42964
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   2040
         TabIndex        =   50
         Top             =   8790
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         Format          =   253362179
         CurrentDate     =   42964
      End
      Begin VB.CheckBox chkZlhis 
         Caption         =   "仅含业务相关SQL"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         ToolTipText     =   "不包括数据库系统用户(SYS,SYSTEM,OGG等)解析的SQL。"
         Top             =   8280
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "保存至Excel"
         Height          =   350
         Left            =   5520
         TabIndex        =   7
         Top             =   25
         Width           =   1335
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   5760
         TabIndex        =   5
         Top             =   8280
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   7215
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6495
         _cx             =   11456
         _cy             =   12726
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
         GridColor       =   32768
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   300
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "请输入SQL文本后按回车，支持按SQLID查找"
         Top             =   60
         Width           =   2415
      End
      Begin VB.CheckBox chkPLSQL 
         Caption         =   "显示PL/SQL中执行的语句"
         Height          =   255
         Left            =   480
         TabIndex        =   68
         Top             =   90
         Width           =   2295
      End
      Begin VB.Label lblInst 
         AutoSize        =   -1  'True
         Caption         =   "当前实例ID：1"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1320
         TabIndex        =   10
         Top             =   127
         Width           =   1170
      End
      Begin VB.Label lblRate 
         AutoSize        =   -1  'True
         Caption         =   "次数"
         Height          =   180
         Left            =   2400
         TabIndex        =   54
         Top             =   8340
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblEndTime 
         AutoSize        =   -1  'True
         Caption         =   "到"
         Height          =   180
         Left            =   3960
         TabIndex        =   53
         Top             =   8835
         Width           =   180
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         Caption         =   "最后执行时间从"
         Height          =   180
         Left            =   720
         TabIndex        =   52
         Top             =   8850
         Width           =   1260
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         Caption         =   "SQL语句列表"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   127
         Width           =   990
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   2880
         TabIndex        =   1
         Top             =   127
         Width           =   360
      End
   End
   Begin TabDlg.SSTab sstPlan 
      Height          =   9855
      Left            =   7200
      TabIndex        =   6
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   17383
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   10
      TabHeight       =   635
      TabCaption(0)   =   "执行计划"
      TabPicture(0)   =   "frmTunning.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pctPlan"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "统计信息"
      TabPicture(1)   =   "frmTunning.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pctStatics"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "执行信息"
      TabPicture(2)   =   "frmTunning.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pctUser"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "SQLProfile及优化器参数"
      TabPicture(3)   =   "frmTunning.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pctProfiles"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "SQL语句AWR"
      TabPicture(4)   =   "frmTunning.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pctAWR"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox pctAWR 
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   9735
         TabIndex        =   56
         Top             =   360
         Width           =   9735
         Begin VB.CommandButton cmdAwr 
            Caption         =   "过滤(&F)"
            Height          =   300
            Left            =   5280
            TabIndex        =   60
            Top             =   120
            Width           =   975
         End
         Begin SHDocVwCtl.WebBrowser webAwr 
            Height          =   4095
            Left            =   0
            TabIndex        =   57
            Top             =   600
            Width           =   6375
            ExtentX         =   11245
            ExtentY         =   7223
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin MSComCtl2.DTPicker dtpStartInterval 
            Height          =   300
            Left            =   1200
            TabIndex        =   59
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd HH:mm"
            Format          =   253362179
            CurrentDate     =   42964
         End
         Begin MSComCtl2.DTPicker dtpEndInterval 
            Height          =   300
            Left            =   3360
            TabIndex        =   61
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd HH:mm"
            Format          =   253362179
            CurrentDate     =   42964
         End
         Begin VB.Label lblEndInterval 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   3120
            TabIndex        =   62
            Top             =   180
            Width           =   180
         End
         Begin VB.Label lblSartInterval 
            AutoSize        =   -1  'True
            Caption         =   "快照生成时间"
            Height          =   180
            Left            =   120
            TabIndex        =   58
            Top             =   180
            Width           =   1080
         End
      End
      Begin VB.PictureBox pctStatics 
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   8535
         Left            =   -74880
         ScaleHeight     =   8535
         ScaleWidth      =   10335
         TabIndex        =   34
         Top             =   380
         Width           =   10335
         Begin VB.CommandButton cmdExecuteAll 
            Caption         =   "收集当前所有表(&A)"
            Height          =   350
            Left            =   8160
            TabIndex        =   63
            Top             =   8040
            Width           =   1815
         End
         Begin VB.TextBox txtAdv 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   6480
            Width           =   9975
         End
         Begin VB.CommandButton cmdExecute 
            Caption         =   "收集当前表(&E)"
            Height          =   350
            Left            =   6480
            TabIndex        =   38
            Top             =   8040
            Width           =   1575
         End
         Begin VB.OptionButton optAuto 
            Caption         =   "Auto"
            Height          =   180
            Left            =   3960
            TabIndex        =   37
            Top             =   8125
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optSkewOnly 
            Caption         =   "SkewOnly"
            Height          =   180
            Left            =   4680
            TabIndex        =   36
            Top             =   8125
            Width           =   1095
         End
         Begin VB.OptionButton optNull 
            Caption         =   "Null"
            Height          =   180
            Left            =   5760
            TabIndex        =   35
            Top             =   8125
            Width           =   735
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfColSta 
            Height          =   3855
            Left            =   0
            TabIndex        =   40
            ToolTipText     =   "颜色加重行标识索引涉及的列。"
            Top             =   2160
            Width           =   4935
            _cx             =   8705
            _cy             =   6800
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vsfTblSta 
            Height          =   1335
            Left            =   0
            TabIndex        =   41
            Top             =   360
            Width           =   9735
            _cx             =   17171
            _cy             =   2355
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vsfIdx 
            Height          =   3855
            Left            =   5040
            TabIndex        =   42
            ToolTipText     =   "颜色加重行标明当前SQL语句使用到的索引。"
            Top             =   2160
            Width           =   4935
            _cx             =   8705
            _cy             =   6800
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VB.Label lblAdv 
            AutoSize        =   -1  'True
            Caption         =   "收集统计信息"
            Height          =   180
            Left            =   0
            TabIndex        =   48
            Top             =   6240
            Width           =   1080
         End
         Begin VB.Label lblSTa 
            AutoSize        =   -1  'True
            Caption         =   "表统计信息"
            Height          =   180
            Left            =   0
            TabIndex        =   47
            Top             =   60
            Width           =   900
         End
         Begin VB.Label lblColSta 
            AutoSize        =   -1  'True
            Caption         =   "列统计信息"
            Height          =   180
            Left            =   0
            TabIndex        =   46
            Top             =   1920
            Width           =   900
         End
         Begin VB.Label lblIdx 
            AutoSize        =   -1  'True
            Caption         =   "索引统计信息"
            Height          =   180
            Left            =   5040
            TabIndex        =   45
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label lblTip2 
            AutoSize        =   -1  'True
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   120
            TabIndex        =   44
            Top             =   8085
            Width           =   90
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   "收集方式"
            Height          =   180
            Left            =   3120
            TabIndex        =   43
            Top             =   8130
            Width           =   720
         End
      End
      Begin VB.PictureBox pctPlan 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   120
         ScaleHeight     =   8655
         ScaleWidth      =   10095
         TabIndex        =   24
         Top             =   380
         Width           =   10095
         Begin VB.CommandButton cmdFree 
            Caption         =   "编辑自定义提示"
            Height          =   350
            Left            =   2400
            TabIndex        =   29
            Top             =   7920
            Width           =   1455
         End
         Begin VB.PictureBox pctHorLine 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   0
            MousePointer    =   7  'Size N S
            ScaleHeight     =   135
            ScaleWidth      =   9015
            TabIndex        =   28
            Top             =   2040
            Width           =   9015
         End
         Begin VB.CommandButton cmdAuto 
            Caption         =   "自动优化(&A)"
            Height          =   350
            Left            =   7680
            TabIndex        =   27
            Top             =   7920
            Width           =   1335
         End
         Begin VB.CommandButton cmdRule 
            Caption         =   "添加RULE提示"
            Height          =   350
            Left            =   6120
            TabIndex        =   26
            Top             =   7920
            Width           =   1455
         End
         Begin VB.CommandButton cmdOptmizer 
            Caption         =   "添加优化器版本提示"
            Height          =   350
            Left            =   3960
            TabIndex        =   25
            Top             =   7920
            Width           =   2055
         End
         Begin MSComctlLib.TabStrip tabPlan 
            Height          =   375
            Left            =   0
            TabIndex        =   30
            Top             =   2280
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
            Height          =   4935
            Index           =   1
            Left            =   0
            TabIndex        =   31
            ToolTipText     =   "颜色加重行标识当前语句引起性能问题的原因。"
            Top             =   2640
            Visible         =   0   'False
            Width           =   9015
            _cx             =   15901
            _cy             =   8705
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
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin RichTextLib.RichTextBox txtFullSql 
            Height          =   2055
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   3625
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmTunning.frx":008C
         End
         Begin VB.Label lblTip1 
            AutoSize        =   -1  'True
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   240
            TabIndex        =   33
            Top             =   8005
            Width           =   90
         End
      End
      Begin VB.PictureBox pctProfiles 
         BorderStyle     =   0  'None
         Height          =   9240
         Left            =   -75000
         ScaleHeight     =   9240
         ScaleWidth      =   10215
         TabIndex        =   16
         Top             =   380
         Width           =   10215
         Begin VB.CommandButton cmdOptExecute 
            Caption         =   "执行(&E)"
            Height          =   350
            Left            =   7680
            TabIndex        =   65
            Top             =   8280
            Width           =   1095
         End
         Begin VB.TextBox txtOptExecute 
            Height          =   300
            Left            =   1320
            TabIndex        =   64
            Top             =   8305
            Width           =   6255
         End
         Begin VB.CommandButton cmdAllProfiles 
            Caption         =   "显示全部SQL PROFILES"
            Height          =   350
            Left            =   4920
            TabIndex        =   19
            Top             =   5040
            Width           =   2175
         End
         Begin VB.CommandButton cmdRProfiles 
            Caption         =   "刷新"
            Height          =   350
            Left            =   9120
            TabIndex        =   18
            Top             =   5040
            Width           =   975
         End
         Begin VB.CommandButton cmdDelProfile 
            Caption         =   "删除SQL PROFILE"
            Height          =   350
            Left            =   7200
            TabIndex        =   17
            Top             =   5040
            Width           =   1815
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfOpt 
            Height          =   1815
            Left            =   0
            TabIndex        =   20
            Top             =   5880
            Width           =   10095
            _cx             =   17806
            _cy             =   3201
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vsfProfiles 
            Height          =   4455
            Left            =   0
            TabIndex        =   21
            Top             =   360
            Width           =   10095
            _cx             =   17806
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VB.Label lblOptExecute 
            Caption         =   "修改优化器参数"
            Height          =   180
            Left            =   0
            TabIndex        =   67
            Top             =   8365
            Width           =   1260
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            Caption         =   "优化器相关参数"
            Height          =   180
            Left            =   0
            TabIndex        =   66
            Top             =   5640
            Width           =   1260
         End
         Begin VB.Label lblProfiles 
            AutoSize        =   -1  'True
            Caption         =   "SQL PROFILES列表"
            Height          =   180
            Left            =   0
            TabIndex        =   23
            Top             =   60
            Width           =   1440
         End
         Begin VB.Label lblTip4 
            AutoSize        =   -1  'True
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   0
            TabIndex        =   22
            Top             =   5130
            Width           =   90
         End
      End
      Begin VB.PictureBox pctUser 
         BorderStyle     =   0  'None
         Height          =   8175
         Left            =   -74880
         ScaleHeight     =   8175
         ScaleWidth      =   9975
         TabIndex        =   11
         Top             =   380
         Width           =   9975
         Begin VSFlex8Ctl.VSFlexGrid vsfUser 
            Height          =   5295
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Width           =   9855
            _cx             =   17383
            _cy             =   9340
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vsfReport 
            Height          =   1455
            Left            =   0
            TabIndex        =   13
            Top             =   6000
            Width           =   9855
            _cx             =   17383
            _cy             =   2566
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
            GridColor       =   32768
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   380
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VB.Label lblUser 
            AutoSize        =   -1  'True
            Caption         =   "当前SQL执行人员"
            Height          =   180
            Left            =   0
            TabIndex        =   15
            Top             =   60
            Width           =   1350
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            Caption         =   "当前SQL相关报表"
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   5760
            Width           =   1350
         End
      End
   End
End
Attribute VB_Name = "frmTunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsAdmins As ADODB.Recordset    '用于记录zlsystems中的所有者
Private mstrCompatible As String    '记录当前数据库兼容性版本
Private mstrOptVision As String       '记录当前数据库优化器版本
Private mstrTbl_Owner() As String    '记录当前SQL Plan涉及的表及其所有者
Private mstrIdx_Owner() As String   '记录当前SQL Plan使用的索引及其所有者
Private mintIns_ID As Integer '记录SQL语句执行环境的Ins_id
Private mstrNewSqlId As String    '记录当前选中的SQL语句的SQL_ID
Private mblnClicked(5) As Boolean '分别记录Tab是否被点击过
Private mintFirPlan As Integer  '记录SQLPLAN的childNumber下界
Private mstrFilePath As String  '存储AWR的临时文件
Private mlngMinSize As Long '中型表大小
Private mlngMaxSize As Long
Private mrsBigTbl As ADODB.Recordset    '需要检查的表
Private mrsBigIdx As ADODB.Recordset
Private mrsLowIdx As ADODB.Recordset

Private WithEvents mfrmComments As frmComments
Attribute mfrmComments.VB_VarHelpID = -1
Private Enum TabNum
    tab1 = 0
    tab2 = 1
    tab3 = 2
    tab4 = 3
    tab5
End Enum
Private Const conCol = "Operation,2000,1;Name,500,1;ID,500,1;Cardinality,500,1;Bytes,500,1;Cost,500,1;Time,500,1;Object_Owner,500,1;Object_Type,500,1"

Public Sub ShowMe()
    Me.Show
End Sub

Private Sub cmdAwr_Click()
    webAwr.Navigate "about:blank"
    GetAWRByTime
End Sub

Private Sub cmdExecuteAll_Click()
    Dim strTmp As String, strTbl As String
    Dim i As Integer, strSql As String
    
    On Error GoTo errH
    With vsfTblSta
        strTmp = "是否要收集以下" & .Rows - 1 & "张表的统计信息，此操作耗时较长，请在业务空闲期进行。" & vbNewLine
        For i = .FixedRows To .Rows - 1
            strTmp = strTmp & .TextMatrix(i, .ColIndex("OWNER")) & "." & .TextMatrix(i, .ColIndex("TABLE_NAME")) & vbNewLine
        Next
    
        If strTmp = "是否要收集以下" & .Rows - 1 & "张表的统计信息，此操作耗时较长，请在业务空闲期进行。" & vbNewLine Then
            MsgBox "没有相关表，无法收集。"
            Exit Sub
        End If

        If MsgBox(strTmp, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then Exit Sub
        
        '开始收集统计信息
        For i = .FixedRows To .Rows - 1
            Call SetCmdEnable(False)
            
            .Select i, 0
            strTbl = .TextMatrix(i, .ColIndex("OWNER")) & "." & .TextMatrix(i, .ColIndex("TABLE_NAME"))
            lblTip2.Caption = "正在收集表" & strTbl & "的统计信息..."
            strSql = "begin " & txtAdv.Text & "end;"
            gcnOracle.Execute strSql
        Next
        
        Call SetCmdEnable(True)
        lblTip2.Caption = "收集统计信息完成！"
    End With
    
    Exit Sub
errH:
    Call SetCmdEnable(True)
    lblTip2.Caption = "收集表" & strTbl & "的统计信息失败！"
    ErrCenter
End Sub

Private Sub cmdFree_Click()
    Dim strTmp As String
    
    If vsfList.Rows = vsfList.FixedRows Or mstrNewSqlId = "" Then
        MsgBox "没有选中SQL语句，无法操作。"
        Exit Sub
    End If
    
    frmEditHint.ShowEdit mstrNewSqlId, Mid(txtFullSql.Text, InStr(1, txtFullSql.Text, vbNewLine) + 2), mintIns_ID, strTmp
    lblTip1.Caption = strTmp
    '刷新列表
    mblnClicked(tab4) = False
End Sub

Private Sub cmdOptExecute_Click()
    Dim strTmp  As String
    
    On Error GoTo errH
    strTmp = UCase(Replace(txtOptExecute.Text, " ", ""))
    
    If strTmp = "" Then
        MsgBox "文本框中没有内容，无法修改。"
        Exit Sub
    End If
    
    If Not strTmp Like "ALTERSYSTEMSET*" Then
        MsgBox "只能修改优化器相关参数，请检查后重新输入。"
        Exit Sub
    End If
    
    strTmp = TrimEx(txtOptExecute.Text)
    gcnOracle.Execute strTmp
        
    LoadParameter
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub Form_load()
    Dim strCol As String
    Dim strSql As String
        
    pctSqlList.Width = 9625
    pctLine.Width = 65
    pctLine.Top = 0
    
    dtpStart.Value = Date: dtpEnd.Value = Date + 1
    dtpStartInterval.Value = Date: dtpEndInterval.Value = Date + 1
    '初始化执行计划表头
    Call InitTable(vsfPlan(1), conCol)
    
    '初始化表统计表头
    strCol = "owner,2000,1;table_name,1500,4;num_rows,1500,1;blocks,1500,1;empty_blocks,1500,1;avg_space,500,1;chain_cnt,500,1;avg_row_len,500,1;Partition_Name,1500,1;Last_Analyzed,1500,1"
    Call InitTable(vsfTblSta, strCol)
    
    '初始化列统计表头
    strCol = "column_name,1500,1;histogram,1500,1;num_buckets,1500,1;last_analyzed,1500,1;num_distinct,1500,1,1500,1;density,1500,1;num_nulls,1500,1;avg_col_len,1500,1"
    Call InitTable(vsfColSta, strCol)
    
    '初始化索引统计表头
    strCol = "index_name,1500,1 ;distinct_keys,1500,1 ;num_rows, 1500,1;clustering_factor, 1500,1;leaf_blocks,1500,1;last_analyzed,1500,1"
    Call InitTable(vsfIdx, strCol)
    
    '初始化相关人员表头
    strCol = "Sid,1500,1 ;Serial#,1500,1 ;姓名, 1500,1;部门, 1500,1;Program,1500,1;Module,1500,1;Sql_Exec_Start,1500,1"
    Call InitTable(vsfUser, strCol)
    
    '初始化涉及报表表头
    strCol = "编号,1500,1;名称,1500,1"
    Call InitTable(vsfReport, strCol)
    
    '初始化SQL PROFILES表头
    strCol = "NAME,1500,1;CATEGORY,1500,1;Flags,1500,1;SIGNATURE,1500,1;CREATED,1500,1;LAST_MODIFIED,1500,1;DESCRIPTION,1500,1; TYPE,1500,1;SQL_TEXT ,1500,1"
    Call InitTable(vsfProfiles, strCol)
    
    '初始化优化器参数表头
    strCol = IIf(gblnRAC, "Inst_ID, 1500,1;", "") & "NAME,1500,1;VALUE,1500,1;DESCRIPTION,1500,1"
    Call InitTable(vsfOpt, strCol)

    If gblnRAC Then
        lblInst.Visible = True
        lblInst.Caption = "当前实例ID：" & gintInstId
    Else
        lblInst.Visible = False
    End If
    
    Call GetOptmizerVision
    
    '初始化SQL语句列表表头
    Call ChangeTableT(1)

    '首次进入，显示信息
    Call ClearVsf(vsfList, "")
    
    LoadParameter
    
    Set mfrmComments = frmComments
    '相关界面效果
    tabPlan.Tabs().Clear
    tabPlan.Tabs().Add 1, , "执行计划"
    Call ClearVsf(vsfPlan(1), "")
    vsfPlan(1).Visible = True
    
    webAwr.Navigate "about:blank"
    
    '设置AWR临时文件保存路径
    mstrFilePath = GetSetting("ZLSOFT", "公共全局", "程序路径", App.Path)
    If mstrFilePath = App.Path Then
        mstrFilePath = mstrFilePath & "\zlsqlawr_temp.html"
    Else
        'C:\APPSOFT\ZLHIS+.exe
        mstrFilePath = Mid(mstrFilePath, 1, InStrRev(mstrFilePath, "\")) & "zlsqlawr_temp.html"
    End If
    
    
    If mlngMinSize = 0 Then
        Call GetMidTabSize(mlngMinSize, mlngMaxSize)
    End If

    Set mrsBigTbl = GetCheckObj(1, mlngMinSize, mlngMaxSize)
    Set mrsBigIdx = GetCheckObj(2, mlngMinSize, mlngMaxSize)
    Set mrsLowIdx = GetCheckObj(3, mlngMinSize, mlngMaxSize)

    sstPlan.TabVisible(tab5) = False
End Sub

Private Sub LoadSqlPlan(vsfPlan As VSFlexGrid, rsPlan As ADODB.Recordset)
'功能：根据传入的数据集和vsfgrid对象，绘制出执行计划
'参数： vsfPlan-需要添加数据的vsfGRID   rsPlan-数据集
    Dim intRowNum As Integer, intVsfIndex As Integer
    
    With vsfPlan
        intVsfIndex = vsfPlan.Index - 1
        .Redraw = flexRDNone
        .FixedCols = 0
        .Editable = flexEDNone
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimpleLeaf
        .SubtotalPosition = flexSTAbove
        .AllowUserResizing = flexResizeColumns
        .Rows = .FixedRows
        .Rows = .FixedRows + rsPlan.RecordCount
        intRowNum = 1
        
        If rsPlan.RecordCount = 0 Then
            Call ClearVsf(vsfPlan, "")
        End If
        Do While Not rsPlan.EOF
                        
            '获取执行计划涉及的表及其所有者，字符串格式如：'table1:owner1,table2:owner2'
            If InStr(rsPlan!Object_Type, "TABLE") > 0 And Trim(rsPlan!Name) <> "" Then
                If mstrTbl_Owner(intVsfIndex) = "" Then
                    mstrTbl_Owner(intVsfIndex) = Trim(rsPlan!Name) & ":" & Trim(rsPlan!Object_Owner)
                Else
                    '判断是否有重复
                    If InStr(1, mstrTbl_Owner(intVsfIndex), Trim(rsPlan!Name) & ":" & Trim(rsPlan!Object_Owner)) = 0 Then
                        mstrTbl_Owner(intVsfIndex) = mstrTbl_Owner(intVsfIndex) & "," & Trim(rsPlan!Name) & ":" & Trim(rsPlan!Object_Owner)
                    End If
                End If
            End If
            
            '获取执行计划涉及的索引及其所有者，字符串格式如：'index1:owner1,index2:owner2'
            If InStr(rsPlan!Object_Type, "INDEX") > 0 And Trim(rsPlan!Name) <> "" Then
                If mstrIdx_Owner(intVsfIndex) = "" Then
                    mstrIdx_Owner(intVsfIndex) = Trim(rsPlan!Name) & ":" & Trim(rsPlan!Object_Owner)
                Else
                    If InStr(1, mstrIdx_Owner(intVsfIndex), Trim(rsPlan!Name) & ":" & Trim(rsPlan!Object_Owner)) = 0 Then
                        mstrIdx_Owner(intVsfIndex) = mstrIdx_Owner(intVsfIndex) & "," & Trim(rsPlan!Name) & ":" & Trim(rsPlan!Object_Owner)
                    End If
                End If
            End If
            
            .TextMatrix(intRowNum, .ColIndex("Operation")) = "" & LTrim(rsPlan!Operation)
            .TextMatrix(intRowNum, .ColIndex("Name")) = "" & rsPlan!Name
            .TextMatrix(intRowNum, .ColIndex("ID")) = "" & rsPlan!Id
            .TextMatrix(intRowNum, .ColIndex("Cardinality")) = "" & rsPlan!Cardinality
            .TextMatrix(intRowNum, .ColIndex("Bytes")) = "" & rsPlan!Bytes
            .TextMatrix(intRowNum, .ColIndex("Cost")) = "" & rsPlan!Cost
            .TextMatrix(intRowNum, .ColIndex("Time")) = "" & rsPlan!Time
            .TextMatrix(intRowNum, .ColIndex("Object_Owner")) = "" & rsPlan!Object_Owner
            .TextMatrix(intRowNum, .ColIndex("Object_Type")) = "" & rsPlan!Object_Type
            
            .RowOutlineLevel(intRowNum) = Len(rsPlan!Operation) - Len(LTrim(rsPlan!Operation)) '以空格个数控制树形结构的等级
            .IsSubtotal(intRowNum) = True
            intRowNum = intRowNum + 1
            rsPlan.MoveNext
        Loop
        .AutoResize = True
        .AutoSize .ColIndex("Operation"), .ColIndex("Object_Owner"), False
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub LoadSqlList(intMod As Integer, Optional strSqlText As String)
'方法用于加载列表
'intMod 取值范围如下
'intMod =1 :加载大表全扫描语句
'intMod =2 :加载AWR中大表全扫描语句
'intMod =3 :加载逻辑读高耗时的SQL语句
'intMod =4 :加载硬解析的SQL语句
'intMod =5 :高频执行的SQL语句
'intMod =6 :加载含有提示字的SQL语句
'intMod =7 :加载查询SQL语句，同时传入需要查找的strSqlText
     
    Dim rsData As ADODB.Recordset, strSql As String
    Dim objVsf As Object, strPsn As String, lngRowNum As Long
    Dim strIdxRange As String, strTblRange As String, strOwnerRange As String
    Dim strQuery As String
    
    On Error GoTo errH
    
    Call ChangeTableT(intMod)
    '判断是否只显示业务系统
    If chkZlhis.Value = 1 Then
       strPsn = " And Parsing_Schema_Name not in ('OGG','SYS','SYSTEM','SCOTT','OUTLN','DBSNMP','MTSSYS','MDSYS','ORDSYS','ORDPLUGINS','CTXSYS','XDB','WMSYS','TSMSYS','SYSMAN','SI_INFORMTN_SCHEMA','OLAPSYS','MGMT_VIEW','MDDATA','EXFSYS','DMSYS','DIP','ANONYMOUS')"
    Else
        strPsn = ""
    End If
    
    '判断是否为ZlHis环境
    If gblnIsZlhis Then
        If gblnHasZltables Then
            '如果有Zltables这张表，则业务数据只需要关联此表
            strTblRange = " p.Object_Name In (Select 表名 From zlTables Where 分类 In ('B1','B2','B3','C1','C2','C3') ) "
            strIdxRange = "p.Object_Name In (Select Index_Name From Dba_Indexes Where Table_Name In (Select 表名 From zlTables where 分类 In ('B1','B2','B3','C1','C2','C3' )))"
        Else
            strTblRange = " p.Object_Name In (Select 表名 From zlBakTables  " & IIf(gblnHasBigtables, "Union All Select 表名 From Zlbigtables ", "") & ")  "
            strIdxRange = "p.Object_Name In((Select Index_Name From Dba_Indexes Where Table_Name In (Select 表名 From zlBakTables " & IIf(gblnHasBigtables, " Union All Select 表名 From Zlbigtables", "") & " )))"
        End If
        strOwnerRange = "p.Object_Owner In (Select 所有者 From zlSystems Union All Select 所有者 From zlbakspaces) And"
    Else
    '非ZLHIS环境
        strTblRange = " p.Object_Name In (Select Table_Name From Dba_Tables Where (Num_Rows Is Null Or Num_Rows > 100000))  "
        strIdxRange = " p.Object_Name In (Select Index_Name From Dba_Indexes Where (Num_Rows Is Null Or Num_Rows > 100000)) "
        strOwnerRange = ""
    End If
    
    strSql = " Select distinct Sql_Id, Executions, Last_Active_Time, Module, Parsing_Schema_Name Schema,Sql_Text , Trunc(Round(t.Buffer_Gets / Decode(t.Executions, 0, 1, t.Executions))) As Per_Buffer_Gets " & IIf(gblnRAC, " ,T.INST_ID", "") & _
               " From " & IIf(gblnRAC, "G", "") & "v$sqlarea T Where " & IIf(chkPLSQL.Value = 1, "", "t.Module Not In ('plsqldev.exe', 'PL/SQL Developer') And")
               
    vsfList.TextMatrix(0, vsfList.ColIndex("Last_Active_Time")) = "Last_Active_Time"
    lblStartTime.Caption = "最后执行时间从"
    Select Case intMod
        Case 1  '加载大表全扫描语句
            strSql = "Select" & vbNewLine & _
                        "      " & IIf(gblnRAC, "T.Inst_ID,", "") & " t.Sql_Id,t.Executions, t.Last_Active_Time, t.Module, t.Parsing_Schema_Name Schema, t.Sql_Text , t.Per_Buffer_Gets," & vbNewLine & _
                        "        f_List2str(Cast(Collect(p.Object_Name) As t_Strlist)) As Object_Name, f_List2str(Cast(Collect( p.Operation || ' ' || p.Options) As t_Strlist)) As Options" & vbNewLine & _
                        "From" & vbNewLine & _
                        "    (Select " & IIf(gblnRAC, "T.Inst_ID,", "") & " t.Sql_Id, t.Executions, t.Last_Active_Time, t.Module, t.Parsing_Schema_Name , t.Sql_Text,t.Hash_Value, Trunc(Round(t.Buffer_Gets / Decode(t.Executions, 0, 1, t.Executions))) As Per_Buffer_Gets " & vbNewLine & _
                        "    From " & IIf(gblnRAC, "G", "") & "V$sqlarea t Where " & IIf(chkPLSQL.Value = 1, "", "t.Module Not In ('plsqldev.exe', 'PL/SQL Developer') And") & "  t.Last_Active_Time between [1] and [2] And t.Executions>0  " & vbNewLine & _
                        strPsn & vbNewLine & _
                                            "        ) T, " & IIf(gblnRAC, "G", "") & "V$sql_Plan P" & vbNewLine & _
                        "Where " & IIf(gblnRAC, "T.Inst_ID = p.Inst_ID And ", "") & " T.Sql_Id = p.Sql_Id And" & vbNewLine & _
                        strOwnerRange & vbNewLine & _
                        "       " & vbNewLine & _
                        "       (( " & strTblRange & " And p.Operation = 'TABLE ACCESS' And p.Options = 'FULL')" & vbNewLine & _
                        "        Or" & vbNewLine & _
                        "        " & vbNewLine & _
                        "       ( " & strIdxRange & vbNewLine & _
                        "        And p.Operation = 'INDEX' And p.Options In ('FAST FULL SCAN', 'FULL SCAN', 'SKIP SCAN')))" & vbNewLine & _
                        "Group By " & IIf(gblnRAC, "T.Inst_ID,", "") & "  t.Sql_Id, t.Executions, t.Last_Active_Time, t.Module, t.Parsing_Schema_Name, t.Sql_Text ,t.Per_Buffer_Gets"

        Case 2  '加载AWR中大表全扫描语句
            vsfList.TextMatrix(0, vsfList.ColIndex("Last_Active_Time")) = "Produced_Time"
            lblStartTime.Caption = "计划产生时间从"
            
            strSql = "Select" & vbNewLine & _
                            "       distinct t.sql_id,t.Inst_Id,t.Executions,t.Schema,t.Sql_Text,t.Module,p.Timestamp Last_Active_Time,T.Per_Buffer_Gets," & vbNewLine & _
                            "       f_List2str(Cast(Collect(p.Object_Name) As t_Strlist)) As Object_Name," & vbNewLine & _
                            "       f_List2str(Cast(Collect(p.Operation || ' ' || p.Options) As t_Strlist)) As Options" & vbNewLine & _
                            "From" & vbNewLine & _
                            "       (Select a.dbid,a.Sql_Id, a.Instance_Number Inst_Id, a.Executions_Total Executions,Trunc(Round(a.buffer_gets_total/decode(a.executions_total,0,1,a.executions_total))) Per_Buffer_Gets," & vbNewLine & _
                            "                     a.Parsing_Schema_Name Schema, To_Char(Dbms_Lob.Substr(b.Sql_Text, 2000)) Sql_Text ,a.Module" & vbNewLine & _
                            "       From Dba_Hist_Sqlstat A, Dba_Hist_Sqltext B" & vbNewLine & _
                            "       Where a.Dbid = b.Dbid And a.Sql_Id = b.Sql_ID And" & vbNewLine & _
                            "       " & IIf(chkPLSQL.Value = 1, "", "a.Module Not In ('plsqldev.exe', 'PL/SQL Developer') And") & "  a.Executions_Total >0" & vbNewLine & _
                            strPsn & vbNewLine & _
                            "       ) T, Dba_Hist_Sql_Plan P" & vbNewLine & _
                            "Where" & vbNewLine & _
                            "      t.Sql_Id = p.Sql_Id And t.dbid = p.dbid And p.Timestamp Between [1] And [2] And" & vbNewLine & _
                            strOwnerRange & vbNewLine & _
                            "       " & vbNewLine & _
                            "       (( " & strTblRange & " And p.Operation = 'TABLE ACCESS' And p.Options = 'FULL')" & vbNewLine & _
                            "        Or" & vbNewLine & _
                            "        " & vbNewLine & _
                            "       ( " & strIdxRange & vbNewLine & _
                            "        And p.Operation = 'INDEX' And p.Options In ('FAST FULL SCAN', 'FULL SCAN', 'SKIP SCAN')))" & vbNewLine & _
                            "Group By t.sql_id,t.Inst_Id,t.Executions,t.Schema,t.Sql_Text,t.Module,p.Timestamp,t.Per_Buffer_Gets"


        Case 3  '加载逻辑读高耗时的SQL语句
            strSql = "Select" & vbNewLine & _
                            "      distinct " & IIf(gblnRAC, "T.Inst_ID,", "") & " t.Sql_Id,t.Executions, t.Last_Active_Time, t.Module, t.Parsing_Schema_Name Schema, t.Sql_Text , Trunc(Round(t.Buffer_Gets / Decode(t.Executions, 0, 1, t.Executions))) As Per_Buffer_Gets" & vbNewLine & _
                            "From" & vbNewLine & _
                            "    (Select " & IIf(gblnRAC, "T.Inst_ID,", "") & " t.Sql_Id, t.Executions, t.Last_Active_Time, t.Module, t.Parsing_Schema_Name , t.Sql_Text,t.Hash_Value,t.BUFFER_GETS,t.Disk_Reads" & vbNewLine & _
                            "    From " & IIf(gblnRAC, "G", "") & "V$sqlarea t Where " & IIf(chkPLSQL.Value = 1, "", "t.Module Not In ('plsqldev.exe', 'PL/SQL Developer') And ") & "  Round(t.Buffer_Gets / Decode(t.Executions, 0, 1, t.Executions)) >" & Val(txtRate.Text) & "  " & vbNewLine & _
                            "     And t.Last_Active_Time Between [1] and [2]" & vbNewLine & _
                            strPsn & vbNewLine & _
                                                "        ) T, " & IIf(gblnRAC, "G", "") & "V$sql_Plan P" & vbNewLine & _
                            "Where " & IIf(gblnRAC, "T.Inst_ID = p.Inst_ID And ", "") & " T.Sql_Id = p.Sql_Id And not (" & vbNewLine & _
                            strOwnerRange & vbNewLine & _
                            "       " & vbNewLine & _
                            "       (( " & strTblRange & " And p.Operation = 'TABLE ACCESS' And p.Options = 'FULL')" & vbNewLine & _
                            "        Or" & vbNewLine & _
                            "        " & vbNewLine & _
                            "       ( " & strIdxRange & vbNewLine & _
                            "        And p.Operation = 'INDEX' And p.Options In ('FAST FULL SCAN', 'FULL SCAN', 'SKIP SCAN'))))"
        Case 4  '加载硬解析的SQL语句
            strSql = strSql + " Force_Matching_Signature <> Exact_Matching_Signature And Last_Active_Time Between [1] And [2] " & vbNewLine & _
                             strPsn & vbNewLine & _
                             "Order By  Last_Active_Time Desc"
                            
        Case 5  '高频执行的SQL语句
            strSql = strSql + " Executions > " & Val(txtRate.Text) & " And Last_Active_Time Between [1] And [2]" & vbNewLine & _
                            strPsn & vbNewLine & _
                            "Order By Last_Active_Time Desc"

        Case 6  '加载含有提示字的SQL语句
            strSql = strSql + "  sql_text like [1]  " & strPsn & " Order By Last_Active_Time Desc"
            strQuery = "%/*+%*/%"
        Case 7  '加载查询SQL语句，同时传入需要查找的strSqlText
            If Len(strSqlText) = 13 And Not IsCharChinese(strSqlText) And InStr(1, ",", strSqlText) = 0 And InStr(1, " ", strSqlText) = 0 And InStr(1, ".", strSqlText) = 0 Then
                strSql = strSql + " sql_id=[1] Order By Last_Active_Time Desc"
                strQuery = Trim(strSqlText)
            Else
                strSql = strSql + "  sql_text like [1]   " & strPsn & "  Order By Last_Active_Time Desc"
                strQuery = "%" & Trim(strSqlText) & "%"
            End If
    End Select
    
    If intMod = 6 Or intMod = 7 Then
        Set rsData = OpenSQLRecord(strSql, Me.Caption, strQuery)
    Else
        Set rsData = OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm")))
    End If
    
    If rsData.RecordCount = 0 Then
        
        Call ClearVsf(vsfList, "")
        Call ClearVsf(vsfTblSta, "")
        Call ClearVsf(vsfColSta, "")
        Call ClearVsf(vsfIdx, "")
        vsfList.Redraw = flexRDNone
        vsfList_AfterRowColChange 0, 0, 0, 0
        vsfList.Tag = intMod
        vsfList.Redraw = flexRDDirect
        Exit Sub
    End If

    With vsfList
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsData.RecordCount
        .Row = 0
        lngRowNum = .FixedRows
        Do While Not rsData.EOF
            .RowHeight(0) = 100
            .TextMatrix(lngRowNum, 0) = lngRowNum
            .TextMatrix(lngRowNum, .ColIndex("Sql_Id")) = rsData!Sql_Id
            .TextMatrix(lngRowNum, .ColIndex("Executions")) = "" & rsData!Executions
            .TextMatrix(lngRowNum, .ColIndex("Module")) = "" & rsData!Module
            .TextMatrix(lngRowNum, .ColIndex("Schema")) = "" & rsData!Schema
            .TextMatrix(lngRowNum, .ColIndex("Last_Active_Time")) = "" & Format(rsData!Last_Active_Time, "yyyy/MM/dd hh:mm:ss")
            .TextMatrix(lngRowNum, .ColIndex("Sql_Text")) = "" & Left(Replace(rsData!sql_text, Chr(10), ""), 60)
            .TextMatrix(lngRowNum, .ColIndex("Per_Buffer_Gets")) = Val(rsData!Per_Buffer_Gets)
            If intMod = 1 Or intMod = 2 Then
                .TextMatrix(lngRowNum, .ColIndex("Object_Name")) = "" & rsData!Object_Name
                .TextMatrix(lngRowNum, .ColIndex("Options")) = "" & rsData!Options
            End If

            If gblnRAC Then
                .RowData(lngRowNum) = "" & rsData!Inst_ID
                .TextMatrix(lngRowNum, .ColIndex("Inst_ID")) = "" & rsData!Inst_ID
            End If
            
            If lngRowNum Mod 2 = 0 Then
                .Cell(flexcpBackColor, lngRowNum, 0, lngRowNum, .Cols - 1) = BackAlterNate_颜色
            Else
                .Cell(flexcpBackColor, lngRowNum, 0, lngRowNum, .Cols - 1) = Back_颜色
            End If
            
            lngRowNum = lngRowNum + 1
            rsData.MoveNext
        Loop
        .AutoResize = True
        .AutoSize 0, .Cols - 1, False
        .ColWidth(.ColIndex("Sql_Text")) = 2625
        .Redraw = flexRDDirect
        .Row = .FixedRows
        .AllowUserResizing = flexResizeColumns
        '获取SQL语句列表的类型存入vsfList的Tag中，用于根据类型进行刷新
        .Tag = intMod
    End With
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
    Call SetCmdEnable(True)
End Sub

Private Sub cmdAllProfiles_Click()
        If cmdAllProfiles.Caption = "显示选中SQL PROFILES" Then
            cmdAllProfiles.Caption = "显示全部SQL PROFILES"
        Else
            cmdAllProfiles.Caption = "显示选中SQL PROFILES"
        End If
        Call RefreshProfiles
End Sub

Private Sub cmdAuto_Click()
    Dim strSql As String, rsData As ADODB.Recordset
    Dim strSqlID As String, strTaskName As String, strQuery As String
    Dim strTmp As String
    
    If vsfList.Rows = vsfList.FixedRows Or mstrNewSqlId = "" Then
        MsgBox "没有选中SQL语句，无法操作。"
        Exit Sub
    End If
    
    strSqlID = mstrNewSqlId
    strTaskName = "ZL_" & strSqlID
    strQuery = "确定要对SQL_ID为 " & strSqlID & " 的SQL语句执行自动优化吗？" & vbNewLine & "该操作可能耗时较长,建议在业务空闲期间运行。"
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then
        Call SetCmdEnable(True)
        Exit Sub
    End If
    
    On Error GoTo errH
    
    lblTip1.Caption = "正在产生优化策略，请耐心等待。"
    strSql = "Select 1 From Dba_Advisor_Tasks Where Task_Name = [1] "
    Set rsData = OpenSQLRecord(strSql, Me.Caption, strTaskName)
    
    If rsData.RecordCount > 0 Then
        gcnOracle.Execute "Dbms_Sqltune.Drop_Tuning_Task('" & strTaskName & "')"
    End If
    
    strSql = "Declare" & vbNewLine & _
                "  v_Sql_Id      V$session.Prev_Sql_Id%Type;" & vbNewLine & _
                "  v_Tuning_Task Varchar2(30);" & vbNewLine & _
                "Begin" & vbNewLine & _
                "  v_Sql_Id      := '" & strSqlID & "';" & vbNewLine & _
                "  v_Tuning_Task := Dbms_Sqltune.Create_Tuning_Task(Sql_Id => v_Sql_Id, Plan_Hash_Value => Null," & vbNewLine & _
                "                                                   Scope => Dbms_Sqltune.Scope_Comprehensive," & vbNewLine & _
                "                                                   Time_Limit => Dbms_Sqltune.Time_Limit_Default," & vbNewLine & _
                "                                                   Task_Name => '" & strTaskName & "', Description => Null);" & vbNewLine & _
                "  Dbms_Sqltune.Execute_Tuning_Task(v_Tuning_Task);" & vbNewLine & _
                "End;"
    gcnOracle.Execute strSql
    
    strSql = "SELECT dbms_sqltune.report_tuning_task([1]) COMMENTS FROM dual"
    Set rsData = OpenSQLRecord(strSql, Me.Caption, strTaskName)
    
    
    
    lblTip1.Caption = "优化策略生成成功。"
    Call SetCmdEnable(True)
    
    
    Call mfrmComments.ShowFrm(rsData!Comments, strTaskName)
    Exit Sub
errH:
    lblTip1.Caption = ""
    If InStr(Err.Description, "ORA-13780") Then
        MsgBox "内存中不存在当前语句，无法优化。"
        Exit Sub
    End If
End Sub

Private Sub cmdDelProfile_Click()
    Dim strSql As String, strProfileName As String, intOldRow As Integer
    Dim strQuery As String
    
    On Error GoTo errH
    strProfileName = vsfProfiles.TextMatrix(vsfProfiles.Row, vsfProfiles.ColIndex("NAME"))
    If strProfileName = "" Or vsfProfiles.Rows = vsfProfiles.FixedRows Then
        MsgBox "当前没有SQL PROFILES，无法删除。"
        Exit Sub
    End If
    
    strQuery = "确定要删除名称为 " & strProfileName & " 的SQL PROFILE吗？" & vbNewLine
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then: Exit Sub
    strSql = "declare" & vbNewLine & _
                "begin" & vbNewLine & _
                "  DBMS_SQLTUNE.drop_sql_profile('" & strProfileName & "');" & vbNewLine & _
                "end;"

    gcnOracle.Execute strSql
    lblTip4.Caption = "删除SQL PROFILE " & strProfileName & "成功！"
    
     
    '将删除的SQL PROFILE从列表移除
    With vsfProfiles
        intOldRow = .Row
        .RemoveItem intOldRow
        If intOldRow >= .Rows - .FixedRows Then '保证选中行的位置不变
            .Select .Rows - .FixedRows, 1
        Else
            .Select intOldRow, 1
        End If
        .TopRow = .Row

        If .Rows = .FixedRows Then '没有数据
            Call ClearVsf(vsfProfiles, "")
        End If
    End With
    Exit Sub
errH:
    Call SetCmdEnable(True)
    
    If InStr(Err.Description, "ORA-22922") Then
        lblTip4.Caption = "数据库中相关数据丢失，删除失败。"
        Exit Sub
    End If
    
    ErrCenter
End Sub

Private Sub cmdRule_Click()
    Dim strGdSQL As String, strBdSQL As String, strQuery As String
    Dim strSqlText, strChild As String, strInstID As String
    
    If vsfList.Rows = vsfList.FixedRows Or mstrNewSqlId = "" Then
        MsgBox "没有选中SQL语句，无法操作。"
        Exit Sub
    End If
    
    Call SetCmdEnable(False)
    strBdSQL = mstrNewSqlId
    
    If cmdRule.Caption = "添加RULE提示" Then
        strQuery = "确定要为SQL_ID为 " & strBdSQL & " 的SQL语句添加RULE提示字吗？" & vbNewLine
    Else
        strQuery = "确定要删除SQL_ID为 " & strBdSQL & " 的SQL语句的RULE提示字吗？" & vbNewLine
    End If
    
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then
        Call SetCmdEnable(True)
        lblTip1.Caption = "操作被取消。"
        Exit Sub
    End If
    strSqlText = TrimEx(Mid(txtFullSql.Text, InStr(1, txtFullSql.Text, vbNewLine) + 2))
    strGdSQL = ChangeSQL(IIf(cmdRule.Caption = "添加RULE提示", 1, 2), strBdSQL, strSqlText, strChild, mintIns_ID)
    
    If strGdSQL = "" Then
        Call SetCmdEnable(True)
        Exit Sub
    End If
    
    If strGdSQL = "5" Or strGdSQL = "1" Or strGdSQL = "2" Then
        lblTip1.Caption = IIf(cmdRule.Caption = "添加RULE提示", "添加", "删除") & "RULE提示失败，建议使用自定义提示操作。"
        Call SetCmdEnable(True)
        Exit Sub
    End If
    
    If CreateSqlProfiles(strBdSQL, strGdSQL, strChild) Then
        lblTip1.Caption = IIf(cmdRule.Caption = "添加RULE提示", "添加RULE提示成功。", "删除RULE提示成功。")
    End If
    
    '刷新列表
    mblnClicked(tab4) = False
    Call SetCmdEnable(True)
End Sub

Private Sub cmdExcel_Click()
    Dim objZlPrint As Object
    Dim objPrint As Object
    
    On Error Resume Next

    Err.Clear: On Error GoTo errH:
    If Not IsInstallExcel Then
        MsgBox "本机未安装Excel。", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    If App.EXEName = "zlSvrStudio.EXE" Then
        Set objZlPrint = CreateObject("zl9PrintMode.zlPrintMethod")
        Set objPrint = CreateObject("zl9PrintMode.zlPrint1Grd")
        If objZlPrint Is Nothing Then
            vsfList.SaveGrid VB.App.Path & "\vsf.xls", flexFileExcel
            MsgBox "保存成功，已经保存至" & VB.App.Path & "\SqlList.xls"
        Else
            Set objPrint.body = vsfList
            objZlPrint.zlPrintOrView1Grd objPrint, 3
        End If
    Else
        vsfList.SaveGrid VB.App.Path & "\vsf.xls", flexFileExcel
        MsgBox "保存成功，已经保存至" & VB.App.Path & "\SqlList.xls"
    End If

    Exit Sub
errH:
    ErrCenter
    Call SetCmdEnable(True)
End Sub

Private Sub cmdExecute_Click()
    Dim strQuery As String, strSql As String, intRow As Integer
    Dim strOwner As String, strTbl As String
    
    strOwner = vsfTblSta.TextMatrix(vsfTblSta.Row, 0)
    strTbl = vsfTblSta.TextMatrix(vsfTblSta.Row, 1)
    
    If strOwner = "没有涉及的表信息" Then
        MsgBox "当前没有选中表，无法收集统计信息。"
        Exit Sub
    End If
    
    If InStr(1, LCase(txtAdv.Text), "dbms_stats.gather_table_stats") = 0 Then
        MsgBox "文本框的内容不包含收集统计信息功能，无法收集统计信息。"
        Exit Sub
    End If
    
    If InStr(1, LCase(txtAdv.Text), LCase(strOwner)) = 0 Or InStr(1, LCase(txtAdv.Text), LCase(strTbl)) = 0 Then
        MsgBox "收集统计信息必须包含所选的Owner、Table_Name，无法收集统计信息。"
        Exit Sub
    End If
    
    strQuery = "你确定要收集" & strOwner & "." & strTbl & "的统计信息吗？" & vbNewLine & "该操作可能耗时较长，并且对系统性能有一定的影响，建议在业务空闲期间运行。"
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then Exit Sub
    
    On Error GoTo errH
    Call SetCmdEnable(False)
    lblTip2.Caption = "正在收集统计信息..."
    strSql = "begin " & txtAdv.Text & "end;"
    gcnOracle.Execute strSql
    lblTip2.Caption = "收集统计信息完成！"
    Call SetCmdEnable(True)
    intRow = vsfTblSta.Row
    Call ReStatTab(True)

    vsfTblSta.Row = intRow
    vsfTblSta.SetFocus
    Exit Sub
errH:
    ErrCenter
    Call SetCmdEnable(True)
End Sub

Private Sub cmdOptmizer_Click()
    Dim strGdSQL As String, strBdSQL As String, strQuery As String
    Dim strSqlText As String, strChild As String

    If vsfList.Rows = vsfList.FixedRows Or mstrNewSqlId = "" Then
        MsgBox "没有选中SQL语句，无法操作。"
        Exit Sub
    End If
    Call SetCmdEnable(False)
    
    strBdSQL = mstrNewSqlId
    
    If cmdOptmizer.Caption = "添加优化器版本提示" Then
reInput:
        strQuery = "确定要为SQL_ID为 " & strBdSQL & " 的SQL语句添加优化器参数提示字吗？" & vbNewLine & vbNewLine & _
                            "当前优化器版本为：" & mstrOptVision & vbNewLine & vbNewLine & _
                            "如果添加提优化器版本错误，请删除对应SQL Profile之后重新添加。"
        strQuery = InputBox(strQuery, "添加优化器版本提示", mstrCompatible, vbYesNo + vbQuestion + vbDefaultButton1)
        
        If strQuery = "" Then
            lblTip1.Caption = "操作被取消。"
            Call SetCmdEnable(True)
            Exit Sub
        End If
        
        If UBound(Split(strQuery, ".")) <> 3 Then
            lblTip1.Caption = "输入有误，添加失败，请重新输入。"
            GoTo reInput
        End If
        
    Else
        strQuery = "确定要删除SQL_ID为 " & strBdSQL & " 的SQL语句的优化器参数提示吗？" & vbNewLine
        If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then
            Call SetCmdEnable(True)
            lblTip1.Caption = "操作被取消。"
            Exit Sub
        End If
    End If
    
    strSqlText = Mid(txtFullSql.Text, InStr(1, txtFullSql.Text, vbNewLine) + 2)
    strGdSQL = ChangeSQL(IIf(cmdOptmizer.Caption = "添加优化器版本提示", 3, 4), strBdSQL, strSqlText, strChild, mintIns_ID, strQuery)
    
    If strGdSQL = "" Then
        Call SetCmdEnable(True)
        Exit Sub
    End If
    
    If strGdSQL = "3" Or strGdSQL = "4" Then
        lblTip1.Caption = "获取语句优化器参数失败，建议使用自定义提示操作。"
        Call SetCmdEnable(True)
        Exit Sub
    End If
    
    If strGdSQL = "5" Then
        lblTip1.Caption = IIf(cmdRule.Caption = "添加优化器版本提示", "添加", "删除") & "优化器版本提示失败，建议使用自定义提示操作。"
        Call SetCmdEnable(True)
        Exit Sub
    End If
    
    If CreateSqlProfiles(strBdSQL, strGdSQL, strChild) Then
        lblTip1.Caption = IIf(cmdOptmizer.Caption = "添加优化器版本提示", "添加", "删除") & "优化器参数提示成功。"
    Else
        lblTip1.Caption = IIf(cmdOptmizer.Caption = "添加优化器版本提示", "添加", "删除") & "优化器参数提示失败，建议使用自定义提示操作。"
    End If
    
    lblTip1.Refresh
    '刷新列表
    mblnClicked(tab4) = False
    Call SetCmdEnable(True)
End Sub

Private Sub cmdRefresh_Click()
    Dim intOldRow As Integer
    
    Call SetCmdEnable(False)
    With vsfList
        intOldRow = .Row
        If .Tag = "" Then
            Call LoadSqlList(1, Trim(txtFind.Text))
            Call SetCmdEnable(True)
            Exit Sub
        Else
            Call LoadSqlList(.Tag, Trim(txtFind.Text))
        End If
        
        If intOldRow >= .Rows - .FixedRows Then '保证选中行的位置不变
            .Select .Rows - .FixedRows, 1
            Call vsfList_AfterRowColChange(vsfList.Row, vsfList.Col, .Rows - .FixedRows, 1)
        Else
            .Select intOldRow, 1
            Call vsfList_AfterRowColChange(vsfList.Row, vsfList.Col, intOldRow, 1)
        End If
        .TopRow = .Row
    End With
    Call SetCmdEnable(True)
End Sub


Private Sub cmdRProfiles_Click()
        Call RefreshProfiles
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    '左边pct
    pctSqlList.Height = Me.ScaleHeight - tabData.Height
    
    tabData.Top = pctSqlList.Height - 15
    tabData.Width = pctSqlList.Width
    
    pctLine.Height = pctSqlList.Height
    pctLine.Left = pctSqlList.Width + pctSqlList.Left
    
    '右边控件
    sstPlan.Height = Me.ScaleHeight
    sstPlan.Width = Abs(Me.ScaleWidth - pctSqlList.Width)
    sstPlan.Left = Abs(pctSqlList.Left + pctSqlList.Width + 82)
    
    Select Case sstPlan.Tab
    Case tab1
        pctPlan.Width = sstPlan.Width
        pctPlan.Height = sstPlan.Height - 400
    Case tab2
        pctStatics.Width = sstPlan.Width
        pctStatics.Height = sstPlan.Height - 400
    Case tab3
        pctUser.Width = sstPlan.Width
        pctUser.Height = sstPlan.Height - 400
    Case tab4
        pctProfiles.Width = sstPlan.Width
        pctProfiles.Height = sstPlan.Height - 400
    Case tab5
        pctAWR.Width = sstPlan.Width
        pctAWR.Height = sstPlan.Height - 400
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '卸载窗体时删除zlsqlawr_temp.html
    On Error Resume Next
    Kill mstrFilePath
End Sub

Private Sub mfrmComments_UpdateStatus(ByVal strStatus As String)
    lblTip1.Caption = strStatus
End Sub

Private Sub optAuto_Click()
    Call ChangeAdvice
End Sub

Private Sub optNull_Click()
    Call ChangeAdvice
End Sub

Private Sub optSkewOnly_Click()
    Call ChangeAdvice
End Sub

Private Sub pctAWR_Resize()
    On Error Resume Next
    webAwr.Width = pctAWR.ScaleWidth
    webAwr.Height = pctAWR.ScaleHeight - webAwr.Top
End Sub

Private Sub pctHorLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objVsf As VSFlexGrid
    Dim intY As Integer, intOldHeight As Integer
    
    If Button <> 1 Then Exit Sub
    '防止拖动过度，导致界面异常
    If pctHorLine.Top + y < 360 Then Exit Sub
    If pctHorLine.Top + y > 10095 Then Exit Sub
    
    intOldHeight = txtFullSql.Height
    pctHorLine.Top = pctHorLine.Top + y
    txtFullSql.Height = Abs(pctHorLine.Top - txtFullSql.Top)

    tabPlan.Top = pctHorLine.Top + 240
    For Each objVsf In vsfPlan
        objVsf.Top = tabPlan.Top + tabPlan.Height
        objVsf.Height = Abs(objVsf.Height - (txtFullSql.Height - intOldHeight))
    Next
    
    Me.Refresh
    
End Sub

Private Sub pctLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If Abs(pctSqlList.Width + x) < 120 Then Exit Sub '左边的pctureBox宽度小于5575会发生显示异常
    
    pctSqlList.Width = Abs(pctSqlList.Width + x)
    tabData.Width = pctSqlList.Width
      
    Call Form_Resize
    
    Me.Refresh
End Sub

Private Sub pctPlan_Resize()
    Dim objVsf As VSFlexGrid
    
    On Error Resume Next
    txtFullSql.Width = Abs(sstPlan.Width - txtFullSql.Left - 265)
    tabPlan.Width = txtFullSql.Width
    For Each objVsf In vsfPlan
        objVsf.Width = txtFullSql.Width
        objVsf.Height = cmdRefresh.Top + cmdRefresh.Height - objVsf.Top - 360
    Next
    cmdAuto.Top = cmdRefresh.Top + cmdRefresh.Height - 315: cmdRule.Top = cmdAuto.Top: cmdOptmizer.Top = cmdAuto.Top: cmdFree.Top = cmdAuto.Top
    cmdAuto.Left = Abs(vsfPlan(1).Width + vsfPlan(1).Left - cmdAuto.Width)
    cmdRule.Left = Abs(cmdAuto.Left - cmdRule.Width - 105)
    cmdOptmizer.Left = Abs(cmdRule.Left - cmdOptmizer.Width - 105)
    cmdFree.Left = Abs(cmdOptmizer.Left - cmdFree.Width - 105)
    
    lblTip1.Top = cmdAuto.Top + lblTip1.Height / 2
    lblTip1.Left = txtFullSql.Left + 60
    
    pctHorLine.Top = txtFullSql.Top + txtFullSql.Height
    pctHorLine.Width = txtFullSql.Width
    pctHorLine.Left = txtFullSql.Left
End Sub


Private Sub pctProfiles_Resize()
    
    On Error Resume Next
    vsfProfiles.Left = lblProfiles.Left
    vsfProfiles.Top = lblProfiles.Top + lblProfiles.Height + 60
    vsfProfiles.Width = pctProfiles.Width - vsfProfiles.Left - 265
    vsfProfiles.Height = pctProfiles.Height / 2
    cmdRProfiles.Top = vsfProfiles.Top + vsfProfiles.Height + 105
    cmdRProfiles.Left = vsfProfiles.Width + vsfProfiles.Left - cmdRProfiles.Width
    cmdDelProfile.Top = cmdRProfiles.Top
    cmdDelProfile.Left = cmdRProfiles.Left - cmdDelProfile.Width - 105
    cmdAllProfiles.Top = cmdRProfiles.Top
    cmdAllProfiles.Left = cmdDelProfile.Left - cmdAllProfiles.Width - 105
    lblTip4.Top = cmdDelProfile.Top + cmdDelProfile.Height / 2 - lblTip4.Height / 2
    lblOpt.Top = cmdDelProfile.Top + cmdDelProfile.Height + 120
    
    vsfOpt.Top = lblOpt.Top + lblOpt.Height + 60
    vsfOpt.Width = vsfProfiles.Width
    vsfOpt.Height = tabData.Top - vsfOpt.Top - txtOptExecute.Height - 220
    
    lblOptExecute.Left = vsfOpt.Left
    txtOptExecute.Top = vsfOpt.Height + vsfOpt.Top + 45
    txtOptExecute.Left = lblOptExecute.Left + lblOptExecute.Width + 45
    lblOptExecute.Top = txtOptExecute.Top + txtOptExecute.Height / 2 - lblOptExecute.Height / 2
    txtOptExecute.Width = vsfOpt.Width - cmdOptExecute.Width - lblOptExecute.Width - 140
    cmdOptExecute.Top = txtOptExecute.Top - 25
    cmdOptExecute.Left = vsfOpt.Left + vsfOpt.Width - cmdOptExecute.Width
End Sub


Private Sub pctSqlList_Resize()

    On Error Resume Next
    cmdRefresh.Top = Abs(pctSqlList.Height - cmdRefresh.Height - 645 + 450)
    dtpEnd.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - dtpEnd.Height / 2: dtpStart.Top = dtpEnd.Top
    lblStartTime.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - lblStartTime.Height / 2: lblEndTime.Top = lblStartTime.Top
    chkZlhis.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - chkZlhis.Height / 2 + 15
    vsfList.Width = Abs(pctSqlList.Width - vsfList.Left)
    cmdRefresh.Left = Abs(vsfList.Width + vsfList.Left - cmdRefresh.Width)
    txtRate.Top = dtpEnd.Top
    lblRate.Top = lblStartTime.Top
    
    '窗体
    dtpEnd.Left = cmdRefresh.Left - dtpEnd.Width - 105
    lblEndTime.Left = dtpEnd.Left - lblEndTime.Width - 45
    
    dtpStart.Left = lblEndTime.Left - dtpStart.Width - 105
    lblStartTime.Left = dtpStart.Left - lblStartTime.Width - 45
    
    txtRate.Left = lblStartTime.Left - txtRate.Width - 75
    lblRate.Left = txtRate.Left - lblRate.Width - 45
    
    chkZlhis.Left = IIf(lblStartTime.Visible, IIf(lblRate.Visible, lblRate.Left, lblStartTime.Left), cmdRefresh.Left) - chkZlhis.Width - 45
    cmdExcel.Left = vsfList.Left
    vsfList.Height = cmdRefresh.Top - vsfList.Top - 105
    
    cmdExcel.Left = vsfList.Left + vsfList.Width - cmdExcel.Width
    txtFind.Left = cmdExcel.Left - txtFind.Width - 105
    lblFind.Left = txtFind.Left - lblFind.Width - 65
    chkPLSQL.Left = lblFind.Left - chkPLSQL.Width - 65
End Sub


Private Sub pctUser_Resize()
    
    On Error Resume Next
    vsfUser.Left = lblUser.Left
    vsfUser.Top = lblUser.Top + lblUser.Height + 60
    vsfUser.Width = pctUser.Width - vsfUser.Left - 265
    vsfUser.Height = pctUser.Height / 10 * 7

    lblReport.Top = vsfUser.Top + vsfUser.Height + 120
    vsfReport.Top = lblReport.Top + lblReport.Height + 60
    vsfReport.Width = vsfUser.Width
    vsfReport.Height = tabData.Top - vsfReport.Top - 220

End Sub


Private Sub pctStatics_Resize()
    
    On Error Resume Next
    txtAdv.Width = pctStatics.Width - txtAdv.Left - 240
    txtAdv.Top = pctStatics.Height - txtAdv.Height - 600
    lblAdv.Left = txtAdv.Left: lblAdv.Top = txtAdv.Top - lblAdv.Height - 60
    vsfTblSta.Width = txtAdv.Width
    vsfTblSta.Height = (lblAdv.Top - lblSTa.Top) / 3 - lblSTa.Height - 285
    vsfTblSta.Top = lblSTa.Top + lblSTa.Height + 60
    lblColSta.Top = vsfTblSta.Height + vsfTblSta.Top + 225
    vsfColSta.Top = lblColSta.Top + lblColSta.Height + 60
    vsfColSta.Width = vsfTblSta.Width / 2 - 30
    vsfColSta.Height = (lblAdv.Top - lblSTa.Top) / 3 * 2.1 - lblSTa.Height - 600
    lblIdx.Top = lblColSta.Top
    lblIdx.Left = vsfColSta.Left + vsfColSta.Width + 60
    vsfIdx.Move lblIdx.Left, vsfColSta.Top, vsfColSta.Width, vsfColSta.Height
    
    cmdExecuteAll.Top = cmdRefresh.Top + cmdRefresh.Height - 315
    cmdExecuteAll.Left = txtAdv.Left + txtAdv.Width - cmdExecuteAll.Width
    cmdExecute.Top = cmdExecuteAll.Top
    cmdExecute.Left = cmdExecuteAll.Left - cmdExecute.Width - 45
    
    optNull.Top = cmdExecute.Top + cmdExecute.Height / 2 - optNull.Height / 2
    optSkewOnly.Top = optNull.Top: optAuto.Top = optNull.Top
    lblType.Top = optNull.Top: lblTip2.Top = optNull.Top
    optNull.Left = cmdExecute.Left - optNull.Width - 45
    optSkewOnly.Left = optNull.Left - optSkewOnly.Width - 45
    optAuto.Left = optSkewOnly.Left - optAuto.Width - 45
    lblType.Left = optAuto.Left - lblType.Width - 45
    
End Sub


Private Sub sstPlan_Click(PreviousTab As Integer)
    
    Screen.MousePointer = vbArrowHourglass
    Call Form_Resize
    Me.Refresh
    Call ReSqlPlanTab
    
    Select Case sstPlan.Tab
    Case tab2
        Call ReStatTab
    Case tab3
        Call ReExecuteTab
    Case tab4
        Call ReProfileTab
    Case tab5
        Call ReAWR
    End Select
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub tabData_Click()

    If Val(vsfList.Tag) = tabData.SelectedItem.Index Then Exit Sub
    vsfList.Tag = tabData.SelectedItem.Index
    
    lblRate.Visible = tabData.SelectedItem.Index = 5 Or tabData.SelectedItem.Index = 3
    txtRate.Visible = lblRate.Visible
    lblRate.Caption = IIf(tabData.SelectedItem.Index = 5, "次数", "逻辑读块数")
    
    sstPlan.TabVisible(4) = tabData.SelectedItem.Index = 2
    Select Case tabData.SelectedItem.Index
        Case 6
            dtpStart.Visible = False: dtpEnd.Visible = False
            lblStartTime.Visible = False: lblEndTime.Visible = False
        Case 1, 2, 3, 4, 5
            dtpStart.Visible = True: dtpEnd.Visible = True
            lblStartTime.Visible = True: lblEndTime.Visible = True
            
    End Select

    pctSqlList_Resize
    
    SetCmdEnable False
    Call LoadSqlList(Val(vsfList.Tag))
    SetCmdEnable True
End Sub

Private Sub tabPlan_Click()
    Dim intPlanNum As Integer, intVsfIndex As Integer
    
    '显示当前选中计划
    If Val(tabPlan.SelectedItem.Index) = Val(tabPlan.Tag) Or tabPlan.Tag = "" Then Exit Sub
    vsfPlan(tabPlan.SelectedItem.Index).Visible = True
    vsfPlan(tabPlan.Tag).Visible = False
    tabPlan.Tag = tabPlan.SelectedItem.Index
    intVsfIndex = tabPlan.SelectedItem.Index
    intPlanNum = tabPlan.SelectedItem.Index - 1 + mintFirPlan
    
    If vsfPlan(tabPlan.SelectedItem.Index).Rows = 1 Then
        Call LoadSqlPlan(vsfPlan(intVsfIndex), GetSQLPlanBySqlID(mstrNewSqlId, intPlanNum))
        Call SetColor
    End If
    
    Call CheckSqlPlan(vsfPlan(intVsfIndex), 0, 1, mrsBigTbl, mrsBigIdx, mrsLowIdx)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       '按下回车
        If Trim(txtFind.Text) = "" Then Exit Sub
        Call LoadSqlList(7, Trim(txtFind.Text))
    End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    Call OnlyIntCK(KeyAscii)
End Sub


Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
       
    If vsfList.Redraw = flexRDNone Then Exit Sub
    For i = 0 To sstPlan.Tabs - 1
        mblnClicked(i) = False
    Next
   
    '获取所选SQL的SQL_ID和实例ID
    If vsfList.TextMatrix(NewRow, vsfList.ColIndex("Sql_ID")) = "" Or vsfList.Rows = vsfList.FixedRows Then
        mstrNewSqlId = ""
    Else
        mstrNewSqlId = vsfList.TextMatrix(NewRow, vsfList.ColIndex("Sql_ID"))
        mintIns_ID = Val(vsfList.RowData(NewRow))
    End If
    
    '刷新任务
    Call sstPlan_Click(1)
    Screen.MousePointer = vbDefault
End Sub


Private Function GetSQLPlanBySqlID(strSqlID As String, intChild_Number As Integer) As ADODB.Recordset
    '功能：根据SQLID取执行计划
    Dim strSql As String, rstmp As ADODB.Recordset
        
    On Error GoTo errH
    If tabData.SelectedItem.Index = 2 Then
        strSql = "Select *" & vbNewLine & _
                        "From (Select /*+ no_merge */" & vbNewLine & _
                        "        LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name, ID, Cardinality, Bytes," & vbNewLine & _
                        "        Cost, Time, Object_Owner, Object_Type" & vbNewLine & _
                        "       From (Select *" & vbNewLine & _
                        "              From Dba_Hist_Sql_Plan" & vbNewLine & _
                        "              Where Sql_Id = [1] And" & vbNewLine & _
                        "                    Plan_Hash_Value = (Select Plan_Hash_Value" & vbNewLine & _
                        "                                       From (Select Plan_Hash_Value, Rownum - 1 Child_Number" & vbNewLine & _
                        "                                              From (Select Distinct Plan_Hash_Value" & vbNewLine & _
                        "                                                     From Dba_Hist_Sqlstat" & vbNewLine & _
                        "                                                     Where Sql_Id = [1] " & vbNewLine & _
                        "                                                     Order By Plan_Hash_Value) A)" & vbNewLine & _
                        "                                       Where Child_Number = [2])) A" & vbNewLine & _
                        "       Start With a.Id = 0" & vbNewLine & _
                        "       Connect By Prior a.Id = a.Parent_Id" & vbNewLine & _
                        "       Order By ID, Position)"
    Else
        '这里v$sql_plan必须要用一个子查询，否则会慢
        '外面要用select *，否则会报ID字段不存在
        strSql = "Select * From (" & vbNewLine & _
                                                    "With A As" & vbNewLine & _
                                                    " (Select Operation, Options, Object_Name, ID, Cardinality, Bytes, Cost, Time, Object_Owner, Object_Type, Position," & vbNewLine & _
                                                    "         Parent_Id" & vbNewLine & _
                                                    "  From " & IIf(gblnRAC And mintIns_ID <> gintInstId, "G", "") & "v$sql_Plan" & vbNewLine & _
                                                    "  Where Sql_Id = [1] And Child_Number = [2] " & IIf(gblnRAC And mintIns_ID <> gintInstId, "And INST_ID = " & mintIns_ID & " ", "") & ")" & vbNewLine & _
                                                    "Select *" & vbNewLine & _
                                                    "From (Select LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name, ID, Cardinality," & vbNewLine & _
                                                    "              Bytes, Cost, Time, Object_Owner, Object_Type" & vbNewLine & _
                                                    "       From A" & vbNewLine & _
                                                    "       Start With a.Id = 0" & vbNewLine & _
                                                    "       Connect By Prior a.Id = a.Parent_Id" & vbNewLine & _
                                                    "       Order By ID, Position))"
    End If
    Set GetSQLPlanBySqlID = OpenSQLRecord(strSql, "GetSQLPlan", strSqlID, intChild_Number)
    Exit Function
errH:
    ErrCenter
    Call SetCmdEnable(True)
End Function


Private Sub GetOptmizerVision()
'功能:获取当前数据库优化器版本和数据库兼容版本
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errH
        strSql = "select NAME,VALUE,DESCRIPTION from v$parameter where (name = 'optimizer_features_enable' or name='compatible')  Order by Name"
        Set rsData = OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount = 0 Then
            Exit Sub
        End If
        
        mstrCompatible = rsData!Value
        rsData.MoveNext
        mstrOptVision = rsData!Value
        
        mstrCompatible = Left(mstrCompatible, Len(mstrOptVision))
    Exit Sub
errH:
    ErrCenter
    Call SetCmdEnable(True)
End Sub


Private Sub vsfOpt_RowColChange()
    Dim strName As String, strValue As String, strType As String
    
    If vsfOpt.Redraw = flexRDNone Then Exit Sub
    strName = vsfOpt.TextMatrix(vsfOpt.Row, vsfOpt.ColIndex("NAME"))
    strValue = vsfOpt.TextMatrix(vsfOpt.Row, vsfOpt.ColIndex("VALUE"))
    strType = vsfOpt.RowData(vsfOpt.Row)
    
    '1 - Boolean 2 - String 3 - Integer 4 - Parameter file 5 - Reserved 6 - Big integer

    Select Case strType
        Case 1, 3, 6
            txtOptExecute.Text = "Alter System Set " & strName & " = " & strValue
        Case 2, 4, 5
            txtOptExecute.Text = "Alter System Set " & strName & " = " & "'" & strValue & "'"
    End Select
    
    
End Sub



Private Sub vsfTblSta_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTblName As String, strSchemaName As String
    
    If vsfTblSta.Rows = 1 Or vsfTblSta.Redraw = flexRDNone Then
        Call ClearVsf(vsfColSta, "")
        Call ClearVsf(vsfIdx, "")
        txtAdv.Text = ""
        Exit Sub
    End If
    
    Call RefreshStatistic
    Call ChangeAdvice

End Sub

Private Sub ChangeAdvice()
    Dim strAdvice As String, strOption As String
    Dim strTblName As String, strSchemaName As String
    
    If vsfTblSta.Rows = 1 Or vsfTblSta.Redraw = flexRDNone Then Exit Sub
    
    lblTip2.Caption = ""
    '生成建议
    With vsfTblSta
        strTblName = .TextMatrix(.Row, .ColIndex("table_name"))
        strSchemaName = .TextMatrix(.Row, .ColIndex("owner"))
    End With
    
    If strTblName = "没有涉及的表信息" Then
        strAdvice = ""
        cmdExecute.Enabled = False
    Else
        cmdExecute.Enabled = True
        
        If optAuto.Value Then
            strOption = " AUTO"
        ElseIf optSkewOnly.Value Then
            strOption = " SKEWONLY"
        Else
            strOption = ""
        End If
    
        strAdvice = "--DEGREE-当前最大并行度可设置为：" & gintCpuMax & ",建议设置为：" & gintCpuAdvise & "；" & vbNewLine & vbNewLine & _
                                "DBMS_STATS.GATHER_TABLE_STATS(OwnName => '" & strSchemaName & "'," & _
                                "TabName => '" & strTblName & "'," & vbNewLine & _
                                "Estimate_Percent => " & IIf(gstrBigVer < 11, 100, "DBMS_STATS.AUTO_SAMPLE_SIZE") & "," & _
                                "Degree => " & gintCpuAdvise & "," & _
                                "Cascade => TRUE," & vbNewLine & _
                                "Method_Opt => 'FOR ALL COLUMNS SIZE" & strOption & "');"
    End If
    
    txtAdv.Text = strAdvice
End Sub

Private Sub LoadVsfGrid(vsfGrid As VSFlexGrid, strSql As String, strInfo As String, ParamArray arrInput() As Variant)
'功能：对不需要修改表头的表格进行初始化，实现代码复用
'传入参数：vsfGrid-需要初始化的表格 ，strSql-查询所用的SQL语句 ，strInfo-查询数据为空时，显示的提示信息 ,strInput -传入参数
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    
    If UBound(arrInput) = 0 Then
        Set rsData = OpenSQLRecord(strSql, Me.Caption, arrInput(0))
    ElseIf UBound(arrInput) = 1 Then
        Set rsData = OpenSQLRecord(strSql, Me.Caption, arrInput(0), arrInput(1))
    ElseIf UBound(arrInput) = 2 Then
        Set rsData = OpenSQLRecord(strSql, Me.Caption, arrInput(0), arrInput(1), arrInput(2))
    Else
        Set rsData = OpenSQLRecord(strSql, Me.Caption)
    End If

    With vsfGrid
        If rsData.RecordCount = 0 Then
            Call ClearVsf(vsfGrid, strInfo)
            Call SetCmdEnable(True)
            Exit Sub
        End If
        .Redraw = flexRDNone
        .Rows = .FixedRows
        Set .DataSource = rsData
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .AutoResize = True
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDDirect
        .Select 1, 1
    End With
    
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
    Call SetCmdEnable(True)
End Sub

Private Sub RefreshProfiles()
'功能： 刷新SQL PROFILES列表
    Dim strSql As String
    
    Call SetCmdEnable(False)
    lblTip4.Caption = ""
    strSql = "Select  Name, Category, Sql_Text, Created, Last_Modified, Type, Status, Force_Matching From Dba_Sql_Profiles"

    If cmdAllProfiles.Caption = "显示全部SQL PROFILES" Then  '根据所选SQL语句刷新SQL概要
        If txtFullSql.Text = "" Then
            Call SetCmdEnable(True)
            Call ClearVsf(vsfProfiles, "")
            Exit Sub
        End If
        
        strSql = strSql + " Where Signature in (Select Force_Matching_Signature From  " & IIf(gblnRAC, "G", "") & "v$SQL Where SQL_ID= [1]) " & vbNewLine & _
                        "Or  Signature in  (Select Exact_Matching_Signature From " & IIf(gblnRAC, "G", "") & "v$SQL Where SQL_ID= [1])  "
       
        Call LoadVsfGrid(vsfProfiles, strSql, "", mstrNewSqlId)
        
    Else    '重新获取全部SQL概要
        Call LoadVsfGrid(vsfProfiles, strSql, "")
    End If
    
    If vsfProfiles.Rows = vsfProfiles.FixedRows Then
        vsfProfiles.Select 0, 0
    Else
        vsfProfiles.Select 1, 0
    End If
    
    Call SetCmdEnable(True)
    
End Sub

Private Sub SetCmdEnable(blnEnable As Boolean)
    Dim objCmd As Object

    For Each objCmd In Me.Controls
        If TypeName(objCmd) = "CommandButton" Then
            objCmd.Enabled = blnEnable
        End If
    Next
    If blnEnable Then
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbArrowHourglass
    End If
End Sub


Private Sub SetColor()
'功能 ：设置SQL语句涉及的表、索引的颜色

    Dim intIdxRow As Integer, intColRow As Integer, strCols As String
    Dim strSql As String, rsData As ADODB.Recordset
    Dim intPlanNum As Integer
    
    '设置索引表格颜色
    On Error GoTo errH
    If tabPlan.Tabs.Count = 0 Then Exit Sub
    
    '清空颜色
    With vsfIdx
        If .Rows <> .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = 0
        End If
    End With
    With vsfColSta
        If .Rows <> .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = 0
        End If
    End With
    
    intPlanNum = tabPlan.SelectedItem.Index - 1
    If mstrIdx_Owner(intPlanNum) = "" Then Exit Sub
    strSql = "Select Index_Owner, Index_Name, Column_Name" & vbNewLine & _
                    "From Dba_Ind_Columns" & vbNewLine & _
                    "Where (Index_Name,Index_Owner) in" & vbNewLine & _
                    " (select /*+ cardinality(a,10)*/ * from table(f_Str2List2([1])) a)"

    Set rsData = OpenSQLRecord(strSql, Me.Caption, Trim(mstrIdx_Owner(intPlanNum)))
    Do While Not rsData.EOF
        strCols = strCols & "'" & rsData!Column_Name & "'-'" & rsData!Index_Name & "',"
        rsData.MoveNext
    Loop
    
    With vsfIdx
        For intIdxRow = 1 To .Rows - 1
            If InStr(1, mstrIdx_Owner(intPlanNum), .TextMatrix(intIdxRow, 0)) > 0 Then
                '设置索引表格颜色
                .Cell(flexcpBackColor, intIdxRow, 0, intIdxRow, .Cols - 1) = Used_颜色
                '设置列表格颜色
                With vsfColSta
                    For intColRow = 1 To .Rows - 1
                        If InStr(1, strCols, "'" & .TextMatrix(intColRow, 0) & "'-'" & vsfIdx.TextMatrix(intIdxRow, 0) & "'") > 0 Then
                            .Cell(flexcpBackColor, intColRow, 0, intColRow, .Cols - 1) = Used_颜色
                        End If
                    Next
                End With
            End If
        Next
    End With
    
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshStatistic()
'功能：刷新所选SQL执行计划对应统计信息
    Dim strTbl As String, strOwner As String, strSql As String
    Dim rstmp As ADODB.Recordset, i As Integer, j As Integer
    
    '加载列统计信息
    strTbl = vsfTblSta.TextMatrix(vsfTblSta.Row, vsfTblSta.ColIndex("table_name"))
    strOwner = vsfTblSta.TextMatrix(vsfTblSta.Row, vsfTblSta.ColIndex("owner"))
    strSql = "select column_name,histogram,num_buckets ,last_analyzed ,num_distinct, Trunc(density,2) density,num_nulls,avg_col_len from dba_tab_col_statistics where table_name = [1] and owner = [2] order by column_name"

    Set rstmp = OpenSQLRecord(strSql, "RefreshStatistic", strTbl, strOwner)
    
    With vsfColSta
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + rstmp.RecordCount
        .ColAlignment(-1) = flexAlignLeftCenter
        i = 0
        Do While Not rstmp.EOF
            For j = 0 To 7
                If .ColIndex("density") = j Then
                    .TextMatrix(i + 1, j) = Format(rstmp.Fields(j).Value, "0.0")
                Else
                    .TextMatrix(i + 1, j) = rstmp.Fields(j).Value
                End If
            Next
            i = i + 1
            rstmp.MoveNext
        Loop
        
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
        
    End With
    
    '加载索引统计信息
     strSql = "select index_name,distinct_keys,num_rows,clustering_factor,leaf_blocks ,last_analyzed from dba_ind_statistics where table_name = [1] and owner = [2] order by index_name"
    Call LoadVsfGrid(vsfIdx, strSql, "", strTbl, strOwner)
    Call SetColor

End Sub


Private Sub ReSqlPlanTab()
'功能：刷新执行计划页卡的内容
    Dim strSql As String, rsData As ADODB.Recordset
    Dim objVsf As VSFlexGrid, strTmp As String
    Dim intPlanNum As Integer
    
    On Error GoTo errH
    '判断SQLID是否发生变化，未发生变化不需要刷新
    If mblnClicked(tab1) = True Then
        If txtFullSql.Text <> "" And tabPlan.Tabs.Count <> 0 Then
            Exit Sub
        End If
    End If
        
    mblnClicked(tab1) = True
    
    '先清空相关信息
    txtFullSql.Text = ""
    tabPlan.Tabs().Clear
    For Each objVsf In vsfPlan
        If objVsf.Index <> 1 Then
            Unload objVsf
        Else
            objVsf.Tag = ""
            objVsf.Rows = objVsf.FixedRows
        End If
    Next
    '换行后将临时存储的统计信息清除
   ReDim mstrTbl_Owner(0): ReDim mstrIdx_Owner(0)
   mintFirPlan = 0
   
    '如果没有相关的SQL语句，退出
    If mstrNewSqlId = "" Then
        If vsfList.Rows = vsfList.FixedRows And tabPlan.Tabs().Count = 0 Then
            tabPlan.Tabs().Add 1, , "执行计划"
            Call ClearVsf(vsfPlan(1), "")
            vsfPlan(1).Visible = True
        End If
        Exit Sub
    End If
    
    If gblnRAC And mintIns_ID = 0 Then
        Exit Sub
    End If
    
    '页卡在第一\三\四页，需要加载SQL文本
    If sstPlan.Tab = tab1 Or sstPlan.Tab = tab3 Or sstPlan.Tab = tab4 Then
    
        If tabData.SelectedItem.Index = 2 Then
            strSql = "Select sql_text Sql_Fulltext  From dba_hist_sqltext  Where Sql_Id = [1]"
        Else
            strSql = "Select Sql_Fulltext  From " & IIf(gblnRAC, "G", "") & "V$sql Where Sql_Id = [1] " & IIf(gblnRAC, "And INST_ID = " & mintIns_ID & " ", "") & ""
        End If
        Set rsData = OpenSQLRecord(strSql, Me.Caption, mstrNewSqlId)
        
        '缓存区内容发生变化，未取到SQL语句
        If rsData.RecordCount = 0 Then
            Exit Sub
        Else
            strTmp = strTmp & rsData!Sql_Fulltext
        End If
        
    
        txtFullSql.Text = "SQLID: " & mstrNewSqlId & IIf(gblnRAC, "  INS_ID:" & vsfList.RowData(vsfList.Row), "") & vbNewLine & strTmp
        
        '修改按钮的Caption
        strTmp = UCase(Replace(strTmp, " ", ""))
        If InStr(1, UCase(strTmp), "/*+RULE*/") > 0 Then
            cmdRule.Caption = "删除RULE提示"
        Else
            cmdRule.Caption = "添加RULE提示"
        End If
        
        If InStr(1, LCase(strTmp), "optimizer_features_enable") > 0 Then
            cmdOptmizer.Caption = "删除优化器版本提示"
        Else
            cmdOptmizer.Caption = "添加优化器版本提示"
        End If
    End If
    
    '页卡在第一页或第二页，需要加载第一个执行计划
    If sstPlan.Tab = tab1 Or sstPlan.Tab = tab2 Then
        '获取执行计划,判断子游标个数,并取出每个子游标对应的执行计划
        If tabData.SelectedItem.Index = 2 Then
            strSql = "Select a.*, Rownum - 1 Child_Number" & vbNewLine & _
                            "From (Select Distinct Plan_Hash_Value From Dba_Hist_Sqlstat Where Sql_Id = [1] Order By Plan_Hash_Value) A"
        Else
            strSql = "select sql_id,child_Number from " & IIf(gblnRAC, "G", "") & "v$sql_plan where sql_id =[1]  " & IIf(gblnRAC, "And INST_ID = " & mintIns_ID & " ", "") & " group by child_number,sql_id Order by child_number"
        End If
        Set rsData = OpenSQLRecord(strSql, Me.Caption, mstrNewSqlId)
        
    
        If rsData.RecordCount > 10 Then
            lblTip1.Caption = "当前SQL语句的执行计划超过10个。"
        Else
            lblTip1.Caption = ""
        End If
    
        If rsData.RecordCount > 0 Then
            mintFirPlan = rsData!child_Number
            ReDim mstrIdx_Owner(rsData.RecordCount - 1)
            ReDim mstrTbl_Owner(rsData.RecordCount - 1)
        Else
            Exit Sub
        End If
        
        Do While Not rsData.EOF
            '添加TAB
            If rsData.RecordCount = 1 Then
                tabPlan.Tabs().Add intPlanNum + 1, , "执行计划"
            Else
                tabPlan.Tabs().Add intPlanNum + 1, , "执行计划" & intPlanNum + 1
            End If

            '添加VSFGRID,Index为1的VSFGRID不用重复加载
            If intPlanNum > 0 Then
                Load vsfPlan(intPlanNum + 1)
                Call InitTable(vsfPlan(intPlanNum + 1), conCol)
            End If
            intPlanNum = intPlanNum + 1
            If intPlanNum = 9 Or intPlanNum = rsData.RecordCount Then Exit Do  '控制最多显示10个子计划
            rsData.MoveNext
        Loop
    
        '加载第一个执行计划,并选中
        Call LoadSqlPlan(vsfPlan(1), GetSQLPlanBySqlID(mstrNewSqlId, mintFirPlan))
        Call CheckSqlPlan(vsfPlan(1), 0, 1, mrsBigTbl, mrsBigIdx, mrsLowIdx)
        
        If tabPlan.Tabs.Count > 0 Then
            tabPlan.Tag = 1
            tabPlan.Tabs(1).Selected = True
            vsfPlan(1).Visible = True
        End If
    End If
    
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub ReStatTab(Optional ByVal blnForceRefresh As Boolean)
'功能：刷新统计信息页卡的内容
'参数：blnForceRefresh: 强制刷新控制，如果为true，那么强制刷新
    Dim strSql As String, intPlanNum As Integer
         
    If mblnClicked(tab2) = True And blnForceRefresh = False Then Exit Sub   '判断是否点击
    
    mblnClicked(tab2) = True

    '没有相关SQL语句或SQL语句不涉及表
    If mstrNewSqlId = "" Or tabPlan.Tabs.Count = 0 Then
        mstrTbl_Owner(0) = "": mstrIdx_Owner(0) = ""
        Call ClearVsf(vsfTblSta, "")
        Call ClearVsf(vsfColSta, "")
        Call ClearVsf(vsfIdx, "")
        Exit Sub
    End If
        
    intPlanNum = tabPlan.SelectedItem.Index - 1
    '含有SQL语句，加载相关信息
    
    
    strSql = "Select  /*+ rule*/Owner , Table_Name, Num_Rows, Blocks, Empty_Blocks, Avg_Space, Chain_Cnt, Avg_Row_Len,Partition_Name, Last_Analyzed" & vbNewLine & _
                    "From Dba_Tab_Statistics" & vbNewLine & _
                    "Where (Table_Name,Owner) In" & vbNewLine & _
                    "  (select /*+ cardinality(a,10)*/ * from table(f_Str2List2([1])) a" & vbNewLine & _
                    "  union all" & vbNewLine & _
                    "  select  table_name, Owner  from  Dba_Indexes where (Index_Name, Owner) In (Select /*+ cardinality(a,10)*/* From Table(f_Str2list2([2])) A))" & vbNewLine & _
                    "Order By Owner, Table_Name"

  
    Call LoadVsfGrid(vsfTblSta, strSql, "", Trim(mstrTbl_Owner(intPlanNum)), Trim(mstrIdx_Owner(intPlanNum)))
   
    
End Sub

Private Sub ReExecuteTab()
'功能：刷新执行信息页卡的内容
    Dim strSql As String, blnIsAdmin As Boolean, rsData As ADODB.Recordset
    Dim strDetialSql  As String, strSqlText As String
    
    
    If mblnClicked(tab3) = True Then Exit Sub  '判断SQLID是否发生变化，未发生变化不需要刷新
    mblnClicked(tab3) = True
    
    Screen.MousePointer = vbArrowHourglass
    If mstrNewSqlId = "" Then
        Call ClearVsf(vsfUser, "")
        Call ClearVsf(vsfReport, "")
    Else
    
        '加载相关人员
        strSql = "  , (Select b.姓名, d.名称 As 部门, a.用户名" & vbNewLine & _
                        " From 上机人员表 A, 人员表 B, 部门人员 C, 部门表 D" & vbNewLine & _
                        " Where a.人员id = b.Id And b.Id = c.人员id And c.缺省 = 1 And c.部门id = d.Id) C "

        strSql = "Select  b.Sid, b.Serial#, " & IIf(gblnIsZlhis, "c.姓名, c.部门, ", "") & " a.Program, a.Module, b.Username," & IIf(gstrBigVer < 11, "max(a.SAMPLE_TIME) Sql_Exec_Start", " a.Machine, max(a.Sql_Exec_Start)  Sql_Exec_Start") & vbNewLine & _
                "From Dba_Hist_Active_Sess_History A, " & IIf(gblnRAC, "G", "") & "v$session B" & vbNewLine & _
                IIf(gblnIsZlhis, strSql, "") & _
                "Where a.Sql_Id =[1] And a.Session_Id = b.Sid And a.Session_Serial# = b.Serial#  " & IIf(gblnIsZlhis, "And b.Username = c.用户名(+)", "") & vbNewLine & _
                " Group By b.Sid, b.Serial#, " & IIf(gblnIsZlhis, "c.姓名, c.部门, ", "") & "  a.Program, a.Module, b.Username " & IIf(gstrBigVer < 11, "", " , a.Machine")

        strSql = "Select Sid, Serial#, " & IIf(gblnIsZlhis, "姓名, 部门, ", "") & "  Program, Module, Username " & IIf(gstrBigVer < 11, "", " , Machine") & ",max(Sql_Exec_Start) from ( " & vbNewLine & _
                        strSql & vbNewLine & _
                        "Union" & vbNewLine & _
                        Replace(strSql, "Dba_Hist_Active_Sess_History", IIf(gblnRAC, "G", "") & "V$active_Session_History") & vbNewLine & _
                        " ) Group By Sid, Serial#, " & IIf(gblnIsZlhis, "姓名, 部门, ", "") & "  Program, Module, Username " & IIf(gstrBigVer < 11, "", " , Machine")
                
        Call LoadVsfGrid(vsfUser, strSql, "", mstrNewSqlId)
    End If
    
    '加载Zlsystems中的所有者
    If gblnIsZlhis Then
        If mrsAdmins Is Nothing Then
            strSql = "Select 所有者" & vbNewLine & _
                            "From zlSystems" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select 所有者 From zlBakspaces"
            Set mrsAdmins = OpenSQLRecord(strSql, Me.Caption)
        Else
            mrsAdmins.MoveFirst
        End If
    Else
        Call ClearVsf(vsfReport, "")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    '判断当前语句是否含有 “所有者字段.”
    strSqlText = Mid(txtFullSql.Text, InStr(1, txtFullSql.Text, vbNewLine) + 2)
    blnIsAdmin = False: mrsAdmins.MoveFirst
    Do While Not mrsAdmins.EOF
        If InStr(1, strSqlText, mrsAdmins!所有者 & ".") > 0 Then
            blnIsAdmin = True
            strSqlText = Replace(strSqlText, mrsAdmins!所有者 & ".", "")
        End If
        mrsAdmins.MoveNext
    Loop

    If blnIsAdmin Then
        
        Screen.MousePointer = vbHourglass
        
        '替换绑定变量为空，用于匹配明细SQL和分类SQL
        strDetialSql = Replace(Replace(strSqlText, " ", ""), Chr(10), "")
        strSql = "select POSITION,NAME,VALUE_STRING ,last_captured ,DataType from " & IIf(gblnRAC, "G", "") & "v$sql_bind_capture where  SQL_ID= [1]  " & _
                 "and CHILD_NUMBER in (select max(CHILD_NUMBER)  from " & IIf(gblnRAC, "G", "") & "v$sql_bind_capture where SQL_ID= [1]  " & IIf(gblnRAC, "And INST_ID = " & mintIns_ID & " ", "") & "  ) order by POSITION"
        Set rsData = OpenSQLRecord(strSql, Me.Caption, mstrNewSqlId)
        
        Do While Not rsData.EOF
            strDetialSql = Replace(strDetialSql, rsData!Name, "", 1, 1)
            rsData.MoveNext
        Loop
        strDetialSql = Left(strDetialSql, 1000)
        
        '去掉换行与空格
        strSqlText = Replace(Replace(strSqlText, " ", ""), Chr(10), "")
        '只取语句中Where之前的内容,用于匹配内容SQL
        If InStr(1, UCase(strSqlText), "WHERE") = 0 Then
            strSqlText = Left(strSqlText, 1000)
        Else
            strSqlText = Left(Left(strSqlText, InStr(1, UCase(strSqlText), "WHERE") - 1), 1000)
        End If
        
        
        '替换后进行查找
        strSql = "Select Distinct a.编号, a.名称" & vbNewLine & _
                        "From zlReports A, zlRPTDatas B," & vbNewLine & _
                        "     (Select Replace(f_List2str(Cast(Collect(内容 order by 行号) As t_Strlist), '', 1, 1000), ' ', '') 内容, 源id" & vbNewLine & _
                        "       From (Select 源id, replace(replace(内容,' ',''),chr(9),'') 内容, 行号 From zlRPTSQLs where substr(trim(内容),1,2) <> '--' )" & vbNewLine & _
                        "       Group By 源id ) C" & vbNewLine & _
                        "Where c.源id = b.Id And b.报表id = a.Id And substr(c.内容,1,instr(Upper(C.内容),'WHERE')-1) = [1]" & vbNewLine & _
                        "Union" & vbNewLine & _
                        "Select Distinct a.编号, a.名称" & vbNewLine & _
                        "From zlReports A, zlRPTDatas B," & vbNewLine & _
                        "     (Select Replace(Translate(Replace(Replace(分类sql, ' ', ''), Chr(9), ''), '[#0123456789', '[#'), '[]', '') 分类sql," & vbNewLine & _
                        "              Replace(Translate(Replace(Replace(明细sql, ' ', ''), Chr(9), ''), '[#0123456789', '[#'), '[]', '') 明细sql, 源id" & vbNewLine & _
                        "       From zlRPTPars where  分类sql is not null or 明细sql is not null ) C" & vbNewLine & _
                        "Where c.源id = b.Id And b.报表id = a.Id And (c.明细sql = [2] Or c.分类sql = [2])"


                        
        Call LoadVsfGrid(vsfReport, strSql, "", strSqlText, strDetialSql)
    Else
        Call ClearVsf(vsfReport, "")
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub ReProfileTab()
'功能：刷新SQLProfile页卡的内容
    Dim strSql As String
    
    lblTip4.Caption = ""
        
    If mblnClicked(tab4) = True Then Exit Sub  '判断是否需要刷新
    mblnClicked(tab4) = True
    
    If cmdAllProfiles.Caption = "显示选中SQL PROFILES" Then Exit Sub '显示全部SQLProfile的情况下不需要刷新
    
    If mstrNewSqlId = "" Then
        Call ClearVsf(vsfProfiles, "")
        Exit Sub
    End If
    Call RefreshProfiles
End Sub

Private Function CheckRepeatation(arrInput() As String, strText As String) As Boolean
    '功能：判断数组中是否存在某一字符串，若存在返回True
    '参数：arrInput-需要查找的数组；strText-进行匹配的字段
    Dim intNum As Integer, blnResult As Boolean
    
    blnResult = False
    intNum = 0
    
    Do While intNum <= UBound(arrInput)
        If (arrInput(intNum) = strText) Then
             blnResult = True
             Exit Do
        End If
        intNum = intNum + 1
    Loop
    
    CheckRepeatation = blnResult
End Function

Private Sub ReAWR()
    Dim strSql As String, rstmp As ADODB.Recordset
    Dim lngBid As Long, lngEid As Long
    Dim i As Long
    Dim objFile As New FileSystemObject, strHtml As String
    
    If mblnClicked(tab5) = True Or tabData.SelectedItem.Index <> 2 Then Exit Sub    '判断是否需要刷新
    mblnClicked(tab5) = True
    
    If mstrNewSqlId = "" Then
        webAwr.Navigate "about:blank"
        Exit Sub
    End If
    
    On Error GoTo errH
    strSql = "Select Distinct a.Dbid, A. Instance_Number, a.Startup_Time," & vbNewLine & _
                    "                First_Value(a.Snap_Id) Over(Partition By a.Startup_Time Order By a.Startup_Time Desc) Bid," & vbNewLine & _
                    "                Last_Value(a.Snap_Id) Over(Partition By a.Startup_Time Order By a.Startup_Time Desc) Eid" & vbNewLine & _
                    "From Dba_Hist_Snapshot A, Dba_Hist_Sqlstat B" & vbNewLine & _
                    "Where a.Dbid = b.Dbid And a.Snap_Id = b.Snap_Id And b.Sql_Id = [1]" & vbNewLine & _
                    "Order By a.Startup_Time Desc"


    Set rstmp = OpenSQLRecord(strSql, "ReAWR", mstrNewSqlId)
    If rstmp.RecordCount = 0 Then Exit Sub
    
    lngBid = Val(rstmp!Bid)
    lngEid = Val(rstmp!eid)
    
    If lngBid = lngEid Then
        lngBid = lngBid - 1
    End If
    
    strSql = "Select Output From Table(Dbms_Workload_Repository.Awr_Sql_Report_Html([1], [2], [3],[4],[5])) A "
    Set rstmp = OpenSQLRecord(strSql, "ReAWR", Val(rstmp!DBID), Val(rstmp!Instance_Number), lngBid, lngEid, mstrNewSqlId)
    
    Do While Not rstmp.EOF
        strHtml = strHtml & rstmp!OutPut & ""
        rstmp.MoveNext
    Loop
    
    objFile.CreateTextFile(mstrFilePath).Write strHtml
    webAwr.Navigate "file:///" & Replace(mstrFilePath, "\", "/")
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub GetAWRByTime()
    '功能：根据修改的快照时间范围，获取AWR
    Dim strSql As String, rstmp As ADODB.Recordset
    Dim strHtml As String, objFile As New FileSystemObject
    Dim lngBid As Long, lngEid As Long
    
    If tabData.SelectedItem.Index <> 2 Then Exit Sub
    If dtpStartInterval.Value = "" Or dtpEndInterval.Value = "" Then Exit Sub
    
    On Error GoTo errH
    
    strSql = "Select Distinct a.Dbid, A. Instance_Number, a.Startup_Time," & vbNewLine & _
                    "                First_Value(a.Snap_Id) Over(Partition By a.Startup_Time Order By a.Startup_Time Desc) Bid," & vbNewLine & _
                    "                Last_Value(a.Snap_Id) Over(Partition By a.Startup_Time Order By a.Startup_Time Desc) Eid" & vbNewLine & _
                    "From Dba_Hist_Snapshot A, Dba_Hist_Sqlstat B" & vbNewLine & _
                    "Where a.Dbid = b.Dbid And a.Snap_Id = b.Snap_Id And b.Sql_Id = [1]" & vbNewLine & _
                    "And a.begin_interval_time Between [2] And [3]" & vbNewLine & _
                    "Order By a.Startup_Time Desc"

    Set rstmp = OpenSQLRecord(strSql, "GetSnapID", mstrNewSqlId, _
                            CDate(Format(dtpStartInterval.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(dtpEndInterval.Value, "yyyy-MM-dd hh:mm:ss")))
    
    If rstmp.RecordCount = 0 Then Exit Sub
    
    If rstmp!Bid = rstmp!eid Then
        lngEid = Val(rstmp!eid) + 1
    Else
        lngEid = Val(rstmp!eid)
    End If
    lngBid = Val(rstmp!Bid)
    
    strSql = "Select Output From Table(Dbms_Workload_Repository.Awr_Sql_Report_Html([1], [2], [3],[4],[5])) A "
    Set rstmp = OpenSQLRecord(strSql, "GetAWRByTime", rstmp!DBID, rstmp!Instance_Number, lngBid, lngEid, mstrNewSqlId)
    
    If rstmp.RecordCount = 0 Then Exit Sub
    
    Do While Not rstmp.EOF
        strHtml = strHtml & rstmp!OutPut & ""
        rstmp.MoveNext
    Loop
    
    objFile.CreateTextFile(mstrFilePath).Write strHtml
    webAwr.Navigate "file:///" & Replace(mstrFilePath, "\", "/")
    
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
    
End Sub

Private Sub LoadParameter()
    '加载参数
    Dim strSql As String, rstmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    'strCol = "NAME,1500,1;VALUE,1500,1;DESCRIPTION,1500,1"
    strSql = "select  " & IIf(gblnRAC, "Inst_ID,", "") & "  NAME,VALUE,DESCRIPTION ,TYPE from " & IIf(gblnRAC, "G", "") & "v$parameter where " & _
                    "(name like 'optimizer%' or name='db_file_multiblock_read_count') Order by 1,2"
                        
    Set rstmp = OpenSQLRecord(strSql, "LoadParameter")

    With vsfOpt
        
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + rstmp.RecordCount
        
        i = 1
        Do While Not rstmp.EOF
            If gblnRAC Then
                .TextMatrix(i, .ColIndex("Inst_ID")) = rstmp!Inst_ID
            End If
            
            .TextMatrix(i, .ColIndex("NAME")) = "" & rstmp!Name
            .TextMatrix(i, .ColIndex("VALUE")) = "" & rstmp!Value
            .TextMatrix(i, .ColIndex("DESCRIPTION")) = "" & rstmp!Description
            .RowData(i) = Val(rstmp!Type)
            rstmp.MoveNext
            i = i + 1
        Loop
        
        .AutoResize = True
        .AutoSize 0, .Cols - 1
        
        .Redraw = flexRDDirect
        On Error Resume Next
        .Select 1, 0
    End With
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub ChangeTableT(ByVal intMod As Integer)
    Dim strCol As String
    
    '重新初始化SQL语句列表表头
    If intMod = 1 Or intMod = 2 Then
        strCol = "Rows,1500,1;" & IIf(gblnRAC, "Inst_ID, 1500,1;", "") & "Sql_Id, 1500,1;Module, 1500,1;Schema,1500,1;Sql_Text,1500,1;" & _
                    "Per_Buffer_Gets,1500,1;Executions,1500,1;Object_Name,1500,1;Options,1500,1;Last_Active_Time,1500,1"
    Else
        strCol = "Rows,1500,1;" & IIf(gblnRAC, "Inst_ID, 1500,1;", "") & "Sql_Id, 1500,1;Module, 1500,1;Schema,1500,1;Sql_Text,1500,1;Per_Buffer_Gets,1500,1;Executions,1500,1;Last_Active_Time,1500,1"
    End If
    
    Call InitTable(vsfList, strCol)
    With vsfList
        .Redraw = flexRDNone
        vsfList.FixedCols = 1
        .Redraw = flexRDDirect
    End With
    
    
End Sub
