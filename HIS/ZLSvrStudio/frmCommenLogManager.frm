VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCommenLogManager 
   BackColor       =   &H80000005&
   Caption         =   "通用日志管理"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmCommenLogManager.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   10740
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstPanel 
      Height          =   6495
      Left            =   1560
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmCommenLogManager.frx":803A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picPage(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmCommenLogManager.frx":8056
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPage(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmCommenLogManager.frx":8072
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PicFind"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   5820
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   5820
         ScaleWidth      =   8655
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   360
         Width           =   8655
         Begin VB.Frame fraLogSetDetailCtrl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3600
            TabIndex        =   48
            Top             =   4920
            Width           =   5055
            Begin VB.CommandButton cmdLogSetCtrlAdd 
               Caption         =   "新增规则(&B)"
               Height          =   375
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Width           =   1575
            End
            Begin VB.CommandButton cmdLogSetCtrlEdit 
               Caption         =   "修改规则(&M)"
               Height          =   375
               Left            =   1680
               TabIndex        =   50
               Top             =   0
               Width           =   1575
            End
            Begin VB.CommandButton cmdLogSetCtrlDel 
               Caption         =   "删除规则(&R)"
               Height          =   375
               Left            =   3360
               TabIndex        =   49
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame fraLogSetCtrl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3600
            TabIndex        =   44
            Top             =   2520
            Width           =   5055
            Begin VB.CommandButton cmdLogSetDel 
               Caption         =   "删除分类"
               Height          =   375
               Left            =   3720
               TabIndex        =   47
               Top             =   0
               Width           =   1100
            End
            Begin VB.CommandButton cmdLogSetEdit 
               Caption         =   "修改分类"
               Height          =   375
               Left            =   2520
               TabIndex        =   46
               Top             =   0
               Width           =   1100
            End
            Begin VB.CommandButton cmdLogSetAdd 
               Caption         =   "新增分类"
               Height          =   375
               Left            =   1320
               TabIndex        =   45
               ToolTipText     =   "新增的日志属于自定义类型，需配合在应用程序中调用日志相关方法，一般用于插件等二次开发"
               Top             =   0
               Width           =   1100
            End
         End
         Begin VB.CommandButton cmdLogServer 
            Caption         =   "日志服务器设置"
            Height          =   375
            Left            =   6480
            TabIndex        =   43
            Top             =   120
            Width           =   1935
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfLogSetDetail 
            Height          =   1680
            Left            =   0
            TabIndex        =   52
            Top             =   3120
            Width           =   8490
            _cx             =   14975
            _cy             =   2963
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
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmCommenLogManager.frx":808E
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
         Begin VSFlex8Ctl.VSFlexGrid vsfLogSet 
            Height          =   1800
            Left            =   0
            TabIndex        =   53
            Top             =   600
            Width           =   8490
            _cx             =   14975
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
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmCommenLogManager.frx":81E2
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
         Begin VB.Label lblLogSetDetail 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "日志记录规则"
            Height          =   180
            Left            =   0
            TabIndex        =   55
            Top             =   2880
            Width           =   1080
         End
         Begin VB.Label lblLogSet 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "日志分类"
            Height          =   180
            Left            =   0
            TabIndex        =   54
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   3660
         Index           =   0
         Left            =   120
         ScaleHeight     =   3660
         ScaleWidth      =   7395
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Width           =   7395
         Begin VB.CheckBox chkAllCols 
            BackColor       =   &H80000005&
            Caption         =   "显示所有列"
            Height          =   180
            Left            =   0
            TabIndex        =   39
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "查找"
            Height          =   345
            Left            =   240
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfLogInfo 
            Height          =   1920
            Left            =   0
            TabIndex        =   40
            Top             =   840
            Width           =   7050
            _cx             =   12435
            _cy             =   3387
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
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   300
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmCommenLogManager.frx":838D
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            ExplorerBar     =   5
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
         Begin VB.Label lblFindDetail 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            Height          =   180
            Left            =   1800
            TabIndex        =   41
            Top             =   210
            Width           =   90
         End
      End
      Begin VB.PictureBox PicFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   -74880
         ScaleHeight     =   4185
         ScaleWidth      =   6465
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   6495
         Begin VB.Frame fraFind 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   4095
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   6375
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   8
               Left            =   4080
               TabIndex        =   25
               Tag             =   "Category_Name"
               Top             =   3150
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   7
               Left            =   960
               TabIndex        =   23
               Tag             =   "Category_Name"
               Top             =   3090
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   6
               Left            =   4080
               TabIndex        =   21
               Tag             =   "Category_Name"
               Top             =   2670
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   5
               Left            =   960
               TabIndex        =   19
               Tag             =   "Category_Name"
               Top             =   2670
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   4
               Left            =   960
               TabIndex        =   17
               Tag             =   "Category_Name"
               Top             =   2250
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   3
               Left            =   4080
               TabIndex        =   15
               Tag             =   "Category_Name"
               Top             =   1830
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   2
               Left            =   960
               TabIndex        =   13
               Tag             =   "Category_Name"
               Top             =   1830
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   1
               Left            =   4080
               TabIndex        =   11
               Tag             =   "Category_Name"
               Top             =   1350
               Width           =   2145
            End
            Begin VB.ComboBox cboGroup 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   9
               Tag             =   "Category_Name"
               Top             =   1320
               Width           =   2145
            End
            Begin VB.CommandButton cmdFindCancel 
               Cancel          =   -1  'True
               Caption         =   "取消(&C)"
               Height          =   350
               Left            =   5125
               TabIndex        =   28
               Top             =   3720
               Width           =   1100
            End
            Begin VB.CommandButton cmdFindOK 
               Caption         =   "确定(&O)"
               Height          =   350
               Left            =   3960
               TabIndex        =   27
               Top             =   3720
               Width           =   1100
            End
            Begin VB.ComboBox cboFindLevel 
               Height          =   300
               Left            =   4080
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Tag             =   "Log_Level"
               Top             =   915
               Width           =   2145
            End
            Begin VB.ComboBox cboFindCategory 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Tag             =   "Category_Name"
               Top             =   915
               Width           =   2145
            End
            Begin VB.CommandButton cmdFindReset 
               Caption         =   "重置条件"
               Height          =   350
               Left            =   960
               TabIndex        =   26
               Top             =   3720
               Width           =   1100
            End
            Begin VB.Frame FraHead 
               BackColor       =   &H80000005&
               Height          =   405
               Left            =   30
               TabIndex        =   34
               Top             =   30
               Width           =   6315
               Begin VB.PictureBox PicClose 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   200
                  Left            =   6000
                  Picture         =   "frmCommenLogManager.frx":855C
                  ScaleHeight     =   195
                  ScaleWidth      =   210
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   215
               End
               Begin VB.Label lblFindHead 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "条件设置"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Left            =   90
                  TabIndex        =   36
                  Top             =   160
                  Width           =   720
               End
            End
            Begin MSComCtl2.DTPicker dtpFindDate 
               Height          =   315
               Index           =   0
               Left            =   960
               TabIndex        =   1
               Top             =   480
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
               Format          =   204210179
               CurrentDate     =   43466
               MinDate         =   43466
            End
            Begin MSComCtl2.DTPicker dtpFindDate 
               Height          =   315
               Index           =   1
               Left            =   4080
               TabIndex        =   3
               Top             =   480
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
               Format          =   204210179
               CurrentDate     =   43466.0416666667
               MinDate         =   43466
            End
            Begin VB.Line linSplit 
               BorderColor     =   &H80000000&
               Index           =   5
               X1              =   0
               X2              =   6720
               Y1              =   3600
               Y2              =   3600
            End
            Begin VB.Label lblFindDateTo 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "至"
               Height          =   180
               Left            =   3502
               TabIndex        =   2
               Top             =   547
               Width           =   180
            End
            Begin VB.Label lblFindDate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "登记时间"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   0
               Top             =   540
               Width           =   720
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "服务名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   8
               Left            =   3420
               TabIndex        =   24
               Top             =   3210
               Width           =   540
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "功能"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   7
               Left            =   480
               TabIndex        =   22
               Top             =   3150
               Width           =   360
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "模块"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   6
               Left            =   3600
               TabIndex        =   20
               Top             =   2715
               Width           =   360
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "部件"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   480
               TabIndex        =   18
               Top             =   2730
               Width           =   360
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "进程名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   300
               TabIndex        =   16
               Top             =   2310
               Width           =   540
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "工作站"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   3420
               TabIndex        =   10
               Top             =   1395
               Width           =   540
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "IP"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   3780
               TabIndex        =   14
               Top             =   1890
               Width           =   180
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "会话ID"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   300
               TabIndex        =   12
               Top             =   1890
               Width           =   540
            End
            Begin VB.Label lblFindInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "用户名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   8
               Top             =   1395
               Width           =   540
            End
            Begin VB.Label lblFindLevel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "日志级别"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   3240
               TabIndex        =   6
               Top             =   976
               Width           =   720
            End
            Begin VB.Label lblFindCategory 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "日志类别"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   4
               Top             =   990
               Width           =   720
            End
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   840
      Left            =   60
      TabIndex        =   30
      Top             =   510
      Width           =   780
      _Version        =   589884
      _ExtentX        =   1376
      _ExtentY        =   1482
      _StockProps     =   64
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "通用日志管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   29
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmCommenLogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PageEnum
    PE_日志分类 = 0
    PE_记录规则 = 1
    PE_查找选项 = 9
End Enum
Private Type MousePoint
    x As Single
    y As Single
End Type
Private Type WindowRect
    Left As Single
    Top As Single
End Type

Private Enum FindInfo
    FI_User = 0
    FI_Station
    FI_SID
    FI_IP
    FI_Process
    FI_Component
    FI_Module
    FI_Function
    FI_Call
End Enum

Private Enum LogSetCol
    LSC_ID = 0
    LSC_Name
    LSC_Builtin
    LSC_Begin_Time
    LSC_End_Time
    LSC_Log_Keep_Days
    LSC_Log_Mode
    LSC_Log_Level
    LSC_Description
End Enum

Private Enum LogSetDetailCol
    LSDC_ID = 0
    LSDC_Condition
    LSDC_Component_Names
    LSDC_Module_Names
    LSDC_Function_Names
    LSDC_Call_Names
End Enum

Private Const MSTR_LOGINFO As String = _
    "登记时间,,3,1900,dt|ID,,0,0,n|日志级别,,3,900|日志分类,,3,1500|服务名,,3,2000|状态,,3,540|IP,,3,1500|工作站,,3,1000|" & _
    "进程ID,,3,1000,n|进程名,,3,1000|服务器,,3,1500|部件,,3,1200|模块,,3,1200|功能,,3,1200|用户,,3,800|会话ID,,3,800,n|" & _
    "内容,,3,2500"

Private mcnOracle As ADODB.Connection             '日志库连接
Private mcurMousePoint As MousePoint
Private mcurWindowRect As WindowRect
Private mblnIgnoreWarn As Boolean
Private WithEvents mobjLogInfo As clsVSFlexGridEx
Attribute mobjLogInfo.VB_VarHelpID = -1

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
    
    Dim objPrint As zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    Set objPrint = New zlPrint1Grd
    If Me.ActiveControl = vsfLogSet Then
        objPrint.Title.Text = "日志分类"
        Set objPrint.Body = vsfLogSet
    ElseIf Me.ActiveControl = vsfLogSetDetail Then
        objPrint.Title.Text = "日志规则"
        Set objPrint.Body = vsfLogSetDetail
    Else
        objPrint.Title.Text = "通用日志"
        Set objPrint.Body = vsfLogInfo
    End If
    Set objRow = New zlTabAppRow
    objRow.Add "时间：" & Format(CurrentDate(), "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add objRow
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
              zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Form_Load()
    Set mcnOracle = ReGetConnection
    
    Call ResetContainer
    Call LoadLogList
    
    '日志等级
    cboFindLevel.addItem ""
    cboFindLevel.addItem "1-错误"
    cboFindLevel.addItem "2-警告"
    cboFindLevel.addItem "3-重要"
    cboFindLevel.addItem "4-跟踪"
    cboFindLevel.addItem "5-全开"
    
    '对应数据表的字段名
    lblFindInfo(FindInfo.FI_User).Tag = "USER_NAME"
    lblFindInfo(FindInfo.FI_Station).Tag = "STATION"
    lblFindInfo(FindInfo.FI_SID).Tag = "SESSION_ID"
    lblFindInfo(FindInfo.FI_IP).Tag = "IP"
    lblFindInfo(FindInfo.FI_Process).Tag = "PROCESS_NAME"
    lblFindInfo(FindInfo.FI_Component).Tag = "COMPONENT_NAME"
    lblFindInfo(FindInfo.FI_Module).Tag = "MODULE_NAME"
    lblFindInfo(FindInfo.FI_Function).Tag = "FUNCTION_NAME"
    lblFindInfo(FindInfo.FI_Call).Tag = "CALL_NAME"
    
    '初始化界面
    With tbPage.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ClientFrame = xtpTabFrameSingleLine
        .BoldSelected = True
        .OneNoteColors = True
        .ShowIcons = False
    End With
    tbPage.InsertItem PE_日志分类, "日志记录", picPage(PE_日志分类).hwnd, 0
    tbPage.InsertItem PE_记录规则, "日志设置", picPage(PE_记录规则).hwnd, 0
    
    Set mobjLogInfo = New clsVSFlexGridEx
    With mobjLogInfo
        .AppTemplate EM_Modify, vsfLogInfo, MSTR_LOGINFO, ""
        .Init
        .Binding.ExplorerBar = flexExNone
        .Binding.ScrollTrack = True
    End With
    vsfLogInfo.ColComboList(vsfLogInfo.ColIndex("内容")) = "..."
    vsfLogInfo.ExplorerBar = flexExSortShow
            
    Call tbPage_SelectedChanged(tbPage.Item(0))
    Call PicFind_Resize
    Call LoadData(PE_日志分类)
    Call LoadData(PE_查找选项)
    Call cmdFindReset_Click
End Sub

Private Sub LoadData(ByVal bytIndex As PageEnum)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo errH
    
    Select Case bytIndex
    Case PE_日志分类
        strSQL = GetFindSQL()
        If strSQL = "" Then Exit Sub
        Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption & "-日志分类")
        
        mobjLogInfo.Recordset = rsTmp
        If chkAllCols.value = 1 Then
            mobjLogInfo.ColsHide = "ID"
        Else
            mobjLogInfo.ColsHide = "ID|会话ID|工作站|进程ID|进程名|部件|模块|日志等级|日志分类"
        End If
        mobjLogInfo.Repaint RT_Rows
        mobjLogInfo.SetColsHide
        rsTmp.Close
        '定位至底
        vsfLogInfo.Row = vsfLogInfo.Rows - 1
        vsfLogInfo.TopRow = vsfLogInfo.Row
        If vsfLogInfo.Visible Then vsfLogInfo.SetFocus
    Case PE_记录规则
        strSQL = _
            "Select ID, Name, Description, Decode(Builtin, 1, '固定', '自定义') Builtin, Begin_Time, End_Time, Log_Keep_Days" & vbCr & _
            "  , Decode(Log_Mode, 0, '不记录', 1, '本地记录', 2, '数据库记录', 3, '本地和数据库记录') Log_Mode" & vbCr & _
            "  , Decode(Log_Level, 0, '0-关闭', 1, '1-错误', 2, '2-警告', 3, '3-重要', 4, '4-跟踪', 5, '5-全开') Log_Level " & vbCr & _
            "From Zllogcategory Order By ID "
        Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption & "-日志记录规则")
        With vsfLogSet
            .Redraw = flexRDNone
            .Rows = .FixedRows
            vsfLogSetDetail.Rows = vsfLogSetDetail.FixedRows
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, LSC_ID) = rsTmp!Id
                .TextMatrix(i, LSC_Name) = rsTmp!name & ""
                .TextMatrix(i, LSC_Builtin) = rsTmp!BuiltIn & ""
                .TextMatrix(i, LSC_Begin_Time) = Format(rsTmp!Begin_Time & "", "yyyy-mm-dd Hh:Nn:Ss")
                .TextMatrix(i, LSC_End_Time) = Format(rsTmp!End_Time & "", "yyyy-mm-dd Hh:Nn:Ss")
                .TextMatrix(i, LSC_Log_Keep_Days) = rsTmp!Log_Keep_Days & ""
                .TextMatrix(i, LSC_Log_Mode) = rsTmp!Log_Mode & ""
                .TextMatrix(i, LSC_Log_Level) = rsTmp!Log_Level & ""
                .TextMatrix(i, LSC_Description) = rsTmp!Description & ""
                rsTmp.MoveNext
            Next
            .Redraw = flexRDDirect
        End With
    Case PE_查找选项
        strSQL = _
            "Select * From (" & vbCr & _
            "  With T1 As" & vbCr & _
            "    (" & vbCr & _
            "     Select Upper(a.User_Name) User_Name, Upper(a.Station) Station, a.Session_Id, a.Ip, Upper(a.Process_Name) Process_Name" & vbCr & _
            "       , Upper(a.Component_Name) Component_Name, Upper(a.Module_Name) Module_Name, Upper(a.Function_Name) Function_Name, Upper(a.Call_Name) Call_Name" & vbCr & _
            "     From Zlloginfo A" & vbCr & _
            "     Where Rownum < 5000" & vbCr & _
            "    )" & vbCr & _
            "  Select Distinct 0 类别, User_Name 内容 From T1 Union All" & vbCr & _
            "  Select Distinct 1, Station From T1 Union All" & vbCr & _
            "  Select Distinct 2, Cast(Session_Id As Varchar2(8)) From T1 Union All" & vbCr & _
            "  Select Distinct 3, Ip From T1 Union All" & vbCr & _
            "  Select Distinct 4, Process_Name From T1 Union All" & vbCr & _
            "  Select Distinct 5, Component_Name From T1 Union All" & vbCr & _
            "  Select Distinct 6, Module_Name From T1 Union All" & vbCr & _
            "  Select Distinct 7, Function_Name From T1 Union All" & vbCr & _
            "  Select Distinct 8, Call_Name From T1" & vbCr & _
            ") Order By 类别, 内容"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption & "-查找选项")
        With rsTmp
            For i = 0 To 8
                cboGroup(i).Clear
                cboGroup(i).addItem ""
                
                .Filter = "类别=" & i
                Do While Not .EOF
                    If Trim$("" & !内容) <> "" Then
                        cboGroup(i).addItem "" & !内容
                    End If
                    .MoveNext
                Loop
            Next
            .Close
        End With
    End Select
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
errH:
    Screen.MousePointer = vbDefault
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub LoadSetDetail(ByVal lngCategoryID As Long)
    Dim strSQL              As String
    Dim rsTmp               As ADODB.Recordset
    Dim i                   As Long
    Dim strCondition        As String
    
    On Error GoTo errH
    
    strSQL = "Select ID, Category_Id, User_Name, Station, Ip, Component_Names, Module_Names, Function_Names, Call_Names" & vbNewLine & _
            "From Zllogset" & vbNewLine & _
            "Where Category_Id = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption, lngCategoryID)
    With vsfLogSetDetail
        .Redraw = flexRDNone
        .Rows = .FixedRows
        vsfLogSetDetail.Rows = vsfLogSetDetail.FixedRows
        Call CtrlSetFunc
        Call CtrlSetDetailFunc
        .Rows = rsTmp.RecordCount + .FixedRows
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, LSC_ID) = rsTmp!Id
                
            If Not IsNull(rsTmp!User_Name) Then
                strCondition = strCondition & " 用户=" & rsTmp!User_Name
            End If
            If Not IsNull(rsTmp!Station) Then
                strCondition = strCondition & " 客户端=" & rsTmp!Station
            End If
            If Not IsNull(rsTmp!IP) Then
                strCondition = strCondition & " IP=" & rsTmp!IP
            End If
            strCondition = Trim(strCondition)
            .TextMatrix(i, LSDC_Condition) = strCondition
            .TextMatrix(i, LSDC_Component_Names) = rsTmp!Component_Names & ""
            .TextMatrix(i, LSDC_Module_Names) = rsTmp!Module_Names & ""
            .TextMatrix(i, LSDC_Function_Names) = rsTmp!Function_Names & ""
            .TextMatrix(i, LSDC_Call_Names) = rsTmp!Call_Names & ""
            rsTmp.MoveNext
        Next
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function GetFindSQL() As String
    Dim strDateStart As String, strDateEnd As String, strSQL As String, strFind As String
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = " Create_Time Between To_Date('" & dtpFindDate(0).value & "','yyyy-MM-dd hh24:mi:ss') " & _
             "And To_date('" & dtpFindDate(1).value & "','yyyy-MM-dd hh24:mi:ss')"
    strFind = lblFindDate.Caption & ":" & _
              Format(dtpFindDate(0).value, "yyyy-mm-dd Hh:Nn:Ss") & _
              "~" & _
              Format(dtpFindDate(1).value, "yyyy-mm-dd Hh:Nn:Ss")
    
    For i = cboGroup.LBound To cboGroup.UBound
        If cboGroup(i).Text <> "" Then
            Select Case i
            Case FindInfo.FI_SID
                strSQL = strSQL & " And " & lblFindInfo(i).Tag & " = " & val(cboGroup(i).Text)
                strFind = strFind & "," & lblFindInfo(i).Caption & ":" & val(cboGroup(i).Text)
            Case Else
                strSQL = strSQL & " And Upper(" & lblFindInfo(i).Tag & ") = '" & UCase(Trim(cboGroup(i).Text)) & "'"
                strFind = strFind & "," & lblFindInfo(i).Caption & ":" & UCase(cboGroup(i).Text)
            End Select
        End If
    Next
    
    If cboFindCategory.Text <> "" Then
        strSQL = strSQL & " And " & cboFindCategory.Tag & " ='" & cboFindCategory.Text & "'"
        strFind = strFind & "," & lblFindCategory.Caption & ":" & cboFindCategory.Text
    End If
    If cboFindLevel.Text <> "" Then
        strSQL = strSQL & " And " & cboFindLevel.Tag & " =" & cboFindLevel.ListIndex
        strFind = strFind & "," & lblFindLevel.Caption & ":" & cboFindLevel.Text
    End If
    
    'Dbms_Lob.Substr(Log_Info_Ex, ?, ?)) 参数2的值过大会引起“ORA-06502 PL/SQL:数字或值错误：字符串缓冲区太小”异常
    GetFindSQL = "Select ID, Server 服务器, User_Name 用户, Session_Id 会话ID, Ip, Station 工作站, Process_Id 进程ID" & vbCr & _
                 "  , Process_Name 进程名, Category_Name 日志分类, Component_Name 部件, Module_Name 模块" & vbCr & _
                 "  , Function_Name 功能, Call_Name 服务名, Decode(Stage, 0, '开始', 1, '结束', Null) 状态, Create_Time 登记时间 " & vbCr & _
                 "  , Decode(Log_Level, 0, '0-关闭', 1, '1-错误', 2, '2-警告', 3, '3-重要', 4, '4-跟踪', 5, '5-全开') 日志级别" & vbCr & _
                 "  , Nvl(Log_Info, Dbms_Lob.Substr(Log_Info_Ex, 2000, 1)) 内容" & vbCr & _
                 "From Zlloginfo " & vbCrLf & _
                 "Where " & strSQL & vbCrLf & _
                 "Order By Create_Time"
    lblFindDetail.Caption = strFind
    Exit Function
    
errH:
    MsgBox "获取数据失败！" & vbCrLf & "错误编号：" & err.Number & vbCrLf & "错误内容：" & err.Description _
        , vbCritical, gstrSysName
End Function

Private Function ReGetConnection() As ADODB.Connection
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    Dim strCommand      As String
    Dim strServer       As String, strUser      As String, strPass  As String, blnTrans As Boolean
    Dim arrTmp          As Variant, i           As Long
    Dim strError        As String
    Dim cnOralce        As ADODB.Connection
    
    On Error GoTo errH
    
    strSQL = "Select Max(内容) 内容 From zlRegInfo Where 项目 = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "日志服务器")
    If Not IsNull(rsTmp!内容) Then
        If rsTmp!内容 & "" Like "ZLSV*:*" Then
            strCommand = Sm4DecryptEcb(rsTmp!内容)
        Else
            strCommand = rsTmp!内容
        End If

        arrTmp = Split(strCommand, " ")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If Trim(arrTmp(i)) <> "" Then
                If arrTmp(i) Like "USER=*" Then
                    strUser = Mid(arrTmp(i), Len("USER=*"))
                ElseIf arrTmp(i) Like "PASS=*" Then
                    strPass = Mid(arrTmp(i), Len("PASS=*"))
                ElseIf arrTmp(i) Like "TRANS=*" Then
                    blnTrans = val(Mid(arrTmp(i), Len("TRANS=*"))) = 1
                ElseIf arrTmp(i) Like "SERVER=*" Then
                    strServer = Mid(arrTmp(i), Len("SERVER=*"))
                Else
                    If LenB(strServer) = 0 Then
                        strServer = arrTmp(i)
                    End If
                End If
            End If
        Next

        '“日志服务器设备”未指定用户、密码，缺省使用ZLUA，与其逻辑保持一致
        If strUser = "" Then
            strUser = "ZLUA"
            strPass = Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY))
            blnTrans = False
        End If
        
        Set cnOralce = gobjRegister.GetConnection(strServer, strUser, strPass, blnTrans, OraOLEDB, strError, False)
        If cnOralce.State = adStateClosed Then
            MsgBox "连接日志服务器出错！错误：" & strError & ",将读取当前数据库的日志配置。", vbInformation, gstrSysName
            Set cnOralce = gcnOracle
        End If
    Else
        Set cnOralce = gcnOracle
    End If
    Set ReGetConnection = cnOralce
    Exit Function
errH:
    MsgBox "连接日志服务器出错！错误：" & err.Description & ",将读取当前数据库的日志配置。", vbInformation, gstrSysName
    Set ReGetConnection = gcnOracle
End Function

Private Sub ResetContainer()
    Dim i As Integer
    
    For i = picPage.LBound To picPage.UBound
        If Not picPage(i) Is Nothing Then
            Set picPage(i).Container = Me
        End If
    Next
    Set PicFind.Container = Me
End Sub

Private Sub LoadLogList()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select ID, Name From Zllogcategory Order By ID"
    Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption)
    cboFindCategory.Clear
    cboFindCategory.addItem ""
    Do While Not rsTmp.EOF
        cboFindCategory.addItem "" & rsTmp!name
        rsTmp.MoveNext
    Loop
    Exit Sub
    
errH:
    MsgBox "加载日志分类出错，请检查数据表或权限。", vbInformation, gstrSysName
End Sub

Private Sub CtrlSetFunc()
    Dim blnEnable As Boolean, blnCanDel As Boolean
    
    With vsfLogSet
        If .Rows = .FixedRows Then
            blnEnable = False
        Else
            blnEnable = .Row >= .FixedRows
        End If
        blnCanDel = blnEnable And .TextMatrix(.Row, LSC_Builtin) <> "固定"
    End With
    cmdLogSetAdd.Enabled = blnEnable
    cmdLogSetEdit.Enabled = blnEnable
    cmdLogSetDel.Enabled = blnCanDel
End Sub

Private Sub CtrlSetDetailFunc()
    Dim blnEnable As Boolean, blnHaveSet As Boolean
    
    With vsfLogSetDetail
        If .Rows = .FixedRows Then
            blnEnable = False
        Else
            blnEnable = .Row >= .FixedRows
        End If
    End With
    With vsfLogSet
        If .Rows = .FixedRows Then
            blnHaveSet = False
        Else
            blnHaveSet = .Row >= .FixedRows
        End If
    End With
    cmdLogSetCtrlAdd.Enabled = blnHaveSet
    cmdLogSetCtrlEdit.Enabled = blnEnable
    cmdLogSetCtrlDel.Enabled = blnEnable
End Sub

Private Sub chkAllCols_Click()
    If chkAllCols.value = 1 Then
        mobjLogInfo.ColsHide = "ID"
    Else
        mobjLogInfo.ColsHide = "ID|会话ID|工作站|进程ID|进程名|部件|模块|日志分类"
    End If
    mobjLogInfo.SetColsHide
End Sub

Private Sub cmdFind_Click()
    PicFind.Visible = True
    CmdFind.Enabled = False
    Call PicFind_Resize
    If cboFindCategory.Visible And cboFindCategory.Enabled Then
        cboFindCategory.SetFocus
    End If
End Sub

Private Sub cmdFindCancel_Click()
    PicFind.Visible = False
    CmdFind.Enabled = True
End Sub

Private Sub cmdFindOK_Click()
    '检查时间范围
    If dtpFindDate(1).value - dtpFindDate(0).value > 1 Then
        If Not mblnIgnoreWarn Then
            Call frmMsgBoxEx.ShowMe(Me, "提醒：查找的日志范围超过一天，可能比较耗时。", , mslInformation, True, gstrSysName)
            mblnIgnoreWarn = frmMsgBoxEx.chkIgnorePrompt.value = 1
            Unload frmMsgBoxEx
        End If
    ElseIf dtpFindDate(1).value - dtpFindDate(0).value < 0 Then
        MsgBox "查找结束时间不能比开始时间早。", vbCritical, gstrSysName
        If dtpFindDate(1).Visible Then dtpFindDate(1).SetFocus
        Exit Sub
    End If
    
    PicFind.Visible = False
    frmMDIMain.stbThis.Panels(2).Text = "正在查找..."
    Call LoadData(PE_日志分类)
    frmMDIMain.stbThis.Panels(2).Text = ""
    CmdFind.Enabled = True
End Sub

Private Sub cmdFindReset_Click()
    Dim i As Long
    Dim dtmCurrent As Date
    
    dtmCurrent = mdlMain.CurrentDate()
    For i = cboGroup.LBound To cboGroup.UBound
        cboGroup(i).Text = ""
    Next
    cboFindCategory.ListIndex = 0
    cboFindLevel.ListIndex = 0
    dtpFindDate(1).value = CDate(Format(dtmCurrent, "yyyy-MM-dd HH:mm"))
    dtpFindDate(0).value = dtpFindDate(1).value - 1 / 24
End Sub

Private Sub cmdLogServer_Click()
    If frmCommenLogServer.ShowMe() Then
        '重新设置连接，并重新刷新数据
        Set mcnOracle = ReGetConnection
        Call LoadLogList
        Call LoadData(PE_记录规则)
    End If
End Sub

Private Sub cmdLogSetAdd_Click()
    If frmCommenLogSetEdit.ShowMe(mcnOracle) Then
        Call LoadData(PE_记录规则)
        vsfLogSet.TopRow = vsfLogSet.Rows - 1
        vsfLogSet.ShowCell vsfLogSet.Rows - 1, LSC_ID
        vsfLogSet.Select vsfLogSet.Rows - 1, LSC_ID
    End If
End Sub

Private Sub cmdLogSetCtrlAdd_Click()
    Dim lngRow As Long, lngCategoryID As Long
    
    lngRow = vsfLogSet.Row
    If lngRow > 0 Then
        lngCategoryID = val(vsfLogSet.TextMatrix(lngRow, LSC_ID))
        If frmCommenLogSetDetailEdit.ShowMe(mcnOracle, lngCategoryID) Then
            Call LoadSetDetail(lngCategoryID)
            lngRow = vsfLogSetDetail.Rows - 1
            vsfLogSetDetail.TopRow = lngRow
            vsfLogSetDetail.ShowCell lngRow, LSDC_ID
            vsfLogSetDetail.Select lngRow, LSDC_ID
        End If
    End If
End Sub

Private Sub cmdLogSetCtrlDel_Click()
    Dim strSQL          As String
    Dim strName         As String
    Dim lngID           As Long
    Dim lngCategoryID   As Long
    Dim lngRow          As Long
    
    On Error GoTo errH
    With vsfLogSetDetail
        If .Row >= .FixedRows Then
            strName = .TextMatrix(.Row, LSDC_Condition)
            If Len(strName) > 20 Then
                strName = Left$(strName, 20) & "..."
            End If
            lngID = val(.TextMatrix(.Row, LSDC_ID))
        Else
            Exit Sub
        End If
    End With
    If MsgBox("你确定要删除日志规则“" & strName & "”？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "Zllogset_Edit(2," & lngID & ")"
    Call ExecuteProcedure(strSQL, Me.Caption, mcnOracle)
    lngRow = vsfLogSet.Row
    If lngRow > 0 Then
        lngCategoryID = val(vsfLogSet.TextMatrix(lngRow, LSC_ID))
        Call LoadSetDetail(lngCategoryID)
        If vsfLogSetDetail.Rows > vsfLogSetDetail.FixedRows Then
            If lngRow >= vsfLogSetDetail.Rows Then
                lngRow = vsfLogSetDetail.Rows - 1
            End If
            vsfLogSetDetail.TopRow = lngRow
            vsfLogSetDetail.ShowCell lngRow, LSDC_ID
            vsfLogSetDetail.Select lngRow, LSDC_ID
        End If
    End If
    Exit Sub
    
errH:
    MsgBox "删除日志规则出现错误：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdLogSetCtrlEdit_Click()
    Dim lngRow      As Long
    Dim lngCategoryID   As Long
    
    lngRow = vsfLogSet.Row
    If lngRow > 0 Then
        lngCategoryID = val(vsfLogSet.TextMatrix(lngRow, LSC_ID))
        lngRow = vsfLogSetDetail.Row
        If lngRow > 0 Then
            If frmCommenLogSetDetailEdit.ShowMe(mcnOracle, lngCategoryID, val(vsfLogSetDetail.TextMatrix(lngRow, LSDC_ID))) Then
                Call LoadSetDetail(lngCategoryID)
                vsfLogSetDetail.TopRow = lngRow
                vsfLogSetDetail.ShowCell lngRow, LSDC_ID
                vsfLogSetDetail.Select lngRow, LSDC_ID
            End If
        End If
    End If
End Sub

Private Sub cmdLogSetDel_Click()
    Dim strSQL          As String
    Dim strName         As String
    Dim lngID           As Long
    Dim lngRow          As Long
    
    On Error GoTo errH
    
    With vsfLogSet
        lngRow = .Row
        If .Row >= .FixedRows Then
            If .TextMatrix(.Row, LSC_Builtin) = "固定" Then
                Exit Sub
            End If
            strName = .TextMatrix(.Row, LSC_Name)
            lngID = val(.TextMatrix(.Row, LSC_ID))
        Else
            Exit Sub
        End If
    End With
    If MsgBox("你确认要删除日志分类“" & strName & "”？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "Zllogcategory_Edit(2, " & lngID & ")"
    Call ExecuteProcedure(strSQL, Me.Caption, mcnOracle)
    Call LoadData(PE_记录规则)
    Call LoadLogList
    If vsfLogSet.Rows > vsfLogSet.FixedRows Then
        If lngRow >= vsfLogSet.Rows Then
            lngRow = vsfLogSet.Rows - 1
        End If
        vsfLogSet.TopRow = lngRow
        vsfLogSet.ShowCell lngRow, LSC_ID
        vsfLogSet.Select lngRow, LSC_ID
    End If
    Exit Sub
    
errH:
    MsgBox "删除日志分类出现错误：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdLogSetEdit_Click()
    Dim lngRow As Long, lngID As Long
    
    lngRow = vsfLogSet.Row
    If lngRow > 0 Then
        lngID = val(vsfLogSet.TextMatrix(lngRow, LSC_ID))
        If frmCommenLogSetEdit.ShowMe(mcnOracle, vsfLogSet, lngID) Then
            Call LoadData(PE_记录规则)
            vsfLogSet.TopRow = lngRow
            vsfLogSet.ShowCell lngRow, LSC_ID
            vsfLogSet.Select lngRow, LSC_ID
            Call LoadLogList
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    On Error Resume Next
    tbPage.Height = Me.ScaleHeight - tbPage.Top + 15
    tbPage.Width = Me.ScaleWidth - tbPage.Left + 15
    For i = 0 To 1
        picPage(i).Left = 0
        picPage(i).Width = tbPage.Width - 60
        picPage(i).Height = tbPage.Height - picPage(i).Top - 60
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjLogInfo = Nothing
    Unload frmCommenLogSetEdit
    Unload frmCommenLogSetDetailEdit
    Unload frmCommenLogServer
    Unload frmShowContent
End Sub

Private Sub FraHead_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    PicFind_MouseDown Button, Shift, x, y
    FraHead.Tag = "1"
End Sub

Private Sub FraHead_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And FraHead.Tag = "1" Then
        PicFind_MouseMove Button, Shift, x, y
    End If
End Sub

Private Sub FraHead_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    FraHead.Tag = ""
End Sub

Private Sub PicClose_Click()
    Call cmdFindCancel_Click
End Sub

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then RaisEffect PicClose, -2
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then RaisEffect PicClose, 0
End Sub

Private Sub PicFind_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        mcurMousePoint.x = x
        mcurMousePoint.y = y
    End If
End Sub

Private Sub PicFind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        With mcurWindowRect
            .Left = PicFind.Left + x - mcurMousePoint.x
            .Top = PicFind.Top + y - mcurMousePoint.y
            
            If .Left < ScaleLeft Then .Left = ScaleLeft
            If .Left + PicFind.Width > ScaleWidth Then .Left = ScaleWidth - PicFind.Width
            If .Top < ScaleTop Then .Top = ScaleTop
            If .Top + PicFind.Height > ScaleHeight Then .Top = ScaleHeight - PicFind.Height
        End With
        
        PicFind.Move mcurWindowRect.Left, mcurWindowRect.Top
    End If
End Sub

Private Sub PicFind_Resize()
    On Error Resume Next
    
    PicFind.Move (Me.ScaleWidth - PicFind.Width) \ 2 + Me.Left, (Me.ScaleHeight - PicFind.Height) \ 2 + Me.Top
    fraFind.Move 0, 0, PicFind.ScaleWidth, PicFind.ScaleHeight
    FraHead.Move 0, -105, fraFind.Width + 30
    PicClose.Move PicFind.ScaleWidth - PicClose.Width - 30
End Sub

Private Sub picPage_Resize(Index As Integer)
    On Error Resume Next
    
    If Index = PE_日志分类 Then
        vsfLogInfo.Move 0, vsfLogInfo.Top, picPage(PE_日志分类).ScaleWidth - vsfLogInfo.Left - 90 _
            , picPage(PE_日志分类).ScaleHeight - vsfLogInfo.Top - 30
        lblFindDetail.Width = picPage(Index).ScaleWidth - lblFindDetail.Left - 30
        lblFindDetail.Height = 180 * val("3-行")
    Else
        If picPage(PE_记录规则).ScaleWidth > 3000 And picPage(PE_记录规则).ScaleHeight > 4000 Then
            vsfLogSet.Move 0, lblLogSet.Top + lblLogSet.Height + 120 _
                , picPage(PE_日志分类).ScaleWidth - 120 _
                , picPage(PE_日志分类).ScaleHeight * 0.5
            fraLogSetCtrl.Left = vsfLogSet.Left + vsfLogSet.Width - fraLogSetCtrl.Width
            fraLogSetCtrl.Top = vsfLogSet.Top + vsfLogSet.Height + 30
            cmdLogSetDel.Left = fraLogSetCtrl.Width - cmdLogSetDel.Width - 15
            cmdLogSetEdit.Left = cmdLogSetDel.Left - cmdLogSetEdit.Width - 45
            cmdLogSetAdd.Left = cmdLogSetEdit.Left - cmdLogSetAdd.Width - 45
            
            lblLogSetDetail.Move 0, fraLogSetCtrl.Top + fraLogSetCtrl.Height - 120
            vsfLogSetDetail.Move vsfLogSet.Left, lblLogSetDetail.Top + lblLogSetDetail.Height + 30 _
                , vsfLogSet.Width _
                , picPage(PE_记录规则).ScaleHeight - lblLogSetDetail.Top - lblLogSetDetail.Height - fraLogSetDetailCtrl.Height - 30 * 2
            fraLogSetDetailCtrl.Left = vsfLogSet.Left + vsfLogSet.Width - fraLogSetDetailCtrl.Width
            fraLogSetDetailCtrl.Top = vsfLogSetDetail.Top + vsfLogSetDetail.Height + 30
            cmdLogSetCtrlDel.Left = fraLogSetDetailCtrl.Width - cmdLogSetCtrlDel.Width
            cmdLogSetCtrlEdit.Left = cmdLogSetCtrlDel.Left - cmdLogSetCtrlEdit.Width - 45
            cmdLogSetCtrlAdd.Left = cmdLogSetCtrlEdit.Left - cmdLogSetCtrlAdd.Width - 45
        End If
        cmdLogServer.Left = vsfLogSet.Left + vsfLogSet.Width - cmdLogServer.Width - 30
        cmdLogServer.Top = lblLogSet.Top - 90
    End If
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub

    If Item.Index = 1 Then
        Call LoadData(Item.Index)
        Call picPage_Resize(Item.Index)
        Call cmdFindCancel_Click
    End If
End Sub

Private Sub txtFindInfo_Change(Index As Integer)

End Sub

Private Sub vsfLogInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = vsfLogInfo.ColIndex("内容") Then
        With frmShowContent
            .Mode = scmCommonLogContent
            .ShowMe vsfLogInfo.TextMatrix(0, Col), vsfLogInfo.TextMatrix(Row, Col)
        End With
    End If
End Sub

Private Sub vsfLogInfo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> vsfLogInfo.ColIndex("内容")
End Sub

Private Sub vsfLogSet_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call CtrlSetFunc
    If NewRow > 0 Then
        lblLogSetDetail.Caption = "日志记录规则（" & vsfLogSet.TextMatrix(NewRow, LSC_Name) & "）"
        Call LoadSetDetail(val(vsfLogSet.TextMatrix(NewRow, LSC_ID)))
    End If
End Sub

Private Sub vsfLogSet_DblClick()
    Call cmdLogSetEdit_Click
End Sub

Private Sub vsfLogSetDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call CtrlSetDetailFunc
End Sub

Private Sub vsfLogSetDetail_DblClick()
    Call cmdLogSetCtrlEdit_Click
End Sub

