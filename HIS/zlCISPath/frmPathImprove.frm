VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPathImprove 
   BackColor       =   &H80000005&
   Caption         =   "辅助改进"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   DrawStyle       =   1  'Dash
   HasDC           =   0   'False
   Icon            =   "frmPathImprove.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   15420
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraContent 
      BackColor       =   &H80000005&
      Caption         =   "概况"
      Height          =   1695
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   24
      Top             =   840
      Width           =   15015
      Begin VSFlex8Ctl.VSFlexGrid vsStep 
         Height          =   1095
         Left            =   7200
         TabIndex        =   8
         Top             =   480
         Width           =   5655
         _cx             =   9975
         _cy             =   1931
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
         BackColorFixed  =   13430215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   4000
         ColWidthMin     =   500
         ColWidthMax     =   4500
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathImprove.frx":6852
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
         BackColorFrozen =   13430215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "请勾选生成新版路径需要调整的阶段。"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   27
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "原标准住院日为：12天，现平均标准住院日为：10天。"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "该临床路径的总病例数：400人，正常结束病例数为：320人。"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   755
      ScaleMode       =   0  'User
      ScaleWidth      =   15420
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8835
      Width           =   15420
      Begin VB.CommandButton cmdSend 
         Caption         =   "新版路径生成(&S)"
         Height          =   309
         Left            =   12600
         TabIndex        =   13
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line linSplit 
         BorderColor     =   &H8000000F&
         Index           =   0
         X1              =   0
         X2              =   20400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linSplit 
         BorderColor     =   &H8000000E&
         Index           =   1
         X1              =   0
         X2              =   20640
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraFilter 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "查询条件"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   15375
      Begin VB.ComboBox cboBranch 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   480
         Width           =   3075
      End
      Begin VB.ComboBox cboCategory 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   60
         Width           =   1815
      End
      Begin VB.CommandButton cmdPathName 
         Appearance      =   0  'Flat
         Caption         =   "…"
         Height          =   250
         Left            =   6880
         Picture         =   "frmPathImprove.frx":68DF
         TabIndex        =   35
         Top             =   75
         Width           =   300
      End
      Begin VB.TextBox txtPathName 
         Height          =   320
         Left            =   4200
         TabIndex        =   1
         Text            =   "声带息肉"
         Top             =   60
         Width           =   3015
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   0
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "95"
         Top             =   495
         Width           =   400
      End
      Begin VB.CommandButton cmdAnalyse 
         Caption         =   "变异分析(F)"
         Height          =   320
         Left            =   13440
         TabIndex        =   7
         Top             =   470
         Width           =   1500
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   12120
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "80"
         Top             =   495
         Width           =   400
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "95"
         Top             =   495
         Width           =   400
      End
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7560
         ScaleHeight     =   300
         ScaleWidth      =   6015
         TabIndex        =   16
         Top             =   60
         Width           =   6015
         Begin MSComCtl2.DTPicker dtpTimeStart 
            Height          =   300
            Left            =   1080
            TabIndex        =   2
            Top             =   0
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   106168323
            CurrentDate     =   41334
         End
         Begin MSComCtl2.DTPicker dtpTimeEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            TabIndex        =   3
            Top             =   0
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   99352579
            CurrentDate     =   41365
         End
         Begin VB.Label lblBetweenTimes 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "导入时间：从                        至                    "
            Height          =   180
            Left            =   0
            TabIndex        =   17
            Top             =   60
            Width           =   5220
         End
      End
      Begin VB.Label lblBranch 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "分支路径"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblRate 
         BackColor       =   &H80000005&
         Caption         =   "阶段提前或延后的比例≥      %"
         Height          =   180
         Index           =   2
         Left            =   4560
         TabIndex        =   23
         Top             =   540
         Width           =   2655
      End
      Begin VB.Label lblRate 
         BackColor       =   &H80000005&
         Caption         =   "路径外项目的比例≥      %"
         Height          =   180
         Index           =   1
         Left            =   10440
         TabIndex        =   21
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label lblRate 
         BackColor       =   &H80000005&
         Caption         =   "路径内未生成比例≥      %"
         Height          =   180
         Index           =   0
         Left            =   7560
         TabIndex        =   20
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "路径名称(&N)"
         Height          =   180
         Left            =   3120
         TabIndex        =   19
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblCategory 
         BackColor       =   &H80000005&
         Caption         =   "分类(&C)"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   240
      ScaleHeight     =   5535
      ScaleWidth      =   14955
      TabIndex        =   14
      Top             =   2880
      Width           =   14985
      Begin VB.Frame fraSplit 
         Height          =   15
         Left            =   -840
         MousePointer    =   7  'Size N S
         TabIndex        =   34
         Top             =   2400
         Width           =   15735
      End
      Begin VB.Frame fraAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3480
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   14775
         Begin VB.Frame fraSplitAdvice 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   6840
            MousePointer    =   9  'Size W E
            TabIndex        =   36
            Top             =   600
            Width           =   135
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
            Height          =   1935
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Tag             =   "取消医嘱"
            Top             =   720
            Width           =   6375
            _cx             =   11245
            _cy             =   3413
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   6
            Cols            =   5
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D131
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
            AutoSearch      =   1
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
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
            Height          =   1935
            Index           =   1
            Left            =   7320
            TabIndex        =   12
            Tag             =   "增加医嘱"
            Top             =   720
            Width           =   6015
            _cx             =   10610
            _cy             =   3413
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   6
            Cols            =   5
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D21F
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
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
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "路径医嘱： 未使用比例=所有使用该阶段的未生成该路径医嘱的病人数/所有使用该阶段的病人数；       "
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   8055
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使用比例=该阶段所有添加该医嘱类路径外项目的病人数/所有使用该阶段的病人数。"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   32
            Top             =   360
            Width           =   6735
         End
      End
      Begin VB.Frame fraItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   14895
         Begin VB.Frame fraSplitItem 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   6960
            MousePointer    =   9  'Size W E
            TabIndex        =   37
            Top             =   720
            Width           =   120
         End
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   1575
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   6615
            _cx             =   11668
            _cy             =   2778
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D2F6
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   110
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
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   1575
            Index           =   1
            Left            =   7080
            TabIndex        =   10
            Top             =   600
            Width           =   6375
            _cx             =   11245
            _cy             =   2778
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D3C8
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   110
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
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "使用比例=该阶段所有添加该非医嘱类路径外项目的病人数/所有使用该阶段的病人数。"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   30
            Top             =   360
            Width           =   7095
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "路径项目： 未使用比例=所有使用该阶段的未生成非医嘱类项目的病人数/所有使用该阶段的病人数；"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   8055
         End
      End
   End
End
Attribute VB_Name = "frmPathImprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng路径ID As Long              '记录当前路径ID
Private mstrPrivs As String             '操作路径的权限
Private mstr分类 As String              '记录当前路径分
Private mstr编码 As String              '当前路径编码
Private mlng版本号 As Long              '路径版本号
Private mrsStep As ADODB.Recordset      '所有阶段
Private mblnSend As Boolean             '生成成功标记

'阶段列号
Private Enum COL_Step
    COL_阶段 = 0
    COL_选择
    COL_详情
End Enum
'
Private Enum INDEX_TAG
    Index_DEL = 0
    Index_Add = 1
End Enum
'项目表单列号
Private Enum COL_Item
    COL_Item_阶段 = 0
    COL_Item_选择
    COL_Item_项目名称
    COL_Item_比例
    COL_Item_分类    '增加项目用到
End Enum
'医嘱表单列号
Private Enum COL_Advice
    COL_Advice_阶段 = 0
    COL_Advice_医嘱ID
    COL_Advice_相关ID
    COL_Advice_选择
    COL_Advice_期效
    COL_Advice_医嘱内容 '取消医嘱：该列cell值存放 医嘱序号;增加医嘱:该列cell存放 执行ID
    COL_Advice_比例
    COL_Advice_诊疗类别 '增加医嘱:cell存放的项目分类
    COL_Advice_诊疗项目ID
    COL_Advice_标本部位
    COL_Advice_名称
    COL_Advice_项目名称
End Enum
'比例值下标索引
Private Enum INDEX_RATE
    RATE_STEP = 0
    RATE_UNSEND
    RATE_PATHOUT
End Enum



Private Sub cboBranch_Click()
    Dim lngId As Long
    
    With cboBranch
        If .ListIndex = Val(.Tag) Then Exit Sub '.tag初始值为-1
        If mrsStep Is Nothing Then Exit Sub  '未进行分析
        lngId = .ItemData(.ListIndex)
        Call SetVSRowHidden(vsStep, lngId)
        Call SetVSRowHidden(vsItem(Index_Add), lngId)
        Call SetVSRowHidden(vsItem(Index_DEL), lngId)
        Call SetVSRowHidden(vsAdvice(Index_Add), lngId)
        Call SetVSRowHidden(vsAdvice(Index_DEL), lngId)
        .Tag = .ListIndex
    End With
End Sub

Private Sub cboCategory_Click()
    Dim lngCmd As Long
    If Trim(cboCategory.Text) = "" Then Exit Sub
    mstr分类 = Trim(cboCategory.Text)
    If cboCategory.Tag = "LOAD" Then
        lngCmd = 0
    Else
        lngCmd = 1
    End If
    Call LoadPathName(lngCmd)
    cboCategory.Tag = ""
End Sub

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
'根据输入   字符检索下拉列表

    Call Cbo.SetIndex(cboCategory.Hwnd, Cbo.MatchIndex(cboCategory.Hwnd, KeyAscii))
    Call cboCategory_Click 'CboSetIndex方法不会触发click事件
End Sub

Private Sub cboCategory_LostFocus()
    If Trim(cboCategory.Text) = "" Then
      cboCategory.SetFocus
      Call cboCategory_KeyPress(vbKeySpace)  '如果为空时，弹出分类下拉狂
    End If
End Sub

Private Sub cmdAnalyse_Click()
    Dim i As Integer
    Dim lngAllPati As Long
    
    '开始时间<结束时间检查
    If DateDiff("s", CDate(dtpTimeStart.Value), CDate(dtpTimeEnd.Value)) < 0 Then
         MsgBox "开始时间晚于结束时间,请重新调整时间。", vbInformation + vbOKOnly, gstrSysName
         dtpTimeStart.SetFocus
         Exit Sub
    End If
    '时间过长检查
    If DateDiff("m", CDate(dtpTimeStart.Value), CDate(dtpTimeEnd.Value)) >= 3 Then
        If MsgBox("您选择的日期间隔超过3个月,您确定按照这个日期间隔进行统计分析吗?", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Exit Sub
        End If
    End If
    '比例检查
    If Val(txtRate(RATE_STEP).Text) < 70 Or Val(txtRate(RATE_UNSEND).Text) < 70 Or Val(txtRate(RATE_PATHOUT).Text) < 70 Then
        If MsgBox("您输入的比例值小于70,你确定要按照这个比例值进行统计吗？", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            txtRate(i).SetFocus
            Exit Sub
        End If
    End If
    '加载阶段
    Call GetPathPhase
    
    '加载汇总信息
    Call SetSummaryInfo
    
    If Val(lblInfo(0).Tag) = 0 Then
        MsgBox "当前没有找到合格的病例,请调整过滤条件再进行变异分析。", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
        Exit Sub
    End If
    
    '加载阶段
    
    Call LoadPhase
    '加载项目
    Call LoadItem
    '加载医嘱
    Call LoadAdvice
    
    '缺省定位到主路径
    cboBranch.Tag = "-1"
    cboBranch.ListIndex = 0
    Call cboBranch_Click
    
End Sub

Private Sub cmdPathName_Click()
'加载路径名称
    Call LoadPathName(2, "")
End Sub

Private Sub cmdSend_Click()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim intVersion As Integer
    Dim blnTrans As Boolean
    Dim str医嘱组ID As String
    Dim blnDo As Boolean
    Dim arrSQL As Variant
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    If mrsStep Is Nothing Then Exit Sub
    ' 执行前检查如果未选择任何项目，提示并禁止
    blnDo = vsAdvice(Index_Add).FindRow("1", , COL_Advice_选择) = -1 And vsAdvice(Index_DEL).FindRow("1", , COL_Advice_选择) = -1 And _
        vsItem(Index_DEL).FindRow("1", , COL_Item_选择) = -1 And vsItem(Index_Add).FindRow("1", , COL_Item_选择) = -1
    If blnDo And vsStep.FindRow("1", , COL_选择) = -1 Then
        If MsgBox("你未选择任何内容,是否需要退出?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Unload Me: Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' 如果存在需要调整的阶段，给予用户提示，要求用户手动调整阶段。
    If vsStep.FindRow("1", , COL_选择) <> -1 Then
        MsgBox "如果您需要调整阶段，请到路径设计界面手动调整。", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
    End If
    
    If blnDo Then Exit Sub  '未选择任何修改内容
    
    '检查临床路径版本审核状态，未审核时，先删除未审核版本再新增。
    arrSQL = Array()
    strSql = "Select 审核时间, 版本号" & _
            "   From (Select t.审核时间, t.版本号 From 临床路径版本 T Where t.路径id = [1] Order By t.版本号 Desc)" & _
            "   Where Rownum < 2"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.RecordCount = 1 Then
        If IsNull(rsTmp!审核时间) Then
            intVersion = rsTmp!版本号
            If MsgBox("当前路径存在未审核的新版," & vbCrLf & "你确定要删除未审核的新版本再新增?", vbOKCancel + vbDefaultButton2 + vbQuestion, gstrSysName) = vbOK Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Delete(" & mlng路径ID & "," & intVersion & ")"
            Else
                Exit Sub
            End If
        End If
    End If
    
    '根据选择，复制旧版路径后，删除添加路径项目和路径医嘱。
    '复制当前选择版本内容产生新版本内容
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Copy(" & mlng路径ID & "," & mlng版本号 & "," & mlng路径ID & ",0)"
   
    '修正数据:项目取消 、增加项目 、取消医嘱 、增加医嘱
    '先处理增加项目再处理删除项目,避免整个阶段的项目删除完后在增加时报错
    '增加项目
    With vsItem(Index_Add)
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Item_选择) = Checked Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Improve(1," & mlng路径ID & "," & .Cell(flexcpData, i, COL_Item_阶段) & ",Null,Null,'" & .Cell(flexcpData, i, COL_Item_分类) & "'," & .Cell(flexcpData, i, COL_Item_项目名称) & ")"  '此处COL_Item_项目名称存的是执行ID值
            End If
        Next
    End With
     '增加医嘱
    With vsAdvice(Index_Add)
        str医嘱组ID = ""
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Advice_选择) = Checked Then
                strTmp = IIf(Val(.TextMatrix(i, COL_Advice_相关ID)) = 0, .TextMatrix(i, COL_Advice_医嘱ID), .TextMatrix(i, COL_Advice_相关ID))
                If InStr(str医嘱组ID & ",", "," & strTmp & ",") = 0 Then
                    str医嘱组ID = str医嘱组ID & "," & strTmp
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Improve(3," & mlng路径ID & "," & .Cell(flexcpData, i, COL_Advice_阶段) & ",'路径改进项目',Null,'" & .Cell(flexcpData, i, COL_Advice_诊疗类别) & "'," & .Cell(flexcpData, i, COL_Advice_医嘱内容) & "," & Val(strTmp) & ")"
                End If
            End If
        Next
    End With
    
    '项目取消
    With vsItem(Index_DEL)
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Item_选择) = Checked Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Improve(0," & mlng路径ID & "," & .Cell(flexcpData, i, COL_Item_阶段) & ",'" & .TextMatrix(i, COL_Item_项目名称) & "')"
            End If
        Next
    End With
    '取消医嘱
    With vsAdvice(Index_DEL)
        str医嘱组ID = "" '记录添加过的组ID
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Advice_选择) = Checked Then
                strTmp = IIf(Val(.TextMatrix(i, COL_Advice_相关ID)) = 0, .TextMatrix(i, COL_Advice_医嘱ID), .TextMatrix(i, COL_Advice_相关ID))
                If InStr(str医嘱组ID & ",", "," & strTmp & ",") = 0 Then
                    str医嘱组ID = str医嘱组ID & "," & strTmp
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Improve(2," & mlng路径ID & "," & .Cell(flexcpData, i, COL_Advice_阶段) & ",'" & .Cell(flexcpData, i, COL_Advice_项目名称) & "'," & .Cell(flexcpData, i, COL_Advice_医嘱内容) & ")"
                End If
            End If
        Next
    End With
    
     '提交数据
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next i
    gcnOracle.CommitTrans: blnTrans = False
    mblnSend = True
    '5)  生成成功后关闭退出窗体。
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpTimeEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRate(RATE_STEP).SetFocus
    End If
End Sub

Private Sub dtpTimeStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTimeEnd.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim DatCurr As Date
    Dim strTmp As String

    '加载分类
    Call Cbo.SetListHeight(cboCategory, 2000)
    cboCategory.Tag = "LOAD"
    Call LoadCategory
    '默认开始时间和结束时间间隔为30天
    DatCurr = zlDatabase.Currentdate
    dtpTimeStart.Value = Format(DateAdd("d", -30, DatCurr), "YYYY-MM-DD 00:00:00")
    dtpTimeEnd.Value = Format(DatCurr, "YYYY-MM-DD 23:59:59")
    
    '
    '表列初始化
    Call InitStep
    
    strTmp = "阶段;取消项目;取消项目;取消项目|阶段,1500,4;选择,500,4;项目名称,3000,4;未使用比例(%),1500,4"
    Call InitVSItem(vsItem(Index_DEL), strTmp)
    
    strTmp = "阶段;增加项目;增加项目;增加项目;增加项目|阶段,1500,4;选择,500,4;项目名称,3000,4;使用比例(%),1500,4;分类,,"
    Call InitVSItem(vsItem(Index_Add), strTmp)
    
    strTmp = "阶段;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱;取消医嘱|" & _
            "阶段,1500,4;相关ID;医嘱ID;选择,500,4;期效,500,4;医嘱内容,2500,4;未使用比例(%),1500,4;诊疗项目ID;诊疗类别;标本部位;名称;项目名称,,"
    Call InitVSAdvice(vsAdvice(Index_DEL), strTmp)
    
    strTmp = "阶段;增加医嘱;增加医嘱;增加医嘱;增加医嘱;增加医嘱;增加医嘱;增加医嘱|" & _
            "阶段,1500,4;相关ID;医嘱ID;选择,500,4;期效,500,4;医嘱内容,2500,4;使用比例(%),1500,4;诊疗类别"
    Call InitVSAdvice(vsAdvice(Index_Add), strTmp)
    
    lblInfo(0).Caption = "该临床路径的总病例数：0 人，正常结束病例数为：0 人。"
    lblInfo(1).Caption = "原标准住院日为：0 天，现平均标准住院日为：0 天。"
End Sub

Private Sub Form_Resize()
    Dim lngWidth As Long
    Dim lngLeft As Long

    On Error Resume Next
    lngLeft = 105
    lngWidth = Me.ScaleWidth - lngLeft * 2

    fraFilter.Move lngLeft, 0, lngWidth
    cmdAnalyse.Move fraFilter.Width - cmdAnalyse.Width - 450, (fraFilter.Height - cmdAnalyse.Height) / 2
    fraContent.Move lngLeft, fraFilter.Top + fraFilter.Height, lngWidth
    lblInfo(2).Left = lngLeft + lngWidth / 2
    vsStep.Left = lngLeft + lngWidth / 2
    picCenter.Move lngLeft, fraContent.Top + fraContent.Height, lngWidth, Me.ScaleHeight - picBottom.Height - (fraContent.Top + fraContent.Height)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    
    If Not mrsStep Is Nothing Then
        Set mrsStep = Nothing
    End If
    
End Sub

Private Sub fraSplitAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next

        If fraSplitAdvice.Left + X <= picCenter.ScaleWidth / 10 * 1 Or fraSplitAdvice.Left + X >= picCenter.ScaleWidth / 10 * 9 Then Exit Sub
        vsAdvice(Index_DEL).Width = vsAdvice(Index_DEL).Width + X
        fraSplitAdvice.Left = fraSplitAdvice.Left + X
        vsAdvice(Index_Add).Left = vsAdvice(Index_Add).Left + X
        vsAdvice(Index_Add).Width = vsAdvice(Index_Add).Width - X
    End If
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next

        If fraSplit.Top + Y <= picCenter.ScaleHeight / 10 * 1 Or fraSplit.Top + Y >= picCenter.ScaleHeight / 10 * 9 Then Exit Sub
        vsItem(Index_DEL).Height = vsItem(Index_DEL).Height + Y
        vsItem(Index_Add).Height = vsItem(Index_Add).Height + Y
        fraItem.Height = fraItem.Height + Y
        fraSplit.Top = fraSplit.Top + Y
        fraAdvice.Top = fraAdvice.Top + Y
        fraAdvice.Height = fraAdvice.Height - Y
        vsAdvice(Index_DEL).Height = vsAdvice(Index_DEL).Height - Y
        vsAdvice(Index_Add).Height = vsAdvice(Index_Add).Height - Y
    End If
End Sub

Private Sub fraSplitItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next

        If fraSplitItem.Left + X <= picCenter.ScaleWidth / 10 * 1 Or fraSplitItem.Left + X >= picCenter.ScaleWidth / 10 * 9 Then Exit Sub
        vsItem(Index_DEL).Width = vsItem(Index_DEL).Width + X
        fraSplitItem.Left = fraSplitItem.Left + X
        vsItem(Index_Add).Left = vsItem(Index_Add).Left + X
        vsItem(Index_Add).Width = vsItem(Index_Add).Width - X
    End If
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    cmdSend.Move picBottom.ScaleWidth - cmdSend.Width - 500
End Sub

Private Sub picCenter_Resize()
    On Error Resume Next
    Dim lngHeight As Long

    With picCenter
        lngHeight = (.ScaleHeight - 30) / 2
        '分隔线
        fraSplit.Move 0, lngHeight, .ScaleWidth, 30
        '分隔线上
        fraItem.Move 0, 0, .ScaleWidth, lngHeight
        vsItem(Index_DEL).Move 0, 600, .ScaleWidth / 2, lngHeight - 600
        fraSplitItem.Move vsItem(Index_DEL).Width, 600, 60, .Height
        vsItem(Index_Add).Move fraSplitItem.Left + 60, 600, .ScaleWidth / 2 - 30, lngHeight - 600
        '分隔线下
        fraAdvice.Move 0, lngHeight + fraSplit.Height + 60, .ScaleWidth, lngHeight
        vsAdvice(Index_DEL).Move 0, 600, .ScaleWidth / 2, lngHeight - 600
        fraSplitAdvice.Move vsAdvice(Index_DEL).Width, 600, 60, .Height
        vsAdvice(Index_Add).Move fraSplitAdvice.Left + 60, 600, .ScaleWidth / 2 - 30, lngHeight - 600
    End With
End Sub

Public Sub ShowMe(frmParent As Object, ByVal lng路径ID As Long, ByRef str分类 As String, ByRef str编码 As String, ByRef blnRefresh As Boolean)
'功能:
'参数:lng路径ID-当前默认选中的路径ID
'     str分类-选中的路径分类
'     blnRefresh=True 需要刷新主窗体
'     gstrPrivs-操作路径的权限
    mlng路径ID = lng路径ID
    mstr分类 = str分类
    mstrPrivs = gstrPrivs
    mblnSend = False

    Me.Show 1, frmParent
    blnRefresh = mblnSend
    str编码 = mstr编码
    str分类 = mstr分类
End Sub

Private Sub LoadPathName(ByVal lngCmd As Long, Optional ByVal strInput As String)
'功能:加载某分类下的路径名称，随分类的变化而变化
'参数:lngcmd 0-初始加载时，根据传人的mlng路径ID定位路径名称
'            1-选择分类时，默认选择该分类下第一个路径名称
'            2-路径名称中输入:汉字，编码 进行右匹配。
'     strInput -当lngCmd=2时，传人；输入的要匹配的字符
    Dim i As Long
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng原路径ID As Long
    Dim blnOK As Boolean
    
    On Error GoTo errH

    strInput = Trim(strInput)
    lng原路径ID = mlng路径ID
    
    If lngCmd = 0 Then
        strTmp = "and ID=" & mlng路径ID
        blnOK = True
    ElseIf lngCmd = 1 Then
        strTmp = "and 分类= '" & mstr分类 & "' and RowNum <2"
        blnOK = True
    Else
        If strInput <> "" Then
            '根据输入的内容判断，如果是是汉字则查找名称，非数字则查找编码
            If zlCommFun.IsCharChinese(strInput) Then
                '包含汉字 查找名称
                strTmp = "and 分类= '" & mstr分类 & "' and 名称 like '" & gstrLike & strInput & "%'"
            Else
                strTmp = "and 分类= '" & mstr分类 & "' and 编码 like '" & gstrLike & UCase(strInput) & "%'"
            End If
        Else
            strTmp = "and 分类= '" & mstr分类 & "'"
        End If
    End If

    strSql = "Select a.Id,a.编码,a.名称,最新版本 From 临床路径目录 A Where a.最新版本 >= 1 " & strTmp
    If InStr(mstrPrivs, "全院路径") = 0 Then
        '没有权限时，只能对只应用于本科的路径进行处理
        strSql = strSql & "And A.通用 = 2 And Exists" & vbNewLine & _
                 "      (Select 1 From 部门人员 C,临床路径科室 D " & vbNewLine & _
                 "       Where C.人员id = [1] and D.科室id = C.部门id And D.路径id = A.ID  )"
    End If
    strSql = strSql & " order by  分类,编码"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    If lngCmd = 2 Then
        If rsTmp.RecordCount = 1 Then
            blnOK = True
        ElseIf rsTmp.RecordCount > 1 Then
            If zlDatabase.zlShowListSelect(Me, glngSys, glngModul, txtPathName, rsTmp, True, , "1", rsTmp) Then
                blnOK = True
            End If
        Else
            MsgBox "未找到与输入内容匹配的路径，缺省选中原路径", vbInformation + vbOKOnly, Me.Caption
        End If
    End If

    If blnOK Then
        txtPathName.Text = rsTmp!名称
        txtPathName.Tag = rsTmp!名称
        mlng路径ID = rsTmp!ID
        mlng版本号 = rsTmp!最新版本
        mstr编码 = rsTmp!编码
    Else
       
        txtPathName.Text = txtPathName.Tag '未选择路径时，保持原来路径名称
    End If
    
    txtPathName.SelStart = Len(Trim(txtPathName.Text))
    If lng原路径ID <> mlng路径ID Or lngCmd = 0 Then
        Call LoadBranch
        If Not mrsStep Is Nothing Then
            Call ClearData
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCategory()
'功能:加载路径分类
    Dim i As Long
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    strSql = "Select distinct a.分类 From 临床路径目录 A Where a.最新版本 >= 1 "
    If InStr(mstrPrivs, "全院路径") = 0 Then
        '没有权限时，只能对只应用于本科的路径进行处理
        strSql = strSql & "And A.通用 = 2 And Exists" & vbNewLine & _
                "      (Select 1 From 部门人员 C,临床路径科室 D " & vbNewLine & _
                 "       Where C.人员id = [1] and D.科室id = C.部门id And D.路径id = A.ID )"
    End If
    strSql = strSql & " order by  分类"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    cboCategory.Clear
    For i = 1 To rsTmp.RecordCount
        cboCategory.AddItem rsTmp!分类
        rsTmp.MoveNext
    Next
    '缺省分类
    Call Cbo.Locate(cboCategory, mstr分类, False)  '会触发cboCategory_Click事件
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPathName_GotFocus()
    Call zlControl.TxtSelAll(txtPathName)
End Sub

Private Sub txtPathName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call LoadPathName(2, txtPathName.Text)
    End If
End Sub

Private Sub txtPathName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then KeyAscii = 0
End Sub

Private Sub InitStep()
'功能:初始化临床路径阶段表单
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    strcol = "阶段,2000,4;选择,500,4;详情,2500,4"
    arrHead = Split(strcol, ";")
    With vsStep
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows + 1    '缺省显示一行空白
        .Editable = flexEDNone

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            If Split(arrHead(i), ",")(0) = "选择" Then .ColDataType(i) = flexDTBoolean
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Row = 0  '不选择任何行
        .Redraw = True
    End With
End Sub

Private Sub InitVSItem(ByRef vsItem As VSFlexGrid, ByVal strHeads As String)
'功能:初始化vsItem临床路径项目表单
    Dim arrHead As Variant
    Dim arrHeads As Variant
    Dim lngRow As Long
    Dim i As Long
    Dim k As Long
    
    arrHeads = Split(strHeads, "|")
    If UBound(arrHeads) < 0 Then Exit Sub
    With vsItem
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 2: .FixedCols = 1
        .Cols = UBound(Split(arrHeads(0), ";")) + 1
        .Rows = .FixedRows + 1    '缺省显示一行空白
        .Editable = flexEDNone  '不容许编辑
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        lngRow = 0
        For k = LBound(arrHeads) To UBound(arrHeads)
            arrHead = Split(arrHeads(k), ";")
            For i = 0 To UBound(arrHead)
                .TextMatrix(lngRow, i) = Split(arrHead(i), ",")(0)
                If Split(arrHead(i), ",")(0) = "选择" Then .ColDataType(i) = flexDTBoolean
                If UBound(Split(arrHead(i), ",")) > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(i) = True
                    .ColWidth(i) = 0
                End If
            Next
            lngRow = lngRow + 1
        Next
        .MergeRow(0) = True
        .MergeCol(0) = True

        .Row = 0
        .Redraw = True
    End With
End Sub

Private Sub InitVSAdvice(ByRef vsAdvice As VSFlexGrid, ByVal strHeads As String)
'功能:初始化临床路径医嘱表单
    Dim arrHead As Variant
    Dim arrHeads As Variant
    Dim lngRow As Long
    Dim i As Long
    Dim k As Long
    
    arrHeads = Split(strHeads, "|")
    If UBound(arrHeads) < 0 Then Exit Sub
    With vsAdvice
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 2: .FixedCols = 2
        .Cols = UBound(Split(arrHeads(0), ";")) + 1
        .Rows = .FixedRows + 1    '缺省显示一行空白
        .MergeCells = flexMergeFree
        .Editable = flexEDKbdMouse
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        lngRow = 0
        For k = LBound(arrHeads) To UBound(arrHeads)
            arrHead = Split(arrHeads(k), ";")
            For i = 0 To UBound(arrHead)
                .TextMatrix(lngRow, i) = Split(arrHead(i), ",")(0)
                If Split(arrHead(i), ",")(0) = "选择" Then .ColDataType(i) = flexDTBoolean
                If UBound(Split(arrHead(i), ",")) > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(i) = True
                    .ColWidth(i) = 0
                End If
            Next
            lngRow = lngRow + 1
        Next
 
        '固定行列合并处理
        .MergeRow(0) = True
        .MergeCol(0) = True

        .Editable = flexEDNone  '不容许编辑
        .Row = 0
        .Redraw = True
    End With
End Sub

Private Sub txtRate_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtRate(Index))
End Sub

Private Sub txtRate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txtRate.UBound Then
            txtRate(Index + 1).SetFocus
        ElseIf Index = txtRate.UBound Then
            cmdAnalyse.SetFocus
        End If
    End If
End Sub

Private Sub txtRate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub

    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    ElseIf IsNumeric(Chr(KeyAscii)) Then
        '第一位不能为0
        If txtRate(Index).SelStart = 0 And Chr(KeyAscii) = "0" Then
            KeyAscii = 0
        ElseIf txtRate(Index).SelStart = 2 And Chr(KeyAscii) <> "0" Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub SetSummaryInfo()
'功能:设置汇总信息,包括总病例数，合格病例数（正常结束的病人），标准住院日，平均标准住院日
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String

    strSql = "  Select Count(1) As 总病例数, Sum(Decode(a.状态, 2, 1, 0)) As 合格病例数, b.标准住院日," & _
             "  To_Char(Sum(Decode(a.状态, 2, a.当前天数)) / Sum(Decode(a.状态, 2, 1, 0)), '99999.0') As 平均标准住院日" & _
             "  From 病人临床路径 A, 临床路径版本 B " & _
             "  Where a.路径id = b.路径id And a.版本号 = b.版本号 And b.路径id = [1] and b.版本号= [2] " & _
             "  and a.导入时间 between [3] and [4] " & _
             "  Group By b.标准住院日"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号, _
                                         CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")))
    If rsTmp Is Nothing Then Exit Sub
    
    If rsTmp.RecordCount = 1 Then
        With rsTmp
            lblInfo(0).Tag = Val(!合格病例数 & "")
            lblInfo(0).Caption = "该临床路径的总病例数：" & !总病例数 & " 人，正常结束病例数为：" & Val(!合格病例数 & "") & " 人。"
            '标准住院日：<=N天；M-N天
            lblInfo(1).Caption = "原标准住院日为：" & !标准住院日 & " 天，现平均标准住院日为：" & Val(!平均标准住院日 & "") & " 天。"
        End With
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPhase()
'功能:加载阶段提前或延后比例高于指定值的阶段
    Dim rsStep As ADODB.Recordset
    Dim strSql As String
    Dim strIDs As String
    Dim i As Long
    
    For i = 1 To mrsStep.RecordCount
        strIDs = strIDs & "," & mrsStep!ID
        mrsStep.MoveNext
    Next
    strIDs = Mid(strIDs, 2)
    
    On Error GoTo errH:
    
    strSql = "Select a.阶段id, To_Char((Sum(Decode(a.时间进度, 1, 1)) / Count(Distinct a.路径记录id)) * 100, '990.00') As 提前率," & vbNewLine & _
            "       To_Char((Sum(Decode(a.时间进度, -1, 1)) / Count(Distinct a.路径记录id)) * 100, '999.00') As 延后率" & vbNewLine & _
            "From (Select Distinct b.阶段id, b.路径记录id, Decode(b.时间进度,2,1,b.时间进度) as 时间进度" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径评估 B, Table(f_Str2list([5])) C" & vbNewLine & _
            "       Where a.Id = b.路径记录id And a.路径id = [1] And a.版本号 = [2] And a.状态 = 2  And b.阶段id = c.Column_Value And" & vbNewLine & _
            "             a.导入时间 Between [3] And [4]) A" & vbNewLine & _
            "Group By a.阶段id" & vbNewLine & _
            "Having(Sum(Decode(a.时间进度, 1, 1)) / Count(Distinct a.路径记录id)) * 100 >= [6] Or (Sum(Decode(a.时间进度, -1, 1)) / Count(Distinct a.路径记录id)) * 100 > [6]"
            
    Set rsStep = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), strIDs, Val(txtRate(RATE_STEP).Text))
    With vsStep
        .Rows = rsStep.RecordCount + .FixedRows
        For i = 1 To rsStep.RecordCount
            mrsStep.Filter = "ID =" & rsStep!阶段id
            .TextMatrix(i, COL_阶段) = mrsStep!名称 & IIf(Nvl(mrsStep!父ID) = "", "", ",分支:" & Nvl(mrsStep!说明, mrsStep!序号))
            .RowData(i) = IIf(IsNull(mrsStep!分支ID), mlng路径ID, Nvl(mrsStep!分支ID))
            .TextMatrix(i, COL_详情) = IIf(IsNull(rsStep!提前率), "（无提前)", rsStep!提前率 & "%提前") & "/" & IIf(IsNull(rsStep!延后率), "(无延后)", rsStep!延后率 & "%延后")
            rsStep.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub LoadItem()
'功能:加载临床路径取消和增加的项目
    Dim strSql As String
    Dim rsItemRate As ADODB.Recordset
    Dim lngTmp As Long
    Dim i As Long, k As Long
    
    On Error GoTo errH
    
    '1、 取消项目部分
    
    '未使用比例=所有使用该阶段的未生成非医嘱类项目的病人数/所有使用该阶段的病人数；
    '未使用比例=(1 - Nvl(d.病人数, 1) / Nvl(e.病人数, 1)) * 100
    '为未生成的比例:Val(txtRate(0).Text)

    strSql = "Select d.阶段id, d.Id, d.项目内容, To_Char((1 - d.病人数 / Nvl(e.病人数, 1)) * 100, '990.00') As 未使用比例" & vbNewLine & _
            "From (Select b.阶段id, b.Id, b.项目内容, Nvl(病人数, 0) As 病人数" & vbNewLine & _
            "       From (Select a.阶段id, a.项目内容, Count(Distinct c.病人id) As 病人数" & vbNewLine & _
            "              From 临床路径项目 A, 病人路径执行 B, 病人临床路径 C" & vbNewLine & _
            "              Where a.Id = b.项目id And b.路径记录id = c.Id And  a.路径id = [1] And a.版本号 = [2] And c.状态 = 2 And" & vbNewLine & _
            "                    c.导入时间 Between [3] And [4] And Not Exists" & vbNewLine & _
            "               (Select 1 From 临床路径医嘱 T Where t.路径项目id = a.Id)" & vbNewLine & _
            "              Group By a.阶段id, a.项目内容) A, 临床路径项目 B" & vbNewLine & _
            "       Where b.路径id = [1] And b.版本号 = [2] And b.阶段id = a.阶段id(+) And b.项目内容 = a.项目内容(+) And Not Exists" & vbNewLine & _
            "        (Select 1 From 临床路径医嘱 T Where t.路径项目id = b.Id)) D," & vbNewLine & _
            "     (Select b.阶段id, Count(Distinct a.病人id) As 病人数" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
            "       Where a.Id = b.路径记录id And a.路径id = [1] And a.版本号 = [2] And a.状态 = 2 And" & vbNewLine & _
            "             a.导入时间 Between [3] And [4] " & vbNewLine & _
            "       Group By b.阶段id) E" & vbNewLine & _
            "Where d.阶段id = e.阶段id(+) And (1 - d.病人数 / Nvl(e.病人数, 1)) * 100 >= [5]"


    Set rsItemRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_UNSEND).Text))
     '将数据插入到取消项目列表中
    With vsItem(Index_DEL)
        .Redraw = flexRDNone
        .Rows = .FixedRows '清空数据
        .Rows = .FixedRows + 1  '无数据时默认空一行
        .Rows = rsItemRate.RecordCount + .FixedRows
        lngTmp = .FixedRows  '记录插入行
        mrsStep.Filter = ""  '恢复所有记录
        For i = 1 To mrsStep.RecordCount  '按阶段顺序填充表列数据
            '当前阶段取消项目填充
            rsItemRate.Filter = "阶段id=" & mrsStep!ID
            For k = 1 To rsItemRate.RecordCount
                '隐藏部分
                .RowData(lngTmp) = IIf(IsNull(mrsStep!分支ID), mlng路径ID, Nvl(mrsStep!分支ID))
                .Cell(flexcpData, lngTmp, COL_Item_项目名称) = CStr(rsItemRate!ID)
                .Cell(flexcpData, lngTmp, COL_Item_阶段) = CStr(rsItemRate!阶段id)
                
                '显示部分
                .TextMatrix(lngTmp, COL_Item_阶段) = mrsStep!名称 & IIf(IsNull(mrsStep!父ID), "", ",分支:" & Nvl(mrsStep!说明, mrsStep!序号))
                .TextMatrix(lngTmp, COL_Item_项目名称) = rsItemRate!项目内容 & ""
                .TextMatrix(lngTmp, COL_Item_比例) = rsItemRate!未使用比例 & ""
                lngTmp = lngTmp + 1
                
                rsItemRate.MoveNext
            Next
            mrsStep.MoveNext
        Next

        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Item_项目名称, .Rows - 1, COL_Item_项目名称) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效
    End With
    
    '2、增加项目部分
    '使用比例=该阶段所有添加该非医嘱类路径外项目的病人数/所有使用该阶段的病人数。
    '使用比例=(Nvl(d.病人数, 1) / Nvl(e.病人数, 1)) * 100
    strSql = " Select d.阶段id, d.分类,d.项目内容, d.执行id, To_Char((Nvl(d.病人数, 1) / Nvl(e.病人数, 1)) * 100, '990.00') As 使用比例" & _
            "   From (Select b.阶段id,b.分类, b.项目内容, Max(b.Id) As 执行id, Count(Distinct c.病人id) As 病人数" & _
            "       From 病人路径执行 B, 病人临床路径 C" & _
            "       Where b.路径记录id = c.Id And c.路径id = [1] And c.版本号 = [2] And c.状态 = 2 And" & _
            "             c.导入时间 Between [3] And [4] And b.项目id Is Null And b.项目内容 <> '未生成任何项目' And" & _
            "             b.项目内容 <> '路径外项目' And Not Exists (Select 1 From 病人路径医嘱 T Where t.路径执行id = b.Id)" & _
            "       Group By b.阶段id, b.分类,b.项目内容) D," & _
            "     (Select b.阶段id, Count(Distinct a.病人id) As 病人数" & _
            "       From 病人临床路径 A, 病人路径执行 B" & _
            "       Where a.Id = b.路径记录id And a.路径id = [1] And a.版本号 = [2] And a.状态 = 2 And" & _
            "             a.导入时间 Between [3] And [4] " & _
            "       Group By b.阶段id) E" & _
            "   Where d.阶段id = e.阶段id And (Nvl(d.病人数, 1) / Nvl(e.病人数, 1)) * 100 >= [5]"

            
    Set rsItemRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_PATHOUT).Text))

    '将数据插入到增加表列中
    With vsItem(Index_Add)
        .Redraw = flexRDNone
        .Rows = .FixedRows '清空数据
        .Rows = .FixedRows + 1  '无数据时默认空一行
        .Rows = rsItemRate.RecordCount + .FixedRows
        
        lngTmp = .FixedRows  '记录插入行
        mrsStep.Filter = ""  '恢复所有记录
        For i = 1 To mrsStep.RecordCount  '按阶段顺序填充表列数据
         '当前阶段增加项目填充
            rsItemRate.Filter = "阶段id=" & mrsStep!ID
            For k = 1 To rsItemRate.RecordCount
                '隐藏部分
                .RowData(lngTmp) = IIf(IsNull(mrsStep!分支ID), mlng路径ID, Nvl(mrsStep!分支ID))
                .Cell(flexcpData, lngTmp, COL_Item_项目名称) = CStr(rsItemRate!执行ID) '增加项目中的项目ID存放路径执行ID,便于SelectVsItem中一并处理
                .Cell(flexcpData, lngTmp, COL_Item_阶段) = rsItemRate!阶段id & ""
                .Cell(flexcpData, lngTmp, COL_Item_分类) = rsItemRate!分类 & ""
                '显示部分
                .TextMatrix(lngTmp, COL_Item_阶段) = mrsStep!名称 & IIf(IsNull(mrsStep!父ID), "", ",分支:" & Nvl(mrsStep!说明, mrsStep!序号))
                .TextMatrix(lngTmp, COL_Item_项目名称) = rsItemRate!项目内容 & ""
                .TextMatrix(lngTmp, COL_Item_比例) = rsItemRate!使用比例 & ""
                lngTmp = lngTmp + 1
                rsItemRate.MoveNext
            Next
            mrsStep.MoveNext
        Next
     
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Item_项目名称, .Rows - 1, COL_Item_项目名称) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Function MakeStepRS() As ADODB.Recordset
'功能:自定义记录集,用来封装阶段信息
'参数:
    Set MakeStepRS = New ADODB.Recordset
    
    MakeStepRS.Fields.Append "ID", adBigInt
    MakeStepRS.Fields.Append "父ID", adBigInt, , adFldIsNullable
    MakeStepRS.Fields.Append "分支ID", adBigInt, , adFldIsNullable
    MakeStepRS.Fields.Append "名称", adVarChar, 100, adFldIsNullable
    MakeStepRS.Fields.Append "序号", adBigInt, , adFldIsNullable
    MakeStepRS.Fields.Append "说明", adVarChar, 100, adFldIsNullable
    
    MakeStepRS.CursorLocation = adUseClient
    MakeStepRS.LockType = adLockOptimistic
    MakeStepRS.CursorType = adOpenStatic
    MakeStepRS.Open
End Function

Private Sub LoadAdvice()
'功能:加载临床路径取消和增加的医嘱
    Dim strSql As String
    Dim rsAdviceRate As ADODB.Recordset
    Dim lngTmp As Long
    
    Dim i As Long, k As Long, j As Long
    
    On Error GoTo errH
    
    '取消医嘱部分
    
    '未使用比例=所有使用该阶段的未生成该路径医嘱的病人数/所有使用该阶段的病人数；
    '未使用比例=(1 - d.病人数 / Nvl(e.病人数, 1)) * 100
            
    strSql = "Select d.阶段id, d.项目内容, d.医嘱id, d.相关id, d.医嘱内容, d.期效, d.序号, d.诊疗项目id, d.标本部位, d.类别, d.名称," & vbNewLine & _
            "       To_Char((1 - d.病人数 / Nvl(e.病人数, 1)) * 100, '990.00') As 未使用比例" & vbNewLine & _
            "From (Select a.阶段id, a.项目内容, e.Id As 医嘱id, e.相关id, e.医嘱内容, e.期效, e.序号, e.诊疗项目id, e.标本部位, f.类别," & vbNewLine & _
            "              Nvl(g.名称 || Decode(g.规格, Null, Null, ' ' || g.规格), f.名称) As 名称, Nvl(病人数, 0) As 病人数" & vbNewLine & _
            "       From (Select a.阶段id, a.项目内容, Count(Distinct c.病人id) As 病人数" & vbNewLine & _
            "              From 临床路径项目 A, 病人路径执行 B, 病人临床路径 C" & vbNewLine & _
            "              Where a.Id = b.项目id And b.路径记录id = c.Id And a.路径id = [1] And a.版本号 = [2] And c.状态 = 2 And" & vbNewLine & _
            "                    c.导入时间 Between [3] And [4] And Exists" & vbNewLine & _
            "               (Select 1 From 临床路径医嘱 T Where t.路径项目id = a.Id)" & vbNewLine & _
            "              Group By a.阶段id, a.项目内容) H, 临床路径项目 A, 临床路径医嘱 D, 路径医嘱内容 E, 诊疗项目目录 F, 收费项目目录 G" & vbNewLine & _
            "       Where a.阶段id = h.阶段id(+) And a.项目内容 = h.项目内容(+) And a.路径id = [1] And a.版本号 = [2] And Exists" & vbNewLine & _
            "        (Select 1 From 临床路径医嘱 T Where t.路径项目id = a.Id) And a.Id = d.路径项目id And d.医嘱内容id = e.Id And" & vbNewLine & _
            "             e.诊疗项目id = f.Id(+) And Nvl(e.收费细目id, -1) = g.Id(+) and Not (e.组合项目ID is not null and f.类别='C')) D," & vbNewLine & _
            "     (Select b.阶段id, Count(Distinct a.病人id) As 病人数" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
            "       Where a.Id = b.路径记录id And a.路径id = [1] And a.版本号 = [2] And a.状态 = 2 And" & vbNewLine & _
            "             a.导入时间 Between [3] And [4]" & vbNewLine & _
            "       Group By b.阶段id) E" & vbNewLine & _
            "Where d.阶段id = e.阶段id(+) And (1 - d.病人数 / Nvl(e.病人数, 1)) * 100 >= [5]" & vbNewLine & _
            "Order By 医嘱id"


    Set rsAdviceRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_UNSEND).Text))

    '将数据插入到vsAdviceDel表列中
    With vsAdvice(Index_DEL)
        .Redraw = flexRDNone
        .Rows = .FixedRows  '清空上一次的数据
        .Rows = .FixedRows + 1 '没数据的时候默认显示一行空白
        .Rows = .FixedRows + rsAdviceRate.RecordCount
        lngTmp = .FixedRows  '记录插入行
        mrsStep.Filter = ""  '恢复所有记录
        For i = 1 To mrsStep.RecordCount  '按阶段顺序填充表列数据
            rsAdviceRate.Filter = "阶段id=" & mrsStep!ID
            For k = 1 To rsAdviceRate.RecordCount
                .RowData(lngTmp) = IIf(IsNull(mrsStep!分支ID), mlng路径ID, Nvl(mrsStep!分支ID))
                .Cell(flexcpData, lngTmp, COL_Advice_阶段) = rsAdviceRate!阶段id & ""
                .Cell(flexcpData, lngTmp, COL_Advice_项目名称) = rsAdviceRate!项目内容 & ""
                .Cell(flexcpData, lngTmp, COL_Advice_医嘱内容) = rsAdviceRate!序号 & "" '医嘱内容隐藏数据存储医嘱序号
                
                .TextMatrix(lngTmp, COL_Advice_阶段) = mrsStep!名称 & IIf(IsNull(mrsStep!父ID), "", ",分支:" & Nvl(mrsStep!说明, mrsStep!序号))
                .TextMatrix(lngTmp, COL_Advice_医嘱ID) = rsAdviceRate!医嘱id
                .TextMatrix(lngTmp, COL_Advice_相关ID) = IIf(IsNull(rsAdviceRate!相关id), 0, rsAdviceRate!相关id)
                .TextMatrix(lngTmp, COL_Advice_期效) = IIf(rsAdviceRate!期效 = 1, "临嘱", "长嘱")
                .TextMatrix(lngTmp, COL_Advice_医嘱内容) = IIf(rsAdviceRate!医嘱内容 & "" = "", rsAdviceRate!名称, rsAdviceRate!医嘱内容)
                .TextMatrix(lngTmp, COL_Advice_诊疗项目ID) = rsAdviceRate!诊疗项目ID & ""
                .TextMatrix(lngTmp, COL_Advice_标本部位) = rsAdviceRate!标本部位 & ""
                .TextMatrix(lngTmp, COL_Advice_比例) = rsAdviceRate!未使用比例 & ""
                .TextMatrix(lngTmp, COL_Advice_诊疗类别) = rsAdviceRate!类别 & ""
                lngTmp = lngTmp + 1
                rsAdviceRate.MoveNext
            Next
            mrsStep.MoveNext
        Next
        
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Advice_医嘱内容, .Rows - 1, COL_Advice_医嘱内容) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        .AutoSize .FixedCols, .Cols - 1, , 45
    End With
    '增加医嘱部分
    '使用比例=该阶段所有添加该医嘱类路径外项目的病人数/所有使用该阶段的病人数。
    '使用比例=(使用病人数/总病人数)*100
    strSql = "Select a.阶段id, a.分类, a.诊疗项目id,a.执行ID,a.医嘱ID,c.相关id,c.医嘱期效 as 期效,c.诊疗类别,c.医嘱内容,c.标本部位,To_Char(a.病人数 / b.病人数 * 100, '900.00') As 使用比例" & vbNewLine & _
            "From (Select b.阶段id, b.分类, d.诊疗项目id,Max(b.ID) as 执行ID, Count(Distinct a.病人id) As 病人数,Max(d.ID) as 医嘱ID" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径执行 B, 病人路径医嘱 C, 病人医嘱记录 D" & vbNewLine & _
            "       Where a.Id = b.路径记录id And b.Id = c.路径执行id And c.病人医嘱id = d.Id And a.路径id = [1] And a.版本号 = [2] And a.状态 = 2 And" & vbNewLine & _
            "             a.导入时间 Between [3] And [4] And b.项目id Is Null And d.诊疗类别 <> 'E' And" & vbNewLine & _
            "             Not (d.相关id Is Not Null And d.诊疗类别 In ('F', 'G', 'D'))" & vbNewLine & _
            "       Group By b.阶段id, b.分类, d.诊疗项目id) A," & vbNewLine & _
            "     (Select b.阶段id, Count(Distinct a.病人id) As 病人数" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
            "       Where a.Id = b.路径记录id And a.路径id = [1] And a.版本号 = [2] And a.状态 = 2 And" & vbNewLine & _
            "             a.导入时间 Between [3] And [4]" & vbNewLine & _
            "       Group By b.阶段id) B,病人医嘱记录 c" & vbNewLine & _
            "Where a.阶段id = b.阶段id and a.医嘱ID=c.id And c.组合项目id Is  Null and a.病人数 / b.病人数 * 100>=[5] " & _
            " order by 医嘱ID "
 
    Set rsAdviceRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_PATHOUT).Text))

     '将数据插入满足要求的数据插入到vsAdviceAdd表列中
    With vsAdvice(Index_Add)
        .Redraw = flexRDNone
        .Rows = .FixedRows  '清空上一次的数据
        .Rows = .FixedRows + 1 '没数据的时候默认显示一行空白
        .Rows = .FixedRows + rsAdviceRate.RecordCount
        lngTmp = .FixedRows  '记录插入行
        mrsStep.Filter = ""  '恢复所有记录
        For i = 1 To mrsStep.RecordCount  '按阶段顺序填充表列数据
            rsAdviceRate.Filter = "阶段id=" & mrsStep!ID
            For k = 1 To rsAdviceRate.RecordCount
                '表格数据填充
                .RowData(lngTmp) = IIf(IsNull(mrsStep!分支ID), mlng路径ID, Nvl(mrsStep!分支ID))
                .Cell(flexcpData, lngTmp, COL_Advice_阶段) = rsAdviceRate!阶段id & ""
                .Cell(flexcpData, lngTmp, COL_Advice_诊疗类别) = rsAdviceRate!分类 & ""
                .Cell(flexcpData, lngTmp, COL_Advice_医嘱内容) = rsAdviceRate!执行ID & ""
                
                .TextMatrix(lngTmp, COL_Advice_阶段) = mrsStep!名称 & IIf(IsNull(mrsStep!父ID), "", ",分支:" & Nvl(mrsStep!说明, mrsStep!序号))
                .TextMatrix(lngTmp, COL_Advice_医嘱ID) = rsAdviceRate!医嘱id
                .TextMatrix(lngTmp, COL_Advice_相关ID) = Nvl(rsAdviceRate!相关id, 0)
                .TextMatrix(lngTmp, COL_Advice_期效) = IIf(rsAdviceRate!期效 = 1, "临嘱", "长嘱")
                If rsAdviceRate!诊疗类别 & "" = "C" Then
                    .TextMatrix(lngTmp, COL_Advice_医嘱内容) = rsAdviceRate!医嘱内容 & "（" & rsAdviceRate!标本部位 & ")"
                Else
                    .TextMatrix(lngTmp, COL_Advice_医嘱内容) = rsAdviceRate!医嘱内容 & ""
                End If
                .TextMatrix(lngTmp, COL_Advice_比例) = rsAdviceRate!使用比例
                lngTmp = lngTmp + 1
                rsAdviceRate.MoveNext
            Next
            mrsStep.MoveNext
        Next
        
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Advice_医嘱内容, .Rows - 1, COL_Advice_医嘱内容) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        .AutoSize .FixedCols, .Cols - 1, , 45
        '一并给药相关行线的处理
    End With
   
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetPathPhase()
'功能:获取当前路径阶段信息
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
        
    On Error GoTo errH
    strSql = "Select a.Id , a.父id,a.分支id, a.名称,Decode(b.序号, Null, 0, a.序号) As 序号, a.说明" & _
            "   From 临床路径阶段 A, 临床路径阶段 B" & _
            "   Where a.父id = b.Id(+)   And a.路径id = [1] And a.版本号 =[2]" & _
            "   Order By Nvl(a.分支ID,0), Nvl(b.序号, a.序号), Decode(b.序号, Null, 0, a.序号)"


    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号)
    Set mrsStep = MakeStepRS
    For i = 1 To rsTmp.RecordCount
        mrsStep.AddNew
        mrsStep!ID = rsTmp!ID
        mrsStep!父ID = rsTmp!父ID
        mrsStep!分支ID = rsTmp!分支ID
        mrsStep!名称 = rsTmp!名称
        mrsStep!序号 = rsTmp!序号
        mrsStep!说明 = rsTmp!说明
        rsTmp.MoveNext
    Next
    If mrsStep.RecordCount > 0 Then mrsStep.Update: mrsStep.MoveFirst
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, ByVal blnIsHide As Boolean, ByVal vsfThis As VSFlexGrid) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'参数：blnIsHide=范围是否包含隐藏的行
    Dim i As Long, blnTmp As Boolean
    
    With vsfThis
        
        If .TextMatrix(lngRow, COL_Advice_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_Advice_诊疗类别)) = 0 Then Exit Function
        
        If Val(.TextMatrix(lngRow - 1, COL_Advice_相关ID)) = Val(.TextMatrix(lngRow, COL_Advice_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_Advice_相关ID)) = Val(.TextMatrix(lngRow, COL_Advice_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_Advice_相关ID)) = Val(.TextMatrix(lngRow, COL_Advice_相关ID)) And Val(.TextMatrix(lngRow, COL_Advice_相关ID)) <> 0 _
                    Or ((Val(.TextMatrix(lngRow, COL_Advice_相关ID)) = Val(.TextMatrix(i, COL_Advice_医嘱ID)) _
                    Or Val(.TextMatrix(i, COL_Advice_相关ID)) = Val(.TextMatrix(lngRow, COL_Advice_医嘱ID))) And blnIsHide) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_Advice_相关ID)) = Val(.TextMatrix(lngRow, COL_Advice_相关ID)) And Val(.TextMatrix(lngRow, COL_Advice_相关ID)) <> 0 _
                    Or ((Val(.TextMatrix(lngRow, COL_Advice_相关ID)) = Val(.TextMatrix(i, COL_Advice_医嘱ID)) _
                    Or Val(.TextMatrix(i, COL_Advice_相关ID)) = Val(.TextMatrix(lngRow, COL_Advice_医嘱ID))) And blnIsHide) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub txtRate_LostFocus(Index As Integer)
    If txtRate(Index).Text = "" Then
        MsgBox "比例值不能为空。", vbOKOnly + vbDefaultButton1, Me.Caption
        Call txtRate(Index).SetFocus
    End If
End Sub

Private Sub vsAdvice_DrawCell(Index As Integer, ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT

    With vsAdvice(Index)
        '擦除一并给药相关行列的边线及内容
        If Row < .FixedRows Then Exit Sub
        lngLeft = COL_Advice_期效: lngRight = COL_Advice_期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_Advice_比例: lngRight = COL_Advice_比例
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub

        If Not RowIn一并给药(Row, lngBegin, lngEnd, False, vsAdvice(Index)) Then Exit Sub

        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
         Call SelectVsAdvice(vsAdvice(Index), vsAdvice(Index).Row, COL_Advice_选择)
    End If
End Sub

Private Sub vsAdvice_LostFocus(Index As Integer)
    vsAdvice(Index).Row = 0
End Sub

Private Sub vsAdvice_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With vsAdvice(Index)
            If .MouseRow < .FixedRows Then Exit Sub
            If .MouseCol <> COL_Advice_选择 Then Exit Sub
            Call SelectVsAdvice(vsAdvice(Index), .Row, .Col)
        End With
    End If
End Sub

Private Sub vsItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call SelectVsItem(vsItem(Index), vsItem(Index).Row, COL_Item_选择)
    End If
End Sub

Private Sub vsItem_LostFocus(Index As Integer)
    vsItem(Index).Row = 0
End Sub

Private Sub vsItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With vsItem(Index)
            If .MouseRow < .FixedRows Then Exit Sub
            If .MouseCol <> COL_Item_选择 Then Exit Sub
            Call SelectVsItem(vsItem(Index), .Row, .Col)
        End With
    End If
End Sub

Private Sub SelectVsStep(ByVal lngRow As Long, ByVal lngCol As Long)
    With vsStep
        If COL_选择 = .Col And lngRow >= .FixedRows Then
            If .Cell(flexcpChecked, lngRow, COL_选择) = flexChecked Then
                .Cell(flexcpChecked, lngRow, COL_选择) = Unchecked
                .TextMatrix(lngRow, COL_选择) = "0"    '未选中
            Else
                .Cell(flexcpChecked, lngRow, COL_选择) = Checked
                .TextMatrix(lngRow, COL_选择) = "1"     '选中
            End If
        End If
    End With
End Sub

Private Sub SelectVsItem(ByVal vsItem As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'功能:用于勾选项目并标记选中项目和未选中项目便于通过FindRow方法确定是否存在未选择项目的情况。
    With vsItem
            If mrsStep Is Nothing Then Exit Sub
            If lngCol = COL_Item_选择 Then
                If .Cell(flexcpChecked, lngRow, lngCol) = flexChecked Then
                    .Cell(flexcpChecked, lngRow, lngCol) = flexUnchecked
                    .TextMatrix(lngRow, lngCol) = "0"       '标记未选中项
                Else
                    .Cell(flexcpChecked, lngRow, lngCol) = flexChecked
                    .TextMatrix(lngRow, lngCol) = "1"       '标记选中项
                End If
            End If
        End With
End Sub

Private Sub SelectVsAdvice(ByVal vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'功能:用于勾选项目并标记选中项目和未选中项目便于通过FindRow方法确定是否存在未选择项目的情况。

    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim i As Long
        
    With vsAdvice
        If Not mrsStep Is Nothing Then
            If lngRow < .FixedRows Then Exit Sub
        
            If lngCol = COL_Advice_选择 Then   '取消的医嘱
                If .Cell(flexcpChecked, lngRow, lngCol) = flexChecked Then
                  
                    Call RowIn一并给药(lngRow, lngBegin, lngEnd, False, vsAdvice)
                    If lngBegin = lngEnd Then lngBegin = lngRow: lngEnd = lngRow '非一并给药
                    For i = lngBegin To lngEnd
                       .Cell(flexcpChecked, i, lngCol) = flexUnchecked   '未选中
                       .TextMatrix(i, lngCol) = "0"     '标记未选中项
                    Next
                Else
                    Call RowIn一并给药(lngRow, lngBegin, lngEnd, False, vsAdvice)
                    If lngBegin = lngEnd Then lngBegin = lngRow: lngEnd = lngRow '非一并给药
                    For i = lngBegin To lngEnd
                        .Cell(flexcpChecked, i, lngCol) = flexChecked     '选中
                        .TextMatrix(i, lngCol) = "1"    '标记选中项
                    Next
                End If
            End If
        End If
    End With
End Sub

Private Sub vsStep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call SelectVsStep(vsStep.Row, COL_选择)
    End If
End Sub

Private Sub vsStep_LostFocus()
    vsStep.Row = 0
End Sub

Private Sub vsStep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsStep.MouseRow < vsStep.FixedRows Then Exit Sub
        If vsStep.MouseCol <> COL_选择 Then Exit Sub
        Call SelectVsStep(vsStep.MouseRow, vsStep.MouseCol)
    End If
End Sub

Private Sub LoadBranch()
'功能:加载分支路径信息
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    strSql = "Select a.id,a.名称 as 分支名称  " & vbNewLine & _
                "From 临床路径分支 A, 临床路径阶段 B, 临床路径阶段 C" & vbNewLine & _
                "Where a.前一阶段id = b.Id And b.父id = c.Id(+)" & vbNewLine & _
                "And a.路径id = [1] And a.版本号 = [2]" & vbNewLine & _
                "Order By Nvl(c.序号, b.序号), a.名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, mlng版本号)
    If rsTmp Is Nothing Then Exit Sub
    cboBranch.Clear
    cboBranch.AddItem "主路径"
    cboBranch.ItemData(0) = mlng路径ID
    For i = 1 To rsTmp.RecordCount
        cboBranch.AddItem "分支名称：" & rsTmp!分支名称
        cboBranch.ItemData(i) = rsTmp!ID
        rsTmp.MoveNext
    Next
    Call Cbo.SetIndex(cboBranch.Hwnd, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetVSRowHidden(ByVal vsGrid As VSFlexGrid, ByVal lngId As Long)
'功能:存在分支路径时，根据路径ID显示对应的阶段行
'参数：vsGrid表格对象
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim lngRow As Long
    Dim strTmp As String, str标本 As String, str煎法 As String, str麻醉 As String
    
    With vsGrid
        lngBegin = .FixedRows  '初始默认第一行
        For i = .FixedRows To .Rows - 1
           
            If cboBranch.ListCount > 1 Then  '存在分支路径
                If .RowData(i) = lngId Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            End If
            
             '再处理一些附加行的隐藏,及相关内容的显示
            If vsGrid.Tag = "取消医嘱" And Not .RowHidden(i) Then
            '给药途径
                If .TextMatrix(i, COL_Advice_诊疗类别) = "E" And Val(.TextMatrix(i, COL_Advice_相关ID)) = 0 _
                   And Val(.TextMatrix(i - 1, COL_Advice_相关ID)) = Val(.TextMatrix(i, COL_Advice_医嘱ID)) _
                   And InStr(",5,6,", .TextMatrix(i - 1, COL_Advice_诊疗类别)) > 0 Then
                    .RowHidden(i) = True
                End If
                
    
                '输血途径
                If .TextMatrix(i, COL_Advice_诊疗类别) = "E" And .TextMatrix(i - 1, COL_Advice_诊疗类别) = "K" _
                   And Val(.TextMatrix(i, COL_Advice_相关ID)) = Val(.TextMatrix(i - 1, COL_Advice_医嘱ID)) Then
                    .RowHidden(i) = True
                End If
    
                '中药配方和检验组合
                If .TextMatrix(i, COL_Advice_诊疗类别) = "E" And Val(.TextMatrix(i, COL_Advice_相关ID)) = 0 _
                   And Val(.TextMatrix(i - 1, COL_Advice_相关ID)) = Val(.TextMatrix(i, COL_Advice_医嘱ID)) _
                   And InStr(",7,E,C,", .TextMatrix(i - 1, COL_Advice_诊疗类别)) > 0 Then
    
                    str煎法 = "": str标本 = "": strTmp = ""
                    j = .FindRow(CStr(Val(.TextMatrix(i, COL_Advice_医嘱ID))), , COL_Advice_相关ID)
    
                    'j--组合检验项目的行号
                    For k = j To i - 1
                        .RowHidden(k) = k <> i
                        If .TextMatrix(k, COL_Advice_诊疗类别) = "C" Then
                            strTmp = strTmp & "," & .TextMatrix(k, COL_Advice_医嘱内容)
                            str标本 = .TextMatrix(j, COL_Advice_标本部位)    '取第一个检验项目的标本
                        ElseIf .TextMatrix(k, COL_Advice_诊疗类别) = "E" And Val(.TextMatrix(k, COL_Advice_相关ID)) <> 0 Then
                            str煎法 = .TextMatrix(k, COL_Advice_医嘱内容)
                        End If
                    Next
    
                    If .TextMatrix(i - 1, COL_Advice_诊疗类别) = "C" Then
                        .TextMatrix(i, COL_Advice_医嘱内容) = Mid(strTmp, 2) & IIf(str标本 <> "", "(" & str标本 & ")", "")
                    Else
                        .TextMatrix(i, COL_Advice_医嘱内容) = "中药配方," & str煎法 & "," & .TextMatrix(i, COL_Advice_医嘱内容)
                    End If
                End If
    
                '检查组合
                If .TextMatrix(i, COL_Advice_诊疗类别) = "D" And Val(.TextMatrix(i, COL_Advice_相关ID)) = 0 Then
                    str标本 = "": str煎法 = "": strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_Advice_相关ID)) = Val(.TextMatrix(i, COL_Advice_医嘱ID)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, COL_Advice_标本部位) <> "" _
                               And Val(.TextMatrix(j, COL_Advice_诊疗项目ID)) = Val(.TextMatrix(i, COL_Advice_诊疗项目ID)) Then    '相同的项目ID才是新方式
                                If .TextMatrix(j, COL_Advice_标本部位) <> strTmp And strTmp <> "" Then
                                    str标本 = str标本 & "," & strTmp & IIf(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                                    str煎法 = ""
                                End If
                                strTmp = .TextMatrix(j, COL_Advice_标本部位)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        str标本 = str标本 & "," & strTmp & IIf(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                    End If
                    If str标本 <> "" Then    '以前的检查方式时不显示详细医嘱内容
                        .TextMatrix(i, COL_Advice_医嘱内容) = .TextMatrix(i, COL_Advice_医嘱内容) & ":" & Mid(str标本, 2)
                    End If
                End If
    
                '手术项目
                If .TextMatrix(i, COL_Advice_诊疗类别) = "F" And Val(.TextMatrix(i, COL_Advice_相关ID)) = 0 Then
                    strTmp = "": str麻醉 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_Advice_相关ID)) = Val(.TextMatrix(i, COL_Advice_医嘱ID)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, COL_Advice_诊疗类别) = "F" Then
                                strTmp = strTmp & "," & .TextMatrix(j, COL_Advice_医嘱内容)
                            ElseIf .TextMatrix(j, COL_Advice_诊疗类别) = "G" Then
                                str麻醉 = .TextMatrix(j, COL_Advice_医嘱内容)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Or str麻醉 <> "" Then
                        If str麻醉 <> "" Then
                            .TextMatrix(i, COL_Advice_医嘱内容) = "在 " & str麻醉 & " 下行 " & .TextMatrix(i, COL_Advice_医嘱内容)
                        Else
                            .TextMatrix(i, COL_Advice_医嘱内容) = "行 " & .TextMatrix(i, COL_Advice_医嘱内容)
                        End If
                        If strTmp <> "" Then
                            .TextMatrix(i, COL_Advice_医嘱内容) = .TextMatrix(i, COL_Advice_医嘱内容) & " 及 " & Mid(strTmp, 2)
                        End If
                    End If
                End If
                     '相同阶段存在医嘱内容相同时，将医嘱内容处理为：医嘱名称（项目名称）的方式进行区分
                If i > .FixedRows And i <= .Rows - 1 Then
                    If .Cell(flexcpData, i, COL_Advice_阶段) <> .Cell(flexcpData, i - 1, COL_Advice_阶段) Or i = .Rows - 1 Then '上一阶段与下一阶段交接处
                        lngEnd = IIf(i = .Rows - 1, i, i - 1)
                        For j = lngBegin To lngEnd
                            If Not .RowHidden(j) Then
                                lngRow = .FindRow(.TextMatrix(j, COL_Advice_医嘱内容), j + 1, COL_Advice_医嘱内容) '由上往下找
                                '相同阶段区域内找到相同医嘱
                                If lngRow <> -1 And lngRow > lngBegin And lngRow <= lngEnd Then
                                    .TextMatrix(j, COL_Advice_医嘱内容) = .TextMatrix(j, COL_Advice_医嘱内容) & "(" & .Cell(flexcpData, j, COL_Advice_项目名称) & ")"
                                    .TextMatrix(lngRow, COL_Advice_医嘱内容) = .TextMatrix(lngRow, COL_Advice_医嘱内容) & "(" & .Cell(flexcpData, lngRow, COL_Advice_项目名称) & ")"
                                End If
                            End If
                        Next
                        lngBegin = lngEnd + 1 '下一阶段首行
                    End If
                End If
            End If
            
        Next
        
   
        .AutoSize .FixedCols, .Cols - 1, , 45
    End With

End Sub

Private Sub ClearData()
'功能:清空所有数据
    
    vsStep.Rows = vsStep.FixedRows
    vsStep.Rows = vsStep.FixedRows + 1
     
    vsItem(Index_DEL).Rows = vsItem(Index_DEL).FixedRows
    vsItem(Index_DEL).Rows = vsItem(Index_DEL).FixedRows + 1
    
    vsItem(Index_Add).Rows = vsItem(Index_Add).FixedRows
    vsItem(Index_Add).Rows = vsItem(Index_Add).FixedRows + 1
    
    vsAdvice(Index_DEL).Rows = vsAdvice(Index_DEL).FixedRows
    vsAdvice(Index_DEL).Rows = vsAdvice(Index_DEL).FixedRows + 1
    
    vsAdvice(Index_Add).Rows = vsAdvice(Index_Add).FixedRows
    vsAdvice(Index_Add).Rows = vsAdvice(Index_Add).FixedRows + 1

    If Not mrsStep Is Nothing Then
        Set mrsStep = Nothing
    End If

End Sub

