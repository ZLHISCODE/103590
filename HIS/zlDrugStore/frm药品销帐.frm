VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm药品销账 
   Caption         =   "药品退药销账"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15510
   Icon            =   "frm药品销帐.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   15510
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton optListType 
      Caption         =   "按药品汇总显示"
      Height          =   180
      Index           =   0
      Left            =   7080
      TabIndex        =   36
      Top             =   1650
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optListType 
      Caption         =   "按病人汇总显示"
      Height          =   180
      Index           =   1
      Left            =   8760
      TabIndex        =   37
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Frame fraCondition 
      Height          =   1575
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11415
      Begin VB.ComboBox cboNode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton opt科室 
         Caption         =   "病区(&T)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   638
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt科室 
         Caption         =   "医技科室(&W)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   23
         Top             =   638
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   22
         ToolTipText     =   "热键：F2"
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cbo科室 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   21
         Text            =   "cbo科室"
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cbo申请人 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   20
         Text            =   "所有申请人"
         Top             =   1020
         Width           =   2775
      End
      Begin VB.TextBox txtPati 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   19
         ToolTipText     =   "输入住院号、病人ID、床号(指定了病区时)、就诊卡号"
         Top             =   1020
         Width           =   2415
      End
      Begin VB.CommandButton cmdAllSelect 
         Caption         =   "全选(&S)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   18
         ToolTipText     =   "热键：F2"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAllUnSelect 
         Caption         =   "全清(&U)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   17
         ToolTipText     =   "热键：F2"
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkNoTime 
         Caption         =   "忽略期间"
         Height          =   180
         Left            =   1200
         TabIndex        =   16
         Tag             =   "1|0"
         Top             =   255
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp开始时间 
         Height          =   315
         Left            =   2400
         TabIndex        =   25
         Top             =   188
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   104267779
         CurrentDate     =   36985
      End
      Begin MSComCtl2.DTPicker Dtp结束时间 
         Height          =   315
         Left            =   5640
         TabIndex        =   26
         Top             =   188
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   104267779
         CurrentDate     =   36985
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7320
         TabIndex        =   40
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblNode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分部"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4440
         TabIndex        =   39
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lbl时间 
         AutoSize        =   -1  'True
         Caption         =   "申请期间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5280
         TabIndex        =   31
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Lbl科室 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申请科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   660
         Width           =   840
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申 请 人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4440
         TabIndex        =   28
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblPatiInputType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号↓"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4920
         TabIndex        =   27
         Top             =   1080
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   11
      ToolTipText     =   "热键：F2"
      Top             =   480
      Width           =   1335
   End
   Begin TabDlg.SSTab sstabList 
      Height          =   6780
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   11959
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "     未审核(&0)     "
      TabPicture(0)   =   "frm药品销帐.frx":06EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl提示(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picHsc(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picBatHsc(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfBatch(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "vsfDetail(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vsfMain(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "vsfList(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "     已审核(&1)     "
      TabPicture(1)   =   "frm药品销帐.frx":0706
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl提示(1)"
      Tab(1).Control(1)=   "vsfBatch(1)"
      Tab(1).Control(2)=   "vsfDetail(1)"
      Tab(1).Control(3)=   "vsfMain(1)"
      Tab(1).Control(4)=   "picHsc(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "picBatHsc(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "vsfList(1)"
      Tab(1).ControlCount=   7
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2295
         Index           =   1
         Left            =   -74520
         TabIndex        =   38
         Tag             =   "明细"
         Top             =   1800
         Visible         =   0   'False
         Width           =   12135
         _cx             =   21405
         _cy             =   4048
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":0722
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2295
         Index           =   0
         Left            =   45
         TabIndex        =   35
         Tag             =   "明细"
         Top             =   1440
         Visible         =   0   'False
         Width           =   15255
         _cx             =   26908
         _cy             =   4048
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":099A
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
         Height          =   2055
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Tag             =   "待处理"
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   3625
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":0C38
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   2295
         Index           =   0
         Left            =   45
         TabIndex        =   6
         Tag             =   "明细"
         Top             =   2400
         Width           =   12735
         _cx             =   22463
         _cy             =   4048
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":0D78
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   1815
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Tag             =   "明细"
         Top             =   4920
         Width           =   12735
         _cx             =   22463
         _cy             =   3201
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":0FBB
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.PictureBox picBatHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   -75000
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4850
         Width           =   12735
      End
      Begin VB.PictureBox picBatHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4850
         Width           =   12735
      End
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   -75000
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12855
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2450
         Width           =   12855
      End
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12855
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2450
         Width           =   12855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   2055
         Index           =   1
         Left            =   -74955
         TabIndex        =   12
         Tag             =   "待处理"
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   3625
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":118A
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   2295
         Index           =   1
         Left            =   -74955
         TabIndex        =   13
         Tag             =   "明细"
         Top             =   2520
         Width           =   12735
         _cx             =   22463
         _cy             =   4048
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":12A5
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   1815
         Index           =   1
         Left            =   -74955
         TabIndex        =   14
         Tag             =   "明细"
         Top             =   4920
         Width           =   12735
         _cx             =   22463
         _cy             =   3201
         Appearance      =   1
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm药品销帐.frx":14E8
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl提示 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已审核销帐记录列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   -70200
         TabIndex        =   34
         Top             =   60
         Width           =   2025
      End
      Begin VB.Label lbl提示 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "未审核销帐记录列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   4800
         TabIndex        =   33
         Top             =   60
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "退药销账(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   2
      ToolTipText     =   "热键：F2"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   0
      ToolTipText     =   "热键：F2"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Menu mnuPati 
      Caption         =   "病人"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiItem 
         Caption         =   "住院号(&0)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "ID(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "床号(&2)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Frm药品销账"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'接口参数
Private mlng库房id As Long              '当前库房ID
Private mstrUnit As String              '当前库房所对于的包装单位
Private mint药品名称 As Integer         '药品名称包含内容
Private mint金额保留位数 As Integer
Private mint打印退药清单 As Integer
Private mstrReceiveMsg As String        '主界面传递的销账申请消息，格式为：申请时间,病人id|申请时间,病人id...

'其它变量
Private mrsDetail As ADODB.Recordset            '未审核明细记录数据集
Private mrsVerifyDetail As ADODB.Recordset      '已审核明细记录数据集
Private mrsBatch As ADODB.Recordset             '未审核批次明细数据集
Private mrsVerifyBatch As ADODB.Recordset       '已审核批次明细数据集

Private mblnDrop As Boolean                     '在KeyDown中判断下拉列表是否弹出

Private mbln审核出院销账申请 As Boolean
Private mint是否可以销帐拒绝 As Integer

Private mlngMainRow As Long                  '当前选中的项目列表行数
Private mlngDetailRow As Long                '当前选中的明细列表行数
Private mlngListRow As Long                  '当前选中的列表行数

Private mblnAllowChange As Boolean              '是否允许修改销帐数量
Private mdblSum As Double

Private mblnStart As Boolean

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private mstrPrivs As String

Private mstrReturnWriteOffInfo As String    '用于记录本次销账审核的信息，并返回主界面：申请时间,病人id|申请时间,病人id...

'消费卡
Private mstrCardType As String   '消费卡/银行卡类别，格式：短名|全名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
Private mintCardCount As Integer  '卡数量
Private mobjSquareCard As Object    '一卡通部件

Private mobjPlugIn As Object             '外挂接口对象

'医保接口
Private gclsInsure As New clsInsure

Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum FindType
    住院号 = 0
    Id = 1
    床号 = 2
End Enum

Private Sub AutoExpendQuantity()
    '考虑到同一费用ID对应多个收发ID的情况，需要将销帐数量分解到多个收发记录上
    '分解的原则是按序号大的优先分配（已按序号降序排序）
    Dim n As Integer
    Dim dbl准退数量 As Double
    Dim dbl剩余数量 As Double
    Dim int收发序号 As Integer
    Dim lng费用id As Long
    Dim lng药品id As Long
    Dim str申请时间 As String

    With mrsBatch
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl准退数量 = !准退数量
            
            If lng费用id = !费用ID And lng药品id = !药品ID And str申请时间 = !申请时间 Then

            Else
                dbl剩余数量 = !销帐数量
            End If
            
            If dbl剩余数量 >= dbl准退数量 Then
                dbl剩余数量 = dbl剩余数量 - dbl准退数量
                !销帐数量 = dbl准退数量
            Else
                !销帐数量 = dbl剩余数量
                dbl剩余数量 = 0
            End If
            
            lng费用id = !费用ID
            lng药品id = !药品ID
            str申请时间 = !申请时间
            
            .Update
            .MoveNext
        Next
    End With
    
    With mrsDetail
        .MoveFirst
        Do While Not .EOF
            mrsBatch.Filter = "单据=" & !单据 & _
                " And No='" & !NO & "' " & _
                " And 药品ID=" & !药品ID & _
                " And 费用ID=" & !费用ID & _
                " And 申请时间='" & !申请时间 & "'"
            
            dbl准退数量 = 0
            If mrsBatch.RecordCount > 0 Then
                Do While Not mrsBatch.EOF
                    dbl准退数量 = dbl准退数量 + mrsBatch!准退数量
                    mrsBatch.MoveNext
                Loop
                
                If dbl准退数量 < !销帐数量 Then
                    mrsBatch.MoveFirst
                    Do While Not mrsBatch.EOF
                        mrsBatch!审核标志 = 2
                        mrsBatch.Update
                        mrsBatch.MoveNext
                    Loop
                End If
            End If
            
            !准退数量 = dbl准退数量
            If dbl准退数量 < !销帐数量 Then
                !审核标志 = 2
            End If
            .Update
            .MoveNext
        Loop
    End With
End Sub

Private Sub AutoExpendQuantityByVerify()
    '考虑到同一费用ID对应多个收发ID的情况，需要将销帐数量分解到多个收发记录上
    '分解的原则是按序号大的优先分配（已按序号降序排序）
    '处理重新审核之前已拒绝的销账记录
    Dim n As Integer
    Dim dbl准退数量 As Double
    Dim dbl剩余数量 As Double
    Dim int收发序号 As Integer
    Dim lng费用id As Long
    Dim lng药品id As Long
    Dim str申请时间 As String
    Dim lng批次 As Long


    With mrsVerifyBatch
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl准退数量 = !准退数量
            
            If lng费用id = !费用ID And lng药品id = !药品ID And str申请时间 = !申请时间 And lng批次 = !批次 Then
               

            Else
                If (lng费用id <> !费用ID Or str申请时间 <> !申请时间) Then
                    dbl剩余数量 = !销帐数量
                End If
            End If
            
            If dbl剩余数量 >= dbl准退数量 Then
                dbl剩余数量 = dbl剩余数量 - dbl准退数量
                !销帐数量 = dbl准退数量
            Else
                !销帐数量 = dbl剩余数量
                dbl剩余数量 = 0
            End If
            
            lng费用id = !费用ID
            lng药品id = !药品ID
            str申请时间 = !申请时间
            lng批次 = !批次
            
            .Update
            .MoveNext
        Next
    End With
    
    With mrsVerifyDetail
        .Filter = "审核标志=2"
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            mrsVerifyBatch.Filter = "单据=" & !单据 & _
                " And No='" & !NO & "' " & _
                " And 药品ID=" & !药品ID & _
                " And 费用ID=" & !费用ID & _
                " And 申请时间='" & !申请时间 & "'"
            
            dbl准退数量 = 0
            If mrsVerifyBatch.RecordCount > 0 Then
                Do While Not mrsVerifyBatch.EOF
                    dbl准退数量 = dbl准退数量 + mrsVerifyBatch!准退数量
                    mrsVerifyBatch.MoveNext
                Loop
                
                If dbl准退数量 < !销帐数量 Then
                    mrsVerifyBatch.MoveFirst
                    Do While Not mrsVerifyBatch.EOF
                        mrsVerifyBatch!审核标志 = 2
                        mrsVerifyBatch.Update
                        mrsVerifyBatch.MoveNext
                    Loop
                End If
            End If
            
'            !准退数量 = dbl准退数量
'            If dbl准退数量 < !销帐数量 Then
'                !审核标志 = 2
'            End If
'            .Update
            .MoveNext
        Loop
    End With
End Sub
Private Sub GetPres(ByVal int部门类型 As Integer)
    'int部门类型：0-病区；1-医技科室
    Dim rstemp As ADODB.Recordset
    Dim strSqlDept As String
    
    On Error GoTo errHandle
    If cbo科室.ListIndex > 0 Then
        strSqlDept = " And B.部门id = [1] "
    End If
        
    gstrSQL = "Select Distinct A.ID, A.简码||'-'||A.姓名 As 姓名 " & _
        " From 人员表 A, 部门人员 B " & _
        " Where A.ID = B.人员id " & strSqlDept & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
        " Order By 姓名"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取部门人员", Val(cbo科室.ItemData(cbo科室.ListIndex)))
        
    cbo申请人.Clear
    
    cbo申请人.AddItem "所有申请人"
    cbo申请人.ItemData(cbo申请人.NewIndex) = 0

    Do While Not rstemp.EOF
        cbo申请人.AddItem rstemp!姓名
        cbo申请人.ItemData(cbo申请人.NewIndex) = rstemp!Id
        rstemp.MoveNext
    Loop
    
    cbo申请人.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetStockName()
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 名称 From 部门表 Where ID = [1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取库房名称", mlng库房id)
    
    If Not rstemp.EOF Then
        Me.Caption = Me.Caption & "(" & rstemp!名称 & ")"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDetailList(ByVal int审核 As Integer, ByVal lng药品id As Long)
    If int审核 = 0 Then
        With mrsDetail
            If mrsDetail Is Nothing Then Exit Sub
            If .RecordCount = 0 Then Exit Sub

            .Filter = "药品ID=" & lng药品id

            If .EOF Then Exit Sub

            Call IniGrid(int审核, 2)
            Do While Not .EOF
                vsfDetail(0).rows = vsfDetail(0).rows + 1
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("审核标志")) = IIf(!审核标志 = 1, "√", IIf(!审核标志 = 2, "×", ""))
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("申请科室")) = !申请科室
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("单据")) = !单据
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("NO")) = !NO
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("药品id")) = !药品ID
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("申请时间")) = Format(!申请时间, "yyyy-mm-dd hh:mm:ss")
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("标识号")) = IIf(IsNull(!标识号), "", !标识号)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("姓名")) = IIf(IsNull(!姓名), "", !姓名)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("床号")) = IIf(IsNull(!床号), "", !床号)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("准退数量")) = FormatEx(!准退数量 / !包装, 5)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("销帐数量")) = FormatEx(!销帐数量 / !包装, 5)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("包装")) = IIf(IsNull(!包装), "", !包装)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("费用id")) = !费用ID
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("病人id")) = !病人ID
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("销帐原因")) = IIf(IsNull(!销帐原因), "", !销帐原因)
                 
                '准退数小于销帐数量时，准退数标记为红色
                If Val(vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("准退数量"))) < Val(vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("销帐数量"))) Then
                    vsfDetail(0).Cell(flexcpForeColor, vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("准退数量"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("准退数量")) = vbRed
                End If
                
               .MoveNext
            Loop
            
            '审核标志列加粗显示
            vsfDetail(0).Cell(flexcpFontBold, 1, vsfDetail(0).ColIndex("审核标志"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("审核标志")) = True
            
            '审核标志列蓝色显示
            vsfDetail(0).Cell(flexcpForeColor, 1, vsfDetail(0).ColIndex("审核标志"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("审核标志")) = vbBlue
            
            '准退数量列加粗显示
            vsfDetail(0).Cell(flexcpFontBold, 1, vsfDetail(0).ColIndex("准退数量"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("准退数量")) = True
            
            '销帐数量列加粗显示
            vsfDetail(0).Cell(flexcpFontBold, 1, vsfDetail(0).ColIndex("销帐数量"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("销帐数量")) = True
            
            '销帐数量标记为蓝色
            vsfDetail(0).Cell(flexcpForeColor, 1, vsfDetail(0).ColIndex("销帐数量"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("销帐数量")) = vbBlue
            
            vsfDetail(0).Row = 1
        End With
    Else
        With mrsVerifyDetail
            Call IniGrid(int审核, 2)
            If mrsVerifyDetail Is Nothing Then Exit Sub
            
            .Filter = "药品ID=" & lng药品id
            If .EOF Then Exit Sub
            Do While Not .EOF
                vsfDetail(1).rows = vsfDetail(1).rows + 1
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核标志")) = IIf(!审核标志 = 1, "√", IIf(!审核标志 = 2, "×", ""))
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("申请科室")) = !申请科室
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("单据")) = !单据
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("NO")) = !NO
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("药品id")) = !药品ID
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("申请时间")) = Format(!申请时间, "yyyy-mm-dd hh:mm:ss")
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核时间")) = Format(!审核时间, "yyyy-mm-dd hh:mm:ss")
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核人")) = !审核人
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("标识号")) = IIf(IsNull(!标识号), "", !标识号)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("姓名")) = IIf(IsNull(!姓名), "", !姓名)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("床号")) = IIf(IsNull(!床号), "", !床号)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("销帐数量")) = FormatEx(!销帐数量 / !包装, 5)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("包装")) = IIf(IsNull(!包装), "", !包装)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("费用id")) = !费用ID
                
                If !审核标志 = 2 Then
                    '审核拒绝标志列加粗显示
                    vsfDetail(1).Cell(flexcpFontBold, vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核标志"), vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核标志")) = True
                    '审核拒绝标志列红色显示
                    vsfDetail(1).Cell(flexcpForeColor, vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核标志"), vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("审核标志")) = vbRed
                End If
                
                .MoveNext
            Loop
            
            vsfDetail(1).Row = 1
        End With
    End If
End Sub

Private Sub LoadList(ByVal int审核 As Integer)
    
    mdblSum = 0
    If int审核 = 0 Then
        With mrsDetail
            If .RecordCount = 0 Then Exit Sub

            .Filter = ""

            If .EOF Then Exit Sub

            Call IniGrid(int审核, 4)
            Do While Not .EOF
                vsfList(0).rows = vsfList(0).rows + 1
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("审核标志")) = IIf(!审核标志 = 1, "√", IIf(!审核标志 = 2, "×", ""))
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("申请科室")) = !申请科室
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("单据")) = !单据
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("NO")) = !NO
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("药品id")) = !药品ID
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("申请时间")) = Format(!申请时间, "yyyy-mm-dd hh:mm:ss")
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("姓名")) = IIf(IsNull(!标识号), "", !标识号 & "-") & IIf(IsNull(!姓名), "", !姓名)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("床号")) = IIf(IsNull(!当前床号), "", !当前床号)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("准退数量")) = FormatEx(!准退数量 / !包装, 5)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("销帐数量")) = FormatEx(!销帐数量 / !包装, 5)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("销账金额")) = FormatEx(!销帐数量 * !销账金额, 2)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("包装")) = IIf(IsNull(!包装), "", !包装)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("费用id")) = !费用ID
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("药品")) = !药品
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("商品名")) = !商品名
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("规格")) = !规格
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("销帐原因")) = IIf(IsNull(!销帐原因), "", !销帐原因)
                
                mdblSum = mdblSum + FormatEx(!销帐数量 * !销账金额, 2)
                
                '准退数小于销帐数量时，准退数标记为红色
                If Val(vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("准退数量"))) < Val(vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("销帐数量"))) Then
                    vsfList(0).Cell(flexcpForeColor, vsfList(0).rows - 1, vsfList(0).ColIndex("准退数量"), vsfList(0).rows - 1, vsfList(0).ColIndex("准退数量")) = vbRed
                End If
                
               .MoveNext
            Loop
            
            '审核标志列加粗显示
            vsfList(0).Cell(flexcpFontBold, 1, vsfList(0).ColIndex("审核标志"), vsfList(0).rows - 1, vsfList(0).ColIndex("审核标志")) = True
            
            '审核标志列蓝色显示
            vsfList(0).Cell(flexcpForeColor, 1, vsfList(0).ColIndex("审核标志"), vsfList(0).rows - 1, vsfList(0).ColIndex("审核标志")) = vbBlue
            
            '准退数量列加粗显示
            vsfList(0).Cell(flexcpFontBold, 1, vsfList(0).ColIndex("准退数量"), vsfList(0).rows - 1, vsfList(0).ColIndex("准退数量")) = True
            
            '销帐数量列加粗显示
            vsfList(0).Cell(flexcpFontBold, 1, vsfList(0).ColIndex("销帐数量"), vsfList(0).rows - 1, vsfList(0).ColIndex("销帐数量")) = True
            
            '销帐数量标记为蓝色
            vsfList(0).Cell(flexcpForeColor, 1, vsfList(0).ColIndex("销帐数量"), vsfList(0).rows - 1, vsfList(0).ColIndex("销帐数量")) = vbBlue
            
            '显示销帐金额合计信息
            vsfList(0).rows = vsfList(0).rows + 1
            vsfList(0).Cell(flexcpText, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = "销账金额合计：" & FormatEx(mdblSum, 5)
            vsfList(0).Cell(flexcpFontBold, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = True
            vsfList(0).Cell(flexcpForeColor, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = vbRed
            vsfList(0).Cell(flexcpAlignment, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = flexAlignLeftCenter
            vsfList(0).MergeCells = flexMergeRestrictRows
            vsfList(0).MergeRow(vsfList(0).rows - 1) = True
            
            vsfList(0).Row = 1
        End With
    Else
        With mrsVerifyDetail
            If .RecordCount = 0 Then Exit Sub

            .Filter = ""

            If .EOF Then Exit Sub

            Call IniGrid(int审核, 4)
            Do While Not .EOF
                vsfList(1).rows = vsfList(1).rows + 1
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("审核标志")) = IIf(!审核标志 = 1, "√", IIf(!审核标志 = 2, "×", ""))
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("申请科室")) = !申请科室
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("单据")) = !单据
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("NO")) = !NO
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("药品id")) = !药品ID
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("申请时间")) = Format(!申请时间, "yyyy-mm-dd hh:mm:ss")
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("审核时间")) = Format(!审核时间, "yyyy-mm-dd hh:mm:ss")
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("审核人")) = !审核人
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("姓名")) = IIf(IsNull(!标识号), "", !标识号 & "-") & IIf(IsNull(!姓名), "", !姓名)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("床号")) = IIf(IsNull(!当前床号), "", !当前床号)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("销帐数量")) = FormatEx(!销帐数量 / !包装, 5)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("包装")) = IIf(IsNull(!包装), "", !包装)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("费用id")) = !费用ID
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("药品")) = !药品
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("商品名")) = !商品名
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("规格")) = !规格
                 
                If !审核标志 = 2 Then
                    '审核拒绝标志列加粗显示
                    vsfList(1).Cell(flexcpFontBold, vsfList(1).rows - 1, vsfList(1).ColIndex("审核标志"), vsfList(1).rows - 1, vsfList(1).ColIndex("审核标志")) = True
                    '审核拒绝标志列红色显示
                    vsfList(1).Cell(flexcpForeColor, vsfList(1).rows - 1, vsfList(1).ColIndex("审核标志"), vsfList(1).rows - 1, vsfList(1).ColIndex("审核标志")) = vbRed
                End If
                
                .MoveNext
            Loop
            
            vsfList(1).Row = 1
        End With
    End If
End Sub


Private Sub LoadBatchList(ByVal int审核 As Integer, ByVal Int单据 As Integer, _
                ByVal strNo As String, ByVal lng药品id As Long, _
                ByVal str时间 As String, ByVal lng费用id As Long, _
                ByVal bln更新标志 As Boolean, ByVal int审核标志 As Integer)
    If int审核 = 0 Then
        With mrsBatch
            .Filter = "单据=" & Int单据 & _
                    " And No='" & strNo & "' " & _
                    " And 药品ID=" & lng药品id & _
                    " And 费用ID=" & lng费用id & _
                    " And 申请时间='" & str时间 & "' "
            .Sort = "收发序号 Desc"
            
            If .EOF Then Exit Sub
            
            picBatHsc(0).Visible = True
            
            Call IniGrid(int审核, 3)
            Do While Not .EOF
                vsfBatch(0).rows = vsfBatch(0).rows + 1
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("单据")) = !单据
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("NO")) = !NO
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("药品id")) = !药品ID
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("申请时间")) = Format(!申请时间, "yyyy-mm-dd hh:mm:ss")
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("产地")) = IIf(IsNull(!产地), "", !产地)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("批号")) = IIf(IsNull(!批号), "", !批号)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("效期")) = Format(!效期, "yyyy-mm-dd")
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("准退数量")) = FormatEx(!准退数量 / !包装, 5)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("销帐数量")) = FormatEx(!销帐数量 / !包装, 5)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("包装")) = IIf(IsNull(!包装), "", !包装)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("收发序号")) = IIf(IsNull(!收发序号), "", !收发序号)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("单价")) = FormatEx(!单价 * !包装, 5)
                
                If bln更新标志 Then
                    !审核标志 = int审核标志
                    .Update
                End If
                
               .MoveNext
            Loop
            vsfBatch(0).Cell(flexcpForeColor, 1, vsfBatch(0).ColIndex("销帐数量"), vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("销帐数量")) = vbBlue
        End With
    Else
        With mrsVerifyBatch
            .Filter = "单据=" & Int单据 & _
                    " And No='" & strNo & "' " & _
                    " And 药品ID=" & lng药品id & _
                    " And 费用ID=" & lng费用id & _
                    " And 审核时间='" & str时间 & "' "
            .Sort = "收发序号 Desc"

            If .EOF Then Exit Sub
        
            picBatHsc(1).Visible = True

            Call IniGrid(int审核, 3)
            Do While Not .EOF
                vsfBatch(1).rows = vsfBatch(1).rows + 1
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("单据")) = !单据
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("NO")) = !NO
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("药品id")) = !药品ID
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("审核时间")) = Format(!审核时间, "yyyy-mm-dd hh:mm:ss")
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("产地")) = IIf(IsNull(!产地), "", !产地)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("批号")) = IIf(IsNull(!批号), "", !批号)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("效期")) = Format(!效期, "yyyy-mm-dd")
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("准退数量")) = FormatEx(!准退数量 / !包装, 5)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("销帐数量")) = FormatEx(!销帐数量 / !包装, 5)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("包装")) = IIf(IsNull(!包装), "", !包装)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("收发序号")) = IIf(IsNull(!收发序号), "", !收发序号)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("单价")) = FormatEx(!单价 * !包装, 5)
                .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub GetRecord(ByVal int审核 As Integer)
    Dim rstemp As ADODB.Recordset
    Dim strSubUnit As String
    Dim strSubName As String
    Dim intRow As Integer
    Dim str费用ID As String
    Dim strSqlCondition As String
    Dim strNo As String
    Dim str药名 As String
    Dim lngSum As Long
    Dim arrExecute As Variant
    Dim i As Integer
    Dim strNos As String
    
'    On Error GoTo errHandle
    
    Call IniRecord(int审核)
    vsfMain(int审核).rows = 1
    vsfDetail(int审核).rows = 1
    vsfBatch(int审核).rows = 1

    
    ''''1、提取汇总数据
    '是否审核
    If int审核 = 0 Then
        strSqlCondition = strSqlCondition & " And A.审核人 Is Null And A.状态 = 0  "
        If chkNoTime.Value = 0 Then
            strSqlCondition = strSqlCondition & " And A.申请时间 Between [3] And [4] "
        End If
    Else
       strSqlCondition = strSqlCondition & " And A.审核人 Is Not Null And A.状态 <> 0  "
        If chkNoTime.Value = 0 Then
            strSqlCondition = strSqlCondition & " And A.审核时间 Between [3] And [4] "
        End If
    End If
    
    '单位，包装换算
    Select Case mstrUnit
    Case "售价单位"
        strSubUnit = "X.计算单位 单位,1 包装,A.数量 As 销帐数量,"
    Case "门诊单位"
        strSubUnit = "D.门诊单位 单位,D.门诊包装 包装,A.数量 As 销帐数量,"
    Case "住院单位"
        strSubUnit = "D.住院单位 单位,D.住院包装 包装,A.数量 As 销帐数量,"
    Case "药库单位"
        strSubUnit = "D.药库单位 单位,D.药库包装 包装,A.数量 As 销帐数量,"
    End Select
    
    '病区/医技科室
    If cbo科室.ListIndex > 0 Then
        strSqlCondition = strSqlCondition & " And A.申请部门id = [2] "
    End If
    
    '申请人
    If cbo申请人.ListIndex > 0 Then
        strSqlCondition = strSqlCondition & " And A.申请人=[5] "
    End If
    
    '病人姓名
    If Val(txtPati.Tag) <> 0 Then
        strSqlCondition = strSqlCondition & " And B.病人ID=[6] "
    End If
    
    If mstrReceiveMsg <> "" Then
        '主界面销账申请消息内容为主要条件，并去掉原来的时间条件
        gstrSQL = "Select /*+ Rule*/ Distinct A.收费细目id, X.规格, " & strSubUnit & " '['||X.编码||']' As 药品编码,X.名称 As 通用名,E.名称 As 商品名,D.药品来源" & IIf(int审核 = 0, ",A.零售价 ", "") & _
            " From (Select A.收费细目id, Sum(A.数量) As 数量" & IIf(int审核 = 0, ",B.标准单价 零售价 ", "") & _
            " From 住院费用记录 B, 病案主页 F, 病人费用销帐 A , Table(f_Str2list2([7], '|', ',')) T " & _
            " Where A.申请类别=1 And A.费用id = B.ID And A.审核部门id = [1] And B.病人id = F.病人id" & IIf(int审核 = 1, "(+)", "") & " And B.主页id = F.主页id" & IIf(int审核 = 1, "(+)", "") & _
            " And a.申请时间 = To_Date(t.C1, 'yyyy-mm-dd hh24:mi:ss') And b.病人id = t.C2 "
            
        gstrSQL = gstrSQL & strSqlCondition
    Else
        gstrSQL = "Select Distinct A.收费细目id, X.规格, " & strSubUnit & " '['||X.编码||']' As 药品编码,X.名称 As 通用名,E.名称 As 商品名,D.药品来源" & IIf(int审核 = 0, ",A.零售价 ", "") & _
            " From (Select A.收费细目id, Sum(A.数量) As 数量" & IIf(int审核 = 0, ",B.标准单价 零售价 ", "") & _
            " From 住院费用记录 B, 病案主页 F, 病人费用销帐 A " & _
            " Where A.申请类别=1 And A.费用id = B.ID And A.审核部门id = [1] And B.病人id = F.病人id" & IIf(int审核 = 1, "(+)", "") & " And B.主页id = F.主页id" & IIf(int审核 = 1, "(+)", "")
        gstrSQL = gstrSQL & strSqlCondition
    End If
        
    If mbln审核出院销账申请 = False Then
        gstrSQL = gstrSQL & " And F.出院日期" & IIf(int审核 = 1, "(+)", "") & " Is Null "
    End If
    
'    gstrSQL = gstrSQL & strSqlCondition & _
'        " And Exists (Select 1 From 药品收发记录 C " & _
'        " Where B.No = C.No And C.费用id = A.费用id And C.审核人 Is Not Null And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0))"

    '排除已在输液配置中心管理中产生的单据
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y,药品收发记录 S Where Y.收发id = S.ID And S.费用id=B.ID) "
    
    gstrSQL = gstrSQL & " Group By A.收费细目id" & IIf(int审核 = 0, ",B.标准单价", "") & ") A,药品规格 D, 收费项目别名 E, 收费项目目录 X " & _
        " Where A.收费细目id = D.药品id And A.收费细目id = X.ID And X.ID = E.收费细目id(+) And E.性质(+) = 3 " & _
        " Order By 药品编码"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取退药申请", mlng库房id, Val(cbo科室.ItemData(cbo科室.ListIndex)), Dtp开始时间.Value, Dtp结束时间.Value, NeedName(cbo申请人.Text), Val(txtPati.Tag), mstrReceiveMsg)
    
    If rstemp.EOF Then
        Call IniGrid(int审核, 0)
'        MsgBox "没有找到满足条件的退药申请记录！", vbInformation, gstrSysName
        cmdAllSelect.Enabled = False
        cmdAllUnSelect.Enabled = False
        Exit Sub
    End If
    
    If sstabList.Tab = 0 Then
        cmdAllSelect.Enabled = True
        cmdAllUnSelect.Enabled = True
    End If
    
    Call IniGrid(int审核, 0)
    
    mdblSum = 0
    Do While Not rstemp.EOF
        With vsfMain(int审核)
            .rows = .rows + 1
   
            .TextMatrix(.rows - 1, .ColIndex("收费细目id")) = rstemp!收费细目id
            
            If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                str药名 = rstemp!通用名
            Else
                str药名 = IIf(IsNull(rstemp!商品名), rstemp!通用名, rstemp!商品名)
            End If
            
            If mint药品名称 = 0 Then
                .TextMatrix(.rows - 1, .ColIndex("药品名称")) = rstemp!药品编码 & str药名
            ElseIf mint药品名称 = 1 Then
                .TextMatrix(.rows - 1, .ColIndex("药品名称")) = rstemp!药品编码
            ElseIf mint药品名称 = 2 Then
                .TextMatrix(.rows - 1, .ColIndex("药品名称")) = str药名
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("商品名")) = IIf(IsNull(rstemp!商品名), "", rstemp!商品名)
                
            .TextMatrix(.rows - 1, .ColIndex("规格")) = IIf(IsNull(rstemp!规格), "", rstemp!规格)
            .TextMatrix(.rows - 1, .ColIndex("销帐数量")) = FormatEx(rstemp!销帐数量 / rstemp!包装, 5)
            If int审核 = 0 Then
                .TextMatrix(.rows - 1, .ColIndex("销账金额")) = FormatEx(rstemp!销帐数量 * rstemp!零售价, 2)
                mdblSum = mdblSum + FormatEx(rstemp!销帐数量 * rstemp!零售价, 2)
            End If
            .TextMatrix(.rows - 1, .ColIndex("单位")) = rstemp!单位
            .TextMatrix(.rows - 1, .ColIndex("来源")) = rstemp!药品来源
            rstemp.MoveNext
        End With
    Loop
    
    '销账金额的合计信息
    If int审核 = 0 Then
        vsfMain(int审核).rows = vsfMain(int审核).rows + 1
        vsfMain(int审核).Cell(flexcpText, vsfMain(int审核).rows - 1, 1, vsfMain(int审核).rows - 1, vsfMain(int审核).Cols - 1) = "销账金额合计：" & FormatEx(mdblSum, 5)
        vsfMain(int审核).Cell(flexcpFontBold, vsfMain(int审核).rows - 1, 1, vsfMain(int审核).rows - 1, vsfMain(int审核).Cols - 1) = True
        vsfMain(int审核).Cell(flexcpForeColor, vsfMain(int审核).rows - 1, 1, vsfMain(int审核).rows - 1, vsfMain(int审核).Cols - 1) = vbRed
        vsfMain(int审核).MergeCells = flexMergeRestrictRows
        vsfMain(int审核).MergeRow(vsfMain(int审核).rows - 1) = True
    End If
    
    ''''2、提取明细数据
    '单位字串
    Select Case mstrUnit
    Case "售价单位"
        strSubUnit = "X.计算单位 单位,1 包装, A.数量 "
    Case "门诊单位"
        strSubUnit = "D.门诊单位 单位,D.门诊包装 包装, A.数量 "
    Case "住院单位"
        strSubUnit = "D.住院单位 单位,D.住院包装 包装, A.数量 "
    Case "药库单位"
        strSubUnit = "D.药库单位 单位,D.药库包装 包装, A.数量 "
    End Select
    
    If int审核 = 0 Then
        If mstrReceiveMsg <> "" Then
            gstrSQL = "Select /*+ Rule*/ 申请科室, 单据, NO, 序号, 药品编码, 通用名, 商品名,规格, 药品ID, 费用id, 申请时间, 标识号, 姓名, 病人id,床号, 单位, 包装, Sum(数量) As 销帐数量,零售价,当前床号,销帐原因 " & _
                " From (Select Distinct E.名称 As 申请科室, C.单据, C.NO,B.标准单价 零售价,b.序号, '[' || x.编码 || ']' As 药品编码, x.名称 As 通用名, w.名称 As 商品名,X.规格, C.药品ID, A.费用id, A.申请时间,A.销帐原因, B.标识号, nvl(F.姓名,b.姓名) 姓名, B.病人id,B.床号,G.当前床号, " & strSubUnit & " " & _
                " From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 药品规格 D, 收费项目别名 W, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E,病人信息 G  , Table(f_Str2list2([7], '|', ',')) T " & _
                " Where A.申请类别=1 And A.费用id = B.ID And B.病人id=G.病人id(+) And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID " & _
                " And x.Id = w.收费细目id(+) And w.性质(+) = 3 And B.病人id = F.病人id And B.主页id = F.主页id " & _
                " And a.申请时间 = To_Date(t.C1, 'yyyy-mm-dd hh24:mi:ss') And b.病人id = t.C2 "
        Else
            gstrSQL = "Select 申请科室, 单据, NO, 序号, 药品编码, 通用名, 商品名,规格, 药品ID, 费用id, 申请时间, 标识号, 姓名, 病人id,床号, 单位, 包装, Sum(数量) As 销帐数量,零售价,当前床号,销帐原因 " & _
                " From (Select Distinct E.名称 As 申请科室, C.单据, C.NO,B.标准单价 零售价,b.序号, '[' || x.编码 || ']' As 药品编码, x.名称 As 通用名, w.名称 As 商品名,X.规格, C.药品ID, A.费用id, A.申请时间,A.销帐原因, B.标识号, nvl(F.姓名,b.姓名) 姓名, B.病人id,B.床号,G.当前床号, " & strSubUnit & " " & _
                " From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 药品规格 D, 收费项目别名 W, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E,病人信息 G  " & _
                " Where A.申请类别=1 And A.费用id = B.ID And B.病人id=G.病人id(+) And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID " & _
                " And x.Id = w.收费细目id(+) And w.性质(+) = 3 And B.病人id = F.病人id And B.主页id = F.主页id "
        End If
        
        If mbln审核出院销账申请 = False Then
            gstrSQL = gstrSQL & " And F.出院日期 Is Null "
        End If
        
        '排除已在输液配置中心管理中产生过的单据
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
        
        gstrSQL = gstrSQL & " And A.申请部门id = E.ID And B.执行部门id = [1]  " & _
            " And C.审核人 Is Not Null And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0) " & IIf(mstrReceiveMsg = "", strSqlCondition, "") & ")" & _
            " Group By 申请科室, 单据, NO, 序号, 药品编码, 通用名, 商品名,规格, 药品ID, 费用id, 申请时间, 标识号, 姓名, 病人id, 床号,当前床号, 单位, 包装,零售价,销帐原因 " & _
            " Order By 申请科室, 标识号, 申请时间, 单据, NO, 序号 "
    Else
        gstrSQL = "Select 申请科室, 单据, NO, 序号, 药品编码, 通用名, 商品名,规格, 药品ID, 费用id, 申请时间, 审核时间, 审核人, 状态, 标识号, 病人id, 姓名, 床号, 单位, 包装, Sum(数量) As 销帐数量,当前床号 " & _
            " From (Select Distinct E.名称 As 申请科室, C.单据, C.NO, b.序号, '[' || x.编码 || ']' As 药品编码, x.名称 As 通用名, w.名称 As 商品名,X.规格, C.药品ID, A.费用id, A.申请时间, A.审核时间, A.审核人, A.状态, B.标识号, nvl(F.姓名,b.姓名) 姓名, B.病人id, B.床号,G.当前床号, " & strSubUnit & " " & _
            " From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 药品规格 D, 收费项目别名 W, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E,病人信息 G " & _
            " Where A.申请类别=1 And A.费用id = B.ID And B.病人id=G.病人id(+) And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID " & _
            " And x.Id = w.收费细目id(+) And w.性质(+) = 3 And B.病人id = F.病人id(+) And B.主页id = F.主页id(+) " & _
            " And A.申请部门id = E.ID And B.执行部门id = [1]  "
            
        '排除已在输液配置中心管理中产生过的单据
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
        
        gstrSQL = gstrSQL & " And C.审核人 Is Not Null And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0) " & strSqlCondition & ")" & _
            " Group By 申请科室, 单据, NO, 序号, 药品编码, 通用名, 商品名,规格,药品ID, 费用id, 申请时间, 审核时间, 审核人, 状态, 标识号, 姓名, 病人id, 床号,当前床号,单位, 包装 " & _
            " Order By 申请科室, 标识号, 审核时间, 审核人, 单据, NO, 序号 "
    End If
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取单据明细", mlng库房id, Val(cbo科室.ItemData(cbo科室.ListIndex)), Dtp开始时间.Value, Dtp结束时间.Value, NeedName(cbo申请人.Text), Val(txtPati.Tag), mstrReceiveMsg)
    
    If rstemp.EOF Then
'        MsgBox "没有找到满足条件的单据明细记录！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If int审核 = 0 Then
        Do While Not rstemp.EOF
            With mrsDetail
                .AddNew
                
                !审核标志 = IIf(optListType(0).Value = True, 1, 0)
                !申请科室 = rstemp!申请科室
                !单据 = rstemp!单据
                !NO = rstemp!NO
                !药品ID = rstemp!药品ID
                !费用ID = rstemp!费用ID
                !申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
                !标识号 = rstemp!标识号
                !姓名 = rstemp!姓名
                !床号 = rstemp!床号
                !销帐数量 = rstemp!销帐数量
                !销账金额 = rstemp!零售价
                !包装 = rstemp!包装
                !单位 = rstemp!单位
                !规格 = rstemp!规格
                !当前床号 = rstemp!当前床号
                !病人ID = rstemp!病人ID
                !销帐原因 = rstemp!销帐原因
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rstemp!通用名
                Else
                    str药名 = IIf(IsNull(rstemp!商品名), rstemp!通用名, rstemp!商品名)
                End If
                
                If mint药品名称 = 0 Then
                    !药品 = rstemp!药品编码 & str药名
                ElseIf mint药品名称 = 1 Then
                    !药品 = rstemp!药品编码
                ElseIf mint药品名称 = 2 Then
                    !药品 = str药名
                End If
                !商品名 = IIf(IsNull(rstemp!商品名), "", rstemp!商品名)
                
                .Update
                
                If InStr(1, strNo, rstemp!NO) = 0 Then
                    strNo = IIf(strNo = "", "", strNo & ",") & rstemp!NO
                End If
                rstemp.MoveNext
            End With
        Loop
    Else
        Do While Not rstemp.EOF
            With mrsVerifyDetail
                .AddNew
                
                !审核标志 = rstemp!状态
                !申请科室 = rstemp!申请科室
                !单据 = rstemp!单据
                !NO = rstemp!NO
                !药品ID = rstemp!药品ID
                !费用ID = rstemp!费用ID
                !申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
                !审核时间 = Format(rstemp!审核时间, "yyyy-mm-dd hh:mm:ss")
                !审核人 = rstemp!审核人
                !标识号 = rstemp!标识号
                !姓名 = rstemp!姓名
                !床号 = rstemp!床号
                !销帐数量 = rstemp!销帐数量
                !包装 = rstemp!包装
                !单位 = rstemp!单位
                !规格 = rstemp!规格
                !当前床号 = rstemp!当前床号
                !病人ID = rstemp!病人ID
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rstemp!通用名
                Else
                    str药名 = IIf(IsNull(rstemp!商品名), rstemp!通用名, rstemp!商品名)
                End If
                
                If mint药品名称 = 0 Then
                    !药品 = rstemp!药品编码 & str药名
                ElseIf mint药品名称 = 1 Then
                    !药品 = rstemp!药品编码
                ElseIf mint药品名称 = 2 Then
                    !药品 = str药名
                End If
                !商品名 = IIf(IsNull(rstemp!商品名), "", rstemp!商品名)
                
                .Update
                
                If InStr(1, strNo, rstemp!NO) = 0 Then
                    strNo = IIf(strNo = "", "", strNo & ",") & rstemp!NO
                End If
                
                rstemp.MoveNext
            End With
        Loop
    End If
    
    ''''3、提取批次明细数据
    If int审核 = 0 Then
        '单位，包装换算
        Select Case mstrUnit
        Case "售价单位"
            strSubUnit = "X.计算单位 单位,1 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
        Case "门诊单位"
            strSubUnit = "D.门诊单位 单位,D.门诊包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
        Case "住院单位"
            strSubUnit = "D.住院单位 单位,D.住院包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
        Case "药库单位"
            strSubUnit = "D.药库单位 单位,D.药库包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
        End Select
        
        ' 'Having Sum(实际数量) > 0
        gstrSQL = "Select /*+ Rule*/ C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, F.险类, P.名称 As 开单科室, " & _
            " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, C.零售价 As 单价, " & strSubUnit & " " & _
            " From 病人费用销帐 A, 住院费用记录 B, " & _
            " (Select A.ID, A.单据, A.NO, A.序号, A.药品id, A.产地, A.批号, A.效期, A.费用id, B.实际数量, A.零售价 " & _
            " From 药品收发记录 A, " & _
            " (Select a.单据, a.NO, a.序号, a.药品id, Sum(Nvl(a.付数, 1) * a.实际数量) As 实际数量 " & _
            " From 药品收发记录 a ,Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
            " Where a.单据 In (9, 10) And a.审核日期 Is Not Null And a.No=b.Column_Value "
        
        '排除已在输液配置中心管理中产生过的单据
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = A.ID) "
            
        gstrSQL = gstrSQL & " Group By 单据, NO, 序号, 药品id " & _
            " ) B" & _
            " Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号 And A.审核人 Is Not Null " & _
            " And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0))C, " & _
            " 药品规格 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
            " Where A.申请类别=1 And A.费用id = B.ID And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID And B.病人id = F.病人id And B.主页id = F.主页id And A.申请部门id = E.ID " & _
            " And B.执行部门id = [1] " & strSqlCondition

        If mbln审核出院销账申请 = False Then
            gstrSQL = gstrSQL & " And F.出院日期 Is Null "
        End If
        
        gstrSQL = gstrSQL & " Order By A.申请时间, C.单据, C.NO, C.序号 Desc "
    Else
        '单位，包装换算
        '单位，包装换算
        Select Case mstrUnit
        Case "售价单位"
            strSubUnit = "X.计算单位 单位,1 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
        Case "门诊单位"
            strSubUnit = "D.门诊单位 单位,D.门诊包装 包装,C.实际数量 As 准退数量, A.数量 As 销帐数量"
        Case "住院单位"
            strSubUnit = "D.住院单位 单位,D.住院包装 包装,C.实际数量 As 准退数量, A.数量 As 销帐数量"
        Case "药库单位"
            strSubUnit = "D.药库单位 单位,D.药库包装 包装,C.实际数量 As 准退数量, A.数量 As 销帐数量"
        End Select
        
        gstrSQL = "Select C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, F.险类, P.名称 As 开单科室,C.批次, " & _
            " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, A.审核时间, C.零售价 As 单价, " & strSubUnit & " " & _
            " From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 药品规格 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
            " Where A.申请类别=1 And A.费用id = B.ID And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID And B.病人id = F.病人id(+) And B.主页id = F.主页id(+) And A.申请部门id = E.ID " & _
            " And B.执行部门id = [1]  " & strSqlCondition & _
            " And C.审核日期 Is Not Null " & _
            " And ((A.状态 = 1 And Mod(C.记录状态, 3) = 2 And A.审核时间 = C.审核日期)) "
            
        '排除已在输液配置中心管理中产生过的单据
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
        
        gstrSQL = gstrSQL & " Union All "
        
        ' 'Having Sum(实际数量) > 0
        gstrSQL = gstrSQL & "Select /*+ Rule*/ C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, F.险类, P.名称 As 开单科室, C.批次," & _
            " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, A.审核时间,  C.零售价 As 单价, " & strSubUnit & " " & _
            " From 病人费用销帐 A, 住院费用记录 B, " & _
            " (Select A.ID, A.单据, A.NO, A.序号, A.药品id, A.产地, A.批号, A.效期, A.费用id, B.实际数量, A.零售价,A.批次 " & _
            " From 药品收发记录 A, " & _
            " (Select a.单据, a.NO, a.序号, a.药品id, Sum(Nvl(a.付数, 1) * a.实际数量) As 实际数量 " & _
            " From 药品收发记录 a ,Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
            " Where a.单据 In (9, 10) And a.审核日期 Is Not Null And a.No=b.Column_Value "
        
        '排除已在输液配置中心管理中产生过的单据
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = A.ID) "
            
        gstrSQL = gstrSQL & " Group By 单据, NO, 序号, 药品id " & _
            " ) B" & _
            " Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号 And A.审核人 Is Not Null " & _
            " And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0))C, " & _
            " 药品规格 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
            " Where A.状态=2 and A.申请类别=1 And A.费用id = B.ID And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID And B.病人id = F.病人id And B.主页id = F.主页id And A.申请部门id = E.ID " & _
            " And B.执行部门id = [1] And A.申请时间 Between [3] And [4] "

        If mbln审核出院销账申请 = False Then
            gstrSQL = gstrSQL & " And F.出院日期 Is Null "
        End If
        
        gstrSQL = gstrSQL & " Order By 审核时间, 单据, NO, 收发序号 Desc"
        
'        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取批次明细", mlng库房id, Val(cbo科室.ItemData(cbo科室.ListIndex)), Dtp开始时间.Value, Dtp结束时间.Value, NeedName(cbo申请人.Text), Val(txtPati.Tag), strNo)
    End If
    
    If int审核 = 0 Then
        'NO串可能超过4K，分解后分别执行SQL再汇总数据集
        arrExecute = GetArrayByStr(strNo, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strNos = arrExecute(i)
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取批次明细", mlng库房id, Val(cbo科室.ItemData(cbo科室.ListIndex)), Dtp开始时间.Value, Dtp结束时间.Value, NeedName(cbo申请人.Text), Val(txtPati.Tag), strNos)
                
            Do While Not rstemp.EOF
                With mrsBatch
                    .AddNew
                    !单据 = rstemp!单据
                    !NO = rstemp!NO
                    !药品ID = rstemp!药品ID
                    !申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
                    !收发序号 = rstemp!收发序号
                    !产地 = rstemp!产地
                    !批号 = rstemp!批号
                    !效期 = rstemp!效期
                    
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And zlStr.Nvl(!效期) <> "" Then
                        '换算为有效期
                        !效期 = Format(DateAdd("D", -1, !效期), "yyyy-mm-dd")
                    End If
                    
                    !准退数量 = rstemp!准退数量
                    !销帐数量 = rstemp!销帐数量
                    !包装 = rstemp!包装
                    !单位 = rstemp!单位
                    !单价 = rstemp!单价
                    !收发Id = rstemp!收发Id
                    !主页id = IIf(IsNull(rstemp!主页id), 0, rstemp!主页id)
                    !费用序号 = rstemp!费用序号
                    !险类 = rstemp!险类
                    !费用ID = rstemp!费用ID
                    !记录性质 = rstemp!记录性质
                    !审核标志 = IIf(optListType(0).Value = True, 1, 0)
                    .Update
                    
                    rstemp.MoveNext
                End With
            Loop
        Next
        
        Call AutoExpendQuantity
    Else
        arrExecute = GetArrayByStr(strNo, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strNos = arrExecute(i)
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取批次明细", mlng库房id, Val(cbo科室.ItemData(cbo科室.ListIndex)), Dtp开始时间.Value, Dtp结束时间.Value, NeedName(cbo申请人.Text), Val(txtPati.Tag), strNos)
                
            Do While Not rstemp.EOF
                With mrsVerifyBatch
                    .AddNew
                    !单据 = rstemp!单据
                    !NO = rstemp!NO
                    !药品ID = rstemp!药品ID
                    !申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
                    !审核时间 = Format(rstemp!审核时间, "yyyy-mm-dd hh:mm:ss")
                    !收发序号 = rstemp!收发序号
                    !产地 = rstemp!产地
                    !批号 = rstemp!批号
                    !效期 = rstemp!效期
                    !批次 = rstemp!批次
                    
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And zlStr.Nvl(!效期) <> "" Then
                        '换算为有效期
                        !效期 = Format(DateAdd("D", -1, !效期), "yyyy-mm-dd")
                    End If
                    
                    !准退数量 = Abs(rstemp!准退数量)
                    !销帐数量 = Abs(rstemp!销帐数量)
                    !包装 = rstemp!包装
                    !单位 = rstemp!单位
                    !单价 = rstemp!单价
                    !收发Id = rstemp!收发Id
                    !主页id = IIf(IsNull(rstemp!主页id), 0, rstemp!主页id)
                    !费用序号 = rstemp!费用序号
                    !险类 = rstemp!险类
                    !费用ID = rstemp!费用ID
                    !记录性质 = rstemp!记录性质
                    !审核标志 = 1
                    .Update
                    
                    rstemp.MoveNext
                End With
            Loop
        Next
        Call AutoExpendQuantityByVerify
    End If
    
    ''''''4、定位到汇总第一行，并提取第一行明细数据
    If vsfMain(int审核).rows > 1 Then
        mlngMainRow = 1
        mlngDetailRow = 1
        Call LoadDetailList(int审核, Val(vsfMain(int审核).TextMatrix(1, vsfMain(int审核).ColIndex("收费细目id"))))
        
        mlngListRow = 1
        Call LoadList(int审核)
    End If

    cmdAllSelect.Enabled = False
    cmdAllUnSelect.Enabled = False
        
    If sstabList.Tab = 0 Then
        If mrsDetail.RecordCount > 0 Then
            cmdAllSelect.Enabled = True
            cmdAllUnSelect.Enabled = True
        End If
    End If
    
    '提取第一行批次明细数据
    If optListType(0).Value = True Then
        If vsfDetail(int审核).rows > 1 Then
            If int审核 = 0 Then
                Call LoadBatchList(int审核, Val(vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("单据"))), vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("NO")), Val(vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("药品id"))), vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("申请时间")), Val(vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("费用id"))), False, IIf(optListType(0).Value = True, 1, 0))
            Else
                Call LoadBatchList(int审核, Val(vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("单据"))), vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("NO")), Val(vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("药品id"))), vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("审核时间")), Val(vsfDetail(int审核).TextMatrix(1, vsfDetail(int审核).ColIndex("费用id"))), False, IIf(optListType(0).Value = True, 1, 0))
            End If
        End If
    Else
        If vsfList(int审核).rows >= 2 Then
            If int审核 = 0 Then
                Call LoadBatchList(int审核, Val(vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("单据"))), vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("NO")), Val(vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("药品id"))), vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("申请时间")), Val(vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("费用id"))), False, IIf(optListType(0).Value = True, 1, 0))
            Else
                Call LoadBatchList(int审核, Val(vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("单据"))), vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("NO")), Val(vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("药品id"))), vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("审核时间")), Val(vsfList(int审核).TextMatrix(1, vsfList(int审核).ColIndex("费用id"))), False, IIf(optListType(0).Value = True, 1, 0))
            End If
        End If
    End If
    
    cmdAllSelect.Enabled = False
    cmdAllUnSelect.Enabled = False
        
    If int审核 = 0 Then
        If mrsBatch.RecordCount > 0 Then
            cmdAllSelect.Enabled = True
            cmdAllUnSelect.Enabled = True
        End If
        
        If mrsDetail.RecordCount > 0 Then
            If optListType(0).Value = True Then
                If vsfMain(int审核).rows > 1 Then
                    vsfMain(int审核).Row = 1
                    vsfMain(int审核).SetFocus
                End If
            Else
                If vsfList(int审核).rows > 1 Then
                    vsfMain(int审核).Row = 1
                    vsfList(int审核).SetFocus
                End If
            End If
        End If
    Else
        If mrsVerifyDetail.RecordCount > 0 Then
            If optListType(0).Value = True Then
                If vsfMain(int审核).rows > 1 Then
                    vsfMain(int审核).Row = 1
                    vsfMain(int审核).SetFocus
                End If
            Else
                If vsfList(int审核).rows > 1 Then
                    vsfMain(int审核).Row = 1
                    vsfList(int审核).SetFocus
                End If
            End If
        End If
    End If
    
    mstrReceiveMsg = ""
    
    Exit Sub
errHandle:
    mstrReceiveMsg = ""
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub


Private Sub IniRecord(ByVal int审核 As Integer)
    '未审核明细记录集
    If int审核 = 0 Then
        Set mrsDetail = New ADODB.Recordset
        With mrsDetail
            If .State = 1 Then .Close
            .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
            .Fields.Append "申请科室", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "单据", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "标识号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "床号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
            .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
            .Fields.Append "销账金额", adDouble, 18, adFldIsNullable
            .Fields.Append "包装", adDouble, 18, adFldIsNullable
            .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
            .Fields.Append "药品", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "商品名", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "规格", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "当前床号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
            .Fields.Append "销帐原因", adLongVarChar, 200, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    Else
        '已审核明细记录集
        Set mrsVerifyDetail = New ADODB.Recordset
        With mrsVerifyDetail
            If .State = 1 Then .Close
            .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
            .Fields.Append "申请科室", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "单据", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "审核时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "审核人", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "标识号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "床号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
            .Fields.Append "包装", adDouble, 18, adFldIsNullable
            .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
            .Fields.Append "药品", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "商品名", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "规格", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "当前床号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
    
    '未审核批次明细记录集
    If int审核 = 0 Then
        Set mrsBatch = New ADODB.Recordset
        With mrsBatch
            If .State = 1 Then .Close
            .Fields.Append "单据", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "收发序号", adDouble, 18, adFldIsNullable
            .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
            .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
            .Fields.Append "包装", adDouble, 18, adFldIsNullable
            .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
            .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
            .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
            .Fields.Append "险类", adDouble, 18, adFldIsNullable
            .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
            .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
            .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
            .Fields.Append "单价", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    Else
        '已审核批次明细记录集
        Set mrsVerifyBatch = New ADODB.Recordset
        With mrsVerifyBatch
            If .State = 1 Then .Close
            .Fields.Append "单据", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "审核时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "收发序号", adDouble, 18, adFldIsNullable
            .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
            .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
            .Fields.Append "包装", adDouble, 18, adFldIsNullable
            .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
            .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
            .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
            .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
            .Fields.Append "险类", adDouble, 18, adFldIsNullable
            .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
            .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
            .Fields.Append "单价", adDouble, 18, adFldIsNullable
            .Fields.Append "批次", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
End Sub
Private Sub IniGrid(ByVal int审核 As Integer, ByVal intGrid As Integer)
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    'int审核：0－未审核；1－已审核
    'intGrid：0－初始所有列表；1－初始汇总列表；2－初始明细列表；3－批次明细列表；4-病人汇总列表
    
    '初始汇总列表
    If intGrid = 0 Or intGrid = 1 Then
        With vsfMain(int审核)
            .Redraw = flexRDNone
            .rows = 1
            
            If gint药品名称显示 = 2 Then
                .ColWidth(.ColIndex("商品名")) = IIf(.ColWidth(.ColIndex("商品名")) = 0, 2000, .ColWidth(.ColIndex("商品名")))
            Else
                .ColWidth(.ColIndex("商品名")) = 0
            End If
                
            .Redraw = flexRDDirect
        End With
    End If
    
    '初始明细列表
    If intGrid = 0 Or intGrid = 2 Then
        With vsfDetail(int审核)
            .Redraw = flexRDNone
            .rows = 1
            .ColDataType(.ColIndex("销帐数量")) = flexDTDouble

            .Redraw = flexRDDirect
        End With
    End If
    
    '初始批次列表
    If intGrid = 0 Or intGrid = 3 Then
        With vsfBatch(int审核)
            .Redraw = flexRDNone
            .rows = 1
            .ColDataType(.ColIndex("销帐数量")) = flexDTDouble
            .TextMatrix(0, .ColIndex("效期")) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")

            .Redraw = flexRDDirect
        End With
    End If
    
    '初始病人汇总列表
    If intGrid = 0 Or intGrid = 4 Then
        With vsfList(int审核)
            .Redraw = flexRDNone
            .rows = 1
            
            If gint药品名称显示 = 2 Then
                .ColWidth(.ColIndex("商品名")) = IIf(.ColWidth(.ColIndex("商品名")) = 0, 2000, .ColWidth(.ColIndex("商品名")))
            Else
                .ColWidth(.ColIndex("商品名")) = 0
            End If
                
            .Redraw = flexRDDirect
        End With
    End If
End Sub

Private Sub Oper_ReVerify()
    '重审之前已拒绝的销账记录
    Dim strCurrent As String
    Dim Int单据 As Integer
    Dim strNo As String
    Dim lng药品id As Long
    Dim lng费用id As Long
    Dim str申请时间 As String
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim bln是否有退药 As Boolean
    Dim str序号数量 As String
    Dim str药品id As String
    Dim arrSql As Variant
    Dim blnBeginTrans As Boolean
    Dim Int退药 As Integer
    Dim lng病人ID As Long
    Dim dbl销帐数量 As Double
    Dim strReturnInfo As String
    Dim strReserve As String
    
    On Error GoTo errHandle
    
    If optListType(0).Value = True Then
        With vsfDetail(1)
            If .Row = 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("费用id"))) = 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("申请时间"))) = "" Then Exit Sub
            
            Int单据 = Val(.TextMatrix(.Row, .ColIndex("单据")))
            strNo = .TextMatrix(.Row, .ColIndex("NO"))
            lng药品id = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
            lng费用id = Val(.TextMatrix(.Row, .ColIndex("费用id")))
            str申请时间 = .TextMatrix(.Row, .ColIndex("申请时间"))
        End With
    Else
        With vsfList(1)
            If .Row = 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("费用id"))) = 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("申请时间"))) = "" Then Exit Sub
            
            Int单据 = Val(.TextMatrix(.Row, .ColIndex("单据")))
            strNo = .TextMatrix(.Row, .ColIndex("NO"))
            lng药品id = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
            lng费用id = Val(.TextMatrix(.Row, .ColIndex("费用id")))
            str申请时间 = .TextMatrix(.Row, .ColIndex("申请时间"))
        End With
    End If
    
    mrsVerifyDetail.Filter = "费用id=" & lng费用id & " And 申请时间='" & str申请时间 & "' "
    If mrsVerifyDetail.RecordCount = 0 Then Exit Sub
    lng病人ID = mrsVerifyDetail!病人ID
    dbl销帐数量 = mrsVerifyDetail!销帐数量
    
    '检查是否结账
    mrsVerifyBatch.Filter = "单据=" & Int单据 & _
        " And No='" & strNo & "' " & _
        " And 药品ID=" & lng药品id & _
        " And 费用ID=" & lng费用id & _
        " And 申请时间='" & str申请时间 & "' " & _
        " And 审核标志=1 " & _
        " And 销帐数量<>0 "
    If mrsVerifyBatch.RecordCount = 0 Then
        MsgBox "病人所剩药品小于当前需要销帐的数量，不能进行销帐审核操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If IsOutPatient(mstrPrivs, mrsVerifyBatch!单据, mrsVerifyBatch!NO, 2, 2) = False Then Exit Sub
    If IsReceiptBalance_Charge(1, mstrPrivs, mrsVerifyBatch!单据, mrsVerifyBatch!NO, mrsVerifyBatch!费用序号, 2, 2) = False Then Exit Sub
    
    '初始化医保部件
    gclsInsure.InitOracle gcnOracle
    
    strCurrent = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    arrSql = Array()
    
    '重审之前已拒绝的销账记录
    gstrSQL = "Zl_病人费用销帐_Cancel("
    '费用ID
    gstrSQL = gstrSQL & lng费用id
    '申请时间
    gstrSQL = gstrSQL & ",To_Date('" & str申请时间 & "','YYYY-MM-DD HH24:MI:SS')"
    '审核人
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '审核时间
    gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
    '操作类型
    gstrSQL = gstrSQL & ",0"
    gstrSQL = gstrSQL & ")"
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL
     
    '退药处理
    Do While Not mrsVerifyBatch.EOF
        gstrSQL = "zl_药品收发记录_部门退药("
        '收发ID
        gstrSQL = gstrSQL & mrsVerifyBatch!收发Id
        '审核人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '审核时间
        gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
        '批号
        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsVerifyBatch!批号), "NULL", IIf(Mid(mrsVerifyBatch!批号, 1, 1) = "(", "NULL", "'" & Mid(mrsVerifyBatch!批号, 1, 8) & "'"))
        '效期
        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsVerifyBatch!效期), "NULL", IIf(mrsVerifyBatch!效期 = "", "NULL", "To_Date('" & Format(mrsVerifyBatch!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
        '产地
        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsVerifyBatch!产地), "NULL", "'" & mrsVerifyBatch!产地 & "'")
        '退药数
        gstrSQL = gstrSQL & "," & mrsVerifyBatch!销帐数量
        '退药库房
        gstrSQL = gstrSQL & ",NULL"
        '退药人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '金额保留位数
        gstrSQL = gstrSQL & "," & mint金额保留位数
        '门诊
        gstrSQL = gstrSQL & ",2"
        '汇总发药号
        gstrSQL = gstrSQL & ",Null"
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
                
        bln是否有退药 = True
        
        If InStr("," & str药品id & ",", "," & mrsVerifyBatch!药品ID & ",") = 0 Then
            str药品id = IIf(str药品id = "", "", str药品id & ",") & mrsVerifyBatch!药品ID
        End If
        
        strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(mrsVerifyBatch!收发Id) & "," & mrsVerifyBatch!销帐数量
        
        '记录当前销账审核的记录的申请时间和病人ID，用于返回给主界面
        If mstrReturnWriteOffInfo = "" Then
            mstrReturnWriteOffInfo = Format(mrsVerifyBatch!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & lng病人ID
        ElseIf InStr(mstrReturnWriteOffInfo & "|", Format(mrsVerifyBatch!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & lng病人ID & "|") = 0 Then
            mstrReturnWriteOffInfo = mstrReturnWriteOffInfo & "|" & Format(mrsVerifyBatch!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & lng病人ID
        End If
        
        mrsVerifyBatch.MoveNext
    Loop
    
    mrsVerifyBatch.MoveFirst
    str序号数量 = mrsVerifyBatch!费用序号 & ":" & dbl销帐数量
    
    '处理费用记账记录
    gstrSQL = "ZL_住院记帐记录_Delete("
    'NO
    gstrSQL = gstrSQL & "'" & mrsVerifyBatch!NO & "'"
    '序号，数量串
    gstrSQL = gstrSQL & ",'" & str序号数量 & "'"
    '操作员编号
    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
    '操作员姓名
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '记录性质
    gstrSQL = gstrSQL & "," & mrsVerifyBatch!记录性质
    '操作状态
    gstrSQL = gstrSQL & ",1"
    gstrSQL = gstrSQL & ")"

    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL

    '医保处理
    If Not IsNull(mrsVerifyBatch!险类) And InStr(1, strMCNO, mrsVerifyBatch!NO) = 0 Then
        MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(mrsVerifyBatch!险类))
        MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(mrsVerifyBatch!险类))
        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & mrsVerifyBatch!NO & "," & mrsVerifyBatch!险类 & _
                "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
    End If
       
    '提示停用药品
    If str药品id <> "" Then
        Int退药 = 1
        Call CheckStopMedi(str药品id, Int退药)
        If Int退药 = 2 Then Exit Sub
    End If

    '集中处理退药销账事务
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "Oper_ReVerify")
    Next
                
    '医保，记帐作废上传，作废时上传
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:
                    Exit Sub
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '医保，记帐作废上传，完成后上传
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    
    If bln是否有退药 = True Then
        frm部门发药管理New.BlnRefresh = True
        If mint打印退药清单 = 2 Then
            If MsgBox("你需要打印退药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & strCurrent, "包装系数=" & IIf(mstrUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
            End If
        ElseIf mint打印退药清单 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & strCurrent, "包装系数=" & IIf(mstrUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
        End If
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退药 Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlng库房id, strReturnInfo, CDate(strCurrent), strReserve
        err.Clear: On Error GoTo 0
    End If

    Call GetRecord(Val(sstabList.Tab))
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub SetAllSelect(ByVal intType As Integer)
    'intType:1-AllSelect;0-AllUnSelect
    Dim n As Integer
    Dim int审核标志 As Integer
    
    If sstabList.Tab = 1 Then Exit Sub
    
    If optListType(0).Value = True Then
        With vsfDetail(0)
            If .rows <= 1 Then Exit Sub
            If .TextMatrix(1, .ColIndex("申请科室")) = "" Then Exit Sub
    
            For n = 1 To .rows - 1
                If .TextMatrix(n, .ColIndex("申请科室")) <> "" Then
                    If intType = 1 Then
                        If Val(.TextMatrix(n, .ColIndex("准退数量"))) >= Val(.TextMatrix(n, .ColIndex("销帐数量"))) Then
                            .TextMatrix(n, .ColIndex("审核标志")) = "√"
                        Else
                            .TextMatrix(n, .ColIndex("审核标志")) = "×"
                        End If
                    Else
                        .TextMatrix(n, .ColIndex("审核标志")) = ""
                    End If
                End If
            Next
        End With
    Else
        With vsfList(0)
            If .rows <= 1 Then Exit Sub
            If .TextMatrix(1, .ColIndex("申请科室")) = "" Then Exit Sub
    
            For n = 1 To .rows - 1
                If .TextMatrix(n, .ColIndex("申请科室")) <> "" Then
                    If intType = 1 Then
                        If Val(.TextMatrix(n, .ColIndex("准退数量"))) >= Val(.TextMatrix(n, .ColIndex("销帐数量"))) Then
                            .TextMatrix(n, .ColIndex("审核标志")) = "√"
                        Else
                            .TextMatrix(n, .ColIndex("审核标志")) = "×"
                        End If
                    Else
                        .TextMatrix(n, .ColIndex("审核标志")) = ""
                    End If
                End If
            Next
        End With
    End If
    
    With mrsDetail
        .Filter = ""
        .MoveFirst
        
        Do While Not .EOF
            If intType = 1 Then
                If !准退数量 >= !销帐数量 Then
                    !审核标志 = 1
                Else
                    !审核标志 = 2
                End If
            Else
                !审核标志 = 0
            End If
            
            .Update
            
            '同步更新批次明细列表
            mrsBatch.Filter = "单据=" & !单据 & _
                " And No='" & !NO & "' " & _
                " And 药品ID=" & !药品ID & _
                " And 费用ID=" & !费用ID & _
                " And 申请时间='" & !申请时间 & "'"
            Do While Not mrsBatch.EOF
                mrsBatch!审核标志 = !审核标志
                mrsBatch.Update
                mrsBatch.MoveNext
            Loop
            
            .MoveNext
        Loop
    End With
End Sub

Public Function ShowForm(FrmMain As Form, ByVal lng库房id As Long, ByVal strUnit As String, ByVal int金额保留位数 As Integer, _
    ByVal strCards As String, ByVal int打印退药清单 As Integer, ByVal strWriteOffMsg As String, _
    ByRef objSquareCard As Object, ByVal objPlugIn As Object) As String
    '返回本次已进行的销账审核信息：申请时间,病人id|申请时间,病人id...
    mlng库房id = lng库房id
    mstrUnit = strUnit
    mint金额保留位数 = int金额保留位数
    mstrCardType = strCards
    mint打印退药清单 = int打印退药清单
    mstrReceiveMsg = strWriteOffMsg
       
    Set mobjSquareCard = objSquareCard
    Set mobjPlugIn = objPlugIn
    
    If mstrCardType <> "" Then
        mintCardCount = UBound(Split(mstrCardType, ";")) + 1
    End If
    
    Me.Show vbModal, FrmMain
    
    ShowForm = mstrReturnWriteOffInfo
End Function
Private Sub Oper_Verify()
    '退药销账
    Dim i As Integer
    Dim strCurrent As String
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim bln是否有退药 As Boolean
    Dim str序号数量 As String
    Dim str药品id As String
    Dim arrSql As Variant
    Dim blnBeginTrans As Boolean
    Dim Int退药 As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    
    arrSql = Array()
    
    On Error GoTo errHandle
    
    If optListType(0).Value = True Then
        If vsfMain(0).rows = 1 Then Exit Sub
        If vsfMain(0).TextMatrix(1, vsfMain(0).ColIndex("收费细目id")) = "" Then Exit Sub
    Else
        If vsfList(0).rows = 1 Then Exit Sub
        If vsfList(0).TextMatrix(1, vsfList(0).ColIndex("药品id")) = "" Then Exit Sub
    End If
    
    strCurrent = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    gclsInsure.InitOracle gcnOracle
    
    With mrsDetail
        If .State = 0 Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
        
        '检查是否结账
        .Filter = ""
        .Sort = "单据, NO, 药品ID, 申请时间"
        Do While Not .EOF
            mrsBatch.Filter = "单据=" & !单据 & _
                " And No='" & !NO & "' " & _
                " And 药品ID=" & !药品ID & _
                " And 费用ID=" & !费用ID & _
                " And 申请时间='" & !申请时间 & "' " & _
                " And 审核标志<>0 "
            If mrsBatch.RecordCount > 0 Then
                If mrsBatch!审核标志 = 1 And !销帐数量 <> 0 Then
                    If IsOutPatient(mstrPrivs, mrsBatch!单据, mrsBatch!NO, 2, 2) = False Then Exit Sub
                    If IsReceiptBalance_Charge(1, mstrPrivs, mrsBatch!单据, mrsBatch!NO, mrsBatch!费用序号, 2, 2) = False Then Exit Sub
                End If
            End If
            
            .MoveNext
        Loop
        
        '销帐处理，按药品ID排序
        .Filter = ""
        .Sort = "药品ID,单据,NO,申请时间"
        
        gclsInsure.InitOracle gcnOracle
        
        Screen.MousePointer = 11
                  
        Do While Not .EOF
            mrsBatch.Filter = "单据=" & !单据 & _
                " And No='" & !NO & "' " & _
                " And 药品ID=" & !药品ID & _
                " And 费用ID=" & !费用ID & _
                " And 申请时间='" & !申请时间 & "' " & _
                " And 审核标志<>0 "
            If mrsBatch.RecordCount > 0 Then
                '费用销帐记录处理
                gstrSQL = "zl_病人费用销帐_Audit("
                '费用ID
                gstrSQL = gstrSQL & mrsBatch!费用ID
                '申请时间
                gstrSQL = gstrSQL & ",To_Date('" & mrsBatch!申请时间 & "','YYYY-MM-DD HH24:MI:SS')"
                '审核人
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '审核时间
                gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
                '审核标志
                gstrSQL = gstrSQL & "," & mrsBatch!审核标志
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                                
                                
                '退药处理
                Do While Not mrsBatch.EOF
                    If mrsBatch!审核标志 = 1 And mrsBatch!销帐数量 <> 0 Then
                        gstrSQL = "zl_药品收发记录_部门退药("
                        '收发ID
                        gstrSQL = gstrSQL & mrsBatch!收发Id
                        '审核人
                        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                        '审核时间
                        gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
                        '批号
                        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsBatch!批号), "NULL", IIf(Mid(mrsBatch!批号, 1, 1) = "(", "NULL", "'" & Mid(mrsBatch!批号, 1, 8) & "'"))
                        '效期
                        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsBatch!效期), "NULL", IIf(mrsBatch!效期 = "", "NULL", "To_Date('" & Format(mrsBatch!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
                        '产地
                        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsBatch!产地), "NULL", "'" & mrsBatch!产地 & "'")
                        '退药数
                        gstrSQL = gstrSQL & "," & mrsBatch!销帐数量
                        '退药库房
                        gstrSQL = gstrSQL & ",NULL"
                        '退药人
                        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                        '金额保留位数
                        gstrSQL = gstrSQL & "," & mint金额保留位数
                        '门诊
                        gstrSQL = gstrSQL & ",2"
                        '汇总发药号
                        gstrSQL = gstrSQL & ",Null"
                        gstrSQL = gstrSQL & ")"
        
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = gstrSQL
                                
                        bln是否有退药 = True
                        
                        If InStr("," & str药品id & ",", "," & !药品ID & ",") = 0 Then
                            str药品id = IIf(str药品id = "", "", str药品id & ",") & !药品ID
                        End If
                        
                        strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(mrsBatch!收发Id) & "," & mrsBatch!销帐数量
                        
                        '记录当前销账审核的记录的申请时间和病人ID，用于返回给主界面
                        If mstrReturnWriteOffInfo = "" Then
                            mstrReturnWriteOffInfo = Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID
                        ElseIf InStr(mstrReturnWriteOffInfo & "|", Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID & "|") = 0 Then
                            mstrReturnWriteOffInfo = mstrReturnWriteOffInfo & "|" & Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID
                        End If
                    End If
                    
                    mrsBatch.MoveNext
                Loop
                
                mrsBatch.MoveFirst
                
                '销帐处理
                If mrsBatch!审核标志 = 1 And !销帐数量 <> 0 Then
                    str序号数量 = mrsBatch!费用序号 & ":" & !销帐数量
            
                    gstrSQL = "ZL_住院记帐记录_Delete("
                    'NO
                    gstrSQL = gstrSQL & "'" & mrsBatch!NO & "'"
                    '序号，数量串
                    gstrSQL = gstrSQL & ",'" & str序号数量 & "'"
                    '操作员编号
                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                    '操作员姓名
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '记录性质
                    gstrSQL = gstrSQL & "," & mrsBatch!记录性质
                    '操作状态
                    gstrSQL = gstrSQL & ",1"
                    gstrSQL = gstrSQL & ")"

                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
    
                    '医保处理
                    If Not IsNull(mrsBatch!险类) And InStr(1, strMCNO, mrsBatch!NO) = 0 Then
                        MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(mrsBatch!险类))
                        MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(mrsBatch!险类))
                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & mrsBatch!NO & "," & mrsBatch!险类 & _
                                "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                    End If
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    '提示停用药品
    If str药品id <> "" Then
        Int退药 = 1
        Call CheckStopMedi(str药品id, Int退药)
        If Int退药 = 2 Then Exit Sub
    End If
 
     '集中处理退药销账事务
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "cmdVerify_Click")
    Next
                
    '医保，记帐作废上传，作废时上传
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '医保，记帐作废上传，完成后上传
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    
    Screen.MousePointer = 0
    
    If bln是否有退药 = True Then
        frm部门发药管理New.BlnRefresh = True
        If mint打印退药清单 = 2 Then
            If MsgBox("你需要打印退药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & strCurrent, "包装系数=" & IIf(mstrUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
            End If
        ElseIf mint打印退药清单 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & strCurrent, "包装系数=" & IIf(mstrUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
        End If
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退药 Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlng库房id, strReturnInfo, CDate(strCurrent), strReserve
        err.Clear: On Error GoTo 0
    End If

    Call GetRecord(Val(sstabList.Tab))
    
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub cboNode_Click()
    Call GetDept(IIf(opt科室(0).Value = True, 0, 1))
End Sub


Private Sub cbo科室_Click()
    If cbo科室.ListIndex = -1 Then Exit Sub
    
    If Val(cbo科室.Tag) <> cbo科室.ItemData(cbo科室.ListIndex) Then
        cbo科室.Tag = cbo科室.ItemData(cbo科室.ListIndex)
        Call GetPres(IIf(opt科室(0).Value, 0, 1))
    End If
End Sub

Private Sub Cbo科室_KeyPress(KeyAscii As Integer)
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim vRect As RECT
    Dim rstemp As ADODB.Recordset
    Dim StrNode As String
    Dim blnCancel As Boolean
    Dim i As Integer
    
    If KeyAscii = 13 Then
        If Trim(cbo科室.Text) = "" Then Exit Sub
            
        If opt科室(0).Value = True Then
            gstrSQL = " Select A.ID,b.名称 As 站点名称, b.编号 As 站点,A.编码||'-'||A.名称 科室 From 部门表 A, Zlnodelist B " & _
                " Where a.站点 = b.编号(+) And A.ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
                " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                " And (A.编码 Like [2] Or A.名称 Like [2] Or A.简码 Like [2])"
        Else
            gstrSQL = " Select A.ID,b.名称 As 站点名称, b.编号 As 站点,A.编码||'-'||A.名称 科室 From 部门表 A, Zlnodelist B " & _
                " Where a.站点 = b.编号(+) And A.ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术') And 服务对象 IN(2,3))" & _
                " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                " And (A.编码 Like [2] Or A.名称 Like [2] Or A.简码 Like [2])"
        End If
        
        If cboNode.Visible Then
            If cboNode.ListIndex > 0 Then
                StrNode = cboNode.ItemData(cboNode.ListIndex)
            End If
        End If
        If StrNode <> "" Then
            gstrSQL = gstrSQL & " And A.站点 = [1] "
        End If
        
        gstrSQL = gstrSQL & " Order By a.站点, a.编码 || '-' || a.名称 "
        
        '有多条记录则显示，供选择
        vRect = zlControl.GetControlRect(cbo科室.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = cbo科室.Height
        
        Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "选择病区或科室", False, "", "选择病区或科室", False, False, True, sngX, sngY, sngH, blnCancel, False, False, StrNode, UCase(cbo科室.Text) & "%")
    
        If blnCancel = True Then Exit Sub
        
        If rstemp Is Nothing Then
            cbo科室.Text = ""
            cbo科室.Tag = ""
            cbo科室.SetFocus
            Exit Sub
        Else
            For i = 1 To cbo科室.ListCount - 1
                If cbo科室.ItemData(i) = rstemp!Id Then
                    cbo科室.ListIndex = i
                    Exit For
                End If
            Next
            
            If cbo科室.ListIndex > 0 Then
                cbo科室.Tag = cbo科室.ItemData(cbo科室.ListIndex)
                
                Call GetPres(IIf(opt科室(0).Value, 0, 1))
            End If
        End If
    End If
End Sub

Private Sub cbo申请人_Click()
'    Exit Sub
End Sub

Private Sub cbo申请人_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo申请人.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo申请人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo申请人.Text)
        If cbo申请人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo申请人.List(cbo申请人.ListIndex) Then Call zlControl.CboSetIndex(cbo申请人.hWnd, -1)
        End If
        If strText = "" Then
            cbo申请人.ListIndex = -1
        ElseIf cbo申请人.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo申请人.ListCount - 1
                If Mid(cbo申请人.List(i), 1, InStr(1, cbo申请人.List(i), "-") - 1) = strText _
                    Or Mid(cbo申请人.List(i), InStr(1, cbo申请人.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo申请人.ListCount - 1
                    If UCase(cbo申请人.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo申请人.ListIndex = intIdx
            SendMessage cbo申请人.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo申请人_Click
            Exit Sub
        End If
        If cbo申请人.ListIndex = -1 Then
            If cbo申请人.ListCount > 1 Then
                cbo申请人.ListIndex = 0
            Else
                cbo申请人.Text = "所有申请人"
            End If
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo申请人_Click
            ElseIf intIdx <> cbo申请人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo申请人.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo申请人_Click
            End If
        End If
    End If
End Sub

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Sub chkNoTime_Click()
    If chkNoTime.Value = 0 Then
        Dtp开始时间.Enabled = True
        Dtp结束时间.Enabled = True
    Else
        Dtp开始时间.Enabled = False
        Dtp结束时间.Enabled = False
    End If
    
    If sstabList.Tab = 0 Then
        chkNoTime.Tag = chkNoTime.Value & Mid(chkNoTime.Tag, 2)
    Else
        chkNoTime.Tag = Mid(chkNoTime.Tag, 1, 2) & chkNoTime.Value
    End If
End Sub

Private Sub cmdAllSelect_Click()
    Call SetAllSelect(1)
End Sub

Private Sub cmdAllUnSelect_Click()
    Call SetAllSelect(0)
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub


Private Sub IniDate()
    Dim dateCurrent As Date
    
    dateCurrent = Sys.Currentdate
    
    Dtp开始时间.Value = CDate(Format(DateAdd("D", -1, dateCurrent), "yyyy-MM-dd 00:00:00"))
    Dtp结束时间.Value = CDate(Format(dateCurrent, "yyyy-MM-dd 23:59:59"))
End Sub

Private Sub GetDept(ByVal int部门类型 As Integer)
    'int部门类型：0-病区；1-医技科室
    Dim rstemp As ADODB.Recordset
    Dim StrNode As String
    
    On Error GoTo errHandle
    Select Case int部门类型
        Case 0
            gstrSQL = " Select b.名称 As 站点名称, b.编号 As 站点,A.编码||'-'||A.名称 科室,A.ID From 部门表 A, Zlnodelist B " & _
                 " Where a.站点 = b.编号(+) And A.ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
                 " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) "
        Case 1
            gstrSQL = " Select b.名称 As 站点名称, b.编号 As 站点,A.编码||'-'||A.名称 科室,A.ID From 部门表 A, Zlnodelist B " & _
             " Where a.站点 = b.编号(+) And A.ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术') And 服务对象 IN(2,3))" & _
             " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) "
    End Select
    
    If cboNode.Visible Then
        If cboNode.ListIndex > 0 Then
            StrNode = cboNode.ItemData(cboNode.ListIndex)
        End If
    End If
    If StrNode <> "" Then
        gstrSQL = gstrSQL & " And A.站点 = [1] "
    End If
    
    gstrSQL = gstrSQL & " Order By a.站点, a.编码 || '-' || a.名称 "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取科室", StrNode)
    
    cbo科室.Clear
    cbo科室.Text = ""
    cbo科室.Tag = ""
    
    If int部门类型 = 0 Then
        cbo科室.AddItem "所有病区"
        cbo科室.ItemData(cbo科室.NewIndex) = 0
    Else
        cbo科室.AddItem "所有科室"
        cbo科室.ItemData(cbo科室.NewIndex) = 0
    End If
    
    Do While Not rstemp.EOF
        cbo科室.AddItem rstemp!科室
        cbo科室.ItemData(cbo科室.NewIndex) = rstemp!Id
        rstemp.MoveNext
    Loop
    
    cbo科室.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetNode()
    Dim rstemp As ADODB.Recordset
    Dim strCurNode As String
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.编号, b.名称 " & _
        " From 部门表 A, Zlnodelist B " & _
        " Where a.站点 = b.编号 And a.Id In " & _
        " (Select 部门id From 部门性质说明 Where 工作性质 In ('检查', '检验', '治疗', '手术', '护理') And 服务对象 In (2, 3)) And " & _
        " (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
        " Order By b.编号 "
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取站点信息")
    
    With cboNode
        .Visible = False
        .Clear
        
        If rstemp.RecordCount > 0 Then
            .Visible = True
            .AddItem "所有站点"
            
            Do While Not rstemp.EOF
                If strCurNode <> rstemp!编号 Then
                    strCurNode = rstemp!编号
                    .AddItem rstemp!名称
                    .ItemData(.NewIndex) = rstemp!编号
                End If
                rstemp.MoveNext
            Loop
            .ListIndex = 0
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub IniDept()
    If Lbl科室.Tag = "" Then
        Lbl科室.Tag = "-1"
        opt科室_Click (0)
    End If
End Sub

Private Sub cmdPrint_Click()
'功能：打印退药通知单
    Dim StrDate As String
    
    If Trim(vsfDetail(1).TextMatrix(vsfDetail(1).Row, vsfDetail(1).ColIndex("申请时间"))) = "" Then Exit Sub
    StrDate = Format(vsfDetail(1).TextMatrix(vsfDetail(1).Row, vsfDetail(1).ColIndex("审核时间")), "yyyy-MM-dd HH:mm:ss")
    
    If Not IsDate(StrDate) Then
        MsgBox "请在中间列表中选择明细记录。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & StrDate, "包装系数=" & IIf(mstrUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
End Sub

Private Sub cmdRefresh_Click()
    If Trim(txtPati.Text) <> "" Then
        Call txtPati_KeyDown(vbKeyReturn, 0)
    Else
        Call GetRecord(Val(sstabList.Tab))
    End If
End Sub

Private Sub cmdVerify_Click()
    '执行预调价
    Call setNOtExcetePrice
    
    If sstabList.Tab = 0 Then
        Call Oper_Verify
    Else
        Call Oper_ReVerify
    End If
End Sub

Private Sub Form_Activate()
    If mstrReceiveMsg <> "" Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Dim objMenu As Menu
    Dim intCount As Integer
    Dim strCardName As String
    Dim int录入方式 As Integer
    Dim int显示方式 As Integer
    
    mblnStart = False
    
    mstrReturnWriteOffInfo = ""
    
    mstrPrivs = GetPrivFunc(glngSys, 1342)
    
    Call IniDate
    Call GetStockName
    Call GetNode
    Call IniDept
    
    mbln审核出院销账申请 = (Val(zldatabase.GetPara("审核出院病人的销账申请", glngSys, 1342, 0)) = 1)
    mint是否可以销帐拒绝 = Val(zldatabase.GetPara("是否可以销帐拒绝", glngSys, 1342, 1))
    
    mint药品名称 = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "药品名称显示方式", 0)))
    If mint药品名称 > 2 Or mint药品名称 < 0 Then mint药品名称 = 0
    
    cbo科室.Tag = "-1"
    
    '消费卡菜单处理
    If mintCardCount > 0 Then
        For intCount = 0 To UBound(Split(mstrCardType, ";"))
            '取银行卡名称
            strCardName = Split(Split(mstrCardType, ";")(intCount), "|")(1)
            
            '动态添加菜单
            Load Me.mnuPatiItem(Me.mnuPatiItem.UBound + 1)
            Set objMenu = Me.mnuPatiItem(Me.mnuPatiItem.UBound)
            objMenu.Caption = strCardName & "(" & 3 + intCount & ")"
            objMenu.Tag = Split(mstrCardType, ";")(intCount)
        Next
    End If
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then
        int录入方式 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\药品销帐", "录入方式", "0"))
        If int录入方式 < 0 Or int录入方式 > mnuPatiItem.count - 1 Then
            int录入方式 = 0
        End If
        mnuPatiItem_Click int录入方式
        
        int显示方式 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\药品销帐", "显示方式", "0"))
        If int显示方式 = 0 Then
            optListType(0).Value = True
        Else
            optListType(1).Value = True
        End If
    End If
    
    Call IniGrid(0, 0)
    Call IniGrid(1, 0)
    
    mblnStart = True
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long

    If WindowState = 1 Then Exit Sub
    On Error Resume Next
        
    If Me.Width < 13140 Then
        Me.Width = 13140
        Me.ScaleWidth = 12900
    End If
    If Me.Height < 8940 Then
        Me.Height = 8940
        Me.ScaleHeight = 8370
    End If
    
    cmdVerify.Left = Me.ScaleWidth - cmdVerify.Width - 50
    cmdExit.Left = cmdVerify.Left
    CmdHelp.Left = cmdVerify.Left
    cmdPrint.Left = cmdVerify.Left
    fraCondition.Left = Me.ScaleLeft + 20
    fraCondition.Width = Me.ScaleWidth - cmdVerify.Width - 100
    
    sstabList.Width = Me.ScaleWidth
    sstabList.Height = Me.ScaleHeight - fraCondition.Height
    
    '未审核
    vsfMain(0).Width = sstabList.Width - 90
    vsfDetail(0).Width = vsfMain(0).Width
    vsfBatch(0).Width = vsfMain(0).Width
    picHsc(0).Width = sstabList.Width
    picBatHsc(0).Width = sstabList.Width
    
    vsfBatch(0).Top = sstabList.Height - vsfBatch(0).Height - 50
    picBatHsc(0).Top = vsfBatch(0).Top - picBatHsc(0).Height - 50
    
    If picBatHsc(0).Top - vsfDetail(0).Top - 50 < 600 Then
        vsfDetail(0).Height = 2295
        vsfDetail(0).Top = picBatHsc(0).Top - vsfDetail(0).Height
        picHsc(0).Top = vsfDetail(0).Top - picHsc(0).Height - 50
    Else
        vsfDetail(0).Height = picBatHsc(0).Top - vsfDetail(0).Top - 50
    End If
    picHsc(0).Top = vsfDetail(0).Top - picHsc(0).Height - 50
    vsfMain(0).Height = picHsc(0).Top - vsfMain(0).Top - 50
    
    With vsfList(0)
        .Top = vsfMain(0).Top
        .Left = vsfMain(0).Left
        .Width = vsfMain(0).Width
        .Height = picBatHsc(0).Top - .Top - 50
    End With
    
    If optListType(0).Value = True Then
        vsfMain(0).Visible = True
        vsfDetail(0).Visible = True
        picHsc(0).Visible = True
        
        vsfList(0).Visible = False
    Else
        vsfMain(0).Visible = False
        vsfDetail(0).Visible = False
        picHsc(0).Visible = False
        
        vsfList(0).Visible = True
    End If
        
    '已审核
    vsfMain(1).Width = sstabList.Width - 90
    vsfDetail(1).Width = vsfMain(1).Width
    vsfBatch(1).Width = vsfMain(1).Width
    picHsc(1).Width = sstabList.Width
    picBatHsc(1).Width = sstabList.Width
    
    vsfBatch(1).Top = sstabList.Height - vsfBatch(1).Height - 50
    picBatHsc(1).Top = vsfBatch(1).Top - picBatHsc(1).Height - 50
    
    If picBatHsc(1).Top - vsfDetail(1).Top - 50 < 600 Then
        vsfDetail(1).Height = 2295
        vsfDetail(1).Top = picBatHsc(1).Top - vsfDetail(1).Height
        picHsc(1).Top = vsfDetail(1).Top - picHsc(1).Height - 50
    Else
        vsfDetail(1).Height = picBatHsc(1).Top - vsfDetail(1).Top - 50
    End If
    picHsc(1).Top = vsfDetail(1).Top - picHsc(1).Height - 50
    vsfMain(1).Height = picHsc(1).Top - vsfMain(1).Top - 50
    
    With vsfList(1)
        .Top = vsfMain(1).Top
        .Left = vsfMain(1).Left
        .Width = vsfMain(1).Width
        .Height = picBatHsc(1).Top - .Top - 50
    End With
    
    If optListType(0).Value = True Then
        vsfMain(1).Visible = True
        vsfDetail(1).Visible = True
        picHsc(1).Visible = True
        
        vsfList(1).Visible = False
    Else
        vsfMain(1).Visible = False
        vsfDetail(1).Visible = False
        picHsc(1).Visible = False
        
        vsfList(1).Visible = True
    End If
    
    If cboNode.Visible Then
        lblNode.Visible = True
    Else
        lblNode.Visible = False
        lblDept.Left = lblNode.Left
        cbo科室.Left = cboNode.Left
    End If
    
    Me.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mblnStart = False
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\药品销帐", "录入方式", Val(lblPatiInputType.Tag)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\药品销帐", "显示方式", IIf(optListType(0).Value = True, 0, 1)
    End If
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PopupMenu mnuPati, 2, lblPatiInputType.Left + lblPatiInputType.Width - 30, txtPati.Top
    End If
End Sub

Private Sub mnuPatiItem_Click(index As Integer)
    Dim i As Integer
    
    lblPatiInputType.Tag = index
    txtPati.Text = ""
    txtPati.PasswordChar = ""
    txtPati.MaxLength = 0
    
    Select Case index
        Case FindType.住院号
            lblPatiInputType.Caption = "住院号↓"
        Case FindType.Id
            lblPatiInputType.Caption = "ID↓"
        Case FindType.床号
            lblPatiInputType.Caption = "床号↓"
        Case Else
            lblPatiInputType.Caption = Split(mnuPatiItem(index).Tag, "|")(gCardFormat.全名) & "↓"
            
            '其他消费卡
            txtPati.MaxLength = Val(Split(mnuPatiItem(index).Tag, "|")(gCardFormat.卡号长度))
            txtPati.PasswordChar = IIf(Trim(Split(mnuPatiItem(index).Tag, "|")(gCardFormat.卡号密文)) <> "", "*", "")
    End Select
    
    For i = 0 To mnuPatiItem.count - 1
        mnuPatiItem(i).Checked = (i = index)
    Next
End Sub
    

Private Sub optListType_Click(index As Integer)
    Dim lngRow As Long
    
    If mblnStart = False Then Exit Sub
    
    If Not mrsDetail Is Nothing Then
        With mrsDetail
            .Filter = ""
            If .RecordCount > 0 Then
                Do While Not .EOF
                    !审核标志 = IIf(index = 0, 1, 0)
                    .Update
                    
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    If Not mrsBatch Is Nothing Then
        With mrsBatch
            .Filter = ""
            If .RecordCount > 0 Then
                Do While Not .EOF
                    !审核标志 = IIf(index = 0, 1, 0)
                    .Update
                    
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    DoEvents
    
    Call Form_Resize
    
    If index = 0 Then
        If vsfDetail(Val(sstabList.Tab)).rows > 1 Then
            mlngMainRow = 1
            mlngDetailRow = 1
            vsfMain(Val(sstabList.Tab)).Row = 1
            vsfMain(Val(sstabList.Tab)).SetFocus
        End If
    Else
        If vsfList(Val(sstabList.Tab)).rows > 1 Then
            If Val(sstabList.Tab) = 0 Then
                For lngRow = 1 To vsfList(Val(sstabList.Tab)).rows - 2
                    vsfList(Val(sstabList.Tab)).TextMatrix(lngRow, vsfList(Val(sstabList.Tab)).ColIndex("审核标志")) = ""
                Next
            End If
            
            mlngListRow = 1
            vsfList(Val(sstabList.Tab)).Row = 1
            vsfList(Val(sstabList.Tab)).SetFocus
        End If
    End If
End Sub

Private Sub opt科室_Click(index As Integer)
    If Val(Lbl科室.Tag) <> index Then
        If index = 1 Then
            mnuPatiItem(2).Enabled = False
            If Val(lblPatiInputType.Tag) = FindType.床号 Then
                Call mnuPatiItem_Click(0)
            End If
        Else
            mnuPatiItem(2).Enabled = True
        End If
        
        Call GetDept(index)
        Lbl科室.Tag = index
    End If
End Sub






Private Sub picHsc_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfMain(index).Height + y <= 500 Or vsfDetail(index).Height - y <= 500 Then Exit Sub
        
        picHsc(index).Top = picHsc(index).Top + y
        vsfMain(index).Height = vsfMain(index).Height + y
        vsfDetail(index).Height = vsfDetail(index).Height - y
        vsfDetail(index).Top = vsfDetail(index).Top + y
        
        Me.Refresh
    End If
End Sub


Private Sub sstabList_Click(PreviousTab As Integer)
    If sstabList.Tab = 0 Then
        lbl时间.Caption = "申请期间"
        If vsfMain(0).rows > 1 Then
            If vsfMain(0).TextMatrix(1, vsfMain(0).ColIndex("药品名称")) <> "" Then
                cmdAllSelect.Enabled = True
                cmdAllUnSelect.Enabled = True
            End If
        End If
        cmdVerify.Caption = "退药销账(&V)"
        cmdVerify.Enabled = True
        chkNoTime.Value = Val(Mid(chkNoTime.Tag, 1, 1))
        cmdPrint.Enabled = False
    ElseIf sstabList.Tab = 1 Then
        lbl时间.Caption = "审核期间"
        cmdAllSelect.Enabled = False
        cmdAllUnSelect.Enabled = False
        cmdVerify.Caption = "重审销账(&C)"
        cmdVerify.Enabled = False
        chkNoTime.Value = Val(Mid(chkNoTime.Tag, 3, 1))
        cmdPrint.Enabled = True
    End If
    
    If optListType(0).Value = True Then
        vsfMain(sstabList.Tab).Visible = True
        vsfDetail(sstabList.Tab).Visible = True
        vsfList(sstabList.Tab).Visible = False
    Else
        vsfMain(sstabList.Tab).Visible = False
        vsfDetail(sstabList.Tab).Visible = False
        vsfList(sstabList.Tab).Visible = True
    End If
End Sub

Private Sub txtPati_Change()
    If Val(lblPatiInputType.Tag) > 2 Then
        If Len(txtPati.Text) = txtPati.MaxLength Then
             Call txtPati_KeyDown(vbKeyReturn, 0)
        End If
    End If
End Sub
Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub
Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim vRect As RECT
    Dim strSqlCon As String
    Dim lng病人ID As Long
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    txtPati.Text = Trim(txtPati.Text)
    
    If txtPati.Text = "" Then
        Call GetRecord(Val(sstabList.Tab))
        Exit Sub
    End If
    
    Select Case Val(lblPatiInputType.Tag)
        Case FindType.住院号
            If Not IsNumeric(txtPati.Text) Then Exit Sub
            strSqlCon = " And A.住院号 = [1] "
        Case FindType.Id
            strSqlCon = " And A.病人ID = [2] "
        Case FindType.床号
            If cbo科室.ListIndex = 0 Then
                MsgBox "请选择病区！"
                Exit Sub
            End If
            strSqlCon = " And A.当前病区id = [3] And A.当前床号 = [1] "
        Case Else
            '其他消费卡，按病人ID查找
            lng病人ID = zlfuncCard_GetPatiID(mobjSquareCard, Val(Split(mnuPatiItem(Val(lblPatiInputType.Tag)).Tag, "|")(gCardFormat.卡类别ID)), txtPati.Text)
            strSqlCon = " And A.病人ID = [4]"
    End Select
    
    gstrSQL = "Select A.病人id As ID, A.姓名, A.住院号, B.名称, A.当前床号 As 床号 " & _
        " From 病人信息 A, 部门表 B  Where A.当前病区id = B.Id" & IIf(mbln审核出院销账申请 = True, "(+)", "")
    
    gstrSQL = gstrSQL & strSqlCon & " Order By B.名称, 住院号"
    
    On Error GoTo errHandle
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", txtPati.Text, Val(txtPati.Text), Val(cbo科室.ItemData(cbo科室.ListIndex)), lng病人ID)
    
    If rstemp.EOF Then
        txtPati.Text = ""
        txtPati.Tag = ""
        txtPati.SetFocus
        Exit Sub
    ElseIf rstemp.RecordCount = 1 Then
        '只有一条记录
        txtPati.Text = rstemp!姓名
        txtPati.Tag = rstemp!Id
    Else
        '有多条记录则显示，供选择
        vRect = zlControl.GetControlRect(txtPati.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = txtPati.Height
        
        Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "选择病人", False, "", "选择病人", False, False, True, sngX, sngY, sngH, blnCancel, False, False, txtPati.Text, Val(txtPati.Text), Val(cbo科室.ItemData(cbo科室.ListIndex)))
    
        If blnCancel = True Then Exit Sub
        
        If rstemp Is Nothing Then
            txtPati.Text = ""
            txtPati.Tag = ""
            txtPati.SetFocus
            Exit Sub
        Else
            txtPati.Text = rstemp!姓名
            txtPati.Tag = rstemp!Id
        End If
    End If
    
    If Trim(txtPati.Text) <> "" Then Call GetRecord(Val(sstabList.Tab))
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub


Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If Val(lblPatiInputType.Tag) = FindType.住院号 Or Val(lblPatiInputType.Tag) = FindType.Id Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <> vbKeyEscape Or KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = 0
    ElseIf Val(lblPatiInputType.Tag) > 2 Then
        '其他的是消费卡
        If InStr(":：;；?？''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    If Trim(txtPati.Text) = "" Then txtPati.Tag = ""
End Sub

Private Sub vsfBatch_EnterCell(index As Integer)
    If index = 1 Then Exit Sub
    
    '通常情况下是不能够修改销帐数量的，除了同一张单据同一药品存在多个批次的情况下，当然还要准退数量足够
    With vsfBatch(0)
        .Editable = flexEDNone
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("单据")) = "" Then Exit Sub
        If .Col <> .ColIndex("销帐数量") Then Exit Sub
        If mblnAllowChange = False Then Exit Sub
        If .rows = 2 Then Exit Sub
        
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub vsfBatch_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 1 Then Exit Sub
    
    With vsfBatch(0)
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("单据")) = "" Then Exit Sub
        If .Col <> .ColIndex("销帐数量") Then Exit Sub
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        Call vsfBatch_ValidateEdit(0, .Row, .Col, True)
    End With
End Sub


Private Sub vsfBatch_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If index = 1 Then Exit Sub
    
    '只能输入数字
    If Col = vsfBatch(index).ColIndex("销帐数量") Then
        If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfBatch_RowColChange(index As Integer)
    '移动第一栏的标记到当前行！
    With vsfBatch(index)
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub

Private Sub vsfBatch_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblNewQuantity As Double
    Dim dblLeavingsQuantity As Double
    Dim dblQuantity As Double
    Dim i As Integer
    
    If index = 1 Then Exit Sub

    With vsfBatch(0)
        dblNewQuantity = Val(.EditText)
        
        If dblNewQuantity = Val(.TextMatrix(Row, .ColIndex("销帐数量"))) Then Exit Sub
        
        If dblNewQuantity > Val(.TextMatrix(Row, .ColIndex("准退数量"))) Or dblNewQuantity < 0 Then
            Cancel = True
            Exit Sub
        End If
        
        '计算差额数量
        dblLeavingsQuantity = Val(.TextMatrix(Row, .ColIndex("销帐数量"))) - dblNewQuantity
        
        '把差额数量分配到其他批次上
        For i = 1 To .rows - 1
            If i <> Row Then
                dblQuantity = Val(.TextMatrix(i, .ColIndex("销帐数量")))
                If dblQuantity + dblLeavingsQuantity <= Val(.TextMatrix(i, .ColIndex("准退数量"))) And dblQuantity + dblLeavingsQuantity > 0 Then
                    .TextMatrix(i, .ColIndex("销帐数量")) = dblQuantity + dblLeavingsQuantity
                    dblLeavingsQuantity = 0
                Else
                    .TextMatrix(i, .ColIndex("销帐数量")) = Val(.TextMatrix(i, .ColIndex("准退数量")))
                    dblLeavingsQuantity = dblLeavingsQuantity - (Val(.TextMatrix(i, .ColIndex("准退数量"))) - dblQuantity)
                End If
                
                If dblLeavingsQuantity = 0 Then Exit For
            End If
        Next
        
        '确认最后的当前输入数量
        .EditText = FormatEx(dblNewQuantity + dblLeavingsQuantity, 5)
        .TextMatrix(Row, .ColIndex("销帐数量")) = FormatEx(dblNewQuantity + dblLeavingsQuantity, 5)
        
        '更新记录集中的销帐数量
        For i = 1 To .rows - 1
            mrsBatch.Filter = "单据=" & Val(.TextMatrix(i, .ColIndex("单据"))) & _
                            " And No='" & .TextMatrix(i, .ColIndex("NO")) & "' " & _
                            " And 药品ID=" & Val(.TextMatrix(i, .ColIndex("药品id"))) & _
                            " And 收发序号=" & Val(.TextMatrix(i, .ColIndex("收发序号"))) & _
                            " And 申请时间='" & .TextMatrix(i, .ColIndex("申请时间")) & "' "
            If mrsBatch.EOF Then Exit Sub
    
            mrsBatch!销帐数量 = Val(.TextMatrix(i, .ColIndex("销帐数量"))) * mrsBatch!包装
            mrsBatch.Update
        Next
    End With
End Sub
Private Sub vsfDetail_Click(index As Integer)
    Dim bln更新标志 As Boolean
    Dim int审核标志 As Integer
    
    With vsfDetail(index)
'        If index = 1 Then Exit Sub
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("申请科室")) = "" Then Exit Sub

        If index = 0 And .MouseCol = .ColIndex("审核标志") Then
            If .TextMatrix(.Row, .ColIndex("审核标志")) = "√" Then
                .TextMatrix(.Row, .ColIndex("审核标志")) = IIf(mint是否可以销帐拒绝 = 1, "×", "")
                int审核标志 = IIf(mint是否可以销帐拒绝 = 1, 2, 0)
            ElseIf .TextMatrix(.Row, .ColIndex("审核标志")) = "×" Then
                .TextMatrix(.Row, .ColIndex("审核标志")) = ""
                int审核标志 = 0
            Else
                If Val(.TextMatrix(.Row, .ColIndex("准退数量"))) < Val(.TextMatrix(.Row, .ColIndex("销帐数量"))) Then
                    If mint是否可以销帐拒绝 = 1 Then
                        .TextMatrix(.Row, .ColIndex("审核标志")) = "×"
                        int审核标志 = 2
                    End If
                Else
                    .TextMatrix(.Row, .ColIndex("审核标志")) = "√"
                    int审核标志 = 1
                End If
            End If
            bln更新标志 = True
            
            '更新记录集审核标记
            mrsDetail.Filter = "单据=" & Val(.TextMatrix(.Row, .ColIndex("单据"))) & _
                " And No='" & .TextMatrix(.Row, .ColIndex("NO")) & "' " & _
                " And 药品ID=" & Val(.TextMatrix(.Row, .ColIndex("药品id"))) & _
                " And 费用ID=" & Val(.TextMatrix(.Row, .ColIndex("费用id"))) & _
                " And 申请时间='" & .TextMatrix(.Row, .ColIndex("申请时间")) & "' "
            If mrsDetail.RecordCount > 0 Then
                mrsDetail!审核标志 = int审核标志
                mrsDetail.Update
            End If
        End If
        
        mblnAllowChange = False
        If Val(.TextMatrix(.Row, .ColIndex("准退数量"))) > Val(.TextMatrix(.Row, .ColIndex("销帐数量"))) Then
            mblnAllowChange = True
        End If
        
        '提取批次明细数据
        If mlngDetailRow <> .Row Or bln更新标志 = True Then
            mlngDetailRow = .Row
            If .rows > 1 Then
                If index = 0 Then
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("申请时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), bln更新标志, int审核标志)
                Else
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("审核时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), bln更新标志, int审核标志)
                End If
            End If
        End If
    End With

End Sub

Private Sub vsfDetail_EnterCell(index As Integer)
    
    If index = 1 Then
        If vsfDetail(index).TextMatrix(vsfDetail(index).Row, vsfDetail(index).ColIndex("审核标志")) = "×" Then
            vsfBatch(index).TextMatrix(0, vsfBatch(index).ColIndex("销帐数量")) = "可销帐数量"
        Else
            vsfBatch(index).TextMatrix(0, vsfBatch(index).ColIndex("销帐数量")) = "销帐数量"
        End If
    
        With vsfDetail(index)
            If .Row = 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("审核标志")) = "×" Then
                cmdVerify.Enabled = True
            Else
                cmdVerify.Enabled = False
            End If
        End With
    End If
End Sub


Private Sub vsfDetail_RowColChange(index As Integer)
    '移动第一栏的标记到当前行！
    With vsfDetail(index)
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
            
            '提取批次明细数据
            If mlngDetailRow <> .Row Then
                mlngDetailRow = .Row
                If .rows > 1 Then
                    If index = 0 Then
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("申请时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), False, 0)
                    Else
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("审核时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), False, 0)
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfList_AfterSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index = 0 Then
        If vsfList(0).rows < 2 Then Exit Sub
        '显示销帐金额合计信息
        vsfList(0).rows = vsfList(0).rows + 1
        vsfList(0).Cell(flexcpText, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = "销账金额合计：" & FormatEx(mdblSum, 5)
        vsfList(0).Cell(flexcpFontBold, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = True
        vsfList(0).Cell(flexcpForeColor, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = vbRed
        vsfList(0).Cell(flexcpAlignment, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = flexAlignLeftCenter
        vsfList(0).MergeCells = flexMergeRestrictRows
        vsfList(0).MergeRow(vsfList(0).rows - 1) = True
    End If
End Sub

Private Sub vsfList_BeforeSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index = 0 Then
        If vsfList(0).rows > 2 Then vsfList(0).RemoveItem vsfList(0).rows - 1
    End If
End Sub

Private Sub vsfList_Click(index As Integer)
    Dim bln更新标志 As Boolean
    Dim int审核标志 As Integer
    
    With vsfList(index)
        If index = 1 Then Exit Sub
        If .Row = .rows - 1 Then Exit Sub
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("申请科室")) = "" Then Exit Sub
        
        If index = 0 And .MouseCol = .ColIndex("审核标志") Then
            If .TextMatrix(.Row, .ColIndex("审核标志")) = "√" Then
                .TextMatrix(.Row, .ColIndex("审核标志")) = IIf(mint是否可以销帐拒绝 = 1, "×", "")
                int审核标志 = IIf(mint是否可以销帐拒绝 = 1, 2, 0)
            ElseIf .TextMatrix(.Row, .ColIndex("审核标志")) = "×" Then
                .TextMatrix(.Row, .ColIndex("审核标志")) = ""
                int审核标志 = 0
            Else
                If Val(.TextMatrix(.Row, .ColIndex("准退数量"))) < Val(.TextMatrix(.Row, .ColIndex("销帐数量"))) Then
                    If mint是否可以销帐拒绝 = 1 Then
                        .TextMatrix(.Row, .ColIndex("审核标志")) = "×"
                        int审核标志 = 2
                    End If
                Else
                    .TextMatrix(.Row, .ColIndex("审核标志")) = "√"
                    int审核标志 = 1
                End If
            End If
            bln更新标志 = True
            
            '更新记录集审核标记
            mrsDetail.Filter = "单据=" & Val(.TextMatrix(.Row, .ColIndex("单据"))) & _
                " And No='" & .TextMatrix(.Row, .ColIndex("NO")) & "' " & _
                " And 药品ID=" & Val(.TextMatrix(.Row, .ColIndex("药品id"))) & _
                " And 费用ID=" & Val(.TextMatrix(.Row, .ColIndex("费用id"))) & _
                " And 申请时间='" & .TextMatrix(.Row, .ColIndex("申请时间")) & "' "
            If mrsDetail.RecordCount > 0 Then
                mrsDetail!审核标志 = int审核标志
                mrsDetail.Update
            End If
        End If
        
        mblnAllowChange = False
        If Val(.TextMatrix(.Row, .ColIndex("准退数量"))) > Val(.TextMatrix(.Row, .ColIndex("销帐数量"))) Then
            mblnAllowChange = True
        End If
        
        '提取批次明细数据
        If mlngListRow <> .Row Or bln更新标志 = True Then
            mlngListRow = .Row
            If .rows > 1 Then
                If index = 0 Then
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("申请时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), bln更新标志, int审核标志)
                Else
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("审核时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), bln更新标志, int审核标志)
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfList_EnterCell(index As Integer)
    If index = 1 Then
        With vsfList(index)
            If .Row = 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("审核标志")) = "×" Then
                cmdVerify.Enabled = True
            Else
                cmdVerify.Enabled = False
            End If
        End With
    End If
End Sub

Private Sub vsfList_RowColChange(index As Integer)
    '移动第一栏的标记到当前行！
    With vsfList(index)
        If .Row < 1 Then Exit Sub
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
            
            '提取批次明细数据
            If mlngListRow <> .Row Then
                mlngListRow = .Row
                If .rows > 1 Then
                    If index = 0 Then
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("申请时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), False, 0)
                    Else
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("药品id"))), .TextMatrix(.Row, .ColIndex("审核时间")), Val(.TextMatrix(.Row, .ColIndex("费用id"))), False, 0)
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfMain_AfterSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index = 0 Then
        If vsfMain(index).rows < 2 Then Exit Sub
        vsfMain(index).rows = vsfMain(index).rows + 1
        vsfMain(index).Cell(flexcpText, vsfMain(index).rows - 1, 1, vsfMain(index).rows - 1, vsfMain(index).Cols - 1) = "销账金额合计：" & FormatEx(mdblSum, 5)
        vsfMain(index).Cell(flexcpFontBold, vsfMain(index).rows - 1, 1, vsfMain(index).rows - 1, vsfMain(index).Cols - 1) = True
        vsfMain(index).Cell(flexcpForeColor, vsfMain(index).rows - 1, 1, vsfMain(index).rows - 1, vsfMain(index).Cols - 1) = vbRed
        vsfMain(index).MergeCells = flexMergeRestrictRows
        vsfMain(index).MergeRow(vsfMain(index).rows - 1) = True
    End If
End Sub

Private Sub vsfMain_BeforeSort(index As Integer, ByVal Col As Long, Order As Integer)
    If vsfMain(index).rows <= 2 Then Exit Sub
    If index = 0 Then
        vsfMain(index).RemoveItem vsfMain(index).rows - 1
    End If
End Sub

Private Sub vsfMain_EnterCell(index As Integer)
    With vsfMain(index)
        If .Row < 1 Or (.Row = .rows - 1 And index = 0) Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("收费细目id")) = "" Then Exit Sub
        
        '提取明细数据
        If mlngMainRow = .Row Then Exit Sub
        mlngMainRow = .Row
        Call LoadDetailList(index, Val(.TextMatrix(.Row, .ColIndex("收费细目id"))))
    End With
    
    '提取批次明细数据
    If vsfDetail(index).rows >= 2 Then
        If index = 0 Then
            Call LoadBatchList(index, Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("单据"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("NO")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("药品id"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("申请时间")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("费用id"))), False, IIf(optListType(0).Value = True, 1, 0))
        Else
            Call LoadBatchList(index, Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("单据"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("NO")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("药品id"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("审核时间")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("费用id"))), False, IIf(optListType(0).Value = True, 1, 0))
        End If
    End If
End Sub


Private Sub vsfMain_RowColChange(index As Integer)
    '移动第一栏的标记到当前行！
    With vsfMain(index)
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub


