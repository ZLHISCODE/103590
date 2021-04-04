VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmStuffPrice 
   Caption         =   "材料调价单"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmStuffPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   11265
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picStoceBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   4680
      ScaleHeight     =   5085
      ScaleWidth      =   8685
      TabIndex        =   32
      Top             =   2520
      Width           =   8685
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   240
         ScaleHeight     =   1890
         ScaleWidth      =   7755
         TabIndex        =   39
         Top             =   3120
         Width           =   7755
         Begin VB.CheckBox chk自动计算 
            Caption         =   "自动根据库存计算应付变动金额"
            Height          =   195
            Left            =   0
            TabIndex        =   40
            Top             =   135
            Width           =   2985
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPay 
            Height          =   1440
            Left            =   0
            TabIndex        =   41
            Top             =   480
            Width           =   6735
            _cx             =   11880
            _cy             =   2540
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
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStuffPrice.frx":058A
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
            ExplorerBar     =   7
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
      Begin VB.PictureBox picStoce 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2850
         Left            =   360
         ScaleHeight     =   2850
         ScaleWidth      =   8190
         TabIndex        =   33
         Top             =   720
         Width           =   8190
         Begin VB.CheckBox chk显示所有材料 
            Caption         =   "显示当前所有卫材批次"
            Height          =   270
            Left            =   3630
            TabIndex        =   36
            Top             =   0
            Width           =   2150
         End
         Begin VB.CheckBox chk批次 
            Caption         =   "按库房批次更改"
            Height          =   210
            Left            =   135
            TabIndex        =   37
            Top             =   75
            Width           =   1620
         End
         Begin VB.CheckBox chk应付 
            Caption         =   "应付帐款调整"
            Height          =   195
            Left            =   1845
            TabIndex        =   35
            Top             =   60
            Width           =   1635
         End
         Begin VB.CommandButton cmdPrintStoce 
            Caption         =   "打印库存变动表(&S)…"
            Height          =   350
            Left            =   4800
            Picture         =   "frmStuffPrice.frx":0676
            TabIndex        =   34
            Top             =   360
            Width           =   1965
         End
         Begin VSFlex8Ctl.VSFlexGrid vsStoce 
            Height          =   2340
            Left            =   480
            TabIndex        =   38
            Top             =   840
            Width           =   6510
            _cx             =   11483
            _cy             =   4128
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
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStuffPrice.frx":07C0
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
            ExplorerBar     =   7
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
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2445
         Left            =   240
         TabIndex        =   42
         Top             =   300
         Width           =   7950
         _Version        =   589884
         _ExtentX        =   14023
         _ExtentY        =   4313
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picSeach 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   150
      ScaleHeight     =   5010
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   1725
      Visible         =   0   'False
      Width           =   4425
      Begin VB.Frame fraCost 
         Caption         =   "成本价调整"
         Height          =   1005
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   4125
         Begin VB.CommandButton cmdPriver 
            Caption         =   "…"
            Height          =   270
            Left            =   3800
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   225
            Width           =   255
         End
         Begin VB.TextBox txtPriver 
            Height          =   300
            Left            =   705
            TabIndex        =   13
            Top             =   210
            Width           =   3090
         End
         Begin VB.TextBox txt加成率 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   705
            TabIndex        =   16
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "供应商"
            Height          =   180
            Left            =   90
            TabIndex        =   12
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "％"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1485
            TabIndex        =   17
            Top             =   645
            Width           =   225
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "加成率"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   15
            Top             =   660
            Width           =   540
         End
      End
      Begin VB.CommandButton cmd调价 
         Caption         =   "按条件过滤调价(&R)"
         Height          =   350
         Left            =   2040
         TabIndex        =   18
         Top             =   4440
         Width           =   2250
      End
      Begin VB.Frame fra调整额 
         Caption         =   "调整方式"
         Height          =   1155
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4125
         Begin VB.TextBox txt调整额 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2760
            TabIndex        =   8
            Top             =   270
            Width           =   735
         End
         Begin VB.ComboBox cbo调整方式 
            Height          =   300
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   270
            Width           =   2580
         End
         Begin VB.Label lbl调整 
            AutoSize        =   -1  'True
            Caption         =   "％"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3540
            TabIndex        =   9
            Top             =   315
            Width           =   225
         End
         Begin VB.Label lblInfor 
            Caption         =   "根据成本价，输入新的加成率重新加成调价"
            Height          =   255
            Left            =   165
            TabIndex        =   10
            Top             =   690
            Width           =   3660
         End
      End
      Begin VB.Frame fra 
         Caption         =   "应用范围"
         Height          =   1005
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4125
         Begin VB.OptionButton opt应用 
            Caption         =   "按制定品种卫材(&2)"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   1860
         End
         Begin VB.OptionButton opt应用 
            Caption         =   "当前分类下所有卫材(&1)"
            Height          =   375
            Index           =   1
            Left            =   1845
            TabIndex        =   5
            Top             =   255
            Width           =   2205
         End
         Begin VB.OptionButton opt应用 
            Caption         =   "当前分类卫材(&0)"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   255
            Width           =   1740
         End
      End
      Begin VB.CommandButton cmdType 
         Caption         =   "…"
         Height          =   270
         Left            =   3960
         TabIndex        =   2
         Top             =   150
         Width           =   255
      End
      Begin VB.TextBox txt分类 
         Height          =   300
         Left            =   645
         TabIndex        =   1
         Top             =   135
         Width           =   3420
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卫材"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.PictureBox picPrice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   4680
      ScaleHeight     =   3450
      ScaleWidth      =   10065
      TabIndex        =   21
      Top             =   -360
      Width           =   10065
      Begin VB.CheckBox chkAppAllColumn 
         Caption         =   "修改价格应用于所有列"
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   360
         Width           =   2295
      End
      Begin VB.PictureBox picBakDown 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   0
         ScaleHeight     =   810
         ScaleWidth      =   8850
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2400
         Width           =   8850
         Begin VB.CheckBox Chk定价 
            Caption         =   "时价材料改为定价销售(&D)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   2505
            TabIndex        =   27
            Top             =   525
            Width           =   2370
         End
         Begin VB.TextBox txt调价人 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   90
            Width           =   2445
         End
         Begin VB.TextBox txt说明 
            Height          =   300
            Left            =   825
            TabIndex        =   25
            Top             =   90
            Width           =   4485
         End
         Begin VB.CheckBox chk立即执行 
            Caption         =   "所有价格立即生效(&I)"
            Height          =   210
            Left            =   75
            TabIndex        =   24
            Top             =   525
            Width           =   2040
         End
         Begin MSComCtl2.DTPicker dtp执行日期 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   6285
            TabIndex        =   28
            Top             =   465
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   184418307
            CurrentDate     =   36846.5833333333
         End
         Begin VB.Label lblValuer 
            AutoSize        =   -1  'True
            Caption         =   "调价人"
            Height          =   180
            Left            =   5655
            TabIndex        =   31
            Top             =   150
            Width           =   540
         End
         Begin VB.Label lblRunDate 
            AutoSize        =   -1  'True
            Caption         =   "执行日期"
            Height          =   180
            Left            =   5475
            TabIndex        =   30
            Top             =   525
            Width           =   720
         End
         Begin VB.Label lblSummary 
            AutoSize        =   -1  'True
            Caption         =   "调价说明"
            Height          =   180
            Left            =   30
            TabIndex        =   29
            Top             =   150
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPrice 
         Height          =   2190
         Left            =   0
         TabIndex        =   22
         Top             =   600
         Width           =   10665
         _cx             =   18812
         _cy             =   3863
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffPrice.frx":09A2
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
         ExplorerBar     =   7
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   6870
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffPrice.frx":0C26
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14790
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   10140
      Top             =   4155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPrice.frx":14BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPrice.frx":180E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStuffPrice.frx":1B62
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1530
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkPane 
      Bindings        =   "frmStuffPrice.frx":1AD48
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmStuffPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngBillId As Long                '功能类型:0-调价处理;其他-显示mlngBillId确定的历史调价单
Private mlngStuffId As Long                '进入类型:0-未指定调价卫材;其他-进入时直接显示mlngStuffId的原价格情况
Private Enum 调价方式
        T_售价调价 = 1
        T_成本价调价 = 2
        T_成本和售价调价 = 3
End Enum
Private m调价方式 As 调价方式

Public Enum BillType
    B_单一调价 = 0
    B_批量调价 = 1              '根据卫材分类批量进行调价
    B_查阅 = 2
End Enum
'---------------------------------

Private mintUnit As Integer      '是否以库房单位显示
Private mblnModify As Boolean
Private mblnFirst  As Boolean
Private mBillType As BillType
Private mblnSucces As Boolean
Private mlngPreRow As Long
Private mlngPrice As Long

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Const conMenu_Popup = 1           '工具栏
Private Const conMenu_Preview = 102         '预览(&V)
Private Const conMenu_Print = 103           '打印(&P)
Private Const conMenu_Excel = 104           '输出到&Excel…
Private Const conMenu_Save = 305           '保存
Private Const conMenu_Cancel = 304           '取消
Private Const conMenu_Lable = 300           '计价方式标题
Private Const conMenu_Combo = 301           '计价方式COMBOX


Private Const conMenu_Help_Help = 901        '帮助主题(&H)


'CommandBar辅助热键
Private Const FSHIFT = 4
Private Const FCONTROL = 8
Private Const FALT = 16

Private Const ID_PANE_SEARCH = 1
Private Const ID_PANE_PRICE = 2
Private Const ID_PANE_STOCE = 3
Private mobjFindKey As CommandBarControl
 
Private mlngModule As Long
Private mstrPrivs As String
Private mdbl加成率 As Double
Private mlng供应商ID As Long

'-----------------------------------------------------------------------------------------------------------------
Private Enum mPageNum
    Page_库存调整 = 0
    Page_应付调整 = 1
End Enum

Private Sub InitCommandBar()
    '-------------------------------------------------------------------------------------------
    '功能:初始化菜单
    '参数:
    '返回:
    '编制:刘兴宏
    '日期:2007/08/07
    '-------------------------------------------------------------------------------------------
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objDeptBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
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
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    
    Set cbsMain.Icons = imgPublic.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
  
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    Dim objComBar As CommandBarComboBox
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Preview, "预览")
        
        Set objControl = .Add(xtpControlButton, conMenu_Save, "确定"): objControl.IconId = conMenu_Save
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Cancel, "取消"):    objControl.IconId = conMenu_Cancel
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        
        If mBillType = B_查阅 Then
            m调价方式 = T_售价调价
        Else
        Set objControl = .Add(xtpControlLabel, conMenu_Lable, "调价方式")
        objControl.Flags = xtpFlagRightAlign
        Set objComBar = .Add(xtpControlComboBox, conMenu_Combo, "调价方式")
        objComBar.Flags = xtpFlagRightAlign
        Dim intIndex As Integer
        intIndex = 1
        If zlStr.IsHavePrivs(mstrPrivs, "售价管理") Then
            objComBar.AddItem "按售价调价"
            objComBar.ItemData(intIndex) = 1
            intIndex = intIndex + 1
        End If
        If InStr(1, mstrPrivs, ";成本价管理;") <> 0 Then
            objComBar.AddItem "按成本价调价"
            objComBar.ItemData(intIndex) = 2
            intIndex = intIndex + 1
        End If
        If intIndex = 3 Then
            objComBar.AddItem "按售价和成本价调价"
            objComBar.ItemData(intIndex) = 3
        End If
        objComBar.ListIndex = 1: objComBar.Width = 120
        m调价方式 = objComBar.ItemData(1)
       End If
   End With

    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlLabel Then
        Else
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
     
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_Print   '打印
    End With
End Sub

Private Function InitPanel()
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置区域信息
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 12:19:20
    '-----------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, objPaneFind As Pane
    
    With DkPane
        .ImageList = imlPaneIcons '
        
        Set objPaneFind = DkPane.CreatePane(ID_PANE_SEARCH, 400, 400, DockLeftOf, Nothing)
        objPaneFind.Title = "批量调价条件设置"
        objPaneFind.Options = PaneNoCloseable
        objPaneFind.MinTrackSize.Width = 295
        objPaneFind.MaxTrackSize.Width = 495
        Set objPane = DkPane.CreatePane(ID_PANE_PRICE, 400, 400, DockRightOf, objPaneFind)
        objPane.Title = "调价信息"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picPrice.hwnd
        objPaneFind.Hide
        Set objPane = DkPane.CreatePane(ID_PANE_STOCE, 400, 400, DockBottomOf, objPane)
        objPane.Title = "库存变动信息"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picStoceBack.hwnd
        
        .SetCommandBars Me.cbsMain
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Function Get调价项目() As Boolean
    '--------------------------------------------------------------------------------------------------
    '功能:重新根据相关条件获取调价项目
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/18
    '--------------------------------------------------------------------------------------------------
    Dim lng分类id As Long
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHandle
    If m调价方式 = T_售价调价 Then
        mlng供应商ID = 0
        mdbl加成率 = 0
    Else
        mlng供应商ID = Val(txtPriver.Tag)
        mdbl加成率 = Val(txt加成率.Text)
    End If
    lng分类id = Val(txt分类.Tag)
    If lng分类id = 0 Then
        ShowMsgBox "未选择分类,请检查!"
        zlControl.ControlSetFocus txt分类, True
        Exit Function
    End If
   
    gstrSQL = "" & _
    "    Select I.ID, I.编码, I.名称, I.规格, I.产地, I.计算单位, P.包装单位, Decode(I.是否变价, 1, '时价', '定价') 类型," & _
    "           P.指导批发价,P.指导零售价, P.成本价,  " & IIf(mintUnit = 0, "1", "nvl(p.换算系数,1)") & " As 换算系数,P.跟踪在用" & _
    "    From 收费项目目录 I, 材料特性 P, 诊疗项目目录 M " & _
    "    Where   I.ID = P.材料id And P.诊疗id = M.ID  And "
    If opt应用(2).Value = True Then
        gstrSQL = gstrSQL & " m.id=[1]"
    Else
        If opt应用(0).Value Then
            gstrSQL = gstrSQL & _
            "          M.分类id =[1]"
        Else
            gstrSQL = gstrSQL & _
            "          M.分类id In (Select ID From 诊疗分类目录 Start With ID = [1] Connect By Prior id = 上级ID)"
        End If
    End If
    If mlng供应商ID <> 0 Then
        gstrSQL = gstrSQL & " And exists(Select 1 From 药品库存 where I.id=药品id and 上次供应商ID=[2])"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng分类id, mlng供应商ID)
    
    With vsPrice
         .Redraw = flexRDNone
         i = 1
         If rsTemp.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .Cell(flexcpData, 1, i) = ""
            Next
            Call InitControl
            .Redraw = flexRDBuffered
            Get调价项目 = True
            Exit Function
         Else
            .Rows = rsTemp.RecordCount + 1
         End If
         Call InitControl
        .Col = .ColIndex("现价")
         Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("品名")) = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!名称)
            .Cell(flexcpData, i, .ColIndex("品名")) = zlStr.Nvl(rsTemp!Id)
            .TextMatrix(i, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
            .TextMatrix(i, .ColIndex("产地")) = zlStr.Nvl(rsTemp!产地)
            .TextMatrix(i, .ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!计算单位), zlStr.Nvl(rsTemp!包装单位))
            .TextMatrix(i, .ColIndex("类型")) = zlStr.Nvl(rsTemp!类型)
            .Cell(flexcpData, i, .ColIndex("类型")) = zlStr.Nvl(rsTemp!跟踪在用)

            .TextMatrix(i, .ColIndex("系数")) = zlStr.Nvl(rsTemp!换算系数)
            .TextMatrix(i, .ColIndex("原成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_成本价)
            .Cell(flexcpData, i, .ColIndex("原成本价")) = Val(zlStr.Nvl(rsTemp!成本价))
            .TextMatrix(i, .ColIndex("现成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_成本价)
            .Cell(flexcpData, i, .ColIndex("现成本价")) = Val(zlStr.Nvl(rsTemp!成本价))
            
            .TextMatrix(i, .ColIndex("原采购限价")) = Format(Val(zlStr.Nvl(rsTemp!指导批发价)) * Val(rsTemp!换算系数), mFMT.FM_成本价)
            .TextMatrix(i, .ColIndex("现采购限价")) = .TextMatrix(i, .ColIndex("原采购限价"))
            .Cell(flexcpData, i, .ColIndex("原采购限价")) = Val(zlStr.Nvl(rsTemp!指导批发价))
            .Cell(flexcpData, i, .ColIndex("现采购限价")) = Val(zlStr.Nvl(rsTemp!指导批发价))
            
            .TextMatrix(i, .ColIndex("指导零售价")) = Format(Val(zlStr.Nvl(rsTemp!指导零售价)) * Val(rsTemp!换算系数), mFMT.FM_零售价)
            .TextMatrix(i, .ColIndex("原指导售价")) = .TextMatrix(i, .ColIndex("指导零售价"))
            .TextMatrix(i, .ColIndex("现指导售价")) = .TextMatrix(i, .ColIndex("原指导售价"))
            
            .Cell(flexcpData, i, .ColIndex("指导零售价")) = Val(zlStr.Nvl(rsTemp!指导零售价))
            .Cell(flexcpData, i, .ColIndex("原指导售价")) = Val(zlStr.Nvl(rsTemp!指导零售价))
            .Cell(flexcpData, i, .ColIndex("现指导售价")) = Val(zlStr.Nvl(rsTemp!指导零售价))
            
            Call zlGetPrice(Val(zlStr.Nvl(rsTemp!Id)), IIf(.TextMatrix(i, .ColIndex("类型")) = "时价", True, False), True, i)
            Call LoadStockData(Val(zlStr.Nvl(rsTemp!Id)), Val(.Cell(flexcpData, i, .ColIndex("原价"))), Val(.Cell(flexcpData, i, .ColIndex("现价"))))
            i = i + 1
            rsTemp.MoveNext
        Loop
        
        '计算应付变动情况
        If (m调价方式 = T_成本价调价 Or m调价方式 = T_成本和售价调价) Then
            Call RefreshPayData
        End If
        mlngPreRow = 0:
        Call vsPrice_RowColChange
        .Redraw = flexRDBuffered
     End With
     Get调价项目 = True
     Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub InitOther()
    '------------------------------------------------------------
    '功能:初始化批量调整额的相关信息
    '------------------------------------------------------------
    With cbo调整方式
        .AddItem "根据成本价按加成调价"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "根据售价按比例调价"
        .ItemData(.NewIndex) = 2
        .AddItem "根据售价按固定金额调价"
        .ItemData(.NewIndex) = 3
    End With
End Sub

Public Function ShowBill(ByVal frmMain As Form, ByVal EditType As BillType, ByVal lngBillId As Long, ByVal lng材料ID As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能:显示调价单的入口
    '参数:frmMain-调用的父窗口
    '     lngBillID-单据ID
    '     lng材料ID-材料ID
    '返回:调价成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/18
    '--------------------------------------------------------------------------------------------------------------
    mlngBillId = lngBillId: mlngStuffId = lng材料ID: mBillType = EditType
    
    Me.Show 1, frmMain
    ShowBill = mblnSucces
End Function
  
Private Sub cbo调整方式_Click()
    If cbo调整方式.ListIndex < 0 Then Exit Sub
    Select Case cbo调整方式.ItemData(cbo调整方式.ListIndex)
    Case 1
        lblInfor.Caption = "根据成本价，输入新的加成率重新加成调价"
        lbl调整.Caption = "％"
        txt调整额.MaxLength = 3
    Case 2
        lblInfor.Caption = "在当前售价基础上按照比例调价"
        lbl调整.Caption = "％"
        txt调整额.MaxLength = 3
    Case 3
        lblInfor.Caption = "在当前售价基础上按固定金额加减调价"
        lbl调整.Caption = "元"
        txt调整额.MaxLength = 10
    End Select
End Sub

Private Sub cbo调整方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim ctrCombox As CommandBarComboBox
    
    Select Case Control.Id
    Case conMenu_Preview
        Call printbill(2)
    Case conMenu_Print
        Call printbill(1)
    Case conMenu_Save   '保存
         '检测相关输入合法性
        If ISValied = False Then Exit Sub
        '如果即时执行，则调用过程zl_材料收发记录_Adjust
        If SaveData() = False Then Exit Sub
        mblnModify = False
        mblnSucces = True
        Unload Me
        Exit Sub
    Case conMenu_Cancel  '取消
        mlngBillId = 0
        mlngStuffId = 0
        Unload Me
    Case conMenu_Combo  '计价方式选择
        Set ctrCombox = Control
        Select Case ctrCombox.ItemData(ctrCombox.ListIndex)
        Case 1      '售价调价
            m调价方式 = T_售价调价
        Case 2      '成本价调价
            m调价方式 = T_成本价调价
        Case 3      '售价与成本价调价
            m调价方式 = T_成本和售价调价
        Case Else   '如果不可预知的话，以售价为准
             m调价方式 = T_售价调价
        End Select
        Call SetControlVisble
        Call SetColor(m调价方式)
        Call picStoce_Resize
    Case conMenu_Help_Help  '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
    End Select
End Sub

Private Sub SetColor(ByVal int方式 As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    
    With vsPrice
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = &H8000000F '灰色
        For intRow = 1 To .Rows - 1
            If int方式 = 2 Then
                .Cell(flexcpBackColor, 1, .ColIndex("品名"), .Rows - 1, .ColIndex("品名")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现采购限价"), .Rows - 1, .ColIndex("现采购限价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现指导售价"), .Rows - 1, .ColIndex("现指导售价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现成本价"), .Rows - 1, .ColIndex("现成本价")) = &H80000005 ' 白色
            ElseIf int方式 = 3 Then
                .Cell(flexcpBackColor, 1, .ColIndex("品名"), .Rows - 1, .ColIndex("品名")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现价"), .Rows - 1, .ColIndex("现价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("收入名称"), .Rows - 1, .ColIndex("收入名称")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现采购限价"), .Rows - 1, .ColIndex("现采购限价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现指导售价"), .Rows - 1, .ColIndex("现指导售价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现成本价"), .Rows - 1, .ColIndex("现成本价")) = &H80000005 ' 白色
            Else    '其他方式都以售价方式进行
                .Cell(flexcpBackColor, 1, .ColIndex("品名"), .Rows - 1, .ColIndex("品名")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现价"), .Rows - 1, .ColIndex("现价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("收入名称"), .Rows - 1, .ColIndex("收入名称")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现采购限价"), .Rows - 1, .ColIndex("现采购限价")) = &H80000005 ' 白色
                .Cell(flexcpBackColor, 1, .ColIndex("现指导售价"), .Rows - 1, .ColIndex("现指导售价")) = &H80000005 ' 白色
            End If
        Next
    End With
End Sub
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnData As Boolean, i As Long

    Select Case Control.Id
    Case conMenu_Preview, conMenu_Print
        With vsPrice
            blnData = False
            For i = 0 To .Rows - 1
                If Val(.Cell(flexcpData, i, .ColIndex("品名"))) <> 0 Then
                    blnData = True
                    Exit For
                End If
            Next
        End With
        Control.Enabled = blnData
    Case conMenu_Save   '保存
        
        With vsPrice
            blnData = False
            For i = 0 To .Rows - 1
                If Val(.Cell(flexcpData, i, .ColIndex("品名"))) <> 0 Then
                    blnData = True
                    Exit For
                End If
            Next
        End With
        Control.Enabled = blnData
        If mBillType = B_查阅 Then Control.Visible = False
    Case conMenu_Cancel  '取消
    End Select
End Sub

Private Sub chkAppAllColumn_Click()
    If chkAppAllColumn.Value = 1 Then
        chk批次.Enabled = False
        chk批次.Value = 0
        chk显示所有材料.Value = 1
    Else
        chk批次.Enabled = True
    End If
End Sub

Private Sub chk显示所有材料_Click()
    Dim i As Long
    With vsStoce
        For i = 1 To .Rows - 1
            .RowHidden(i) = IIf(chk显示所有材料.Value = 1, False, True)
        Next
    End With
    mlngPreRow = 0
    Call vsPrice_RowColChange
End Sub

Private Sub chk自动计算_Click()
    '计算应付变动情况
    If (m调价方式 = T_成本价调价 Or m调价方式 = T_成本和售价调价) And chk应付.Value = 1 And chk自动计算.Value = 1 Then
        Call RefreshPayData
    End If
End Sub

Private Sub cmdPriver_Click()
    If Select供应商(Me, txtPriver, "") = False Then Exit Sub
End Sub

Private Sub cmdType_Click()
   Call Select诊疗分类("")
   If txt分类.Enabled Then txt分类.SetFocus
End Sub

Private Sub cmd调价_Click()
    If Get调价项目 = False Then Exit Sub
End Sub

Private Sub DkPane_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Or Action = PaneActionExpanding Then
        If Pane.Id = ID_PANE_SEARCH And Pane.Hidden = False Then
            Cancel = True
        End If
    ElseIf Action = PaneActionPinning Or Action = PaneActionCollapsing Then
    Else
        Cancel = True
    End If
End Sub

Private Sub opt应用_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub picBakDown_Resize()
    err = 0: On Error Resume Next
    Me.txt调价人.Left = picBakDown.ScaleWidth - Me.txt调价人.Width
    Me.lblValuer.Left = txt调价人.Left - lblValuer.Width - 50
    Me.txt说明.Width = lblValuer.Left - txt说明.Left - 300
    Me.dtp执行日期.Left = picBakDown.ScaleWidth - Me.dtp执行日期.Width
    Me.lblRunDate.Left = dtp执行日期.Left - lblRunDate.Width - 50
End Sub
 
Private Sub picPrice_Resize()
    err = 0: On Error Resume Next
    With vsPrice
        .Left = picPrice.ScaleLeft
        .Top = picPrice.ScaleTop + chkAppAllColumn.Height
         picBakDown.Top = picPrice.ScaleHeight - picBakDown.Height
        picBakDown.Left = .Left
        picBakDown.Width = picPrice.ScaleWidth
        .Height = picBakDown.Top - .Top
        .Width = picPrice.ScaleWidth
    End With
End Sub

Private Sub picSeach_Resize()
    err = 0: On Error Resume Next
    With cmdType
        .Left = picSeach.ScaleWidth - .Width - 50
        txt分类.Width = .Left - txt分类.Left
    End With
    With fra
        .Width = picSeach.ScaleWidth - .Left - 50
    End With
    With fra调整额
        .Width = picSeach.ScaleWidth - .Left - 50
        
    End With
    With fraCost
        .Width = picSeach.ScaleWidth - .Left - 50
        cmdPriver.Left = .Width - cmdPriver.Width - 100
        txtPriver.Width = cmdPriver.Left - txtPriver.Left
    End With
    cmd调价.Left = picSeach.ScaleWidth - cmd调价.Width - 50
    
End Sub

Private Sub picStoceBack_Resize()
    err = 0: On Error Resume Next
    With tbPage
        .Left = picStoceBack.ScaleLeft
        .Width = picStoceBack.ScaleWidth
        .Top = picStoceBack.ScaleTop
        .Height = picStoceBack.ScaleHeight
        
    End With
End Sub


Private Sub txtPriver_Change()
    txtPriver.Tag = ""
End Sub

Private Sub txtPriver_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtPriver
End Sub

Private Sub txtPriver_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtPriver.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If txtPriver.Tag = "" And Trim(txtPriver.Text) = "" Then OS.PressKey vbKeyTab: Exit Sub
    If Select供应商(Me, txtPriver, Trim(txtPriver.Text)) = False Then Exit Sub
End Sub

Private Sub txt调整额_KeyPress(KeyAscii As Integer)
    If cbo调整方式.ItemData(cbo调整方式.ListIndex) = 3 Then
        Call zlControl.TxtCheckKeyPress(txt调整额, KeyAscii, m负金额式)
    Else
        Call zlControl.TxtCheckKeyPress(txt调整额, KeyAscii, m金额式)
    End If
End Sub

Private Sub txt分类_Change()
    txt分类.Tag = ""
End Sub
Private Sub txt分类_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txt分类.Tag) <> "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If Trim(txt分类.Text) = "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    
    If Select诊疗分类(Trim(txt分类.Text)) = False Then
        Exit Sub
    End If
    OS.PressKey vbKeyTab
End Sub
Private Sub FullStoce现价(ByVal lng材料ID As Long, ByVal dbl现价 As Double)
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据现价,填充库存变动的现价及调整额
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-07 10:32:13
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl调整额 As Double
    With vsStoce
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息"))) = lng材料ID Then
                .Cell(flexcpData, lngRow, .ColIndex("现价")) = dbl现价
                .TextMatrix(lngRow, .ColIndex("现价")) = Format(dbl现价 * Val(.Cell(flexcpData, lngRow, .ColIndex("单位"))), mFMT.FM_零售价)
                '调整额=数量*(现价-原价)
                dbl调整额 = (dbl现价 - Val(.Cell(flexcpData, lngRow, .ColIndex("原价")))) * Val(.Cell(flexcpData, lngRow, .ColIndex("数量")))
                .TextMatrix(lngRow, .ColIndex("调整额")) = Format(dbl调整额, mFMT.FM_金额)
                .Cell(flexcpData, lngRow, .ColIndex("调整额")) = dbl调整额
                '需要根据加成率重新计算调整的成本价
                 Call AutoCalcStoce(lngRow, .ColIndex("现价"))
            End If
        Next
    End With
End Sub

Private Sub FullStoce成本价(ByVal lng材料ID, ByVal dbl成本价 As Double)
    '成本价
    Dim lngRow As Long, dbl调整额 As Double
    With vsStoce
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息"))) = lng材料ID Then
                .Cell(flexcpData, lngRow, .ColIndex("现成本价")) = dbl成本价
                .TextMatrix(lngRow, .ColIndex("现成本价")) = dbl成本价
                 Call AutoCalcStoce(lngRow, .ColIndex("现成本价"))
            End If
        Next
    End With
End Sub

Private Sub txt加成率_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txt加成率
    
End Sub

Private Sub txt加成率_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
    
End Sub

Private Sub txt加成率_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt加成率, KeyAscii, m金额式)
    
End Sub

Private Sub vsPay_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    zl_VsGridRowChange vsPay, OldRow, NewRow, OldCol, NewCol
    
End Sub

Private Sub vsPay_GotFocus()
'    zl_VsGridGotFocus vsPay
    
End Sub

Private Sub vsPay_LostFocus()
'    zl_VsGridLOSTFOCUS vsPay
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '设置相关的格式
    '刘兴宏:2007/09/17
    '--------------------------------------------------------------------------------
    Dim lngRow As Long, dbl现价 As Double
    
    With vsPrice
        Select Case Col
        Case .ColIndex("原价"), .ColIndex("现指导售价")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_零售价)
        Case .ColIndex("现价")
            If chkAppAllColumn.Value = 0 Then
                '要换算成最小单位
                dbl现价 = Val(.TextMatrix(Row, Col)) / Val(.TextMatrix(Row, .ColIndex("系数")))
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_零售价)
                Call FullStoce现价(Val(.Cell(flexcpData, Row, .ColIndex("品名"))), dbl现价)
            Else
                Call AutoCalc所有库存价格
            End If
        Case .ColIndex("现指导批价")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_成本价)
        Case .ColIndex("品名")
            .ColComboList(Col) = "..."
        Case .ColIndex("收入名称")
            .ColComboList(Col) = "..."
        Case .ColIndex("现成本价")
            '成本价调整和按成本价售价一起调整时可以修改成本价
            If chkAppAllColumn.Value = 1 Then
                Call AutoCalc所有库存价格
            Else
                Call FullStoce成本价(Val(.Cell(flexcpData, Row, .ColIndex("品名"))), Val(.TextMatrix(.Row, .Col)))
            End If
        End Select
    End With
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    zl_VsGridRowChange vsPrice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mBillType = B_查阅 Then Cancel = True: Exit Sub
    mlngPrice = Val(vsPrice.TextMatrix(vsPrice.Row, vsPrice.Col))
    With vsPrice
        Select Case Col
        Case .ColIndex("品名"), .ColIndex("现采购限价"), .ColIndex("现指导售价") '
            .FocusRect = flexFocusSolid
            .HighLight = flexHighlightNever
            If Val(.Cell(flexcpData, Row, .ColIndex("品名"))) = 0 And (Col = .ColIndex("现采购限价") Or Col = .ColIndex("现指导售价")) Then Cancel = True
        Case .ColIndex("现价"), .ColIndex("收入名称")
            If m调价方式 = T_成本价调价 Then
                .FocusRect = flexFocusHeavy
                Cancel = True
                Exit Sub
            Else
                .FocusRect = flexFocusSolid
                .HighLight = flexHighlightNever
                If Val(.Cell(flexcpData, Row, .ColIndex("品名"))) = 0 Then Cancel = True
            End If
        Case .ColIndex("现成本价")
            If m调价方式 = T_成本和售价调价 Or m调价方式 = T_成本价调价 Then
                .FocusRect = flexFocusSolid
                .HighLight = flexHighlightNever
                If Val(.Cell(flexcpData, Row, .ColIndex("品名"))) = 0 Then Cancel = True
            Else
                .FocusRect = flexFocusHeavy
                Cancel = True
            End If
        Case Else
            .FocusRect = flexFocusHeavy
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '功能:按钮选择
    '参数:
    '
    '--------------------------------------------------------------------------
    With vsPrice
        Select Case Col
        Case .ColIndex("品名")
            If SelectStuff("") = False Then Exit Sub
        Case .ColIndex("收入名称")
            If Select收入项目("") = False Then Exit Sub
        Case Else
        End Select
    End With
End Sub

Private Sub vsPrice_ChangeEdit()
    mblnModify = True
End Sub
Private Sub vsPrice_EnterCell()
    If mBillType = B_查阅 Then Exit Sub
    
    With vsPrice
        Select Case .Col
        Case .ColIndex("品名")
             .ColComboList(.Col) = "..."
        Case .ColIndex("收入名称")
            .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsPrice_GotFocus()
'    zl_VsGridGotFocus vsPrice
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, lngRow As Long
    Dim i As Integer
    With vsPrice
        If (.Col = .ColIndex("品名") Or .Col = .ColIndex("收入名称")) And KeyCode <> vbKeyReturn Then
            vsPrice.ColComboList(.Col) = ""
        End If
        
        If KeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行的调价项目吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
           Call MoveStockData(Val(.Cell(flexcpData, .Row, .ColIndex("品名"))))
            If .Row = .Rows - 1 And .Row = 1 Then
                .Clear 1
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                Call InitControl
            Else
                .RemoveItem .Row
                Call RefreshPayData
            End If
        End If
        
        For i = 1 To vsStoce.Rows - 1
            If Val(vsStoce.Cell(flexcpData, i, vsStoce.ColIndex("卫材信息"))) = Val(vsPrice.Cell(flexcpData, vsPrice.Row, vsPrice.ColIndex("品名"))) Then
                vsStoce.RowHidden(i) = False
            Else
                vsStoce.RowHidden(i) = True
            End If
        Next
        
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsPrice
        If Val(.Cell(flexcpData, vsPrice.Row, .ColIndex("品名"))) = 0 Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsPrice, , , mBillType <> B_查阅, lngRow)
    End With
End Sub

Private Sub vsPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPrice
        Select Case Col
        Case .ColIndex("品名")
        
            strKey = Trim(vsPrice.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If SelectStuff(strKey) = False Then Exit Sub
            vsPrice.EditText = vsPrice.TextMatrix(Row, Col)
        Case .ColIndex("收入名称")
            strKey = Trim(vsPrice.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If strKey = "" Then Exit Sub
            If Select收入项目(strKey) = False Then
                vsPrice.TextMatrix(Row, Col) = vsPrice.EditText
                vsPrice.Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            vsPrice.EditText = vsPrice.TextMatrix(Row, Col)
        Case Else
            Call zlVsMoveGridCell(vsStoce, , , False)
        End Select
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    With vsPrice
        Select Case Col
        Case .ColIndex("品名")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("现价")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m金额式)
        Case .ColIndex("现指导批价")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m金额式)
        Case .ColIndex("现指导零价")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m金额式)
        Case .ColIndex("收入名称")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m文本式)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPrice_LostFocus()
'    zl_VsGridLOSTFOCUS vsPrice
End Sub

Private Sub vsPrice_RowColChange()
    '找到指定的卫生材料
    Dim lng材料ID As Long
    With vsPrice
'        .FocusRect = IIf(.Editable = flexEDKbdMouse, flexFocusHeavy, flexFocusSolid)
        If mlngPreRow = .Row Then Exit Sub
        mlngPreRow = .Row
        lng材料ID = Val(.Cell(flexcpData, .Row, .ColIndex("品名")))
        If lng材料ID = 0 Then Exit Sub
        Call Find材料(lng材料ID)
    End With
End Sub
Private Sub Find材料(ByVal lng材料ID As Long, Optional FindNext As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '功能:查找指定的材料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-12-08 15:18:19
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim BlnFind As Boolean
    BlnFind = False
    With vsStoce
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("卫材信息"))) <> lng材料ID Then
                If chk显示所有材料.Value = 0 Then
                    .RowHidden(i) = True
                Else
                    .RowHidden(i) = False
                End If
            Else
                If chk显示所有材料.Value = 1 Then
                    .RowHidden(i) = False
                    .Row = i
                    .TopRow = .Row
                    Exit Sub
                Else
                    .RowHidden(i) = False
                    If Not BlnFind Then
                    .Row = i
                    BlnFind = True
                    End If
                End If
            End If
        Next
    End With
    
End Sub


Private Sub vsPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    Dim intRow As Integer
    If mBillType = B_查阅 Then Cancel = True: Exit Sub
    
    strKey = Trim(vsPrice.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    With vsPrice
        Select Case Col
        Case .ColIndex("品名")
        Case .ColIndex("收入名称") '
        Case .ColIndex("现价")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "现价") = False Then Cancel = True: Exit Sub
                If Val(.Cell(flexcpData, .Row, .ColIndex("品名"))) = 0 Then
                    vsPrice.EditText = Format(Val(strKey), mFMT.FM_零售价)
                    Exit Sub
                End If
                If Val(strKey) > Val(.TextMatrix(.Row, .ColIndex("现指导售价"))) And Val(.TextMatrix(.Row, .ColIndex("现指导售价"))) <> 0 Then
                    MsgBox "现价不能大于指导零售价！（" & Format(Val(.TextMatrix(.Row, .ColIndex("现指导售价"))), mFMT.FM_零售价) & "）", vbQuestion + vbDefaultButton1, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_零售价)
                mblnModify = True
            End If
            If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                For intRow = 1 To .Rows - 1
                    .TextMatrix(intRow, .ColIndex("现价")) = vsPrice.EditText
                Next
            End If
        Case .ColIndex("现指导批价")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "现指导批价") = False Then Cancel = True: Exit Sub
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_成本价)
            End If
        Case .ColIndex("现指导售价")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "现指导售价") = False Then Cancel = True: Exit Sub
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_零售价)
                
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        .TextMatrix(intRow, .ColIndex("现指导售价")) = vsPrice.EditText
                    Next
                End If
            End If
        Case .ColIndex("现采购限价")
            If strKey <> "" Then
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_成本价)
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        .TextMatrix(intRow, .ColIndex("现采购限价")) = vsPrice.EditText
                    Next
                End If
            End If
        Case .ColIndex("现成本价")
            If strKey <> "" Then
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_成本价)
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        .TextMatrix(intRow, .ColIndex("现成本价")) = vsPrice.EditText
                    Next
                End If
            End If
        End Select
    End With
End Sub
Private Function Select诊疗分类(ByVal strSeach As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择指定的卫生材料
    '参数:strKey-多选择的条件
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    
    Set objCtl = txt分类
    vRect = zlControl.GetControlRect(txt分类.hwnd)
    lngH = txt分类.Height
    strKey = GetMatchingSting(strSeach)
      
    strTittle = "卫生材料分类选择"
    If strSeach = "" Then
'        gstrSQL = "" & _
'                "   Select ID,上级ID, 编码,名称,简码 From 诊疗分类目录 a " & _
'                "   where  类型=7 start with 上级id is null connect by prior id=上级id"
        
        gstrSQL = "Select ID, 上级id, 编码, 名称, 类别" & _
                " From (Select ID, 上级id, 编码, 名称, '分类' 类别" & _
                       " From 诊疗分类目录" & _
                       " Where 类型 = 7" & _
                       " Start With 上级id Is Null" & _
                       " Connect By Prior ID = 上级id" & _
                       " Union All" & _
                       " Select a.Id, a.分类id As 上级id, a.编码, a.名称, '品种' 类别" & _
                       " From 诊疗项目目录 A," & _
                       "     (Select ID From 诊疗分类目录 Where 类型 = 7 Start With 上级id Is Null Connect By Prior ID = 上级id) B" & _
                       " Where a.类别 = '4' And a.分类id = b.Id)" & _
                " Start With 上级id Is Null" & _
                " Connect By Prior ID = 上级id"
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False)
    Else
        If opt应用(2).Value = False Then
            gstrSQL = "" & _
                    "   Select ID,上级ID, 编码,名称,简码,'分类' 类别 From 诊疗分类目录 a " & _
                    "   Where (名称 like [1] or  编码  like [1] or  简码  like  [1]) and 类型=7  " & _
                    "   order by 编码"
        Else
            gstrSQL = "select a.分类id,a.id,a.编码,a.名称,'品种' 类别 from 诊疗项目目录 a,诊疗项目别名 b " & _
            " where a.类别 ='4' and a.id=b.诊疗项目id and (a.名称 like [1] or a.编码 like [1] OR b.简码 like [1])"
        End If
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    End If
    
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "没有满足条件的材料分类,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    objCtl.Text = zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!名称)
    objCtl.Tag = zlStr.Nvl(rsTemp!Id)
    If InStr(1, rsTemp!类别, "分类") > 0 Then '分类
        opt应用(0).Enabled = True
        opt应用(1).Enabled = True
        opt应用(2).Enabled = False
        opt应用(2).Value = False
    Else '品种
        opt应用(0).Enabled = False
        opt应用(1).Enabled = False
        opt应用(2).Enabled = True
        opt应用(2).Value = True
    End If
    
    Select诊疗分类 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SelectStuff(ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择指定的卫生材料
    '参数:strKey-多选择的条件
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim int系数 As Integer
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    err = 0: On Error GoTo ErrHand:
    Call CalcPosition(sngX, sngY, vsPrice)
              
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位,P.换算系数,P.包装单位," & _
            "         decode(I.是否变价,1,'时价','定价') 类型," & _
            "         P.成本价 as 成本价ID,P.指导批发价 as 指导批发价ID,P.指导零售价 as 指导零售价ID," & _
            "         to_char(p.成本价," & mOraFMT.FM_成本价 & ") as 成本价," & _
            "         to_char(p.指导批发价," & mOraFMT.FM_成本价 & ") 指导批发价," & _
            "         to_char(p.指导零售价," & mOraFMT.FM_零售价 & ") 指导零售价," & _
            "          P.跟踪在用" & _
            "   From 收费项目目录 I,收费项目别名 N,材料特性 P" & _
            "   Where I.ID=N.收费细目ID and I.类别='4' And I.ID=P.材料ID " & _
            "       and (I.编码 like [1] or N.简码 Like [1] or N.名称 Like [1])" & _
            "       and (I.撤档时间 Is Null Or I.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
     Else
        gstrSQL = "" & _
            "   Select  I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位,P.换算系数,P.包装单位, " & _
            "           decode(I.是否变价,1,'时价','定价') 类型," & _
            "           P.成本价 as 成本价ID,P.指导批发价 as 指导批发价ID,P.指导零售价 as 指导零售价ID," & _
            "           to_char(p.成本价," & mOraFMT.FM_成本价 & ") as 成本价," & _
            "           to_char(p.指导批发价," & mOraFMT.FM_成本价 & ") 指导批发价," & _
            "           to_char(p.指导零售价," & mOraFMT.FM_零售价 & ") 指导零售价," & _
            "           P.跟踪在用" & _
            "   From 收费项目目录 I,材料特性 P" & _
            "   Where I.类别='4' And I.ID=P.材料ID" & _
            "           and (I.撤档时间 Is Null Or I.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
            
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "卫生材料选择", False, "", "", False, False, True, sngX, sngY - vsPrice.CellHeight, vsPrice.CellHeight, blnCancel, False, False, strKey)
    If blnCancel = True Then Exit Function
    
    If rsTemp Is Nothing Then
        ShowMsgBox "不存在指定的卫生材料,请检查!"
        Exit Function
    End If
    
    With Me.vsPrice
        '检查是否选择了同一个品种的卫生材料
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("品名"))) <> 0 Then
                If Val(.Cell(flexcpData, i, .ColIndex("品名"))) = Val(zlStr.Nvl(rsTemp!Id)) And i <> .Row Then
                    ShowMsgBox "该卫生材料已经正在，不能进行调价！"
                    Exit Function
                End If
            End If
        Next
        
        '检查是否改变了原来已经存在的卫生材料
        If Val(.Cell(flexcpData, .Row, .ColIndex("品名"))) <> Val(zlStr.Nvl(rsTemp!Id)) And Val(.Cell(flexcpData, .Row, .ColIndex("品名"))) <> 0 Then
            '需要移除该卫生材料的库房变动情况，后才能更新
             Call MoveStockData(Val(.Cell(flexcpData, .Row, .ColIndex("品名"))))
        End If
        
        .Redraw = flexRDNone
        .TextMatrix(.Row, .ColIndex("品名")) = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!名称)
        .Cell(flexcpData, .Row, .ColIndex("品名")) = zlStr.Nvl(rsTemp!Id)
        .TextMatrix(.Row, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
        .TextMatrix(.Row, .ColIndex("产地")) = zlStr.Nvl(rsTemp!产地)
        .TextMatrix(.Row, .ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!计算单位), zlStr.Nvl(rsTemp!包装单位))
        .TextMatrix(.Row, .ColIndex("类型")) = zlStr.Nvl(rsTemp!类型)
        .Cell(flexcpData, .Row, .ColIndex("类型")) = zlStr.Nvl(rsTemp!跟踪在用)
        
        int系数 = IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!换算系数))
        .TextMatrix(.Row, .ColIndex("系数")) = int系数
        
        
        .TextMatrix(.Row, .ColIndex("现成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价ID)) * int系数, mFMT.FM_成本价)
        .Cell(flexcpData, .Row, .ColIndex("现成本价")) = Val(zlStr.Nvl(rsTemp!成本价ID))
        
        .TextMatrix(.Row, .ColIndex("原采购限价")) = Format(Val(zlStr.Nvl(rsTemp!指导批发价ID)) * int系数, mFMT.FM_成本价)
        .TextMatrix(.Row, .ColIndex("现采购限价")) = .TextMatrix(.Row, .ColIndex("原采购限价"))
        .Cell(flexcpData, .Row, .ColIndex("原采购限价")) = Val(zlStr.Nvl(rsTemp!指导批发价ID))
        .Cell(flexcpData, .Row, .ColIndex("现采购限价")) = Val(zlStr.Nvl(rsTemp!指导批发价ID))
        
        
        .TextMatrix(.Row, .ColIndex("指导零售价")) = Format(Val(zlStr.Nvl(rsTemp!指导零售价ID)) * int系数, mFMT.FM_零售价)
        .TextMatrix(.Row, .ColIndex("原指导售价")) = .TextMatrix(.Row, .ColIndex("指导零售价"))
        .TextMatrix(.Row, .ColIndex("现指导售价")) = .TextMatrix(.Row, .ColIndex("指导零售价"))
        
        .Cell(flexcpData, .Row, .ColIndex("指导零售价")) = Val(zlStr.Nvl(rsTemp!指导零售价ID))
        .Cell(flexcpData, .Row, .ColIndex("原指导售价")) = Val(zlStr.Nvl(rsTemp!指导零售价ID))
        .Cell(flexcpData, .Row, .ColIndex("现指导售价")) = Val(zlStr.Nvl(rsTemp!指导零售价ID))
        Call zlGetPrice(Val(zlStr.Nvl(rsTemp!Id)), IIf(.TextMatrix(.Row, .ColIndex("类型")) = "时价", True, False))
        Call LoadStockData(Val(zlStr.Nvl(rsTemp!Id)), Val(.Cell(flexcpData, .Row, .ColIndex("原价"))), Val(.Cell(flexcpData, .Row, .ColIndex("现价"))))
        
        .Col = .ColIndex("现价")
        .Redraw = flexRDBuffered
        zlControl.ControlSetFocus vsPrice, True
        mlngPreRow = 0:
        Call vsPrice_RowColChange
    End With
    SelectStuff = True
    Exit Function
ErrHand:
    vsPrice.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlGetPrice(lng材料ID As Long, bln实价 As Boolean, Optional bln批量 As Boolean = False, Optional lngRow As Long = -1)
    '----------------------------------------------------
    '功能：填写指定卫材id的对应价格信息
    '入参：lng材料ID-材料ID
    '      bln实价:是否时价卫材
    '      bln批量-False不根据条件算现价,true-根据条件计算现价
    '编制:刘兴宏
    '日期:2007/09/17
    '----------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim bytType As Byte
    Dim dbl比率 As Double
    
    On Error GoTo ErrHandle
    If bln批量 Then
        bytType = cbo调整方式.ItemData(cbo调整方式.ListIndex)
    End If
    
    If bln实价 Then
        Me.Chk定价.Enabled = True
        '表示时价卫材调价，取库存金额/库存数量做为其价格
        gstrSQL = "" & _
            "   Select  P.id,Decode(Nvl(K.库存数量,0),0,P.现价,K.库存金额/Nvl(K.库存数量,1)) 现价," & _
            "           P.执行日期,P.收入项目id,I.名称 as 收入名称, " & IIf(mintUnit = 0, "1", " Nvl(M.换算系数,1)") & " as  系数,nvl(m.成本价,0) as 成本价,m.跟踪在用" & _
            "   From 收费价目 P,收入项目 I,材料特性 M," & _
            "       (   Select Sum(实际金额) 库存金额,Sum(实际数量) 库存数量" & _
            "           From 药品库存 " & _
            "           Where  性质=1 and 药品ID=[1] " & _
            "        ) K" & _
            " where p.收费细目id=M.材料id and P.收入项目id=I.id and P.收费细目id=[1] " & _
            "       and (P.终止日期 is null or P.终止日期=to_date('3000-01-01','YYYY-MM-DD'))" & _
            GetPriceClassString("P")
    Else
        '非时价卫材调价，取其价格记录中的价格
        gstrSQL = "" & _
            "   Select P.id,P.现价,P.执行日期,P.收入项目id,I.名称 as 收入名称," & IIf(mintUnit = 0, "1", " Nvl(M.换算系数,1)") & " as  系数,nvl(m.成本价,0) as 成本价,m.跟踪在用" & _
            "   From 收费价目 P,收入项目 I,材料特性 M" & _
            "   Where p.收费细目id=M.材料id and P.收入项目id=I.id and P.收费细目id=[1]  " & _
            "           and (P.终止日期 is null or P.终止日期=to_date('3000-01-01','YYYY-MM-DD'))" & _
            GetPriceClassString("P")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
    With vsPrice
        If lngRow < 0 Then lngRow = .Row
        If rsTemp.RecordCount > 0 Then
            .RowData(lngRow) = Val(zlStr.Nvl(rsTemp!Id))
            .Cell(flexcpData, lngRow, .ColIndex("类型")) = zlStr.Nvl(rsTemp!跟踪在用)
            
            .TextMatrix(lngRow, .ColIndex("上次日期")) = Format(rsTemp!执行日期, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("原价")) = Format(Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!系数)), mFMT.FM_零售价)
            .Cell(flexcpData, lngRow, .ColIndex("原价")) = Val(zlStr.Nvl(rsTemp!现价))
            If bln批量 = False Then
                .TextMatrix(lngRow, .ColIndex("现价")) = Format(Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!系数)), mFMT.FM_零售价)
                .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(zlStr.Nvl(rsTemp!现价))
            Else
                If Val(txt调整额.Text) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("现价")) = Format(Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!系数)), mFMT.FM_零售价)
                    .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!系数))
                Else
                    Select Case bytType
                    Case 1      '根据成本价加成
                        dbl比率 = 1 + Val(txt调整额.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("现价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * dbl比率 * Val(zlStr.Nvl(rsTemp!系数)), mFMT.FM_零售价)
                        .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(zlStr.Nvl(rsTemp!成本价)) * dbl比率
                    Case 2      '根据零售价按比例
                        dbl比率 = 1 + Val(txt调整额.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("现价")) = Format(Val(zlStr.Nvl(rsTemp!现价)) * dbl比率 * Val(zlStr.Nvl(rsTemp!系数)), mFMT.FM_零售价)
                        .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(zlStr.Nvl(rsTemp!现价)) * dbl比率
                    Case 3      '根据零售价按固定金额加减
                        dbl比率 = Val(txt调整额.Text)
                        .TextMatrix(lngRow, .ColIndex("现价")) = Format((Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!系数))) + dbl比率, mFMT.FM_零售价)
                        .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(zlStr.Nvl(rsTemp!现价)) + dbl比率
                    End Select
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("现价"))) > Val(.TextMatrix(lngRow, .ColIndex("指导零售价"))) And Val(.TextMatrix(lngRow, .ColIndex("指导零售价"))) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("现价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("指导零售价"))), mFMT.FM_零售价)
                    .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(.Cell(flexcpData, lngRow, .ColIndex("指导零售价")))
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("现成本价")) = Format(rsTemp!成本价 * Val(zlStr.Nvl(rsTemp!系数)), mFMT.FM_成本价)
            .Cell(flexcpData, lngRow, .ColIndex("现成本价")) = Val(zlStr.Nvl(rsTemp!成本价))
            
            .TextMatrix(lngRow, .ColIndex("原收入id")) = Val(zlStr.Nvl(rsTemp!收入项目id))
            .TextMatrix(lngRow, .ColIndex("收入名称")) = zlStr.Nvl(rsTemp!收入名称)
            .Cell(flexcpData, lngRow, .ColIndex("收入名称")) = Val(zlStr.Nvl(rsTemp!收入项目id))
        Else
            .RowData(lngRow) = -1
            .TextMatrix(lngRow, .ColIndex("上次日期")) = ""
            .TextMatrix(lngRow, .ColIndex("原价")) = Format(0, mFMT.FM_零售价)
            .TextMatrix(lngRow, .ColIndex("现价")) = Format(0, mFMT.FM_零售价)
            .TextMatrix(lngRow, .ColIndex("现成本价")) = Format(0, mFMT.FM_成本价)
            .Cell(flexcpData, lngRow, .ColIndex("原价")) = 0
            .Cell(flexcpData, lngRow, .ColIndex("现价")) = 0
            .Cell(flexcpData, lngRow, .ColIndex("现成本价")) = 0
            If bln批量 Then
                '第一次批量:
                If Val(txt调整额.Text) = 0 Then
                    '如果没设置调整额,则为0
                    .TextMatrix(lngRow, .ColIndex("现价")) = Format(0, mFMT.FM_零售价)
                    .Cell(flexcpData, lngRow, .ColIndex("现价")) = 0
                Else
                    Select Case bytType
                    Case 1      '根据成本价加成
                        dbl比率 = 1 + Val(txt调整额.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("现价")) = Format(0, mFMT.FM_零售价)
                        .Cell(flexcpData, lngRow, .ColIndex("现价")) = 0 * dbl比率
                    Case 2      '根据零售价按比例
                        dbl比率 = 1 + Val(txt调整额.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("现价")) = Format(0 * dbl比率 * Val(.TextMatrix(lngRow, .ColIndex("系数"))), mFMT.FM_零售价)
                        .Cell(flexcpData, lngRow, .ColIndex("现价")) = 0 * dbl比率
                    Case 3      '根据零售价按固定金额加减
                        dbl比率 = Val(txt调整额.Text)
                        .TextMatrix(lngRow, .ColIndex("现价")) = Format(0 + dbl比率 * Val(.TextMatrix(lngRow, .ColIndex("系数"))), mFMT.FM_零售价)
                        .Cell(flexcpData, lngRow, .ColIndex("现价")) = 0 + dbl比率
                    End Select
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("现价"))) > Val(.TextMatrix(lngRow, .ColIndex("指导零售价"))) And Val(.TextMatrix(lngRow, .ColIndex("指导零售价"))) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("现价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("指导零售价"))), mFMT.FM_零售价)
                    .Cell(flexcpData, lngRow, .ColIndex("现价")) = Val(.Cell(flexcpData, lngRow, .ColIndex("指导零售价")))
                End If
            End If
            If lngRow > 1 Then
                .TextMatrix(lngRow, .ColIndex("原收入id")) = .TextMatrix(lngRow - 1, .ColIndex("原收入id"))
                .TextMatrix(lngRow, .ColIndex("收入名称")) = .TextMatrix(lngRow - 1, .ColIndex("收入名称"))
                .Cell(flexcpData, lngRow, .ColIndex("收入名称")) = .Cell(flexcpData, lngRow - 1, .ColIndex("收入名称"))
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Select收入项目(ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择指定的收入项目信息
    '参数:strKey-多选择的条件
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim sngX As Single, sngY As Single
    
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    err = 0: On Error GoTo ErrHand:
    Call CalcPosition(sngX, sngY, vsPrice)
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select id,编码,名称,简码,收据费目,病案费目" & _
            "   From 收入项目" & _
            "   Where (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) and 末级=1 " & _
            "         and (编码 like [1] or 简码 Like [1] or 名称 Like [1])"
     Else
        gstrSQL = "" & _
            "   Select id,编码,名称,简码,收据费目,病案费目" & _
            "   From 收入项目" & _
            "   Where (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) and 末级=1 "
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "收入项目选择器", False, "", "", False, False, True, sngX, sngY - vsPrice.CellHeight, vsPrice.CellHeight, blnCancel, False, False, strKey)
    If blnCancel = True Then Exit Function
    
    If rsTemp Is Nothing Then
        ShowMsgBox "不存在指定的收入项目,请检查!"
        Exit Function
    End If
    
    With Me.vsPrice
        .Redraw = flexRDNone
        .TextMatrix(.Row, .ColIndex("收入名称")) = zlStr.Nvl(rsTemp!名称)
        .Cell(flexcpData, .Row, .ColIndex("收入名称")) = zlStr.Nvl(rsTemp!Id)
        .Redraw = flexRDBuffered
    End With
    
    Select收入项目 = True
    Exit Function
ErrHand:
    vsPrice.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function
  
 

Private Sub chk立即执行_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, lng材料ID As Long
    
    Dim mlngStuffIdThis As Long, IntCheck As Integer
    
    On Error GoTo ErrHandle
    If chk立即执行.Value = 1 Then
        
        '循环判断所有材料
        With vsPrice
            For i = 1 To .Rows - 1
                lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("品名")))
                gstrSQL = "Select count(*) as 未执行 From 收费价目 where 变动原因=0 and  收费细目id=[1]" & _
                        GetPriceClassString("")
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
                If Not rsTemp.EOF Then
                    If Val(zlStr.Nvl(rsTemp!未执行)) <> 0 Then
                        MsgBox "卫生材料" & .TextMatrix(i, .ColIndex("品名")) & "存在未执行价格，不能设置为立即执行！", vbInformation, gstrSysName
                        chk立即执行.Value = 0
                        Exit Sub
                    End If
                End If
            Next
        End With
    End If
    If Me.chk立即执行.Value Then
        Me.dtp执行日期.Enabled = False
    Else
        Me.dtp执行日期.Enabled = True
    End If
    err = 0: On Error Resume Next
    Me.vsPrice.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
 
Private Function ISValied() As Boolean
    '-------------------------------------------------------------------------------------------
    '功能:检查输入的合法性
    '参数:
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/15
    '-------------------------------------------------------------------------------------------
    '检测各执行价格是否正确
    '以及收入项目相同的情况下现价是否与原价相同
    
    Dim i As Long, blnZero As Boolean, lng材料ID As Long
    Dim strOldID As String, strNewID As String, strTemp As String
    Dim blnHaving As Boolean
    
    ISValied = False
    
    strNewID = "": strOldID = ""
    With vsPrice
        blnZero = False
        For i = 1 To .Rows - 1
        
            lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("品名")))
            If lng材料ID <> 0 Then
                blnHaving = True
                If Not IsNumeric(Trim(.TextMatrix(i, .ColIndex("现价")))) Then
                    MsgBox "第" & i & "行的卫生材料现价中含有非法字符！", vbInformation, gstrSysName
                    Exit Function
                End If
                                
                If m调价方式 <> T_成本价调价 Then
                    '刘兴宏:主要是解决可以为零的情况,比如：疫苗.是免费的
                    '问题:9569 2006-11-20
                    If Val(.TextMatrix(i, .ColIndex("现价"))) = 0 And blnZero = False Then
                        If MsgBox("第" & i & "行的卫生材料现价为零了,是否继续?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
                            blnZero = True
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, .ColIndex("原收入ID"))) = Val(.Cell(flexcpData, i, .ColIndex("收入名称"))) And _
                   Val(.TextMatrix(i, .ColIndex("现价"))) = Val(.TextMatrix(i, .ColIndex("原价"))) Then
                   '需要检查相关的调价信息
                   If m调价方式 <> T_成本价调价 Then
                        '肯定需要进行成本价调价
                        'If .TextMatrix(i, .ColIndex("类型")) = "时价" And .Cell(flexcpData, i, .ColIndex("类型")) <> "1" Then
                        '    '是范围，非跟踪卫生材料的实价卫材
                        'Else
                            MsgBox "第" & i & "行的卫生材料现价与原价相同，不能执行调价！", vbInformation, gstrSysName
                            Exit Function
                        'End If
                   End If
                End If
                
                If m调价方式 <> T_成本价调价 Then
                    If .TextMatrix(i, .ColIndex("类型")) = "时价" And Me.chk立即执行.Value <> 1 Then
                        MsgBox "第" & i & "行为时价卫生材料，必须设置为立即执行！", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    If chk立即执行.Value = 0 Then
                        ShowMsgBox "为成本价调价时,必需立即执行,请检查!"
                        Exit Function
                    End If
                End If
                
                If .RowData(i) <> -1 Then
                    If InStr(1, strOldID & ",", "," & .RowData(i) & ",") > 0 Then
                        ShowMsgBox "在第" & i & "行中,不能对相同品种(" & .TextMatrix(i, .ColIndex("品名")) & ")重复调价"
                        .Row = i: .Col = .ColIndex("品名")
                        .SetFocus
                        Exit Function
                    End If
                    strOldID = strOldID & "," & .RowData(i)
                Else
                    If InStr(1, strNewID & ",", "," & lng材料ID & ",") > 0 Then
                        MsgBox "不能对相同品种(" & .TextMatrix(i, .ColIndex("品名")) & ")重复设置价格", vbExclamation, gstrSysName
                        .Row = i: .Col = .ColIndex("品名")
                        .SetFocus
                        Exit Function
                    End If
                    strNewID = strNewID & "," & lng材料ID
                End If
                
                If Val(.TextMatrix(i, .ColIndex("现价"))) > Val(.TextMatrix(i, .ColIndex("现指导售价"))) And Val(.TextMatrix(i, .ColIndex("现指导售价"))) <> 0 Then
                    ShowMsgBox "在第" & i & "行中,品种(" & .TextMatrix(i, .ColIndex("品名")) & ")的现价超过了指导零售价(" & Val(.TextMatrix(i, .ColIndex("指导零售价"))) & ")"
                    .Row = i: .Col = .ColIndex("现价")
                    .SetFocus
                    Exit Function
                End If
                If IsValied成本价(lng材料ID) = False Then
                    .Row = i: .Col = .ColIndex("现价")
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
        
        If blnHaving = False Then
            MsgBox "未设置调价项目,请检查!", vbInformation, gstrSysName
            .Row = 1: .Col = .ColIndex("品名")
            .SetFocus
            Exit Function
        End If
    End With
    If IsValied应付信息 = False Then Exit Function
    ISValied = True
End Function
Public Function IsValied成本价(ByVal lng材料ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查成本价调价是否合法
    '入参:
    '出参:
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-10 10:04:24
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, blnHaveData As Boolean
    Dim i As Long
    
    On Error GoTo ErrHandle

    '不存成本价调价，就直接返回了
    If m调价方式 = T_售价调价 Then IsValied成本价 = True: Exit Function
    
    '检查是否还有未执行的成本价调价计划
    gstrSQL = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
    If rsTemp.RecordCount = 0 Then
        '需要检查该材料是否存在未审核单据
        If zl存在未审核单据(lng材料ID) = True Then
            gstrSQL = "Select 名称 From 收费项目目录 where id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
            If rsTemp.EOF Then Exit Function
            If MsgBox(rsTemp!名称 & "存在未审核单据，调整成本价可能会造成差价误差。" & _
                vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        IsValied成本价 = True:
        Exit Function    '表示未存在执行价格，则可以直接退出(因为不管调价与不调价，都没什么问题)
    End If
    
    '看是否有相应的成本价调价
    With vsStoce
            blnHaveData = False
            For i = 1 To .Rows - 1
                '存在成本价调整，因此返回True
                If lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("卫材信息"))) Then
                    If Val(.TextMatrix(i, .ColIndex("差价调整额"))) <> 0 Then
                        
                        gstrSQL = "Select 名称 From 收费项目目录 where id=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
                        If rsTemp.EOF Then Exit Function
                        MsgBox "卫生材料“" & zlStr.Nvl(rsTemp!名称) & "”存在未执行成本价，不能设置本次调价！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    blnHaveData = True
                End If
            Next
    End With
    If blnHaveData Then
        '存在该材料,还需要检查该材料是否存在未审核单据
        If zl存在未审核单据(lng材料ID) = True Then
            gstrSQL = "Select 名称 From 收费项目目录 where id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID)
            If rsTemp.EOF Then Exit Function
            If MsgBox(rsTemp!名称 & "存在未审核单据，调整成本价可能会造成差价误差。" & _
                vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    IsValied成本价 = True:
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------
    '功能:保存数据
    '参数:
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/17
    '------------------------------------------------------------------------------
    Dim dtToDay As Date, lng收费价目id As Long, lng材料ID As Long, lngId As Long, strID As String
    Dim ArrayID As Variant, strTemp As String
    Dim strNo As Variant, i As Long, lng序号 As Long
    Dim cllProc As Collection
    
    Set cllProc = New Collection
    
    err = 0: On Error GoTo ErrInfor:
    dtToDay = sys.Currentdate
    
    If m调价方式 = T_成本价调价 Then
    Else
        lng收费价目id = sys.NextId("收费价目")
        strNo = sys.GetNextNo(9)
        If IsNull(strNo) Then Exit Function
    End If
    With Me.vsPrice
        strID = ""
        lng序号 = 1
        For i = 1 To .Rows - 1
            lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("品名")))
            If lng材料ID <> 0 Then
                If Val(.TextMatrix(i, .ColIndex("原收入id"))) <> Val(.Cell(flexcpData, i, .ColIndex("收入名称"))) Or _
                    Val(.TextMatrix(i, .ColIndex("原价"))) <> Val(.TextMatrix(i, .ColIndex("现价"))) _
                    And m调价方式 <> T_成本价调价 Then
                        lngId = sys.NextId("收费价目")
                        If Me.chk立即执行.Value = 1 Then
                            strID = strID & "," & lngId
                        ElseIf .RowData(i) = -1 Then
                            strID = strID & "," & lngId
                        End If
                        If .RowData(i) <> 0 Then
                            '刘兴宏:主要是解决可以为零的情况,比如：疫苗.是免费的
                            '问题:9569 2006-11-20
                            'If Val(.TextMatrix(i, col现价)) <> 0 Then
                                '设置上一次的价格记录终止执行
                                ' zl_收费价目_stop (
                                gstrSQL = "zl_收费价目_stop("
                                '    收费细目ID_IN IN 收费价目.收费细目ID%TYPE,
                                gstrSQL = gstrSQL & "" & lng材料ID & ","
                                '    终止日期_IN IN 收费价目.终止日期%TYPE := NULL
                                If Me.chk立即执行.Value Then
                                    gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToDay), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                Else
                                    gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtp执行日期.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                End If
                                gstrSQL = gstrSQL & ")"
                                AddArray cllProc, gstrSQL
                                
                                'Zl_收费价目_Insert
                                gstrSQL = "zl_收费价目_Insert("
                                '  Id_In         In 收费价目.ID%Type,
                                gstrSQL = gstrSQL & "" & lngId & ","
                                '  原价id_In     In 收费价目.原价id%Type := Null,
                                gstrSQL = gstrSQL & "" & IIf(.RowData(i) = -1, "NUll", .RowData(i)) & ","
                                '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                gstrSQL = gstrSQL & "" & lng材料ID & ","
                                '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                gstrSQL = gstrSQL & "" & IIf(Val(.Cell(flexcpData, i, .ColIndex("收入名称"))) = 0, "NULL", Val(.Cell(flexcpData, i, .ColIndex("收入名称")))) & ","
                                '  原价_In       In 收费价目.原价%Type := Null,
                                If .TextMatrix(i, .ColIndex("类型")) = "时价" And Val(.Cell(flexcpData, i, .ColIndex("类型"))) = 0 Then
                                    '非跟踪卫生材料的实价卫材，是以范围决定的，（主要是医嘱应用),始终填为零
                                    gstrSQL = gstrSQL & "" & 0 & ","
                                Else
                                    gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, .ColIndex("原价"))) / Val(.TextMatrix(i, .ColIndex("系数"))), g_小数位数.obj_散装小数.零售价小数) & ","
                                End If
                                
                                '  现价_In       In 收费价目.现价%Type := Null,
                                gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, .ColIndex("现价"))) / Val(.TextMatrix(i, .ColIndex("系数"))), g_小数位数.obj_散装小数.零售价小数) & ","
                                '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                gstrSQL = gstrSQL & "NULL,"
                                '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                gstrSQL = gstrSQL & "NULL,"
                                '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                gstrSQL = gstrSQL & "'" & Me.txt说明.Text & "',"
                                '  调价id_In     In 收费价目.调价id%Type := Null,
                                gstrSQL = gstrSQL & "" & lng收费价目id & ","
                                '  调价人_In     In 收费价目.调价人%Type := Null,
                                gstrSQL = gstrSQL & "'" & Me.txt调价人.Text & "',"
                                '  执行日期_In   In 收费价目.执行日期%Type := Null,
                                If Me.chk立即执行.Value Then
                                    gstrSQL = gstrSQL & "to_date('" & Format(dtToDay, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                Else
                                    gstrSQL = gstrSQL & "to_date('" & Format(Me.dtp执行日期.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                End If
                                '  变动原因_In   In 收费价目.变动原因%Type := 1,
                                gstrSQL = gstrSQL & "" & 0 & ","
                                '  No_In         In 收费价目.NO%Type := Null,
                                gstrSQL = gstrSQL & "'" & strNo & "',"
                                '  序号_In       In 收费价目.序号%Type := 1
                                gstrSQL = gstrSQL & "" & lng序号 & ","
                                '缺省价格_In
                                If .TextMatrix(i, .ColIndex("类型")) = "时价" And Val(.Cell(flexcpData, i, .ColIndex("类型"))) = 0 Then
                                        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, .ColIndex("现价"))) / Val(.TextMatrix(i, .ColIndex("系数"))), g_小数位数.obj_散装小数.零售价小数) & ")"
                                Else
                                        gstrSQL = gstrSQL & "NULL)"
                                End If
                                AddArray cllProc, gstrSQL
                                lng序号 = lng序号 + 1
                        End If
                End If
                '是否存在指导价格的调整，如果存在，调整指导价格
                '更新指导零售价
                If lng材料ID <> 0 Then
                    If Val(.TextMatrix(i, .ColIndex("原指导售价"))) <> Val(.TextMatrix(i, .ColIndex("现指导售价"))) Then
                        strTemp = Round(Val(.TextMatrix(i, .ColIndex("现指导售价"))) / Val(.TextMatrix(i, .ColIndex("系数"))), g_小数位数.obj_散装小数.零售价小数)
                        'zl_材料特性_UpdateCustom ( 材料ID_IN ,SQL_IN)
                        gstrSQL = "zl_材料特性_UpdateCustom(" & lng材料ID & ",'指导零售价=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                    '更新采购限价
                    If Val(.TextMatrix(i, .ColIndex("原采购限价"))) <> Val(.TextMatrix(i, .ColIndex("现采购限价"))) Then
                        strTemp = Round(Val(.TextMatrix(i, .ColIndex("现采购限价"))) / Val(.TextMatrix(i, .ColIndex("系数"))), g_小数位数.obj_散装小数.成本价小数)
                        'zl_材料特性_UpdateCustom ( 材料ID_IN ,SQL_IN)
                        gstrSQL = "zl_材料特性_UpdateCustom(" & lng材料ID & ",'指导批发价=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                End If
            End If
        Next
    End With
    
    Dim lng供应商ID As Long, lng批次 As Long, lng库房ID As Long, dbl成本价 As Double
    Dim str发票号 As String, str发票日期 As String, dbl发票金额 As Double, lng系数 As Long, j As Long
    
    Dim cllTemp As Collection
    
    '成本价调价处理
    If m调价方式 = T_成本价调价 Or m调价方式 = T_成本和售价调价 Then
        With vsStoce
            For i = 1 To .Rows - 1
                lng库房ID = Val(.Cell(flexcpData, i, .ColIndex("库房")))
                lng供应商ID = Val(.Cell(flexcpData, i, .ColIndex("供应商")))
                lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("卫材信息")))
                lng批次 = Val(.Cell(flexcpData, i, .ColIndex("批号")))
                lng系数 = Val(.Cell(flexcpData, i, .ColIndex("单位")))
                If lng材料ID <> 0 Then
                    str发票号 = "": str发票日期 = "": dbl发票金额 = 0
                    If chk应付.Value = 1 Then
                        With vsPay
                            For j = 1 To .Rows - 1
                                If Val(.Cell(flexcpData, j, .ColIndex("卫材信息"))) = lng材料ID And _
                                    Val(.Cell(flexcpData, j, .ColIndex("供应商"))) = lng供应商ID Then
                                    '看是否有此卫生材料库存变动情况
                                    str发票号 = Trim(.TextMatrix(j, .ColIndex("发票号")))
                                    str发票日期 = Trim(.TextMatrix(j, .ColIndex("发票日期")))
                                    dbl发票金额 = Val(.TextMatrix(j, .ColIndex("发票金额")))
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                    
                    dbl成本价 = Round(Val(.TextMatrix(i, .ColIndex("现成本价"))) / lng系数, g_小数位数.obj_散装小数.成本价小数)
                    
                    ' Zl_材料成本调价_Insert
                    gstrSQL = "Zl_材料成本调价_Insert("
                    '  供药单位id_In In 成本价调价信息.供药单位id%Type,
                    gstrSQL = gstrSQL & IIf(lng供应商ID = 0, "Null", lng供应商ID) & ","
                    '  库房id_In     In 成本价调价信息.库房id%Type,
                    gstrSQL = gstrSQL & "" & lng库房ID & ","
                    '  材料id_In     In 成本价调价信息.药品id%Type,
                    gstrSQL = gstrSQL & "" & lng材料ID & ","
                    '  批次_In       In 成本价调价信息.批次%Type := Null,
                    gstrSQL = gstrSQL & "" & lng批次 & ","
                    '  原成本价_In   In 成本价调价信息.原成本价%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(Val(.Cell(flexcpData, i, .ColIndex("原成本价"))), g_小数位数.obj_散装小数.成本价小数) & ","
                    '  新成本价_In   In 成本价调价信息.新成本价%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl成本价 & ","
                    '  发票号_In     In 成本价调价信息.发票号%Type := Null,
                    gstrSQL = gstrSQL & "'" & str发票号 & "',"
                    '  发票日期_In   In 成本价调价信息.发票日期%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "NULL", "to_date('" & str发票日期 & "','yyyy-mm-dd') ") & ","
                    '  发票金额_In   In 成本价调价信息.发票金额%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                    '  应付款变动_In In 成本价调价信息.应付款变动%Type := 0
                    gstrSQL = gstrSQL & "" & IIf(chk应付.Value = 1 And lng供应商ID <> 0 And dbl发票金额 <> 0, 1, 0) & ")"
                    AddArray cllProc, gstrSQL
                End If
            Next
        End With
    End If
    
    
    '分两种情况下对成本价进行调整:
    '1.当仅为成本价调价及立即执行时，立即对成本价进行调整
    '2.当非立即执行和非成本价(即成本价调价方式)调价时，在卫生材料调价时，再执行。
     '单独成本价调价时
    If m调价方式 = T_成本价调价 And Me.chk立即执行.Value = 1 Then
        With vsPrice
            For i = 1 To .Rows - 1
                lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("品名")))
                If lng材料ID <> 0 Then
                  ' Zl_材料收发记录_Adjust
                  gstrSQL = "Zl_材料收发记录_Adjust("
                  '  调价id_In In Number, --调价记录的ID
                  gstrSQL = gstrSQL & "" & 0 & ","
                  '  定价_In   In Number := 0, --是否转为定价销售（更新材料特性、收费细目中的变价）
                  gstrSQL = gstrSQL & "" & 0 & ","
                  '  材料id_In In Number := 0 --当不为0时表示是成本价调价，不处理售价相关内容
                    gstrSQL = gstrSQL & "" & lng材料ID & ")"
                  AddArray cllProc, gstrSQL
                End If
            Next
        End With
    End If
    
    If strID <> "" Then strID = Mid(strID, 2)
    '循环执行过程
    ArrayID = Split(strID, ",")
    For i = 0 To UBound(ArrayID)
        If Val(ArrayID(i)) <> 0 Then
            'Zl_材料收发记录_Adjust
            gstrSQL = "zl_材料收发记录_adjust("
            '  Adjustid In Number, --调价记录的ID
            gstrSQL = gstrSQL & "" & ArrayID(i) & ","
            '  Bln定价  In Number := 0 --是否转为定价销售（更新药品目录、收费细目中的变价）
            gstrSQL = gstrSQL & "" & Me.Chk定价.Value & ")"
            AddArray cllProc, gstrSQL
        End If
    Next
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, Me.Caption
    mlngBillId = 0
    mlngStuffId = 0
    SaveData = True
    Exit Function
ErrInfor:
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 

 Private Sub printbill(ByVal intPrintMode As Byte)
    '-------------------------------------------------------------------------------------
    '功能:打印
    '参数:intPrintMode-1-打印,2-预览,3-Excel
    '-------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.vsPrice.TextMatrix(1, 0)) = "" Then Exit Sub
    objPrint.Title.Text = "卫材调价通知单"
    
    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txt说明.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(Me.chk立即执行.Value, sys.Currentdate, Me.dtp执行日期.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txt调价人.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.vsPrice
    objPrint.PageFooter = 2
     
    If intPrintMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, intPrintMode
    End If
    Set objPrint = Nothing
End Sub

Private Sub cmdPrintStoce_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim i As Long
    
    
    
    If Trim(vsStoce.TextMatrix(1, vsStoce.ColIndex("卫材信息"))) = "" Then Exit Sub

    objPrint.Title.Text = "调价库存变动表"

    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txt说明.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(Me.chk立即执行.Value, sys.Currentdate, Me.dtp执行日期.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txt调价人.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    '先设置相关隐藏列的宽度
    With vsStoce
        For i = 0 To .Cols - 1
            If .ColHidden(i) Then
                .ColData(i) = .ColWidth(i)
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Set objPrint.Body = vsStoce
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
    '打印完成后,恢复相关隐藏列的宽度
    With vsStoce
        For i = 0 To .Cols - 1
            If .ColHidden(i) Then
                .ColWidth(i) = Val(.ColData(i))
                .ColData(i) = ""
            End If
        Next
    End With

End Sub
  
Private Sub SetCtlEnabled()
    '---------------------------------------------------------------------------------------------
    '功能:设置相关控件的Enabled属性
    '返回:
    '编制:刘兴宏
    '日期:2007/07/17
    '---------------------------------------------------------------------------------------------
    If mBillType <> B_查阅 Then
        With vsPrice
            .Editable = flexEDKbdMouse
            
        End With
        Exit Sub
    End If
    vsPrice.Editable = flexEDNone
    DkPane.Panes(ID_PANE_SEARCH).Close
    
    'DkPane.CloseAll
    Me.txt说明.Enabled = False
    Me.chk立即执行.Value = 0
    Me.chk立即执行.Enabled = False
    Me.dtp执行日期.Enabled = False
End Sub
Private Sub InitBill()
    '---------------------------------------------------------------------------------------------
    '功能:初始化调价信息
    '返回:
    '编制:刘兴宏
    '日期:2007/07/17
    '---------------------------------------------------------------------------------------------
    Dim dtDate As Date, rsTemp As New ADODB.Recordset, i As Long
    dtDate = sys.Currentdate
    
    On Error GoTo ErrHandle

    If mlngBillId = 0 Then
        '进入调价编辑状态
        stbThis.Panels(2).Text = "库存变动表：(由于调价未保存，反映的库存可能不准确)"
         Me.dtp执行日期.MinDate = DateAdd("s", 1, dtDate)
        Me.dtp执行日期.Value = DateAdd("d", 1, dtDate)
        Me.txt调价人.Text = gstrUserName
        
        If mlngStuffId = 0 Then Exit Sub
        '如果指定首先调价的卫材，则直接将该卫材调入
        gstrSQL = "" & _
            "   Select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位,M.分类ID,J.编码 ||'-'||J.名称 as 分类," & _
            "           P.包装单位,decode(I.是否变价,1,'时价','定价') 类型,p.指导零售价," & _
            "           P.指导批发价,P.指导零售价,p.成本价," & _
                        IIf(mintUnit = 0, "1", "nvl(p.换算系数,1)") & " as 换算系数" & _
            "   From 收费项目目录 I,材料特性 P,诊疗项目目录 M,诊疗分类目录 J" & _
            "   Where I.ID=[1] And I.ID=P.材料ID And P.诊疗ID=M.id and M.分类ID=J.id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngStuffId)
        With vsPrice
            If rsTemp.EOF Then
                Exit Sub
            End If
            txt分类.Text = zlStr.Nvl(rsTemp!分类)
            txt分类.Tag = zlStr.Nvl(rsTemp!分类id)
            .Redraw = flexRDNone
            .TextMatrix(.Row, .ColIndex("品名")) = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!名称)
            .Cell(flexcpData, .Row, .ColIndex("品名")) = zlStr.Nvl(rsTemp!Id)
            .TextMatrix(.Row, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
            .TextMatrix(.Row, .ColIndex("产地")) = zlStr.Nvl(rsTemp!产地)
            .TextMatrix(.Row, .ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!计算单位), zlStr.Nvl(rsTemp!包装单位))
            .TextMatrix(.Row, .ColIndex("类型")) = zlStr.Nvl(rsTemp!类型)
            .TextMatrix(.Row, .ColIndex("系数")) = zlStr.Nvl(rsTemp!换算系数)
            .TextMatrix(.Row, .ColIndex("原成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_成本价)
            .Cell(flexcpData, .Row, .ColIndex("原成本价")) = Val(zlStr.Nvl(rsTemp!成本价))
            .TextMatrix(.Row, .ColIndex("现成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_成本价)
            .Cell(flexcpData, .Row, .ColIndex("现成本价")) = Val(zlStr.Nvl(rsTemp!成本价))
            
            .TextMatrix(.Row, .ColIndex("原采购限价")) = Format(Val(zlStr.Nvl(rsTemp!指导批发价)) * Val(rsTemp!换算系数), mFMT.FM_成本价)
            .TextMatrix(.Row, .ColIndex("现采购限价")) = .TextMatrix(.Row, .ColIndex("原采购限价"))
            .Cell(flexcpData, .Row, .ColIndex("原采购限价")) = Val(zlStr.Nvl(rsTemp!指导批发价))
            .Cell(flexcpData, .Row, .ColIndex("现采购限价")) = Val(zlStr.Nvl(rsTemp!指导批发价))
            
            .TextMatrix(.Row, .ColIndex("指导零售价")) = Format(Val(zlStr.Nvl(rsTemp!指导零售价)) * Val(rsTemp!换算系数), mFMT.FM_零售价)
            .TextMatrix(.Row, .ColIndex("原指导售价")) = .TextMatrix(.Row, .ColIndex("指导零售价"))
            .TextMatrix(.Row, .ColIndex("现指导售价")) = .TextMatrix(.Row, .ColIndex("指导零售价"))
            
            .Cell(flexcpData, .Row, .ColIndex("指导零售价")) = Val(zlStr.Nvl(rsTemp!指导零售价))
            .Cell(flexcpData, .Row, .ColIndex("原指导售价")) = Val(zlStr.Nvl(rsTemp!指导零售价))
            .Cell(flexcpData, .Row, .ColIndex("现指导售价")) = Val(zlStr.Nvl(rsTemp!指导零售价))
            Call zlGetPrice(Val(zlStr.Nvl(rsTemp!Id)), IIf(.TextMatrix(.Row, .ColIndex("类型")) = "时价", True, False), False, .Row)
            Call LoadStockData(Val(zlStr.Nvl(rsTemp!Id)), Val(.Cell(flexcpData, .Row, .ColIndex("原价"))), Val(.Cell(flexcpData, .Row, .ColIndex("现价"))))
            .Col = .ColIndex("现价")
            .Redraw = flexRDBuffered
            mlngPreRow = 0:
            Call vsPrice_RowColChange
            Exit Sub
        End With
    End If
    
    '进入调价显示状态
    Dim strBills As String
    strBills = ""
    gstrSQL = "" & _
        "   Select P.ID,M.id as 材料id,'['||M.编码||']'||M.名称 as 品名 ,decode(M.是否变价,1,'时价','定价') 类型,M.规格,M.产地,M.计算单位 as 单位," & _
                IIf(mintUnit = 0, "1", " nvl(j.换算系数,1) ") & " as 换算系数 ,j.包装单位," & _
        "        P.原价,P.现价,P.收入项目id,I.名称 as 收入名称," & _
        "        To_Char(P.执行日期,'yyyy-MM-dd hh24:mi:ss') 执行日期,P.变动原因,P.调价说明,P.调价人,j.成本价,j.指导零售价" & _
        "   From 收费价目 P,收费项目目录 M,收入项目 I,材料特性 J" & _
        "   Where P.收费细目id=M.id and P.收入项目id=I.id And M.ID=J.材料ID and P.ID=[1] " & _
        GetPriceClassString("P") & _
        "   Order by P.id"                            '因调价ID取的是价格记录ID的上一个ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngBillId)
    i = 1
    With vsPrice
        .Redraw = flexRDNone
        If rsTemp.EOF = False Then
            Me.txt说明 = zlStr.Nvl(rsTemp!调价说明)
            Me.txt调价人.Text = zlStr.Nvl(rsTemp!调价人)
            Me.dtp执行日期.Value = zlStr.Nvl(rsTemp!执行日期)
            .Rows = rsTemp.RecordCount + 1
        Else
            .Rows = 2
        End If
        Do While Not rsTemp.EOF
            strBills = strBills & "," & rsTemp!Id
            .RowData(i) = Val(zlStr.Nvl(rsTemp!Id))
            .TextMatrix(i, .ColIndex("品名")) = zlStr.Nvl(rsTemp!品名)
            .Cell(flexcpData, i, .ColIndex("品名")) = zlStr.Nvl(rsTemp!材料ID)
            .TextMatrix(i, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
            .TextMatrix(i, .ColIndex("产地")) = zlStr.Nvl(rsTemp!产地)
            .TextMatrix(i, .ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!单位), zlStr.Nvl(rsTemp!包装单位))
            .TextMatrix(i, .ColIndex("类型")) = zlStr.Nvl(rsTemp!类型)
            .TextMatrix(i, .ColIndex("系数")) = zlStr.Nvl(rsTemp!换算系数)
            .TextMatrix(i, .ColIndex("现成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * Val(rsTemp!换算系数), mFMT.FM_成本价)
            .TextMatrix(i, .ColIndex("指导零售价")) = Format(Val(zlStr.Nvl(rsTemp!指导零售价)) * Val(rsTemp!换算系数), mFMT.FM_零售价)
            .TextMatrix(i, .ColIndex("上次日期")) = Format(rsTemp!执行日期, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(i, .ColIndex("原价")) = Format(Val(zlStr.Nvl(rsTemp!原价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_零售价)
            .TextMatrix(i, .ColIndex("现价")) = Format(Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_零售价)
            .TextMatrix(i, .ColIndex("原收入id")) = Val(zlStr.Nvl(rsTemp!收入项目id))
            .TextMatrix(i, .ColIndex("收入名称")) = zlStr.Nvl(rsTemp!收入名称)
            .Cell(flexcpData, i, .ColIndex("收入名称")) = Val(zlStr.Nvl(rsTemp!收入项目id))
            
            If zlStr.Nvl(rsTemp!执行日期) <= Format(dtDate, "yyyy-mm-dd HH:MM:SS") And rsTemp!变动原因 = 0 Then       '未进行调价计算,则执行计算
                gstrSQL = "zl_材料收发记录_Adjust(" & rsTemp!Id & ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
            rsTemp.MoveNext
        Loop
        .Col = .ColIndex("现价")
        .Redraw = flexRDBuffered
    End With
    If strBills <> "" Then strBills = Mid(strBills, 2)
    
    If rsTemp.RecordCount = 0 Then
        mlngPreRow = 0:
        Call vsPrice_RowColChange
        Exit Sub
    End If
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst

    If rsTemp!执行日期 > dtDate Then
        '如果执行时间未到，则只能模拟显示库存变动
        Me.stbThis.Panels(2).Text = "库存变动表：(由于执行时间未到，反映的库存可能不准确)"
    Else
        '执行时间已到，肯定也进行了调价计算，直接从收发记录提取调价变动情况
        gstrSQL = "" & _
        "   Select S.ID,S.药品ID as 材料ID,D.名称 as 库房,'['||M.编码||']'||M.名称 as 材料信息,M.规格,M.产地,M.计算单位 as 单位, " & _
        "       P.包装单位,P.换算系数,S.批号,S.数量,S.原价,S.现价,S.调整金额" & _
        "   From (  Select ID,库房ID,药品ID,批号,填写数量 as 数量,成本价 as 原价,零售价 as 现价,零售金额 as 调整金额" & _
        "           From (  Select P.ID,N.库房ID,N.药品ID,N.批号,N.填写数量,N.成本价,N.零售价,N.零售金额" & _
        "                   From 药品收发记录 N, (select ID,收费细目ID,执行日期,终止日期 from 收费价目 where ID=[1]" & _
        GetPriceClassString("") & ") P" & _
        "                   where   N.药品ID=P.收费细目ID and N.单据=13 and N.费用ID is null " & _
        "                           and N.审核日期 Between P.执行日期 and nvl(P.终止日期,sysdate))) S," & _
        "       部门表 D,收费项目目录 M,材料特性 P" & _
        " where S.库房id+0=D.id and S.药品ID=M.ID And M.ID=P.材料ID" & _
        " order by M.编码,S.批号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngBillId)
        With vsStoce
            .Rows = 2
            .Clear 1
            If rsTemp.RecordCount > 0 Then .Rows = rsTemp.RecordCount + 1
            i = 1
            Do While Not rsTemp.EOF
                .TextMatrix(i, .ColIndex("库房")) = zlStr.Nvl(rsTemp!库房)
                .TextMatrix(i, .ColIndex("卫材信息")) = zlStr.Nvl(rsTemp!材料信息)
                .TextMatrix(i, .ColIndex("规格|产地")) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格) & IIf(IsNull(rsTemp!产地), "", "|" & rsTemp!产地)
                If mintUnit = 0 Then
                    .TextMatrix(i, .ColIndex("单位")) = IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
                Else
                    .TextMatrix(i, .ColIndex("单位")) = IIf(IsNull(rsTemp!包装单位), "", rsTemp!包装单位)
                End If
                .TextMatrix(i, .ColIndex("批号")) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                .TextMatrix(i, .ColIndex("数量")) = Format(Val(zlStr.Nvl(rsTemp!数量)) / Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_数量)
                .TextMatrix(i, .ColIndex("原价")) = Format(Val(zlStr.Nvl(rsTemp!原价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_零售价)
                .TextMatrix(i, .ColIndex("现价")) = Format(Val(zlStr.Nvl(rsTemp!现价)) * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_零售价)
                .TextMatrix(i, .ColIndex("调整额")) = Format(Val(zlStr.Nvl(rsTemp!调整金额)), mFMT.FM_金额)
                i = i + 1
                rsTemp.MoveNext
            Loop
        End With
        mlngPreRow = 0:
        Call vsPrice_RowColChange
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '-----------------界面显示调整---------------------------------
    Me.Caption = "卫生材料调价"
    Call InitOther
    Call InitCommandBar
    Call InitPanel
    
    '初始页面
    Call InitPage
    Call InitControl
    Call InitBill
    Call SetControlVisble
    Call SetCtlEnabled
    Call SetColor(1)
    '-----------------------------------------------------------
     
    zl_vsGrid_Para_Restore mlngModule, vsPay, Me.Caption, "应付变动"
    zl_vsGrid_Para_Restore mlngModule, vsPay, Me.Caption, "库存变动"
    DkPane.RecalcLayout
    With vsStoce
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = &H8000000F
    End With
    With vsPay
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = &H8000000F
    End With
    
    mblnSucces = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFirst = True
    Call RestoreWinState(Me)
    
    mlng供应商ID = 0
    mdbl加成率 = 0
    '判断是否以库房单位显示
    mintUnit = Get定价单位
    mlngModule = 1711
    mstrPrivs = ";" & GetPrivFunc(glngSys, mlngModule) & ";"
    
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    Call vsPrice_LostFocus
    Call vsStoce_LostFocus
    Call vsPay_LostFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    chkAppAllColumn.Move 0, 0
    
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    If Me.Width < 9720 Then
        Me.Width = 9720
    End If
    Dim panKind As Pane
    Set panKind = Me.DkPane.FindPane(ID_PANE_SEARCH)
    If Not panKind Is Nothing Then
        panKind.MinTrackSize.SetSize 295, Me.ScaleHeight / Screen.TwipsPerPixelY
        panKind.MaxTrackSize.SetSize 400, Me.ScaleHeight / Screen.TwipsPerPixelY
    End If
    Set panKind = Me.DkPane.FindPane(ID_PANE_STOCE)
    If Not panKind Is Nothing Then
        panKind.MinTrackSize.Height = 50
        panKind.MaxTrackSize.Height = (Me.ScaleHeight * 0.7) / Screen.TwipsPerPixelY
    End If
    Me.DkPane.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnModify Then If MsgBox("你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    SaveWinState Me
    DkPane.SaveState Me.Caption & "_Search", App.Title, "Layout"
    
     
    zl_vsGrid_Para_Save mlngModule, vsPay, Me.Caption, "应付变动"
    zl_vsGrid_Para_Save mlngModule, vsStoce, Me.Caption, "库存变动"
End Sub
Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.dtp执行日期.Enabled Then Me.dtp执行日期.SetFocus
End Sub

Private Sub DkPane_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case ID_PANE_SEARCH
        If Item.Handle = 0 Then Item.Handle = picSeach.hwnd
    Case ID_PANE_PRICE
        If Item.Handle = 0 Then Item.Handle = picPrice.hwnd
    Case ID_PANE_STOCE
        If Item.Handle = 0 Then Item.Handle = picStoceBack.hwnd
    End Select
End Sub
 
'***************************************************************************************************************
'**库存变动及应付变动处理
Private Sub SetControlVisble()
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置控件的Eanbled和Visble属性
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-05 15:40:21
    '-----------------------------------------------------------------------------------------------------------
    Dim bln存在成本价调价 As Boolean, bln指导价管理 As Boolean
    
    bln存在成本价调价 = (m调价方式 = T_成本价调价 Or m调价方式 = T_成本和售价调价)
    If mBillType = B_查阅 Then
        bln存在成本价调价 = False
    End If
    
    If m调价方式 = T_成本价调价 Then
        chk立即执行.Value = 1
        chk立即执行.Enabled = False
        Chk定价.Visible = False
        fra调整额.Enabled = False
        txt调整额.Text = ""
    Else
        chk立即执行.Enabled = True
        Chk定价.Visible = True
        fra调整额.Enabled = True
    End If
        
    With vsStoce
        '如果是成本价调价,需要允许编辑
        .Editable = IIf(bln存在成本价调价 And mBillType <> B_查阅, flexEDKbdMouse, flexEDNone)
        .ColHidden(.ColIndex("原成本价")) = Not bln存在成本价调价
        .ColHidden(.ColIndex("现成本价")) = Not bln存在成本价调价
        .ColHidden(.ColIndex("加成率")) = Not bln存在成本价调价
        .ColHidden(.ColIndex("差价调整额")) = Not bln存在成本价调价
        
        .ColHidden(.ColIndex("调整额")) = m调价方式 = T_成本价调价
        '.ColHidden(.ColIndex("原价")) = m调价方式 = T_成本价调价
        '.ColHidden(.ColIndex("现价")) = m调价方式 = T_成本价调价
    End With
    
    chk应付.Visible = bln存在成本价调价
    chk批次.Visible = bln存在成本价调价
    '看是否有此页信息没有
    tbPage.Item(1).Visible = bln存在成本价调价 And chk应付.Value = 1
    With vsPay
        .Editable = IIf(bln存在成本价调价, flexEDKbdMouse, flexEDNone)
    End With
    '检查是否存在指导价格管理权限
    bln指导价管理 = InStr(1, mstrPrivs, ";指导价格管理;") > 0 And Not (mBillType = B_查阅)
    With vsPrice
        .ColHidden(.ColIndex("原采购限价")) = Not bln指导价管理
        .ColHidden(.ColIndex("现采购限价")) = Not bln指导价管理
        .ColHidden(.ColIndex("原指导售价")) = Not bln指导价管理
        .ColHidden(.ColIndex("现指导售价")) = Not bln指导价管理
    End With
    fraCost.Enabled = bln存在成本价调价
    If fraCost.Enabled = False Then
        txtPriver.BackColor = &H8000000F
        cmdPriver.BackColor = &H8000000F
        txt加成率.BackColor = &H8000000F
    End If
End Sub
Private Sub InitControl()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化控件的默认属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-05 15:33:54
    '-----------------------------------------------------------------------------------------------------------
    With vsPrice
        .GridLines = flexGridInset
    End With
    With vsPay
        .Clear 1
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .SelectionMode = flexSelectionByRow
        .GridLines = flexGridInset
    End With
    With vsStoce
        .Clear 1
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .GridLines = flexGridInset
    End With
    
End Sub

Private Function MoveStockData(ByVal lng材料ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:移除指定材料的数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-05 12:00:34
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim lngRow As Long
    err = 0: On Error GoTo ErrHand:
    
    With vsStoce
        lngRow = 1
ReDo:
        If .Rows > 2 Then
            For i = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, i, .ColIndex("卫材信息"))) = lng材料ID Then
                    lngRow = i
                    .RemoveItem i
                    GoTo ReDo
                End If
            Next
        End If
        If .Rows <= 2 Then
            If Val(.Cell(flexcpData, 1, .ColIndex("卫材信息"))) = lng材料ID Then
                .Rows = 2
                .Clear 1
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            End If
        End If
        .Row = 1
    End With
    MoveStockData = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function LoadStockData(ByVal lng材料ID As Long, ByVal dbl原价 As Double, ByVal dbl现价 As Double, Optional bln批量 As Boolean = False) As Boolean
   '-----------------------------------------------------------------------------------------------------------
    '功能:加载库房数据
    '入参:
    '出参:
    '返回: 返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-05 11:57:09
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lng供应商ID As Long
    Dim dbl原成本价 As Double, dbl现成本价 As Double, dbl加成率 As Double, dbl发票金额 As Double
    Dim dblTemp As Double
    
    err = 0: On Error GoTo ErrHand:
        
    '先移出行数据
    Call MoveStockData(lng材料ID)
 
    gstrSQL = "" & _
    "   Select S.库房ID, S.药品ID as 材料ID,S.批次, " & _
    "           D.名称 as 库房,decode(L.编码,NULL ,'','['||L.编码||']') ||L.名称 as 供应商, " & _
    "           '['||M.编码||']'||M.名称 as 材料,M.规格,M.产地,M.计算单位," & _
    "           Nvl(M.是否变价, 0) 变价,S.批号,S.数量,S.时价售价,S.成本价,S.上次供应商ID," & _
    "           p.包装单位,P.指导差价率 As 差价率," & IIf(mintUnit = 0, "1", "nvl(p.换算系数,1)") & " as 换算系数" & _
    "   From (  Select  S.库房ID,S.药品ID,S.上次供应商ID,S.上次批号 批号,S.实际数量 as 数量,S.批次, " & _
    "                 decode(nvl(零售价,0),0,decode(nvl(实际数量,0),0,0,S.实际金额 / S.实际数量) ,零售价) 时价售价, " & _
    "                   s.平均成本价 As 成本价" & _
    "           From 药品库存 S" & _
    "           Where S.实际数量<>0 and S.性质=1 and S.药品id=[1] " & IIf(mlng供应商ID = 0, "", " And Nvl(S.上次供应商ID,0)=[2]") & ") S, " & _
    "       部门表 D,收费项目目录 M,材料特性 P,供应商 L" & _
    " where S.库房id=D.id and S.药品ID=M.ID And M.ID=P.材料ID and S.上次供应商ID=L.ID(+)" & _
    " order by M.编码,S.批号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID, mlng供应商ID)
 
    With vsStoce
        lngRow = .Rows - 1
        If Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息"))) <> 0 Then lngRow = lngRow + 1
        If lngRow = 1 Then
            .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        Else
            .Rows = .Rows + rsTemp.RecordCount
        End If
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("库房")) = zlStr.Nvl(rsTemp!库房)
            .Cell(flexcpData, lngRow, .ColIndex("库房")) = zlStr.Nvl(rsTemp!库房ID)
            .TextMatrix(lngRow, .ColIndex("供应商")) = zlStr.Nvl(rsTemp!供应商):
            .Cell(flexcpData, lngRow, .ColIndex("供应商")) = zlStr.Nvl(rsTemp!上次供应商id)
            .TextMatrix(lngRow, .ColIndex("卫材信息")) = zlStr.Nvl(rsTemp!材料)
            .Cell(flexcpData, lngRow, .ColIndex("卫材信息")) = zlStr.Nvl(rsTemp!材料ID)
            .TextMatrix(lngRow, .ColIndex("规格|产地")) = zlStr.Nvl(rsTemp!规格) & IIf(IsNull(rsTemp!产地), "", "|" & rsTemp!产地)
            .TextMatrix(lngRow, .ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!计算单位), zlStr.Nvl(rsTemp!包装单位))
            
            .Cell(flexcpData, lngRow, .ColIndex("单位")) = zlStr.Nvl(rsTemp!换算系数)
            .TextMatrix(lngRow, .ColIndex("批号")) = zlStr.Nvl(rsTemp!批号)
            .Cell(flexcpData, lngRow, .ColIndex("批号")) = zlStr.Nvl(rsTemp!批次)
            .TextMatrix(lngRow, .ColIndex("数量")) = Format(Val(zlStr.Nvl(rsTemp!数量)) / Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_数量)
            .Cell(flexcpData, lngRow, .ColIndex("数量")) = zlStr.Nvl(rsTemp!数量)
            
            '计算原价
            dblTemp = IIf(Val(zlStr.Nvl(rsTemp!变价)) = 1, Val(zlStr.Nvl(rsTemp!时价售价)), dbl原价) * Val(zlStr.Nvl(rsTemp!换算系数))
            .TextMatrix(lngRow, .ColIndex("原价")) = Format(dblTemp, mFMT.FM_零售价)
            .Cell(flexcpData, lngRow, .ColIndex("原价")) = IIf(Val(zlStr.Nvl(rsTemp!变价)) = 1, Val(zlStr.Nvl(rsTemp!时价售价)), dbl原价)
            
            .TextMatrix(lngRow, .ColIndex("现价")) = Format(dbl现价 * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_零售价)
            .Cell(flexcpData, lngRow, .ColIndex("现价")) = dbl现价
            .TextMatrix(lngRow, .ColIndex("调整额")) = Format(Val(zlStr.Nvl(rsTemp!数量)) * (dbl现价 - Val(.Cell(flexcpData, lngRow, .ColIndex("原价")))), mFMT.FM_金额)
            .Cell(flexcpData, lngRow, .ColIndex("调整额")) = Val(zlStr.Nvl(rsTemp!数量)) * (dbl现价 - Val(.Cell(flexcpData, lngRow, .ColIndex("原价"))))
             
             dbl原成本价 = Val(zlStr.Nvl(rsTemp!成本价))
            
            If mdbl加成率 > 0 Then
                dbl加成率 = Round(mdbl加成率 / 100, 7)
            ElseIf dbl原成本价 > 0 Then
                dbl加成率 = Round(dbl原价 / dbl原成本价 - 1, 7)
            Else
                dbl加成率 = Round(1 / (1 - rsTemp!差价率 / 100) - 1, 7)
            End If
            
            If 1 + dbl加成率 = 0 Then
                dbl现成本价 = 0
            Else
                dbl现成本价 = dbl现价 / (1 + dbl加成率)
            End If
            If dbl加成率 = -1 Then dbl加成率 = 0
            
            .TextMatrix(lngRow, .ColIndex("原成本价")) = Format(dbl原成本价 * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_成本价)
            .Cell(flexcpData, lngRow, .ColIndex("原成本价")) = dbl原成本价
            
            .TextMatrix(lngRow, .ColIndex("加成率")) = Format(dbl加成率 * 100, GFM_VBJCL)
            .Cell(flexcpData, lngRow, .ColIndex("加成率")) = dbl加成率 * 100
            
            
            .TextMatrix(lngRow, .ColIndex("现成本价")) = Format(dbl现成本价 * Val(zlStr.Nvl(rsTemp!换算系数)), mFMT.FM_成本价)
            .Cell(flexcpData, lngRow, .ColIndex("现成本价")) = dbl现成本价
            
            .TextMatrix(lngRow, .ColIndex("差价调整额")) = Format((dbl原成本价 - dbl现成本价) * Val(zlStr.Nvl(rsTemp!数量)), mFMT.FM_金额)
            .Cell(flexcpData, lngRow, .ColIndex("差价调整额")) = (dbl原成本价 - dbl现成本价) * Val(zlStr.Nvl(rsTemp!数量))
            .RowHidden(lngRow) = IIf(chk显示所有材料.Value = 1, False, True)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("卫材信息"))) = 0 And .Rows - 1 <> 1 Then
            .Rows = .Rows - 1
        End If
        
        '计算应付变动情况
        If (m调价方式 = T_成本价调价 Or m调价方式 = T_成本和售价调价) And bln批量 = False Then
            Call RefreshPayData
        End If
    End With
    LoadStockData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function RefreshPayData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:重新获取应付情况变动数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-05 15:03:46
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dbl发票金额 As Double
    Dim lng供应商ID As Long, lng材料ID As Long, blnData As Boolean
    
    err = 0: On Error GoTo ErrHand:
    If chk自动计算.Value <> 1 Then RefreshPayData = True: Exit Function
    
    With vsPay
        .Rows = 2
        .Clear 1
         .Cell(flexcpData, 1, .ColIndex("发票金额"), .Rows - 1, .ColIndex("发票金额")) = ""
    End With
    
    With vsStoce
        For i = 1 To .Rows - 1
            lng供应商ID = Val(.Cell(flexcpData, i, .ColIndex("供应商")))
            lng材料ID = Val(.Cell(flexcpData, i, .ColIndex("卫材信息")))
            If lng供应商ID <> 0 And lng材料ID <> 0 Then
                dbl发票金额 = Val(.Cell(flexcpData, i, .ColIndex("差价调整额")))
                If dbl发票金额 <> 0 Then
                    '先找相关的供应商是否存在
                    With vsPay
                        blnData = False
                        For j = 1 To .Rows - 1
                            If lng材料ID = Val(.Cell(flexcpData, j, .ColIndex("卫材信息"))) And _
                               lng供应商ID = Val(.Cell(flexcpData, j, .ColIndex("供应商"))) Then
                               .Cell(flexcpData, j, .ColIndex("发票金额")) = Val(.Cell(flexcpData, j, .ColIndex("发票金额"))) + dbl发票金额
                                .TextMatrix(j, .ColIndex("发票金额")) = Format(Val(.Cell(flexcpData, j, .ColIndex("发票金额"))), mFMT.FM_金额)
                               blnData = True
                               Exit For
                            End If
                        Next
                        If blnData = False Then
                            '没有此供应商或材料,因此需要额外增加
                            If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("供应商"))) <> 0 Then
                                .Rows = .Rows + 1
                            End If
                            .TextMatrix(.Rows - 1, .ColIndex("供应商")) = vsStoce.TextMatrix(i, vsStoce.ColIndex("供应商"))
                             .Cell(flexcpData, .Rows - 1, .ColIndex("供应商")) = vsStoce.Cell(flexcpData, i, vsStoce.ColIndex("供应商"))
                            .TextMatrix(.Rows - 1, .ColIndex("卫材信息")) = vsStoce.TextMatrix(i, vsStoce.ColIndex("卫材信息"))
                             .Cell(flexcpData, .Rows - 1, .ColIndex("卫材信息")) = vsStoce.Cell(flexcpData, i, vsStoce.ColIndex("卫材信息"))
                            .TextMatrix(.Rows - 1, .ColIndex("规格|产地")) = vsStoce.TextMatrix(i, vsStoce.ColIndex("规格|产地"))
                            .Cell(flexcpData, .Rows - 1, .ColIndex("发票金额")) = dbl发票金额
                            .TextMatrix(.Rows - 1, .ColIndex("发票金额")) = Format(Val(.Cell(flexcpData, .Rows - 1, .ColIndex("发票金额"))), mFMT.FM_金额)
                        End If
                    End With
                End If
            End If
        Next
    End With
    
    RefreshPayData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitPage()
    '------------------------------------------------------------------------------
    '功能:初始化页面控件
    '返回:
    '编制:刘兴宏
    '日期:2007/08/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As TabControlItem
    
    Set objItem = tbPage.InsertItem(mPageNum.Page_库存调整, "库存变动", picStoce.hwnd, 0)
    objItem.Tag = mPageNum.Page_库存调整
    Set objItem = tbPage.InsertItem(mPageNum.Page_应付调整, "应付变动", picPay.hwnd, 0)
    objItem.Tag = mPageNum.Page_应付调整
    
    With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk应付_Click()
    Call SetControlVisble
End Sub

Private Sub picPay_Resize()
    err = 0: On Error Resume Next
    With picPay
        vsPay.Width = .ScaleWidth
        vsPay.Left = .ScaleLeft
        chk自动计算.Move .ScaleLeft, .ScaleTop + 100
        vsPay.Top = chk自动计算.Top + chk自动计算.Height + 100
        vsPay.Height = .ScaleHeight - .Top
    End With
End Sub
Private Sub picStoce_Resize()
    err = 0: On Error Resume Next
    With picStoce
        vsStoce.Width = .ScaleWidth
        vsStoce.Left = .ScaleLeft
        chk批次.Move .ScaleLeft + 100, .ScaleTop + 100
        chk应付.Move chk批次.Left + chk批次.Width + 100, .ScaleTop + 100
        If chk应付.Visible = False And chk批次.Visible = False Then
            chk显示所有材料.Move chk批次.Left, .ScaleTop + 100
        ElseIf chk应付.Visible = False And chk批次.Visible = True Then
            chk显示所有材料.Move chk应付.Left, .ScaleTop + 100
        ElseIf chk应付.Visible = True And chk批次.Visible = False Then
            chk应付.Left = chk批次.Left
            chk显示所有材料.Move chk应付.Left + chk应付.Width + 100, .ScaleTop + 100
        Else
            chk显示所有材料.Move chk应付.Left + chk应付.Width + 100, .ScaleTop + 100
        End If
        
        vsStoce.Top = chk批次.Top + chk批次.Height + 100
        
        cmdPrintStoce.Top = .ScaleTop
        cmdPrintStoce.Left = .Width - cmdPrintStoce.Width - 15
        vsStoce.Height = .ScaleHeight - vsStoce.Top
    End With

End Sub

Private Sub vsStoce_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsStoce
        Select Case Col
        Case .ColIndex("原成本价"), .ColIndex("现成本价")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_成本价)
            '计算相关的值
            Call AutoCalcStoce(Row, Col)
        Case .ColIndex("加成率")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), GFM_VBJCL)
            '计算相关的值
            Call AutoCalcStoce(Row, Col)
        Case .ColIndex("品名")
            .ColComboList(Col) = "..."
        Case .ColIndex("收入名称")
            .ColComboList(Col) = "..."
        End Select
    End With
End Sub

Private Sub vsStoce_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    zl_VsGridRowChange vsStoce, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsStoce_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mBillType = B_查阅 Then Cancel = True: Exit Sub
    If m调价方式 = T_售价调价 Then Cancel = True: Exit Sub
    
    With vsStoce
        Select Case Col
        Case .ColIndex("现成本价"), .ColIndex("加成率")
             If Val(.Cell(flexcpData, Row, .ColIndex("卫材信息"))) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
    
End Sub


Private Sub vsStoce_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsPrice
        Select Case Col
        Case .ColIndex("卫材信息")
            '暂无
        Case Else
        End Select
    End With
End Sub

Private Sub vsStoce_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnModify = True
End Sub
Private Sub vsStoce_EnterCell()
    If mBillType = B_查阅 Then Exit Sub
    With vsStoce
        Select Case .Col
        Case .ColIndex("卫材信息")
        End Select
    End With
End Sub

Private Sub vsStoce_GotFocus()
'        zl_VsGridGotFocus vsStoce
End Sub

Private Sub vsStoce_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, lngRow As Long
    
    With vsStoce
        If (.Col = .ColIndex("卫材信息")) And KeyCode <> vbKeyReturn Then
            .ColComboList(.Col) = ""
        End If
        
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("卫材信息"))) = 0 Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsStoce, , , False, lngRow)
    End With
End Sub

Private Sub vsStoce_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsStoce
        Select Case Col
        Case .ColIndex("现成本价")
            strKey = Trim(vsStoce.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
             
        Case Else
        End Select
    End With
    Call zlVsMoveGridCell(vsStoce, , , False)
End Sub

Private Sub vsStoce_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsStoce_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    With vsStoce
        Select Case Col
        Case .ColIndex("卫材信息")
            Call VsFlxGridCheckKeyPress(vsStoce, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("现成本价")
            Call VsFlxGridCheckKeyPress(vsStoce, Row, Col, KeyAscii, m金额式)
        Case Else
        End Select
    End With
End Sub

Private Sub vsStoce_LostFocus()
'    zl_VsGridLOSTFOCUS vsStoce
End Sub

Private Sub vsStoce_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    If mBillType = B_查阅 Then Cancel = True: Exit Sub
    
    strKey = Trim(vsStoce.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    With vsStoce
        Select Case Col
        Case .ColIndex("现成本价")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "现指导批价") = False Then Cancel = True: Exit Sub
                If Val(.Cell(flexcpData, .Row, .ColIndex("卫材信息"))) <> 0 Then
                    If Check成本价(Val(.Cell(flexcpData, .Row, .ColIndex("卫材信息"))), Val(strKey)) = False Then
                        Cancel = True: Exit Sub
                    End If
                End If
                vsStoce.EditText = Format(Val(strKey), mFMT.FM_成本价)
            End If
        Case .ColIndex("加成率")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 5, , False, , "加成率") = False Then Cancel = True: Exit Sub
                vsStoce.EditText = Format(Val(strKey), GFM_VBJCL)
            End If
        End Select
    End With
End Sub


'*****************************************************************************************************************
'**应付变动处理
Private Sub vsPay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsPay
        Select Case Col
        Case .ColIndex("发票金额")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_金额)
        Case .ColIndex("发票日期")
            .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mBillType = B_查阅 Then Cancel = True: Exit Sub
    If m调价方式 = T_售价调价 Then Cancel = True: Exit Sub
    
    With vsPay
        Select Case Col
        Case .ColIndex("发票金额"), .ColIndex("发票号"), .ColIndex("发票日期")
             If Val(.Cell(flexcpData, Row, .ColIndex("卫材信息"))) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
    
End Sub
 
Private Sub vsPay_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsPay
        Select Case Col
        Case .ColIndex("发票日期")
            Call SelDate

        Case Else
        End Select
    End With
End Sub
Private Function SelDate() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择发票日期
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-07 11:59:54
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String, blnreturn As Boolean
    Dim sngX As Single, sngY As Single, lngH As Long
    strDate = vsPay.TextMatrix(vsPay.Row, vsPay.ColIndex("发票日期"))
    lngH = vsPay.CellHeight
    Call CalcPosition(sngX, sngY, vsPay)
      
    blnreturn = frmDateSel.SelectDate(Me, sngX, sngY, lngH, strDate)
    If blnreturn = False Then Exit Function
    With vsPay
        .TextMatrix(.Row, .ColIndex("发票日期")) = strDate
    End With
    SelDate = True
End Function
Private Sub vsPay_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnModify = True
End Sub
Private Sub vsPay_EnterCell()
    If mBillType = B_查阅 Then Exit Sub
    With vsPay
        Select Case .Col
        Case .ColIndex("发票日期")
            .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsPay_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, lngRow As Long
    
    With vsPay
        If (.Col = .ColIndex("发票日期")) And KeyCode <> vbKeyReturn And KeyCode <> Asc("*") And KeyCode <> vbKeySpace And KeyCode <> vbKeyShift Then
            If Shift = 1 And KeyCode = 56 Then
                vsPay_CellButtonClick .Row, .Col
            Else
                .ColComboList(.Col) = ""
            End If
        End If
        If KeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行的应付变动记录吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                .Clear 1
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""

            Else
                .RemoveItem .Row
            End If
        End If
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("卫材信息"))) = 0 Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsPay, , , False, lngRow)
    End With
End Sub

Private Sub vsPay_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPay
        Select Case Col
        Case .ColIndex("发票金额")
            strKey = Trim(vsPay.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
    End With
    Call zlVsMoveGridCell(vsPay, , , False)
End Sub
Private Sub vsPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub vsPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    With vsPay
        Select Case Col
        Case .ColIndex("发票号"), .ColIndex("发票日期")
            Call VsFlxGridCheckKeyPress(vsPay, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("发票金额")
            Call VsFlxGridCheckKeyPress(vsPay, Row, Col, KeyAscii, m负金额式)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPay_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    If mBillType = B_查阅 Then Cancel = True: Exit Sub
    
    strKey = Trim(vsPay.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    
    With vsPay
        Select Case Col
        Case .ColIndex("发票金额")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "发票金额") = False Then Cancel = True: Exit Sub
                vsPay.EditText = Format(Val(strKey), mFMT.FM_成本价)
            End If
        Case .ColIndex("发票日期")
            If strKey = "" Then Exit Sub
            strKey = zlCheckIsDate(strKey, "发票日期")
            If strKey = "" Then Cancel = True: Exit Sub
            .EditText = strKey
        Case .ColIndex("发票号")
            If strKey = "" Then Exit Sub
            If zlCommFun.StrIsValid(strKey, 200, 0, "发票号") = False Then Cancel = True: Exit Sub
        End Select
    End With
End Sub
'*************************************************************************************************************************

Private Sub AutoCalcStoce(ByVal lngEditRow As Long, ByVal lngEditCol As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:自动计算相关信息(根据加成率计算现成本价及差额,根据现成本价计算差额及加成率)
    '入参:lngEditRow-当前编辑的行
    '     lngEditCol-当前编辑的列
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 17:03:02
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl现成本价 As Double, dbl加成率 As Double, dbl成本差价 As Double, dbl差价调整额 As Double
    Dim lng材料ID As Long, bln库房分批 As Boolean, lng供应商ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long
    
    err = 0: On Error GoTo ErrHand:
    With vsStoce
        bln库房分批 = chk批次.Value = 1
        lngStep = IIf(bln库房分批, lngEditRow, 1)
        lngSteps = IIf(bln库房分批, lngEditRow, .Rows - 1)
        Select Case lngEditCol
        Case .ColIndex("加成率")
            dbl加成率 = Val(.TextMatrix(lngEditRow, lngEditCol)) / 100
            If dbl加成率 = -1 Then dbl加成率 = 0
            '现成本价=现零售价/(1+加成率)
            dbl现成本价 = Round(Val(.Cell(flexcpData, lngEditRow, .ColIndex("现价"))) / (1 + dbl加成率), 7)
            '差价调整额=(原成本价-现成本价)
            dbl成本差价 = (Val(.Cell(flexcpData, lngEditRow, .ColIndex("原成本价"))) - dbl现成本价)
        Case .ColIndex("现成本价")
            '因为存在包装换算问题，因此，目前按最小单位进行设置单价
            dbl现成本价 = Val(.TextMatrix(lngEditRow, lngEditCol)) / Val(.Cell(flexcpData, lngEditRow, .ColIndex("单位")))
            '加成率=现零售价/现成本价-1
            If dbl现成本价 <> 0 Then
                dbl加成率 = Round(Val(.Cell(flexcpData, lngEditRow, .ColIndex("现价"))) / dbl现成本价 - 1, 7)
            Else
                dbl加成率 = 0
            End If
            '差价调整额=(现成本价-原成本价)
            dbl成本差价 = Round((Val(.Cell(flexcpData, lngEditRow, .ColIndex("原成本价"))) - dbl现成本价), 7)
        Case .ColIndex("差价调整额")
            Exit Sub
        Case .ColIndex("现价")
            '现价发生改变时,需要重新根据加成率计算相关的现成本价
            dbl加成率 = Round(Val(.TextMatrix(lngEditRow, .ColIndex("加成率"))) / 100, 7)
            If dbl加成率 = -1 Then dbl加成率 = 0
            '现成本价=现零售价/(1+加成率)
            dbl现成本价 = Round(Val(.Cell(flexcpData, lngEditRow, .ColIndex("现价"))) / (1 + dbl加成率), 7)
            '差价调整额=(现成本价-原成本价)
            dbl成本差价 = (dbl现成本价 - Val(.Cell(flexcpData, lngEditRow, .ColIndex("原成本价"))))
            lngStep = lngEditRow
            lngSteps = lngEditRow
        Case Else
            Exit Sub
        End Select

        lng材料ID = Val(.Cell(flexcpData, lngEditRow, .ColIndex("卫材信息")))
        lng供应商ID = Val(.Cell(flexcpData, lngEditRow, .ColIndex("供应商")))
        Dim cllData As New Collection
        For lngRow = lngStep To lngSteps
            If lng材料ID = Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息"))) Then
                If dbl加成率 = -1 Then dbl加成率 = 0
                .TextMatrix(lngRow, .ColIndex("加成率")) = Format(dbl加成率 * 100, GFM_VBJCL)
                '该成本价是以最小单位为准的，因此要乘小换算系数.
                .TextMatrix(lngRow, .ColIndex("现成本价")) = Format(dbl现成本价 * Val(.Cell(flexcpData, lngRow, .ColIndex("单位"))), mFMT.FM_成本价)
                dbl成本差价 = (Val(.Cell(flexcpData, lngRow, .ColIndex("原成本价"))) - dbl现成本价)
                 '差价调整额=(现成本价-原成本价)*数量
                 dbl差价调整额 = Round(dbl成本差价 * Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))), 7)
                .TextMatrix(lngRow, .ColIndex("差价调整额")) = Format(dbl差价调整额, mFMT.FM_金额)
                .Cell(flexcpData, lngRow, .ColIndex("差价调整额")) = dbl差价调整额
                lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息")))
                lng供应商ID = Val(.Cell(flexcpData, lngRow, .ColIndex("供应商")))
                
                If lng供应商ID <> 0 Then
                    err = 0: On Error Resume Next
                    cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, .ColIndex("供应商")), .TextMatrix(lngRow, .ColIndex("卫材信息")), .TextMatrix(lngRow, .ColIndex("规格|产地"))), "K" & lng供应商ID & "_" & lngTemp
                    If err <> 0 Then
                        '累计差价调整额
                        dbl差价调整额 = Val(cllData("K" & lng供应商ID & "_" & lngTemp)(2)) + dbl差价调整额
                        cllData.Remove "K" & lng供应商ID & "_" & lngTemp
                         err = 0: On Error GoTo ErrHand:
                        cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, .ColIndex("供应商")), .TextMatrix(lngRow, .ColIndex("卫材信息")), .TextMatrix(lngRow, .ColIndex("规格|产地"))), "K" & lng供应商ID & "_" & lngTemp
                       
                    End If
                    On Error GoTo ErrHand:
                End If
            End If
        Next
        If chk自动计算.Value = 1 Then
            '需要自动计算相关的应付变动记录
            For i = 1 To cllData.Count
                With vsPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息")))
                        lng供应商ID = Val(.Cell(flexcpData, lngRow, .ColIndex("供应商")))
                        If lngTemp = Val(cllData(i)(0)) _
                            And lng供应商ID = Val(cllData(i)(1)) Then
                            '卫材及供应商相同,清空相关的值
                            .TextMatrix(lngRow, .ColIndex("发票金额")) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                            .Cell(flexcpData, lngRow, .ColIndex("发票金额")) = Val(cllData(i)(2))
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '需要增加该项供应商的物资
                        If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("卫材信息"))) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, .ColIndex("供应商")) = cllData(i)(3)
                        .Cell(flexcpData, lngRow, .ColIndex("供应商")) = cllData(i)(1)
                        .TextMatrix(lngRow, .ColIndex("卫材信息")) = cllData(i)(4)
                        .Cell(flexcpData, lngRow, .ColIndex("卫材信息")) = cllData(i)(0)
                        .TextMatrix(lngRow, .ColIndex("规格|产地")) = cllData(i)(5)
                        .TextMatrix(lngRow, .ColIndex("发票金额")) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                        .Cell(flexcpData, lngRow, .ColIndex("发票金额")) = Val(cllData(i)(2))
                    End If
                End With
            Next
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AutoCalc所有库存价格()
    '-----------------------------------------------------------------------------------------------------------
    '功能:自动计算所有库存的价格
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl现成本价 As Double, dbl现价 As Double, dbl加成率 As Double, dbl成本差价 As Double, dbl差价调整额 As Double, dbl调整额 As Double
    Dim lng材料ID As Long, bln库房分批 As Boolean, lng供应商ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long
    Dim intCol As Integer
    Dim cllData As New Collection
    
    err = 0: On Error GoTo ErrHand:
    
    '因为存在包装换算问题，因此，目前按最小单位进行设置单价
    dbl现成本价 = Val(vsPrice.TextMatrix(vsPrice.Row, vsPrice.ColIndex("现成本价")))
    dbl现价 = Val(vsPrice.TextMatrix(vsPrice.Row, vsPrice.ColIndex("现价")))
    
    With vsStoce
        For lngRow = 1 To .Rows - 1
            If vsPrice.Col = vsPrice.ColIndex("现成本价") Then
                .TextMatrix(lngRow, .ColIndex("现成本价")) = dbl现成本价
                '加成率=现零售价/现成本价-1
                If dbl现成本价 <> 0 Then
                    dbl加成率 = Round(Val(.Cell(flexcpData, lngRow, .ColIndex("现价"))) / dbl现成本价 - 1, 7)
                Else
                    dbl加成率 = 0
                End If
                '差价调整额=(现成本价-原成本价)
                dbl成本差价 = Round((Val(.Cell(flexcpData, lngRow, .ColIndex("原成本价"))) - dbl现成本价), 7)
            ElseIf vsPrice.Col = vsPrice.ColIndex("现价") Then
                .TextMatrix(lngRow, .ColIndex("现价")) = dbl现价
                '现价发生改变时,需要重新根据加成率计算相关的现成本价
                dbl加成率 = Round(Val(.TextMatrix(lngRow, .ColIndex("加成率"))) / 100, 7)
                If dbl加成率 = -1 Then dbl加成率 = 0
                '现成本价=现零售价/(1+加成率)
                dbl现成本价 = Round(dbl现价 / (1 + dbl加成率), 7)
                '差价调整额=(现成本价-原成本价)
                dbl成本差价 = (dbl现成本价 - Val(.Cell(flexcpData, lngRow, .ColIndex("原成本价"))))
                
                '调整额=数量*(现价-原价)
                dbl调整额 = (dbl现价 - Val(.Cell(flexcpData, lngRow, .ColIndex("原价")))) * Val(.Cell(flexcpData, lngRow, .ColIndex("数量")))
                .TextMatrix(lngRow, .ColIndex("调整额")) = Format(dbl调整额, mFMT.FM_金额)
                .Cell(flexcpData, lngRow, .ColIndex("调整额")) = dbl调整额
            End If
            
            lng材料ID = Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息")))
            lng供应商ID = Val(.Cell(flexcpData, lngRow, .ColIndex("供应商")))
            
            If dbl加成率 = -1 Then dbl加成率 = 0
            .TextMatrix(lngRow, .ColIndex("加成率")) = Format(dbl加成率 * 100, GFM_VBJCL)
            dbl成本差价 = (Val(.Cell(flexcpData, lngRow, .ColIndex("原成本价"))) - dbl现成本价)
             '差价调整额=(现成本价-原成本价)*数量
             dbl差价调整额 = Round(dbl成本差价 * Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))), 7)
            .TextMatrix(lngRow, .ColIndex("差价调整额")) = Format(dbl差价调整额, mFMT.FM_金额)
            .Cell(flexcpData, lngRow, .ColIndex("差价调整额")) = dbl差价调整额
            lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息")))
            lng供应商ID = Val(.Cell(flexcpData, lngRow, .ColIndex("供应商")))
            
            If lng供应商ID <> 0 Then
                err = 0: On Error Resume Next
                cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, .ColIndex("供应商")), .TextMatrix(lngRow, .ColIndex("卫材信息")), .TextMatrix(lngRow, .ColIndex("规格|产地"))), "K" & lng供应商ID & "_" & lngTemp
                If err <> 0 Then
                    '累计差价调整额
                    dbl差价调整额 = Val(cllData("K" & lng供应商ID & "_" & lngTemp)(2)) + dbl差价调整额
                    cllData.Remove "K" & lng供应商ID & "_" & lngTemp
                     err = 0: On Error GoTo ErrHand:
                    cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, .ColIndex("供应商")), .TextMatrix(lngRow, .ColIndex("卫材信息")), .TextMatrix(lngRow, .ColIndex("规格|产地"))), "K" & lng供应商ID & "_" & lngTemp
                   
                End If
                On Error GoTo ErrHand:
            End If
        Next
        
        If chk自动计算.Value = 1 Then
            '需要自动计算相关的应付变动记录
            For i = 1 To cllData.Count
                With vsPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("卫材信息")))
                        lng供应商ID = Val(.Cell(flexcpData, lngRow, .ColIndex("供应商")))
                        If lngTemp = Val(cllData(i)(0)) _
                            And lng供应商ID = Val(cllData(i)(1)) Then
                            '卫材及供应商相同,清空相关的值
                            .TextMatrix(lngRow, .ColIndex("发票金额")) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                            .Cell(flexcpData, lngRow, .ColIndex("发票金额")) = Val(cllData(i)(2))
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '需要增加该项供应商的物资
                        If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("卫材信息"))) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, .ColIndex("供应商")) = cllData(i)(3)
                        .Cell(flexcpData, lngRow, .ColIndex("供应商")) = cllData(i)(1)
                        .TextMatrix(lngRow, .ColIndex("卫材信息")) = cllData(i)(4)
                        .Cell(flexcpData, lngRow, .ColIndex("卫材信息")) = cllData(i)(0)
                        .TextMatrix(lngRow, .ColIndex("规格|产地")) = cllData(i)(5)
                        .TextMatrix(lngRow, .ColIndex("发票金额")) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                        .Cell(flexcpData, lngRow, .ColIndex("发票金额")) = Val(cllData(i)(2))
                    End If
                End With
            Next
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function IsValied应付信息() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查应付信息是否否正确
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-10 10:28:17
    '-----------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, str发票日期 As String, dbl发票金额 As Double, lng系数 As Long, j As Long
 
    
    IsValied应付信息 = False
    If m调价方式 = T_售价调价 Then IsValied应付信息 = True: Exit Function
    If chk应付.Value <> 1 Then IsValied应付信息 = True: Exit Function
    
    With vsPay
        For j = 1 To .Rows - 1
            If Val(.Cell(flexcpData, j, .ColIndex("卫材信息"))) <> 0 Then
                '看是否有此卫生材料库存变动情况
                str发票号 = Trim(.TextMatrix(j, .ColIndex("发票号")))
                str发票日期 = Trim(.TextMatrix(j, .ColIndex("发票日期")))
                dbl发票金额 = Val(.TextMatrix(j, .ColIndex("发票金额")))
                If str发票日期 <> "" Then
                    str发票日期 = zlCheckIsDate(str发票日期, "发票日期")
                    If str发票日期 = "" Then
                        tbPage.Item(1).Selected = True
                        .Row = j: .Col = .ColIndex("发票日期")
                        zlControl.ControlSetFocus vsPay, True
                        Exit Function
                    End If
                Else
                    ShowMsgBox "在第" & j & "行中的发票日期未输入，请检查!"
                    If tbPage.Item(1).Visible Then tbPage.Item(1).Selected = True
                    .Row = j: .Col = .ColIndex("发票日期")
                    zlControl.ControlSetFocus vsPay, True
                    Exit Function
                End If
                
                If zlCommFun.StrIsValid(str发票号, 100, 0, "发票号") = False Then
                        If tbPage.Item(1).Visible Then tbPage.Item(1).Selected = True
                        .Row = j: .Col = .ColIndex("发票号")
                        zlControl.ControlSetFocus vsPay, True
                        Exit Function
                End If
                If str发票号 = "" Then
                    ShowMsgBox "在第" & j & "行中的发票号未输入，请检查!"
                   If tbPage.Item(1).Visible Then tbPage.Item(1).Selected = True
                    .Row = j: .Col = .ColIndex("发票号")
                    zlControl.ControlSetFocus vsPay, True
                    Exit Function
                End If
            End If
        Next
    End With
    IsValied应付信息 = True
End Function
Private Function Check成本价(ByVal lng材料ID As Long, ByVal dbl成本价 As Double, Optional ByRef dblOut成本价 As Double) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查成本价是否有效
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-10 14:27:56
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    dblOut成本价 = dbl成本价
    With vsPrice
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("品名"))) = lng材料ID Then
                If Val(.TextMatrix(lngRow, .ColIndex("现指导售价"))) < dbl成本价 Then
                    dblOut成本价 = Val(.TextMatrix(lngRow, .ColIndex("现指导售价")))
                    If MsgBox("注意：" & vbCrLf & "    卫生材料“" & .TextMatrix(lngRow, .ColIndex("品名")) & "”" & vbCrLf & _
                        "的成本价(" & Format(dbl成本价, mFMT.FM_成本价) & ")大于了指导零售价(" & Format(Val(.TextMatrix(lngRow, .ColIndex("现指导售价"))), mFMT.FM_零售价) & ")" & _
                        "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        Check成本价 = True
                        Exit Function
                    Else
                        Check成本价 = False
                        Exit Function
                    End If
                Else
                    Check成本价 = True: Exit Function
                End If
            End If
        Next
    End With
    '未找到相关的调价信息，也返回true
    Check成本价 = True
End Function





