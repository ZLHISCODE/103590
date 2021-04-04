VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanArrange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "挂号计划安排"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10950
   Icon            =   "frmRegistPlanArrange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   10950
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   9540
      Left            =   540
      ScaleHeight     =   9540
      ScaleWidth      =   8985
      TabIndex        =   32
      Top             =   120
      Width           =   8985
      Begin VB.OptionButton opt生效时间 
         Caption         =   "立即执行"
         Height          =   360
         Index           =   0
         Left            =   1110
         TabIndex        =   27
         Top             =   7860
         Width           =   1170
      End
      Begin VB.OptionButton opt生效时间 
         Caption         =   "指定时间"
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   28
         Top             =   7935
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6345
         TabIndex        =   39
         Top             =   8715
         Width           =   2370
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6345
         TabIndex        =   38
         Top             =   8325
         Width           =   2370
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1110
         TabIndex        =   37
         Top             =   8715
         Width           =   2370
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1110
         TabIndex        =   36
         Top             =   8265
         Width           =   2370
      End
      Begin VB.Frame Frame2 
         Caption         =   "应诊时间"
         Height          =   2010
         Left            =   60
         TabIndex        =   35
         Top             =   1635
         Width           =   8685
         Begin VB.TextBox txt限约 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5145
            MaxLength       =   5
            TabIndex        =   18
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox txt限号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3030
            MaxLength       =   5
            TabIndex        =   16
            Top             =   270
            Width           =   1215
         End
         Begin VB.ComboBox cbo天 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   270
            Width           =   1365
         End
         Begin VB.OptionButton opt周 
            Caption         =   "每周(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   19
            Top             =   630
            Width           =   930
         End
         Begin VB.OptionButton opt天 
            Caption         =   "每天(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   13
            Top             =   285
            Width           =   960
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   1275
            Left            =   1200
            TabIndex        =   20
            Top             =   600
            Width           =   7440
            _cx             =   13123
            _cy             =   2249
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
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
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
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanArrange.frx":06EA
            ScrollTrack     =   0   'False
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "限约"
            Height          =   180
            Left            =   4710
            TabIndex        =   17
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "限号"
            Height          =   180
            Left            =   2595
            TabIndex        =   15
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "应诊诊室:"
         Height          =   3850
         Left            =   60
         TabIndex        =   34
         Top             =   3840
         Width           =   8670
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   3390
            Left            =   150
            TabIndex        =   25
            Top             =   300
            Width           =   8415
            _cx             =   14843
            _cy             =   5980
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
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
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
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
         Begin VB.OptionButton opt分诊 
            Caption         =   "平均分诊"
            Height          =   180
            Index           =   3
            Left            =   4335
            TabIndex        =   24
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "动态分诊"
            Height          =   180
            Index           =   2
            Left            =   3180
            TabIndex        =   23
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "指定诊室"
            Height          =   180
            Index           =   1
            Left            =   2010
            TabIndex        =   22
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "不分诊"
            Height          =   180
            Index           =   0
            Left            =   1020
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   6345
         TabIndex        =   30
         Top             =   7890
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   134545411
         CurrentDate     =   401769
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   3360
         TabIndex        =   29
         Top             =   7890
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   134545411
         CurrentDate     =   38091
      End
      Begin VB.Frame Frame1 
         Caption         =   "基本信息"
         Height          =   1455
         Left            =   60
         TabIndex        =   33
         Top             =   105
         Width           =   8670
         Begin VB.CheckBox chkAppoint 
            Caption         =   "允许预约"
            Height          =   300
            Left            =   6195
            TabIndex        =   12
            Top             =   1027
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.CheckBox chk序号控制 
            Caption         =   "序号控制"
            Height          =   255
            Left            =   1750
            TabIndex        =   2
            Top             =   293
            Width           =   1095
         End
         Begin VB.ComboBox cbo号类 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   2595
         End
         Begin VB.CheckBox chk病案 
            Caption         =   "挂号时必须建病案"
            Height          =   195
            Left            =   3420
            TabIndex        =   11
            Top             =   1080
            Width           =   1845
         End
         Begin VB.ComboBox cbo科室 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   660
            Width           =   2400
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   660
            TabIndex        =   10
            Top             =   1035
            Width           =   2400
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   675
            Width           =   2580
         End
         Begin VB.TextBox txt号别 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            MaxLength       =   5
            TabIndex        =   1
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "号类"
            Height          =   180
            Left            =   3405
            TabIndex        =   3
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "医生"
            Height          =   180
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "项目"
            Height          =   180
            Left            =   3420
            TabIndex        =   7
            Top             =   750
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "科室"
            Height          =   180
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "号别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   210
            TabIndex        =   0
            Top             =   330
            Width           =   390
         End
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   7
         Left            =   5715
         TabIndex        =   47
         Top             =   7950
         Width           =   180
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "计划时间"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   7935
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "审核时间"
         Height          =   180
         Index           =   3
         Left            =   5535
         TabIndex        =   43
         Top             =   8775
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "审核人"
         Height          =   180
         Index           =   2
         Left            =   5715
         TabIndex        =   42
         Top             =   8385
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "安排时间"
         Height          =   180
         Index           =   1
         Left            =   345
         TabIndex        =   41
         Top             =   8775
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "安排人"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   40
         Top             =   8385
         Width           =   540
      End
   End
   Begin VB.PictureBox picTimeSet 
      BorderStyle     =   0  'None
      Height          =   7320
      Left            =   60
      ScaleHeight     =   7320
      ScaleWidth      =   8580
      TabIndex        =   50
      Top             =   1500
      Width           =   8580
      Begin VB.CommandButton cmdAuto 
         Caption         =   "自动计算(&A)"
         Height          =   350
         Left            =   5415
         TabIndex        =   56
         ToolTipText     =   "通过输入的限号数,自动分配时间间隔进行计算"
         Top             =   30
         Width           =   1150
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   6810
         TabIndex        =   68
         ToolTipText     =   "点击重新计算时段"
         Top             =   30
         Width           =   1150
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "全清(&D)"
         Height          =   350
         Left            =   8250
         TabIndex        =   67
         ToolTipText     =   "点击重新计算时段"
         Top             =   30
         Width           =   1150
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   3540
         Index           =   0
         Left            =   3420
         ScaleHeight     =   3540
         ScaleWidth      =   2535
         TabIndex        =   63
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Frame fra应用于 
         Caption         =   "应用于…"
         Height          =   615
         Left            =   0
         TabIndex        =   58
         Top             =   6720
         Width           =   7755
         Begin VB.OptionButton opt所有 
            Caption         =   "所有号别"
            Height          =   255
            Left            =   5685
            TabIndex        =   62
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "本科室(内科)"
            Height          =   255
            Left            =   3870
            TabIndex        =   61
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt应用于 
            Caption         =   "本号码"
            Height          =   255
            Index           =   0
            Left            =   795
            TabIndex        =   59
            Top             =   255
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton opt应用于 
            Caption         =   "本医生(张三)"
            Height          =   255
            Index           =   1
            Left            =   2100
            TabIndex        =   60
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtTimeOut 
         Height          =   300
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "10"
         Top             =   60
         Width           =   465
      End
      Begin VB.CommandButton cmd设置时段 
         Caption         =   "辅助计算(&F)"
         Height          =   350
         Left            =   2220
         TabIndex        =   53
         ToolTipText     =   "点击重新计算时段"
         Top             =   30
         Width           =   1150
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "其他辅助计算(&T)"
         Height          =   350
         Left            =   3675
         TabIndex        =   55
         ToolTipText     =   "点击重新计算时段"
         Top             =   30
         Width           =   1515
      End
      Begin MSComCtl2.UpDown udTime 
         Height          =   300
         Left            =   1635
         TabIndex        =   51
         Top             =   60
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtTimeOut"
         BuddyDispid     =   196650
         OrigLeft        =   2025
         OrigTop         =   3
         OrigRight       =   2280
         OrigBottom      =   348
         Max             =   1440
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl tbSubPage 
         Height          =   4875
         Left            =   285
         TabIndex        =   64
         Top             =   1410
         Width           =   2535
         _Version        =   589884
         _ExtentX        =   4471
         _ExtentY        =   8599
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTime 
         Height          =   5475
         Index           =   0
         Left            =   2175
         TabIndex        =   57
         Top             =   1185
         Width           =   5100
         _cx             =   8996
         _cy             =   9657
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegistPlanArrange.frx":07D0
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
         Begin VB.CommandButton cmd预约 
            Caption         =   "预"
            Height          =   255
            Index           =   0
            Left            =   2685
            TabIndex        =   66
            Top             =   2535
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmd删除 
            Caption         =   "删"
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   65
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "时间间隔(分)"
         Height          =   180
         Left            =   60
         TabIndex        =   54
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9360
      TabIndex        =   45
      Top             =   1425
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9360
      TabIndex        =   31
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   9360
      TabIndex        =   44
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CheckBox chk立即生效 
      Caption         =   "立即生效"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   49
      Top             =   120
      Width           =   1650
   End
   Begin VB.CheckBox chk立即审核 
      Caption         =   "保存后立即审核"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   48
      Top             =   80
      Width           =   1650
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   780
      Left            =   -120
      TabIndex        =   46
      Top             =   0
      Width           =   9015
      _Version        =   589884
      _ExtentX        =   15901
      _ExtentY        =   1376
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmRegistPlanArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstr计划ID As String, mlng安排ID As Long, mblnSucces As Boolean, mblnFirst As Boolean
Private mlngModule As Long, mstrPrivs As String
Private mblnActive As Boolean
Private Enum mPageIndex
    EM_计划 = 0
    EM_时段 = 1
End Enum
Private mrsRegOldData As ADODB.Recordset '本地数据集保存,原始挂号安排
Private mrsRegNewData As ADODB.Recordset '本地数据集保存 重新设置后的安排
Private mrsRegHistory As ADODB.Recordset '历次挂号的数据集
Private mrs上班时间段 As ADODB.Recordset
Private mrsLongPlan As ADODB.Recordset '长期计划
Private mdatBegin As Date
Private mdatEnd As Date
Private mdatOriBegin As Date
Private mdatOriEnd As Date

Public Enum gPlanEditType
    EM_安排_增加 = 0
    EM_安排_修改
    EM_安排_查阅
    EM_计划_增加 = 11
    EM_计划_修改
    EM_计划_查阅
End Enum
Private mPlanEditType As gPlanEditType

Private mblnChangeByCode As Boolean
Public Enum mRegEditType
    ed_计划安排 = 0
    Ed_安排修改 = 1
    Ed_安排删除 = 2
    Ed_安排审核 = 3
    Ed_安排取消 = 4
    ed_安排查阅 = 5
End Enum
Private Enum midxTxt
    idx_安排人 = 0
    idx_安排时间 = 1
    idx_审核人 = 2
    idx_审核时间 = 3
End Enum
'对外上班时间
Private Type t_上班时间
  dat_上午上班 As Date
  dat_上午下班 As Date
  dat_下午上班 As Date
  dat_下午下班 As Date
End Type
Private t_时间 As t_上班时间
Private mEditType As mRegEditType
Private mstr原排班 As String '"周一,上午;周二,下午;..."
Private mstr科室ID As String
Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
Private mrsDoctor As ADODB.Recordset

Private Type PlanInfo               '安排改变需要对比的信息
    str应诊时段     As String
    str排班         As String       '排班信息
    str限号         As String       '限号信息
    bln序号         As Boolean      '是否序号控制
    bln时间段       As Boolean      '是否设置了时间段
End Type

Private Type TimeSet
    bln序号控制 As Boolean
    blnIsInit As Boolean
    lngSelIndex As Long
    blnChange As Boolean
    str安排 As String
    str应诊时段 As String
    rsAssign As ADODB.Recordset
    rsHistory As ADODB.Recordset
    rsRegPlan As ADODB.Recordset
    blnNotBrush As Boolean
    lng计划ID As Long
    lng安排ID As Long
    blnOnChange As Boolean
End Type

Private mTimeSet As TimeSet
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther
Attribute mfrmOtherCalc.VB_VarHelpID = -1
Private mblnSaveMinorChange As Boolean

Private mPlanInfo As PlanInfo '新增时用于保存原始安排信息  修改时 保存原始的计划信息 在保存时 比较相应信息
Private Enum mPgIndex
    Pg_计划安排 = 1
    Pg_计划时段 = 2
End Enum
Private mbln自动默认限约数 As Boolean '45519 自动默认限约数控制
Private mbln限制修改 As Boolean '是否限制修改
Private mstr已约限制 As String '保存那些排班限制更改
Private mdtMinCustom As Date '如果存在预约号,最小的时间

Private Sub LoadTimeSetControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载页控件
    '编制:刘兴洪
    '日期:2012-06-15 13:33:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    For i = 1 To 6
        Load picPage(i): Load vsTime(i)
        Load cmd预约(i): Load cmd删除(i)
       ' cmd预约(i).Visible = True
        Set cmd预约(i).Container = vsTime(i)
        Set cmd删除(i).Container = vsTime(i)
        'cmd删除(i).Visible = True
        picPage(i).Visible = True: vsTime(i).Visible = True
        Set vsTime(i).Container = picPage(i)
    Next
    Set vsTime(0).Container = picPage(0)
    Call LoadTimeSet
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadTimeSet() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载页
    '编制:刘兴洪
    '日期:2012-06-15 13:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim strTemp As String
    On Error GoTo errHandle
    
    tbSubPage.RemoveAll
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        Set ObjItem = tbSubPage.InsertItem(i + 1, strTemp, picPage(i).Hwnd, 0)
        ObjItem.Tag = strTemp
    Next
     With tbSubPage
        tbSubPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    LoadTimeSet = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_计划安排, "计划安排", picBack.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_计划安排

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_计划时段, "时段设置", picTimeSet.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_计划时段
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowCard(ByVal mfrmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As mRegEditType, Optional lng安排ID As Long, Optional ByVal str计划Id As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示所要修改的计划安排
    '入参:mfrmMain-调用的主窗口
    '     lngModule-模块号
    '     strPrivs-权限串
    '     EditType-编辑的类型
    '     lng安排ID-挂号安排ID.
    '     str计划Id-安排时为空,否则,否则为指定的计划ID
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-14 14:31:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs: mstr计划ID = str计划Id: mblnSucces = False: mlng安排ID = lng安排ID
    Me.Show 1, mfrmMain
    ShowCard = mblnSucces
End Function

Private Function LoadData(Optional blnNoChangeTime As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载计划安排数据信息
    '编制:刘兴洪
    '日期:2009-09-14 14:40:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp          As New ADODB.Recordset
    Dim rsDept          As New ADODB.Recordset
    Dim strSQL          As String
    Dim i               As Long
    Dim j               As Long
    Dim rs限号          As ADODB.Recordset
    Dim strTemp         As String
    Dim bln每周         As Boolean
    Dim bln限号         As Boolean
    Dim str限号         As String
    Dim bln限约         As Boolean
    Dim str限约         As String
    Dim dtSys           As Date
    Dim dtTmp           As Date
    Dim blnExitFor      As Boolean
    Err = 0: On Error GoTo Errhand:
    
    '加载安排
    If mEditType = ed_计划安排 Then
       '新增安排
        strSQL = " " & _
        "   Select A.Id as 安排ID,0 as 计划ID,A.号类,A.项目ID as 计划项目ID,   A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id ,   " & _
        "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,A.默认时段间隔 As 默认时段间隔, " & _
        "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室,NULL　as 生效时间,'3000-01-01 00:00:00' as 失效时间 ," & _
        "           NULL as 安排人,NULL as 安排时间,NULL 审核人,NULL 审核时间" & _
        "   From 挂号安排 A,收费项目目录 B,挂号安排计划 C,部门表 D " & _
        "   Where A.Id=C.安排ID(+) And A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
        "         And A.Id=[1]"
    Else
         '非新增
        strSQL = " " & _
        "Select a.安排id, a.Id As 计划id, a.号类, 计划项目id, a.号码, a.科室id, a.项目id, a.医生姓名, a.医生id,   a.周日, a.周一, a.周二, a.周三," & _
        "  a.周四, a.周五, a.周六, a.病案必须, a.分诊方式, a.序号控制, a.开始时间, a.终止时间, b.名称 As 项目, d.名称 As 科室, 生效时间, a.失效时间, a.安排人, a.安排时间," & _
        " a.审核人 , 审核时间,A.默认时段间隔 As 默认时段间隔" & _
        " From (Select c.安排id, c.Id, a.号类, Nvl(c.项目id, a.项目id) As 计划项目id, c.号码, a.科室id, Nvl(c.项目id, a.项目id) As 项目id, C.医生姓名, C.医生id," & _
        "       c.周日, c.周一, c.周二, c.周三, c.周四, c.周五, c.周六, a.病案必须, c.分诊方式, c.序号控制, a.开始时间, a.终止时间, Nvl(C.默认时段间隔,5) as 默认时段间隔," & _
        "      To_Char(c.生效时间, 'yyyy-mm-dd hh24:mi:ss') As 生效时间, To_Char(c.失效时间, 'yyyy-mm-dd hh24:mi:ss') As 失效时间, c.安排人," & _
        "      To_Char(c.安排时间, 'yyyy-mm-dd hh24:mi:ss') As 安排时间, c.审核人, To_Char(c.审核时间, 'yyyy-mm-dd hh24:mi:ss') As 审核时间" & _
        " From 挂号安排 A, 挂号安排计划 C " & _
        " Where a.Id = c.安排id) A, 收费项目目录 B, 部门表 D " & _
        " Where a.项目id = b.Id(+) And a.科室id = d.Id(+) " & _
        "  and a.id=[2]"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
    If rsTemp.EOF Then
        If mEditType = ed_计划安排 Then
            MsgBox "注意:" & vbCrLf & _
                   "    挂号安排可能已经被他人删除,不能再进行计划安排", vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "注意:" & vbCrLf & _
                   "    挂号计划安排可能已经被他人删除,请检查!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    If mEditType = ed_计划安排 Then
        strSQL = "Select 限制项目,限号数,  限约数 From  挂号安排限制 where 安排ID=[1]       "
    Else
        strSQL = "Select 限制项目,限号数,  限约数 From  挂号计划限制 where 计划ID=[2]       "
    End If
    Set rs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
    
    chkAppoint.Value = 0
    Do While Not rs限号.EOF
        If IsNull(rs限号!限约数) Then
            mblnChangeByCode = True
            chkAppoint.Value = 1
            mblnChangeByCode = False
            Exit Do
        Else
            If Val(Nvl(rs限号!限约数)) <> 0 Then
                mblnChangeByCode = True
                chkAppoint.Value = 1
                mblnChangeByCode = False
                Exit Do
            End If
        End If
        rs限号.MoveNext
    Loop
    If rs限号.RecordCount <> 0 Then rs限号.MoveFirst
    
    '检查其他一些功能
    If mEditType = Ed_安排修改 And Nvl(rsTemp!审核时间) <> "" Then
        mblnSaveMinorChange = True
    Else
        mblnSaveMinorChange = False
    End If
    If mEditType = Ed_安排删除 And Nvl(rsTemp!审核时间) <> "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    挂号计划安排已经被他人审核,不能再进行计划删除！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
    End If
    
    If mEditType = Ed_安排审核 And Nvl(rsTemp!审核时间) <> "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    挂号计划安排已经被他人审核,不能再进行计划审核！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
    End If

    If mEditType = Ed_安排取消 And Nvl(rsTemp!审核时间) = "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    挂号计划安排已经被他人取消审核,不能再进行计划审核取消！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
    End If
    
    '加载数据到控件中
    txt号别.Text = Nvl(rsTemp!号码)
    cbo号类.AddItem Nvl(rsTemp!号类): cbo号类.ListIndex = cbo号类.NewIndex
    chk序号控制.Value = IIf(Val(Nvl(rsTemp!序号控制)) = 1, 1, 0)
    
    mTimeSet.bln序号控制 = Val(rsTemp!序号控制) = 1
    
    '获取的安排或者计划是否序号控制
    mPlanInfo.bln序号 = IIf(Val(Nvl(rsTemp!序号控制)) = 1, True, False)
    
    chk病案.Value = IIf(Val(Nvl(rsTemp!病案必须)) = 1, 1, 0)
    
    
    txtEdit(midxTxt.idx_安排人).Text = Nvl(rsTemp!安排人)
    txtEdit(midxTxt.idx_安排时间).Text = Nvl(rsTemp!安排时间)
    If mEditType = ed_计划安排 Then
        txtEdit(midxTxt.idx_安排人) = UserInfo.姓名
        txtEdit(midxTxt.idx_安排时间) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    txtEdit(midxTxt.idx_审核人) = Nvl(rsTemp!审核人)
    txtEdit(midxTxt.idx_审核时间) = Nvl(rsTemp!审核时间)
    If mEditType = Ed_安排审核 Then
        txtEdit(midxTxt.idx_审核人) = UserInfo.姓名
        txtEdit(midxTxt.idx_审核时间) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    
    With cbo科室
        .AddItem Nvl(rsTemp!科室): .ItemData(.NewIndex) = Val(Nvl(rsTemp!科室ID)): .ListIndex = .NewIndex
    End With
    With cboItem
         If mEditType = Ed_安排修改 Or mEditType = ed_计划安排 Then
            zlControl.CboSetText cboItem, rsTemp!项目
        Else
            .AddItem Nvl(rsTemp!项目): .ItemData(.NewIndex) = Val(Nvl(rsTemp!项目ID)): .ListIndex = .NewIndex
        End If
         
    End With
    With cboDoctor
       If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
          LoadDoctor
          zlControl.CboLocate cboDoctor, Nvl(rsTemp!医生姓名)
'          cboDoctor.Text = Nvl(rsTemp!医生姓名)
        Else
            .AddItem Nvl(rsTemp!医生姓名): .ItemData(.NewIndex) = Val(Nvl(rsTemp!医生ID)): .ListIndex = .NewIndex
        End If
    End With
   ' mstr已约限制 = Get已约限制(mlng安排ID)
    'mbln限制修改 = CheckExistsBooking(Nvl(rsTemp!号码), mdtMinCustom)
    
 
    '加载原始数据到数据集
     With mrsRegOldData
        Set mrsRegOldData = New ADODB.Recordset
        mrsRegOldData.Fields.Append "ID", adBigInt, 18
        mrsRegOldData.Fields.Append "限制项目", adVarChar, 20
        mrsRegOldData.Fields.Append "限号数", adBigInt, 10
        mrsRegOldData.Fields.Append "限约数", adBigInt, 18
        mrsRegOldData.Fields.Append "序号控制", adBigInt, 18
        mrsRegOldData.CursorLocation = adUseClient
        mrsRegOldData.LockType = adLockOptimistic
        mrsRegOldData.CursorType = adOpenStatic
        mrsRegOldData.Open


        rs限号.Filter = 0
        If rs限号.RecordCount > 0 Then rs限号.MoveFirst
        Do While Not rs限号.EOF
            With mrsRegOldData
                .AddNew
                !ID = Val(mstr计划ID)
                !限制项目 = Nvl(rs限号!限制项目)
                !限号数 = Val(Nvl(rs限号!限号数))
                !限约数 = Val(Nvl(rs限号!限约数))
                !序号控制 = Val(Nvl(rsTemp!序号控制))
                .Update
            End With
            rs限号.MoveNext
        Loop
    End With
    
    If blnNoChangeTime = False Then
        '-------------------------------
        dtSys = zlDatabase.Currentdate
        If mEditType = Ed_安排修改 Or mEditType = ed_计划安排 Then
           dtpBegin.MinDate = dtSys
             '默认下一天生效
           dtSys = DateAdd("d", 1, Format(dtSys, "yyyy-mm-dd"))
        End If
        If IsNull(rsTemp!生效时间) Then
            dtpBegin.Value = Format(zlGetNextWeekDate, "yyyy-mm-dd HH:MM:SS")
        Else
            If mEditType = Ed_安排修改 Or mEditType = ed_计划安排 Then
                '59754
                dtpBegin.Value = IIf(Format(dtSys, "yyyy-mm-dd HH:MM:SS") > Format(CDate(Nvl(rsTemp!生效时间, "1900-01-01")), "yyyy-mm-dd HH:MM:SS"), dtSys, CDate(Nvl(rsTemp!生效时间)))
            Else
                dtpBegin.Value = CDate(Nvl(rsTemp!生效时间))
            End If
        End If
        dtpEndDate.Value = CDate(Nvl(rsTemp!失效时间, "3000-01-01"))
        mdatOriBegin = CDate(Nvl(rsTemp!生效时间, "2000-01-01"))
        mdatOriEnd = CDate(Nvl(rsTemp!失效时间, "3000-01-01"))
        
        If mEditType = ed_计划安排 Then
            strSQL = "Select nvl(生效时间,Sysdate) as 生效时间 ,nvl(失效时间,to_date('3000-01-01','yyyy-mm-dd')) as 失效时间 From 挂号安排计划 where ID=(Select Max(ID) From 挂号安排计划 where 安排ID=[1]) "
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
            If Not rsDept.EOF Then
                If Format(rsDept!失效时间, "yyyy-mm-dd") < "3000-01-01" Then
                    '上一条计划的终止日期,就是本条的生效时间
                    dtTmp = CDate(Format(rsDept!失效时间, "yyyy-mm-dd HH:MM:SS"))
                    '59754
                    dtpBegin.Value = IIf(Format(dtSys, "yyyy-mm-dd HH:MM:SS") > Format(dtTmp, "yyyy-mm-dd HH:MM:SS"), dtSys, dtTmp)
                Else '以上一条的生效时间的下一周为准
                    dtTmp = zlGetNextWeekDate(Format(rsDept!生效时间, "yyyy-mm-dd HH:MM:SS"))
                    '新增加计划时,以下一个星期开始计算
                    dtSys = zlGetNextWeekDate(Format(DateAdd("d", -1, dtSys), "yyyy-mm-dd"))
                     '59754
                    dtpBegin.Value = IIf(Format(dtSys, "yyyy-mm-dd HH:MM:SS") > Format(dtTmp, "yyyy-mm-dd HH:MM:SS"), dtSys, dtTmp)
                End If
            End If
            
            strSQL = "Select 号表ID as ID,门诊诊室　From 挂号安排诊室 Where 号表ID=[1]"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
        Else
            strSQL = "Select 计划ID as ID,门诊诊室　From 挂号计划诊室 Where 计划ID=[2]"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
        End If
    End If
    
    Call LoadLongPlan
    Call LoadRegHistory
    '---------------------------------------------------
    '判断 每日安排 限号数 限约数 等是否一致
    '---------------------------------------------------
    rs限号.Filter = 0
    If rs限号.RecordCount > 0 Then rs限号.MoveFirst
    bln每周 = Nvl(rsTemp!周日) <> Nvl(rsTemp!周一) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周二) _
        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周三) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周四) _
        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周五) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周六)
        
    If bln每周 = False Then
             rs限号.Filter = "限制项目='周日'"
             If Not rs限号.EOF Then
                str限号 = Nvl(rs限号!限号数)
                str限约 = Nvl(rs限号!限约数)
             End If
            For i = 1 To 6
                strTemp = Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                rs限号.Filter = "限制项目='" & "周" & strTemp & "'"
                If Not rs限号.EOF Then
                    bln限号 = Nvl(rs限号!限号数) = str限号
                    bln限约 = Nvl(rs限号!限约数) = str限约
                    If bln限约 = False Or bln限号 = False Then Exit For
                End If
            Next
          bln每周 = True
         If bln限号 And bln限约 Then bln每周 = False
     
    End If
    mstr原排班 = ""
    If bln每周 Or mrsRegHistory.RecordCount > 0 Then
        '每周
        opt周.Value = True:
        txt限号.Enabled = False: txt限约.Enabled = False
        With vsPlan
            For i = 1 To 7
                '不知什么原因,将.colkey(i)的日,要更改成日日了.
                strTemp = "周" & Replace(.ColKey(i), "日日", "日")
                .TextMatrix(1, i) = Nvl(rsTemp.Fields(strTemp))
                mstr原排班 = mstr原排班 & ";" & strTemp & "," & Nvl(rsTemp.Fields(strTemp)) '"周一,上午;周二,下午;..."
                rs限号.Filter = "限制项目='" & strTemp & "'"
                If Not rs限号.EOF Then
                    .TextMatrix(2, i) = Nvl(rs限号!限号数)
                    If IsNull(rs限号!限约数) Then
                        .TextMatrix(3, i) = ""
                    Else
                        If Val(Nvl(rs限号!限约数)) = 0 Then
                            .TextMatrix(3, i) = "0"
                        Else
                            .TextMatrix(3, i) = Nvl(rs限号!限约数)
                        End If
                    End If
                End If
            Next
            If mstr原排班 <> "" Then mstr原排班 = Mid(mstr原排班, 2)
        End With
    Else
        '每天
        opt天.Value = True:  cbo天.ListIndex = cbo.FindIndex(cbo天, Nvl(rsTemp!周日), True): cbo天.Enabled = True
        mstr原排班 = Nvl(rsTemp!周日)
        If rs限号.RecordCount <> 0 Then rs限号.MoveFirst
        If rs限号.EOF = False Then
            txt限号.Text = Nvl(rs限号!限号数)
            If IsNull(rs限号!限约数) Then
                txt限约.Text = ""
            Else
                If Val(Nvl(rs限号!限约数)) = 0 Then
                    txt限约.Text = "0"
                Else
                    txt限约.Text = Nvl(rs限号!限约数)
                End If
            End If
        End If
        If chkAppoint.Value = 0 Then
            txt限约.Enabled = False
            txt限约.Text = ""
        Else
            txt限约.Enabled = True
        End If
    End If
    
     '------------------------------
    '获取修改或者新增前的 时间段和 限号数
    '用于在保存时 对比限号限约、序号控制以及时间段是否发生了变化
    '如果发生了变化则需要提示  操作员重新设置时段信息
    '------------------------------
   mPlanInfo.str排班 = ""
   mPlanInfo.str限号 = ""
   mPlanInfo.str应诊时段 = ""
    If bln每周 = False Or mrsRegHistory.RecordCount > 0 Then
        For i = 1 To 7
             mPlanInfo.str排班 = mPlanInfo.str排班 & ",'" & Trim(cbo天.Text) & "'"
             mPlanInfo.str应诊时段 = mPlanInfo.str应诊时段 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六") & "-" & Trim(cbo天.Text)
             mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
             mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(txt限号.Text) & "," & txt限约.Text
             mPlanInfo.str应诊时段 = mPlanInfo.str应诊时段 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六") & "-" & Trim(vsPlan.TextMatrix(1, i))
             
        Next
    Else
        For i = 1 To vsPlan.Cols - 1
            mPlanInfo.str排班 = mPlanInfo.str排班 & ",'" & Trim(vsPlan.TextMatrix(1, i)) & "'"
            If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
                mPlanInfo.str应诊时段 = mPlanInfo.str应诊时段 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六") & "-" & Trim(vsPlan.TextMatrix(1, i))
                mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                     mPlanInfo.str限号 = mPlanInfo.str限号 & ",0,0"
                Else
                     mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Trim(vsPlan.TextMatrix(3, i))
                End If
            End If
        Next
    End If
    If mPlanInfo.str限号 <> "" Then mPlanInfo.str限号 = Mid(mPlanInfo.str限号, 2)
    If mPlanInfo.str应诊时段 <> "" Then mPlanInfo.str应诊时段 = Mid(mPlanInfo.str应诊时段, 2)
    
    Select Case Val(Nvl(rsTemp!分诊方式))     '0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
        Case 0  '"不分诊"
            opt分诊(0).Value = True
        Case 1  ' "指定诊室"
            opt分诊(1).Value = True
        Case 2 '"动态分诊"
            opt分诊(2).Value = True
        Case 3 ' "平均分诊"
            opt分诊(3).Value = True
    End Select
    
    '71253 李南春 2014-04-15 14:23:10 将listView 替换为vsflexGrid
    If blnNoChangeTime = False Then
    With vsDept
        blnExitFor = False
        Do While Not rsDept.EOF
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If Nvl(rsDept!门诊诊室) = .TextMatrix(j, i) Then
                        .Cell(flexcpChecked, j, i) = 1
                        blnExitFor = True
                        Exit For
                    End If
                Next
                If blnExitFor Then blnExitFor = False: Exit For
            Next
            rsDept.MoveNext
        Loop
    End With
    rsDept.Close
    End If
    
    If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then mPlanInfo.bln时间段 = Check时段()
    If mrsRegHistory.RecordCount > 0 Then opt天.Enabled = False
    If mEditType = Ed_安排删除 Then
        picTimeSet.Enabled = False
    End If
    mdatBegin = dtpBegin
    mdatEnd = dtpEndDate
    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub chkAppoint_Click()
    Dim i As Integer
    If mblnChangeByCode Then Exit Sub
    If chkAppoint.Value = 0 Then
        If opt天.Value = True Then
            txt限约.Enabled = False
            txt限约.BackColor = &H8000000F
        End If
        txt限约.Text = ""
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(3, i) = ""
        Next i
    Else
        If opt天.Value = True Then
            txt限约.Enabled = True
            txt限约.BackColor = vbWhite
        End If
        If Val(txt限约.Text) = 0 Then txt限约.Text = ""
        For i = 1 To vsPlan.Cols - 1
            If Val(vsPlan.TextMatrix(3, i)) = 0 Then vsPlan.TextMatrix(3, i) = ""
        Next i
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载初始化数据
    '编制:刘兴洪
    '日期:2009-09-14 15:50:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, i As Long
    Dim intRow As Integer
    Dim lngColsWidth As Long
    
    Err = 0: On Error GoTo Errhand:

    strSQL = "Select '    ' 时间段 From dual Union All  " & _
             " Select 时间段 From 时间段"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        cbo天.AddItem rsTemp!时间段
        rsTemp.MoveNext
    Loop
    
    With vsPlan
        .ColComboList(1) = .BuildComboList(rsTemp, "时间段")
        For i = 2 To .Cols - 1
            .ColComboList(i) = .ColComboList(1)
        Next
        .Tag = .ColComboList(1)
    End With
 
    
    '门诊诊室
    strSQL = "Select 编码,名称　From 门诊诊室 Where (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    '71253 李南春 2014-04-15 14:23:10 诊室名称显示不全
    vsDept.Clear
    If rsTemp.RecordCount <> 0 Then
        With vsDept
            Do While Not rsTemp.EOF
                If intRow = .Rows Then .Cols = .Cols + 1: intRow = 0
                .Cell(flexcpChecked, intRow, .Cols - 1) = 2
                .TextMatrix(intRow, .Cols - 1) = Nvl(rsTemp!名称)
                intRow = intRow + 1
                rsTemp.MoveNext
            Loop
            .AutoSize 0, .Cols - 1
            'vsDept下方拖动条显示控制
            For i = 0 To .Cols - 1 '列宽之和
                lngColsWidth = lngColsWidth + .Cell(flexcpWidth, 0, i)
            Next
            If lngColsWidth > .ClientWidth Then .Height = .Height + 230: Frame3.Height = Frame3.Height + 150
            .Editable = flexEDKbdMouse
        End With
    End If
 
    '挂号项目
    If mEditType = Ed_安排修改 Or mEditType = ed_计划安排 Then
        strSQL = "Select ID as 序号,名称 From 收费项目目录 " & _
            " Where 类别='1' And (Sysdate Between 建档时间 And 撤档时间 Or 建档时间<Sysdate And 撤档时间 Is Null)" & _
            " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
            " Order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
        If rsTemp.EOF Then
            MsgBox "没有可用的挂号项目信息,请先到挂号项目设置中初始！", vbInformation, gstrSysName
            Exit Function
        End If
    
        cboItem.Clear
        For i = 1 To rsTemp.RecordCount
            cboItem.AddItem rsTemp!名称
            cboItem.ItemData(cboItem.NewIndex) = rsTemp!序号
            rsTemp.MoveNext
        Next
    End If
    
    'cmdCancel.Caption = "退出(&X)"
    If mEditType = Ed_安排审核 Then
        Me.Caption = Me.Caption & "——审核"
    ElseIf mEditType = Ed_安排删除 Then
        Me.Caption = Me.Caption & "——删除"
        'cmdOK.Caption = "删除(&D)"
    ElseIf mEditType = Ed_安排取消 Then
        Me.Caption = Me.Caption & "——取消审核"
    ElseIf mEditType = ed_安排查阅 Then
        cmdOK.Visible = False
        cmdCancel.Top = cmdOK.Top
    End If
    
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Sub cmdAuto_Click()
    If AutoAssignReapportion(tbSubPage.Item(mTimeSet.lngSelIndex).Caption) = False Then Exit Sub
    Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
End Sub

Private Function AutoAssignReapportion(ByVal str限制项目 As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng限号 As Long
    Dim lng限约 As Long
    Dim dat开始时间 As Date
    Dim dat结束时间 As Date
    Dim lng序号 As Long
    Dim strTmp As String
    Dim str时段 As String
    Dim str限制时间 As String
    Dim lng默认间隔 As Long
    Dim lng分配个数 As Long
    Dim lng固定数量 As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim dat时点 As Date
    Dim lng可用时间 As Long
    Dim lng初始间隔 As Long
    Dim lng待分摊个数 As Long
    Dim lng不分摊个数 As Long
    Dim lng间隔时间 As Long
    If mrs上班时间段 Is Nothing Then
        Call Init时间段
    End If

    If mrs上班时间段 Is Nothing Then Exit Function
    mTimeSet.rsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
    If mTimeSet.rsRegPlan.RecordCount = 0 Then mTimeSet.rsRegPlan.Filter = 0: Exit Function
    lng限号 = Nvl(mTimeSet.rsRegPlan!限号数, 0): lng限约 = Nvl(mTimeSet.rsRegPlan!限约数, 0)
    If lng限约 = 0 Then lng限约 = lng限号
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbInformation, Me.Caption
        Exit Function
    End If

    str时段 = mTimeSet.rsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbInformation, Me.Caption
        Exit Function
    End If
    
    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "' And 已使用=0"
    Do While Not mTimeSet.rsAssign.EOF
        mTimeSet.rsAssign.Delete adAffectCurrent
        mTimeSet.rsAssign.MoveNext
    Loop
    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "'"
    If mTimeSet.rsAssign.RecordCount <> 0 Then
        lng固定数量 = mTimeSet.rsAssign.RecordCount
        lng默认间隔 = Val(Nvl(mTimeSet.rsAssign!时间间隔, lng间隔时间))
        Do While Not mTimeSet.rsAssign.EOF
            lng分配个数 = lng分配个数 + Val(Nvl(mTimeSet.rsAssign!限制数量))
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    lng可用时间 = 0
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If
        lng可用时间 = lng可用时间 + DateDiff("n", dat开始时间, dat结束时间)
        mrs上班时间段.MoveNext
    Loop
    lng可用时间 = lng可用时间 - (lng固定数量 * lng默认间隔)
    If mTimeSet.bln序号控制 Then
        If lng限号 - lng分配个数 = 0 Then Exit Function
        lng初始间隔 = Int(lng可用时间 / (lng限号 - lng分配个数))
        If lng初始间隔 = 0 Then
            MsgBox "设置的限号数过大,无法自动计算时段!", vbInformation, gstrSysName
            Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
            Exit Function
        End If
        lng待分摊个数 = lng可用时间 - lng初始间隔 * (lng限号 - lng分配个数)
        lng不分摊个数 = (lng限号 - lng分配个数) - lng待分摊个数
    Else
        If lng限约 - lng分配个数 = 0 Then Exit Function
        lng初始间隔 = Int(lng可用时间 / (lng限约 - lng分配个数))
        If lng初始间隔 = 0 Then
            MsgBox "设置的限约数过大,无法自动计算时段!", vbInformation, gstrSysName
            Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
            Exit Function
        End If
        lng待分摊个数 = lng可用时间 - lng初始间隔 * (lng限约 - lng分配个数)
        lng不分摊个数 = (lng限约 - lng分配个数) - lng待分摊个数
    End If
    mrs上班时间段.MoveFirst
    
    mTimeSet.rsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        dat时点 = dat开始时间
        mrs上班时间段.MoveNext

        If mTimeSet.bln序号控制 Then
            For i = j To lng限号
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                If i > lng固定数量 Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        If lng不分摊个数 > 0 Then
                            If Format(DateAdd("n", lng初始间隔, dat时点), "yyyy-MM-dd hh:mm:00") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                                !结束时间 = Format(dat结束时间, "hh:mm:ss")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                            Else
                                !结束时间 = Format(DateAdd("n", lng初始间隔, dat时点), "hh:mm:00")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng初始间隔, dat时点), "hh:mm")
                            End If
                        Else
                            If Format(DateAdd("n", lng初始间隔 + 1, dat时点), "yyyy-MM-dd hh:mm:00") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                                !结束时间 = Format(dat结束时间, "hh:mm:ss")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                            Else
                                !结束时间 = Format(DateAdd("n", lng初始间隔 + 1, dat时点), "hh:mm:00")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng初始间隔 + 1, dat时点), "hh:mm")
                            End If
                        End If
                        If lng不分摊个数 > 0 Then
                            !时间间隔 = lng初始间隔
                        Else
                            !时间间隔 = lng初始间隔 + 1
                        End If
                        !限制数量 = IIf(lng分配个数 >= lng限号, 0, 1)
                        !是否预约 = 0
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限号, 0, 1)
                    End With
                    If lng不分摊个数 > 0 Then
                        dat时点 = DateAdd("n", lng初始间隔, dat时点)
                        lng不分摊个数 = lng不分摊个数 - 1
                    Else
                        dat时点 = DateAdd("n", lng初始间隔 + 1, dat时点)
                    End If
                Else
                    mTimeSet.rsAssign.Filter = "序号=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mTimeSet.rsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng初始间隔
                    End If
                    dat时点 = DateAdd("n", lng默认间隔, dat时点)
                End If
            Next

        Else    '非序号控制

            Do While Not Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss")
                ' If lngStart > lng限约 Then blnExit = True: Exit For
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then Exit Do

                If i > lng固定数量 Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        If lng不分摊个数 > 0 Then
                            If Format(DateAdd("n", lng初始间隔, dat时点), "yyyy-MM-dd hh:mm:00") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                                !结束时间 = Format(dat结束时间, "hh:mm:ss")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                            Else
                                !结束时间 = Format(DateAdd("n", lng初始间隔, dat时点), "hh:mm:00")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng初始间隔, dat时点), "hh:mm")
                            End If
                        Else
                            If Format(DateAdd("n", lng初始间隔 + 1, dat时点), "yyyy-MM-dd hh:mm:00") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                                !结束时间 = Format(dat结束时间, "hh:mm:ss")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                            Else
                                !结束时间 = Format(DateAdd("n", lng初始间隔 + 1, dat时点), "hh:mm:00")
                                !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng初始间隔 + 1, dat时点), "hh:mm")
                            End If
                        End If
                        
                        If lng不分摊个数 > 0 Then
                            !时间间隔 = lng初始间隔
                        Else
                            !时间间隔 = lng初始间隔 + 1
                        End If
                        !限制数量 = IIf(lng分配个数 >= lng限约, 0, 1)
                        !是否预约 = 1
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限约, 0, 1)
                    End With
                    If lng不分摊个数 > 0 Then
                        dat时点 = DateAdd("n", lng初始间隔, dat时点)
                        lng不分摊个数 = lng不分摊个数 - 1
                    Else
                        dat时点 = DateAdd("n", lng初始间隔 + 1, dat时点)
                    End If
                Else
                    mTimeSet.rsAssign.Filter = "序号=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mTimeSet.rsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng间隔时间
                    End If
                    dat时点 = DateAdd("n", lng默认间隔, dat时点)
                End If
                i = i + 1
            Loop
        End If
        If i > lng限号 And mTimeSet.bln序号控制 Then
            blnExit = True
        End If
    Loop
    AutoAssignReapportion = True
End Function

Private Sub dtpBegin_Validate(Cancel As Boolean)
    Dim strStartTime As String
    
    If Format(dtpEndDate.Value, "YYYY-MM-DD") <> "3000-01-01" Then Exit Sub
    
    If chk立即生效.Value = 1 Then
        strStartTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
    Else
        strStartTime = Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
        If Not mrsLongPlan Is Nothing Then
            If mrsLongPlan.RecordCount > 0 Then
                If Format(Nvl(mrsLongPlan!生效时间), "yyyy-MM-dd hh:mm:ss") > strStartTime Then
                    dtpEndDate.Value = CDate(Nvl(mrsLongPlan!生效时间, "3000-01-01"))
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call InitPage
    opt生效时间(0).Enabled = True: opt生效时间(1).Enabled = True
    mblnFirst = True
    Call LoadTimeSetControl
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadvsDept
    If InitData = False Then Unload Me: Exit Sub
    If LoadData = False Then Unload Me: Exit Sub
    Call SetCtrlEnabled
    If IsValidation() = False Then Unload Me: Exit Sub
    If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
        zlControl.ControlSetFocus chk序号控制
    Else
        zlControl.ControlSetFocus cmdOK
    End If
    If mblnSaveMinorChange Then
        tbPage.Item(1).Visible = False
        txt号别.Enabled = False
        chk序号控制.Enabled = False
        chk病案.Enabled = False
        Frame2.Enabled = False
        Frame1.Enabled = False
        opt天.Enabled = False
        opt周.Enabled = False
        cbo天.Enabled = False
        txt限号.Enabled = False
        txt限约.Enabled = False
        cbo号类.Enabled = False
        cboDoctor.Enabled = False
        cbo科室.Enabled = False
        vsPlan.Enabled = False
        cboItem.Enabled = False
        chk立即生效.Visible = False
        vsPlan.HighLight = flexHighlightNever
        opt生效时间(0).Enabled = False
        opt生效时间(1).Enabled = False
        dtpBegin.Enabled = False
        dtpEndDate.Enabled = False
        chk立即审核.Visible = False
        dtpBegin.MinDate = mdatOriBegin
        dtpBegin.Value = mdatOriBegin
        dtpEndDate.MaxDate = mdatOriEnd
        dtpEndDate.Value = mdatOriEnd
    End If
End Sub

Private Sub SaveMinorChange()
    Dim strSQL As String, intCount As Integer
    Dim str诊室 As String
    Dim i As Long
    Dim j As Long
    Dim rsTemp As ADODB.Recordset
    Dim intSync As Integer
    Dim lng执行计划ID As Long
    '诊室判断
    If opt分诊(1).Value Or opt分诊(2).Value Or opt分诊(3).Value Then
        intCount = 0
        With vsDept
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If .Cell(flexcpChecked, j, i) = 1 Then intCount = intCount + 1
                Next
            Next
        End With
        If opt分诊(1).Value Then
            If intCount = 0 Then
                MsgBox "指定诊室时必须选择一个对应的门诊诊室！", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            ElseIf intCount > 1 Then
                MsgBox "指定诊室时只能选择一个对应的门诊诊室！", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            End If
        ElseIf opt分诊(2).Value Or opt分诊(3).Value Then
            If intCount < 2 Then
                MsgBox "动态分诊或平均分诊时至少要选择两个对应的门诊诊室！", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            End If
        End If
    End If
    
    '取分诊方式
    intCount = 0
    For i = 0 To opt分诊.UBound
        If opt分诊(i).Value Then intCount = i: Exit For
    Next
    
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str诊室 = str诊室 & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str诊室 = Mid(str诊室, 2)
    
    strSQL = "Zl_挂号安排计划_Modify("
    strSQL = strSQL & mstr计划ID & ",'"
    strSQL = strSQL & str诊室 & "',"
    strSQL = strSQL & intCount & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '112585，谢荣，修改已审核的计划，未及时更改挂号安排的诊室信息
    If txtEdit(3).Text <> "" Then
        '检测当前修改计划的ID是否为挂号安排的执行计划ID
        gstrSQL = "Select 执行计划ID From 挂号安排 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng安排ID)
        lng执行计划ID = Val("" & rsTemp!执行计划ID)
        If lng执行计划ID = Val(mstr计划ID) Then
            '已审核且已生效，更新挂号安排的诊室信息
            strSQL = "Zl_挂号安排_Modify("
            strSQL = strSQL & mlng安排ID & ",'"
            strSQL = strSQL & str诊室 & "',"
            strSQL = strSQL & "Null,"
            strSQL = strSQL & intCount & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub SetCtrlEnabled()
    '设置控件的Enabled属性
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX"
            ctl.Enabled = False
            '修改或者新增计划时 开放限号、限约文本框 供修改
            If ctl Is Me.txt限号 Or ctl Is txt限约 Or ctl Is txtTimeOut Then
               ctl.Enabled = mEditType = Ed_安排修改 Or mEditType = ed_计划安排
            End If
        Case UCase("ComboBox")
            If ctl Is cbo天 And mEditType = ed_计划安排 Then
                   ctl.Enabled = opt天.Value = 1
              ElseIf ctl Is cboItem Or ctl Is cboDoctor Then
                 '-----------------------------------------------------
                 '为修改或者 新增模式时 开放对 项目和医生的更改

                 '------------------------------------------------------
                   If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
                       ctl.Enabled = True
                   Else
                       ctl.Enabled = False
                   End If
               Else:
                   ctl.Enabled = False
               End If
        Case UCase("ListView")
            ctl.Enabled = False
        Case UCase("DTPicker")
            ctl.Enabled = False
        Case UCase("optionbutton"), UCase("CheckBox")
            ctl.Enabled = False
            If (ctl.Name = "opt生效时间" Or ctl.Name = "opt应用于" Or ctl.Name = "opt科室" Or ctl.Name = "opt所有" Or ctl Is chkAppoint) And (mEditType = Ed_安排修改 Or mEditType = ed_计划安排) Then
               ctl.Enabled = True: ctl.Visible = True
            End If
        Case Else
        End Select
    Next
    
    Select Case mEditType
    Case ed_计划安排, Ed_安排修改
        chk序号控制.Enabled = True
        txt限号.Enabled = IIf(opt天.Value = True, True, False): txt限约.Enabled = IIf(opt天.Value = True And chkAppoint.Value = 1, True, False)
        cbo天.Enabled = IIf(opt天.Value = True, True, False)
        dtpBegin.Enabled = IIf(opt生效时间(1).Value = 1, True, False)
        dtpEndDate.Enabled = True
        vsDept.Enabled = True
        chkAppoint.Enabled = True
        opt分诊(0).Enabled = True: opt分诊(1).Enabled = True: opt分诊(2).Enabled = True: opt分诊(3).Enabled = True
        opt天.Enabled = True: opt周.Enabled = True
        dtpBegin.Enabled = True:
        
        '对分诊进行设置:
        '   指定医生时，不能设置成,动态分诊或平均分诊
        If Trim(cboDoctor.Text) <> "" Then
            opt分诊(2).Enabled = False: opt分诊(3).Enabled = False
            If opt分诊(2).Value Or opt分诊(3).Value Then opt分诊(0).Value = True
        Else
            opt分诊(2).Enabled = True: opt分诊(3).Enabled = True
        End If
        If opt天.Value = True Then cbo天.Enabled = True
        chk立即生效.Enabled = False: chk立即生效.Visible = False
        chk立即审核.Enabled = True
    Case Ed_安排审核
        chk立即审核.Enabled = False: chk立即审核.Visible = False
        chk立即生效.Enabled = True: chk立即生效.Visible = True
    Case Else
    End Select
    
    '设置编辑背景色
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX", UCase("ComboBox")
            Call zlSetCtrolBackColor(ctl)
        Case UCase("ListView")
        Case UCase("DTPicker")
        Case Else
        End Select
    Next
    
End Sub
 
Private Sub chk立即生效_Click()
'    dtpBegin.Enabled = chk立即生效.Value = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub
Private Function CheckPlanValied() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查计划的合法性
    '返回：计划安排合法,返回True,否则返回False
    '编制：刘兴洪
    '日期：2010-07-21 17:49:30
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    If mEditType <> Ed_安排修改 And mEditType <> ed_计划安排 Then
        CheckPlanValied = True: Exit Function
    End If
    
    If dtpBegin.Value > dtpEndDate.Value Then
        ShowMsgbox "注意:" & vbCrLf & "    生效时间小于了失效时间,请检查!"
        If dtpEndDate.Enabled And dtpEndDate.Visible Then dtpEndDate.SetFocus
        Exit Function
    End If
    If zlDatabase.Currentdate > dtpBegin.Value Then
        ShowMsgbox "注意:" & vbCrLf & "    生效时间小于了当前系统时间,请检查!"
        If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
        Exit Function
    End If
    Set rsTemp = Nothing
     CheckPlanValied = True: Exit Function
End Function

Private Function IsValied(Optional ByVal blnSave As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据的合法性
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-14 16:31:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long, intCount As Integer
    Dim strTmp As String, lng医生ID As Long, lng已约数 As Long
    Dim j As Integer
    Dim str限制项目 As String
    
    Err = 0: On Error GoTo Errhand:
    If Trim(txt号别) = "" Then
        MsgBox "号别不能为空！", vbInformation, gstrSysName
        txt号别.SetFocus: Exit Function
    End If
    If cbo科室.ListIndex = -1 Then
        MsgBox "未设置号别所对应的科室！", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Function
    End If
    If cboItem.ListIndex = -1 Then
        MsgBox "未设置号别所对应的挂号项目！", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Function
    End If
    
    If opt天.Value Then
        If cbo天.ListIndex = -1 Then
            MsgBox "该号别每天的应诊时间未设置！", vbInformation, gstrSysName
            If txt限号.Enabled Then txt限号.SetFocus
            Exit Function
        End If
        If chk序号控制.Value = 1 Then
            If Val(txt限号.Text) = 0 And Val(txt限约.Text) = 0 Then
                MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
                If txt限号.Enabled Then txt限号.SetFocus
                Exit Function
            End If
        Else
            If Not blnSave Then
                If chkAppoint.Value = 0 Or (chkAppoint.Value = 1 And txt限约.Text <> "" And Val(txt限约.Text) = 0) Then
                    MsgBox "非序号控制的安排设置时段时,必须是可预约的安排！", vbInformation, gstrSysName
                    If txt限号.Enabled Then txt限号.SetFocus: Exit Function
                End If
            End If
        End If
        '限号限约规则
        If Trim(txt限号.Text) <> "" Then
            If Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
               If txt限约.Enabled Then txt限约.SetFocus
                Exit Function
            End If
        ElseIf Trim(txt限约.Text) <> "" Then
            MsgBox "限约必须限号！", vbInformation, gstrSysName
            If txt限号.Enabled Then txt限号.SetFocus
            Exit Function
        End If
    Else
        If Not blnSave Then
            If chkAppoint.Value = 0 And chk序号控制.Value = 0 Then
                MsgBox "非序号控制的安排设置时段时,必须是可预约的安排！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
     With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
                    If chk序号控制.Value = 1 Then
                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                              MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
                              .Row = 2: .Col = i
                              .SetFocus: Exit Function
                          End If
                      End If
                        '限号限约规则
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Function
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            MsgBox "限约必须限号！", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "该号别每周的应诊时间未设置！", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Function
            End If
        End With
    End If
    
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    If lng医生ID = 0 And cboDoctor.Text <> "" Then
        strSQL = "Select 1 From 人员表 Where 姓名 = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDoctor.Text)
        If Not rsTemp.EOF Then
            MsgBox "医生""" & cboDoctor.Text & """不属于科室""" & cbo科室.Text & """,请重新设置该号别的科室与医生信息！", vbInformation, gstrSysName
            cboDoctor.SetFocus: Exit Function
        End If
    End If
    
    '诊室判断
    If opt分诊(1).Value Or opt分诊(2).Value Or opt分诊(3).Value Then
        intCount = 0
        With vsDept
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If .Cell(flexcpChecked, j, i) = 1 Then intCount = intCount + 1
                Next
            Next
        End With
        If opt分诊(1).Value Then
            If intCount = 0 Then
                MsgBox "指定诊室时必须选择一个对应的门诊诊室！", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Function
            ElseIf intCount > 1 Then
                MsgBox "指定诊室时只能选择一个对应的门诊诊室！", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Function
            End If
        ElseIf opt分诊(2).Value Or opt分诊(3).Value Then
            If intCount < 2 Then
                MsgBox "动态分诊或平均分诊时至少要选择两个对应的门诊诊室！", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Function
            End If
        End If
    End If
     
    '项目价格判断
    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
        MsgBox "项目""" & cboItem.Text & """未设置有效价格,请先到收费项目管理中设置！", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Function
    End If
    If opt生效时间(1).Value = 0 Then
        If Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
            ShowMsgbox "生效时间不能小于当前系统时间,请检查!"
            Exit Function
        End If
    End If
    '检查相关的计划
    If CheckPlanValied = False Then Exit Function
    If mEditType = ed_计划安排 Then
        '新增加计划时,检查
        If Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss") < Format(mdtMinCustom, "yyyy-mm-dd hh:mm:ss") Then
            If MsgBox("该计划在生效日期后已存在预约号,是否继续?", vbYesNo + vbDefaultButton1 + vbInformation, Me.Caption) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If CheckUsedCount() = False Then Exit Function
    IsValied = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    If 1 = 2 Then
        Resume
    End If
End Function

Private Function LoadLongPlan() As Boolean
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select 生效时间 From 挂号安排计划" & _
            " Where 安排id = [1] And 失效时间 = To_Date('3000-01-01', 'yyyy-mm-dd') And 审核时间 Is Not Null"
    Set mrsLongPlan = zlDatabase.OpenSQLRecord(strSQL, "长期计划", mlng安排ID)
    LoadLongPlan = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LongPlanIsValied(ByRef lng上次计划ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查长期计划的有效性
    ' 入参 :
    ' 出参 : lng上次计划ID-被调整的长期计划
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2018/11/9 11:29
    ' 问题 :133584
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsPlan As ADODB.Recordset
    Dim strStartTime As String
    
    On Error GoTo errH
    If Format(dtpEndDate.Value, "YYYY-MM-DD") <> "3000-01-01" Then LongPlanIsValied = True: Exit Function
    
    If chk立即生效.Value = 1 Then
        strStartTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
    Else
        strStartTime = Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
        If Not mrsLongPlan Is Nothing Then
            If mrsLongPlan.RecordCount > 0 Then
                If Format(Nvl(mrsLongPlan!生效时间), "yyyy-MM-dd hh:mm:ss") > strStartTime Then
                    MsgBox "长期计划(" & Format(Nvl(mrsLongPlan!生效时间), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01)的生效时间比本次的生效时间晚，本次只能做为短期计划。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    If chk立即审核.Value = 0 And Not mEditType = Ed_安排审核 Then LongPlanIsValied = True: Exit Function
    
    strSQL = "Select 0 as 是否已审, ID, 生效时间 From 挂号安排计划" & _
            " Where 安排id = [1] And 失效时间 = To_Date('3000-01-01', 'yyyy-mm-dd') And 生效时间 < To_Date([2],'YYYY-MM-DD hh24:mi:ss') " & _
            " And 审核时间 Is Null" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 1 as 是否已审, ID, 生效时间 From 挂号安排计划" & _
            " Where 安排id = [1] And 失效时间 = To_Date('3000-01-01', 'yyyy-mm-dd') And 审核时间 Is Not Null"
    Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, "检查长期计划", mlng安排ID, strStartTime)
    If rsPlan.RecordCount = 0 Then LongPlanIsValied = True: Exit Function
    rsPlan.Filter = "是否已审 = " & 0
    If rsPlan.RecordCount > 0 Then
        MsgBox "还有未审核的长期计划(" & Format(Nvl(rsPlan!生效时间), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01)，请依次审核或删除。", vbInformation, gstrSysName
        Exit Function
    End If
    rsPlan.Filter = "是否已审 = 1"
    If rsPlan.RecordCount = 0 Then LongPlanIsValied = True: Exit Function
    If Format(Nvl(rsPlan!生效时间), "yyyy-MM-dd hh:mm:ss") < strStartTime Then
        If MsgBox("存在长期计划(" & Format(Nvl(rsPlan!生效时间), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01), 是否将其失效时间调整为本次的生效时间？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
        lng上次计划ID = Val(Nvl(rsPlan!ID))
    Else
        MsgBox "长期计划(" & Format(Nvl(rsPlan!生效时间), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01)的生效时间比本次的生效时间晚，请先将本次修改为短期计划。", vbInformation, gstrSysName
        Exit Function
    End If
    LongPlanIsValied = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePlan(ByVal lng上次计划ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存计划安排
    '返回:保存成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-14 16:41:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str时间段 As String, str诊室 As String, i As Long, int分诊 As Integer
    Dim lng计划ID As Long, str限号 As String
    Dim str医生姓名         As String
    Dim str医生ID           As String
    Dim blnChange           As Boolean
    Dim bytType             As Byte
    Dim vMsgResult          As VbMsgBoxResult
    Dim strMsg              As String
    Dim colPro              As Collection
    Dim blnTrans            As Boolean
    Dim j                   As Integer
    'bytType 0-新增时 对时段不进行处理 修改时 对时段只删除已经去掉的排班信息
    '        1-新增时 提取原安排的时段信息  修改时 对计划的时段进行删除
    
    Err = 0: On Error GoTo Errhand:
    
    str时间段 = "": str限号 = ""
    If opt天.Value Then
        For i = 1 To 7
            str时间段 = str时间段 & ",'" & Trim(cbo天.Text) & "'"
            str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
            str限号 = str限号 & "," & Val(txt限号.Text) & "," & IIf(chkAppoint.Value = 0, "0", txt限约.Text)
        Next
    Else
        With vsPlan
            For i = 1 To .Cols - 1
                str时间段 = str时间段 & ",'" & Trim(.TextMatrix(1, i)) & "'"
                If Trim(.TextMatrix(1, i)) <> "" Then
                    str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                    str限号 = str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & IIf(chkAppoint.Value = 0, "0", Trim(vsPlan.TextMatrix(3, i)))
                End If
            Next
        End With
    End If
    If str限号 <> "" Then str限号 = Mid(str限号, 2)
            
    If mPlanInfo.bln时间段 Then
      '判断是已经改变 计划信息
        'blnChange = (mPlanInfo.str排班 <> str时间段) Or (mPlanInfo.str限号 <> str限号) Or (IIf(mPlanInfo.bln序号, 1, 0) <> chk序号控制.Value)
         blnChange = True
    End If
    '71253 李南春 2014-04-15 14:23:10 将listView 替换为vsflexGrid
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str诊室 = str诊室 & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str诊室 = Mid(str诊室, 2)
    
    '取分诊方式
    int分诊 = 0
    For i = 0 To opt分诊.UBound
        If opt分诊(i).Value Then int分诊 = i: Exit For
    Next
     '问题号:52275
    '在计划或者安排设置了时段时 对时段处理的处理类型
'    If mPlanInfo.bln时间段 And mEditType = ed_计划安排 And blnChange = False Then
'        '如果原计划或者安排时 设置了时段 提示操作原进行处理
'        strMsg = "安排中设置了时段,是否提取安排的时段做为计划的时段信息? " & vbCrLf
'        strMsg = strMsg & "[是(Y)]提取安排的时段信息作为计划的时段" & vbCrLf
'        strMsg = strMsg & "[否(N)]不提取安排的时段,重新设置时段" & vbCrLf
'        vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
'        bytType = IIf(vMsgResult = vbYes, 1, 0)
'    End If
    If mEditType = Ed_安排修改 Then
      bytType = IIf(IIf(mPlanInfo.bln序号, 1, 0) <> chk序号控制.Value, 1, 0)
    End If
    '取时间范围
    If mEditType = ed_计划安排 Then
        lng计划ID = zlDatabase.GetNextId("挂号安排计划")
    Else
        lng计划ID = Val(mstr计划ID)
    End If
     If cboDoctor.ListIndex = -1 Then
        str医生姓名 = ""
        str医生ID = "0"
     Else
        str医生姓名 = cboDoctor.Text
        str医生ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
     End If
    'Zl_挂号安排计划_Insert
    strSQL = "Zl_挂号安排计划_Insert("
    '  Id_In       In 挂号安排计划.ID%Type,
    strSQL = strSQL & "" & lng计划ID & ","
    '  安排id_In   In 挂号安排计划.安排id%Type,
    strSQL = strSQL & "" & mlng安排ID & ","
    '  号码_In     In 挂号安排计划.号码%Type,
    strSQL = strSQL & "'" & txt号别.Text & "',"
    '  生效时间_In In 挂号安排计划.生效时间%Type,
    If opt生效时间(0).Value = True Then
        strSQL = strSQL & "Sysdate ,"
    Else
        strSQL = strSQL & "to_date('" & dtpBegin.Value & "','yyyy-mm-dd hh24:mi:ss'),"
    End If
    '  失效时间_In In 挂号安排计划.失效时间%Type
    strSQL = strSQL & "to_date('" & dtpEndDate.Value & "','yyyy-mm-dd hh24:mi:ss') "
    '  周日_In     In 挂号安排计划.周日%Type,
    '  周一_In     In 挂号安排计划.周一%Type,
    '  周二_In     In 挂号安排计划.周二%Type,
    '  周三_In     In 挂号安排计划.周三%Type,
    '  周四_In     In 挂号安排计划.周四%Type,
    '  周五_In     In 挂号安排计划.周五%Type,
    '  周六_In     In 挂号安排计划.周六%Type,
    strSQL = strSQL & str时间段 & ","
    '   限号控制_In In Varchar2,
    strSQL = strSQL & "'" & str限号 & "',"
    '  分诊方式_In In 挂号安排计划.分诊方式%Type,
    strSQL = strSQL & "" & int分诊 & ","
    '  序号控制_In In 挂号安排计划.序号控制%Type,
    strSQL = strSQL & "" & IIf(chk序号控制.Value = 1, 1, 0) & ","
    '  项目ID_In   In 挂号安排计划.项目ID%Type,
    strSQL = strSQL & Me.cboItem.ItemData(cboItem.ListIndex) & ","
    '医生姓名_In In 挂号安排计划.医生姓名%Type,
    strSQL = strSQL & IIf(str医生姓名 = "", "NULL,", "'" & str医生姓名 & "',")
    '医生id_In   In 挂号安排计划.医生id%Type,
    strSQL = strSQL & str医生ID & ","
    '  诊室_In     Varchar2,
    strSQL = strSQL & "'" & str诊室 & "',"
    '  新增_In Number:=1,处理类型
    strSQL = strSQL & "" & IIf(mEditType = ed_计划安排, 1, 0) & "," & bytType & ")"
     

    Set colPro = New Collection
    zlAddArray colPro, strSQL
    If Not mTimeSet.blnIsInit Then
         Call LoadTimePlan
    End If
    If SaveTimeSetData(lng计划ID, colPro) = False Then Exit Function
    '立即审核立即生效
    If chk立即审核.Value = 1 Then
        strSQL = "Zl_挂号安排计划_Verify(" & lng计划ID & "," & IIf(opt生效时间(0).Value, 1, 0) & "," & ZVal(lng上次计划ID) & ")"
        zlAddArray colPro, strSQL
    End If
    gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy colPro, Me.Caption, True, True
    gcnOracle.CommitTrans: blnTrans = False
    SavePlan = True
    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Function SaveTimeSetData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '出参:cllPro-返回相关保存数据的SQL
    '编制:刘兴洪
    '日期:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mTimeSet.blnIsInit Then Exit Function
    If zl_CheckMoveAssign() = False Then Exit Function
    If VsTimeValidate(-1) = False Then Exit Function
    
    If mPlanEditType = EM_安排_修改 Or mPlanEditType = EM_安排_增加 Then
        If SaveSetData(lngID, cllPro) = False Then Exit Function
    Else
        If SavePlanData(lngID, cllPro) = False Then Exit Function
    End If
    
    SaveTimeSetData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AssignManage() As Boolean
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    Dim lng限号数 As Long, lng限约数 As Long, lng分配数量 As Long
    Dim lng分配预约 As Long, lngTmp  As Long, lngTemp As Long
    Dim str最大时间 As String, blnChange As Boolean
     
    varData = Split(mTimeSet.str安排, "|")
    lngIndex = -1
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        '如果当天应诊时间改变
        If InStr("|" & mTimeSet.str安排, "|" & strTemp & ",") = 0 Or InStr("|" & mTimeSet.str应诊时段 & "|", "|" & strTemp & "|") = 0 Then
            mTimeSet.rsAssign.Filter = "限制项目='" & strTemp & "'"
            Do While Not mTimeSet.rsAssign.EOF
                mTimeSet.rsAssign.Delete adAffectCurrent
                mTimeSet.rsAssign.Update
                mTimeSet.rsAssign.MoveNext
            Loop
        End If
    Next
    For i = 0 To UBound(varData)
        ''周一,限号数,限约数|周二,限号数,限约数|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            lng限号数 = Val(varTemp(1)): lng限约数 = Val(varTemp(2))
            If lng限约数 = 0 Then lng限约数 = lng限号数
            str最大时间 = ""
            If Not mTimeSet.rsHistory Is Nothing Then
                mTimeSet.rsHistory.Filter = "限制项目='" & varTemp(0) & "'"
                If mTimeSet.rsHistory.RecordCount = 0 Then
                   str最大时间 = ""
                Else
                   str最大时间 = Nvl(mTimeSet.rsHistory!发生时间)
                End If
            End If
            mTimeSet.rsAssign.Filter = "限制项目='" & varTemp(0) & "'"
            mTimeSet.rsAssign.Sort = "序号"
             
'             If str最大时间 <> "" Then
'                Do While Not mTimeSet.rsAssign.EOF
'                   If str最大时间 > Nvl(mTimeSet.rsAssign!开始时间) Then mTimeSet.rsAssign.Delete adAffectCurrent
'                   mTimeSet.rsAssign.MoveNext
'                Loop
'             End If

              lng分配数量 = 0
              blnChange = False
             Do While Not mTimeSet.rsAssign.EOF
                If lng分配数量 + Val(Nvl(mTimeSet.rsAssign!限制数量)) > IIf(mTimeSet.bln序号控制, lng限号数, lng限约数) Then
                    blnChange = True
                    If Val(Nvl(mTimeSet.rsAssign!已使用)) = 0 Then
                        lngTmp = Val(mTimeSet.rsAssign!限制数量)
                        lngTemp = lng分配数量 + lngTmp - IIf(mTimeSet.bln序号控制, lng限号数, lng限约数)
                        If lngTmp <= lngTemp Then
                            lngTmp = 0
                        Else
                            lngTmp = lngTmp - lngTemp
                            lng分配数量 = lng限号数
                        End If
                        mTimeSet.rsAssign!限制数量 = lngTmp
                        mTimeSet.rsAssign.Update
                        If mTimeSet.bln序号控制 Then
                            mTimeSet.rsAssign.Delete adAffectCurrent
                        End If
                    End If
                Else
                    lng分配数量 = lng分配数量 + Val(Nvl(mTimeSet.rsAssign!限制数量))
                End If
                mTimeSet.rsAssign.MoveNext
             Loop
             If blnChange Then
                mTimeSet.rsAssign.Filter = "限制项目='" & varTemp(0) & "' And 限制数量>0"
                lng分配数量 = 0
                If mTimeSet.rsAssign.RecordCount = 0 Then mTimeSet.rsAssign.Filter = 0: AssignManage = True: Exit Function
                mTimeSet.rsAssign.Sort = "序号 desc"
                mTimeSet.rsAssign.MoveFirst
                'lng分配数量
                Do While Not mTimeSet.rsAssign.EOF
                   lng分配数量 = lng分配数量 + Val(Nvl(mTimeSet.rsAssign!限制数量))
                   mTimeSet.rsAssign.MoveNext
                Loop
                mTimeSet.rsAssign.MoveFirst
                If lng分配数量 > IIf(mTimeSet.bln序号控制, lng限号数, lng限约数) Then
                   Do While Not mTimeSet.rsAssign.EOF
                      If Val(Nvl(mTimeSet.rsAssign!已使用)) = 0 Then
                           lngTmp = Val(Nvl(mTimeSet.rsAssign!限制数量))
                           lngTemp = lng分配数量 - lng限号数
                           If lngTemp >= lngTmp Then
                               mTimeSet.rsAssign!限制数量 = 0
                               mTimeSet.rsAssign.Update
                               lng分配数量 = lng分配数量 - lngTmp
                           Else
                               lngTmp = lngTmp - lngTemp
                               mTimeSet.rsAssign!限制数量 = lngTmp
                               mTimeSet.rsAssign.Update
                               lng分配数量 = lng分配数量 - lngTemp
                           End If
                      End If
                      If lng分配数量 <= lng限号数 Then Exit Do
                      mTimeSet.rsAssign.MoveNext
                   Loop
                End If
             End If
        End If
    Next
    mTimeSet.rsAssign.Filter = 0
    If Not mTimeSet.rsHistory Is Nothing Then mTimeSet.rsHistory.Filter = 0
    AssignManage = True
End Function

Private Function SavePlanData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    Dim i As Long, str星期 As String, lng序号 As String, strSQL As String
    Dim str序号s As String, bytType As Byte '应用于
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer, cllPage As Collection
    Dim str时段 As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
   
    On Error GoTo errHandle
    
    Call AssignManage  '序号分配处理
    '68499,刘尔旋,2014-1-8,安排没有时段信息计划有时段信息时时段信息没有保存的错误
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        mTimeSet.blnChange = True
        Call MoveAssign(strTemp)
    Next i
    
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_挂号计划时段_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        mTimeSet.rsAssign.Filter = "限制项目='" & strTemp & "'"
        If mTimeSet.rsAssign.RecordCount > 0 Then
            Do While Not mTimeSet.rsAssign.EOF
    '            序号,开始时间,结束时间,限制数量,预约标志|...
                strTmp = mTimeSet.rsAssign!序号
                strTmp = strTmp & "," & mTimeSet.rsAssign!开始时间 & "," & mTimeSet.rsAssign!结束时间 & "," & mTimeSet.rsAssign!限制数量 & "," & mTimeSet.rsAssign!是否预约
                If Len(str时段 & "|" & strTmp) > 4000 Then
                    str时段 = Mid(str时段, 2)
                    strSQL = "  Zl_挂号计划时段_Insert("
                    '  安排id_In 挂号安排时段.安排id%Type,
                    strSQL = strSQL & lngID & ","
                    '  星期_In   挂号安排时段.星期%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  时段_In   Varchar2,
                    strSQL = strSQL & "'" & str时段 & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    str时段 = ""
                End If
                str时段 = str时段 & "|" & strTmp
                mTimeSet.rsAssign.MoveNext
            Loop
            If str时段 <> "" Then
                 
                str时段 = Mid(str时段, 2)
                strSQL = "  Zl_挂号计划时段_Insert("
                '  安排id_In 挂号安排时段.安排id%Type,
                strSQL = strSQL & lngID & ","
                '  星期_In   挂号安排时段.星期%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  时段_In   Varchar2,
                strSQL = strSQL & "'" & str时段 & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                str时段 = ""
            End If
        
        End If
    Next
    SavePlanData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Function SaveSetData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:安排时段数据保存
    '入参:lngID-安排ID
    '出参:cllPro-返回相关保存数据的SQL
    '编制:刘兴洪
    '日期:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str星期 As String, lng序号 As String, strSQL As String
    Dim str序号s As String, bytType As Byte '应用于
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer
    Dim str时段 As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
     
   
    On Error GoTo errHandle
      
    Call AssignManage  '序号分配处理

    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_挂号安排时段_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        mTimeSet.rsAssign.Filter = "限制项目='" & strTemp & "'"
        If mTimeSet.rsAssign.RecordCount > 0 Then
            Do While Not mTimeSet.rsAssign.EOF
    '            序号,开始时间,结束时间,限制数量,预约标志|...
                strTmp = mTimeSet.rsAssign!序号
                strTmp = strTmp & "," & mTimeSet.rsAssign!开始时间 & "," & mTimeSet.rsAssign!结束时间 & "," & mTimeSet.rsAssign!限制数量 & "," & mTimeSet.rsAssign!是否预约
                If Len(str时段 & "|" & strTmp) > 4000 Then
                    str时段 = Mid(str时段, 2)
                    strSQL = "  Zl_挂号安排时段_Insert("
                    '  安排id_In 挂号安排时段.安排id%Type,
                    strSQL = strSQL & lngID & ","
                    '  星期_In   挂号安排时段.星期%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  时段_In   Varchar2,
                    strSQL = strSQL & "'" & str时段 & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    str时段 = ""
                End If
                str时段 = str时段 & "|" & strTmp
                mTimeSet.rsAssign.MoveNext
            Loop
            If str时段 <> "" Then
                 
                str时段 = Mid(str时段, 2)
                strSQL = "  Zl_挂号安排时段_Insert("
                '  安排id_In 挂号安排时段.安排id%Type,
                strSQL = strSQL & lngID & ","
                '  星期_In   挂号安排时段.星期%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  时段_In   Varchar2,
                strSQL = strSQL & "'" & str时段 & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                str时段 = ""
            End If
        
        End If
    Next
    SaveSetData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

 Private Function GetVsGridIndex(ByVal str星期 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取相关索引
    '编制:刘兴洪
    '日期:2012-06-15 14:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    str星期 = Switch(str星期 = "周日", 0, str星期 = "周一", 1, str星期 = "周二", 2, str星期 = "周三", 3, str星期 = "周四", 4, str星期 = "周五", 5, str星期 = "周六", 6, True, 0)
    GetVsGridIndex = Val(str星期)
 End Function

Private Function MoveAssign(ByVal str限制项目 As String) As Boolean
    '分配调整的序号到数据集中
    Dim nIndex As Long
    Dim lng序号 As Long
    Dim i As Long, j As Long
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim lng限制 As Long
    Dim bln预约 As Boolean
    Dim str最大时间 As String
    
    If Not mTimeSet.blnChange Then MoveAssign = True: Exit Function
    
    nIndex = GetVsGridIndex(str限制项目)
    
    '删掉没有使用部分
    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "' and 已使用=0"
    If mTimeSet.rsAssign.RecordCount > 0 Then
        Do While Not mTimeSet.rsAssign.EOF
            mTimeSet.rsAssign.Delete
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    
    If Not mTimeSet.bln序号控制 Then
        With vsTime(nIndex)
          lng序号 = 0
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                   If .TextMatrix(i, j) <> "" Then
                    
                    str开始时间 = Split(.TextMatrix(i, j), "-")(0)
                    str结束时间 = Split(.TextMatrix(i, j), "-")(1)
                    lng限制 = Val(.TextMatrix(i, j + 1))
                    lng序号 = lng序号 + 1
                    bln预约 = True
                    
                    str最大时间 = ""
                    If Not mTimeSet.rsHistory Is Nothing Then
                        mTimeSet.rsHistory.Filter = "限制项目='" & str限制项目 & "'"
                        If mTimeSet.rsHistory.RecordCount = 0 Then
                            str最大时间 = ""
                            mTimeSet.rsHistory.Filter = 0
                        Else
                            str最大时间 = Nvl(mTimeSet.rsHistory!发生时间)
                            mTimeSet.rsHistory.Filter = 0
                        End If
                    End If
                    
                    If (str最大时间 <> "" And str开始时间 > str最大时间) Or str最大时间 = "" Then
                        With mTimeSet.rsAssign
                            .AddNew
                            !限制项目 = str限制项目
                            !开始时间 = str开始时间
                            !结束时间 = str结束时间
                            !时间段 = str开始时间 & "-" & str结束时间
                            !限制数量 = lng限制
                            !序号 = lng序号
                            !已使用 = 0
                            !是否预约 = 1
                            .Update
                        End With
                    End If
                   End If
                Next
            Next
        End With
        mTimeSet.blnChange = False
        MoveAssign = True
        Exit Function
    End If
    
    
    '序号控制
    
    With vsTime(nIndex)
        For i = 0 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                If Trim(.TextMatrix(i, j)) <> "" Then
                        str开始时间 = Split(.TextMatrix(i + 1, j) & "-", "-")(0)
                        str结束时间 = Split(.TextMatrix(i + 1, j) & "-", "-")(1)
                        lng序号 = Val(.TextMatrix(i, j))
                        lng限制 = 1
                        bln预约 = .Cell(flexcpForeColor, i, j) = vbBlue
                    If .Cell(flexcpFontUnderline, i, j) = False Then
                       
                        With mTimeSet.rsAssign
                            .AddNew
                            !限制项目 = str限制项目
                            !开始时间 = str开始时间
                            !结束时间 = str结束时间
                            !时点 = Format(str开始时间, "hh:00:00")
                            !时间段 = str开始时间 & "-" & str结束时间
                            !限制数量 = lng限制
                            !序号 = lng序号
                            !已使用 = 0
                            !是否预约 = IIf(bln预约, 1, 0)
                            .Update
                        End With
                    ElseIf .Cell(flexcpFontUnderline, i, j) Then
                        ' 固定的信息,可能改变是否预约,现在也只可改变是否预约
                        With mTimeSet.rsAssign
                            .Filter = "序号=" & lng序号 & " And 开始时间='" & Format(str开始时间, "hh:mm:00") & "'"
                            If .RecordCount > 0 Then
                                !是否预约 = IIf(bln预约, 1, 0)
                                .Update
                            End If
                        End With
                    End If
                End If
            Next
        Next
    End With
    mTimeSet.blnChange = False
    MoveAssign = True
    Exit Function
End Function

Private Function zl_CheckMoveAssign(Optional ByVal lngIndex As Long = -1) As Boolean
    Dim str限制项目 As String
    If lngIndex = -1 Then lngIndex = mTimeSet.lngSelIndex
    If lngIndex = -1 Then zl_CheckMoveAssign = True: Exit Function
    If Not mTimeSet.blnChange Then zl_CheckMoveAssign = True: Exit Function
    
    If lngIndex < 0 Or lngIndex > 6 Then Exit Function
    If Not VsTimeValidate(lngIndex) Then Exit Function
    
    str限制项目 = GetVsGridCaption(lngIndex)
    zl_CheckMoveAssign = MoveAssign(str限制项目)
End Function

 Private Function GetVsGridCaption(ByVal nIndex As Integer) As String
    '功能:根据索引获取限制项目
    Dim str星期 As String
    str星期 = Switch(nIndex = 0, "周日", nIndex = 1, "周一", nIndex = 2, "周二", nIndex = 3, "周三", nIndex = 4, "周四", nIndex = 5, "周五", nIndex = 6, "周六", True, "")
    GetVsGridCaption = str星期
 End Function

Private Function VsTimeValidate(ByVal lngIndex As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证设置的限约数是否符合要求
    '入参:lngIndex-指定的页面(星期对应的索引):-1时,表示按所有的页面进行检查
    '出参:
    '返回:校对成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-15 10:17:37
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long, i As Long, j  As Long
    Dim lng预约数   As Long, lng限号数 As Long, lng限约数 As Long, lng号数 As Long
    Dim str星期   As String, str限制项目 As String
    Dim lngPage As Long, lngPages As Long, lngStartPage As Long
    Dim blnNotSetTime As Boolean '允许不设置时间段
    Dim blnAllowNums As Boolean '允许限号数不一致
    Dim blnAllowYYNums As Boolean '允许预约数与设置的预约数不一致
    Dim strCommand As String, bln时段 As Boolean '判断设置了时段的,需要检查其他时段页是否设置
    On Error GoTo errHandle
        
    lngStartPage = 0: lngPages = tbSubPage.ItemCount - 1
    If lngIndex <> -1 Then lngStartPage = lngIndex: lngPages = lngIndex
    bln时段 = False
    For lngPage = lngStartPage To lngPages
        If mTimeSet.bln序号控制 Then
            With vsTime(lngPage)
                For i = 0 To .Rows - 1 Step 2
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                           bln时段 = True: Exit For
                       End If
                    Next
                Next
            End With
        Else
                With vsTime(lngPage)
                    For i = 1 To .Rows - 1
                        For j = 1 To .Cols - 1 Step 2
                            If .TextMatrix(i, j) <> "" Then
                               bln时段 = True: Exit For
                            End If
                        Next
                    Next
                End With
        End If
    Next
    '未启用时段
    If bln时段 = False Then VsTimeValidate = True: Exit Function
    
    For lngPage = lngStartPage To lngPages
        tbSubPage(lngPage).Selected = True
        str限制项目 = GetVsGridCaption(lngPage)
        mTimeSet.rsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
        If mTimeSet.rsRegPlan.RecordCount = 0 Then
            mTimeSet.rsRegPlan.Filter = 0
        Else
                lng限号数 = Val(Nvl(mTimeSet.rsRegPlan!限号数)): lng限约数 = Val(Nvl(mTimeSet.rsRegPlan!限约数))
                If lng限约数 = 0 Then lng限约数 = lng限号数
                lng号数 = 0: lng预约数 = 0
                
                If mTimeSet.bln序号控制 Then
                    '专家号检查限约数是否大于限号数
                    With vsTime(lngPage)
                        For i = 0 To .Rows - 1 Step 2
                            For j = 1 To .Cols - 1
                               If .TextMatrix(i, j) <> "" Then
                                     If .Cell(flexcpForeColor, i, j, i, j) = vbBlue Then
                                         lng预约数 = lng预约数 + 1
                                     End If
                                     lng号数 = lng号数 + 1
                               End If
                            Next
                        Next
                    End With
                    If lng号数 < lng限号数 Then
                        If lng号数 = 0 Then
                           If lngIndex = -1 Then
                                If blnNotSetTime = False And bln时段 Then
                                        strCommand = zlCommFun.ShowMsgbox("提醒", "    在分时段页面中未设置『" & str限制项目 & "』的时段,你确定不设置时间段?" & vbCrLf & vbCrLf & _
                                         "『是』:表示允许不设置时间段进行保存" & vbCrLf & vbCrLf & _
                                         "『忽略』:表示遇到类似的未设置时间段的问题允许保存,但不再提示。" & vbCrLf & vbCrLf & _
                                         "『否』:表示不允许不设置时间段,返回重新设置" & vbCrLf, "是(&O),忽略(&I),否(&C)", Me, vbQuestion)
                                        Select Case strCommand
                                        Case "是"
                                        Case "忽略"
                                             blnNotSetTime = True
                                         Case Else
                                            Call zlSaveTimePageSelected(str限制项目)
                                            mTimeSet.blnNotBrush = True
                                            tbSubPage.Item(lngPage).Selected = True
                                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                            mTimeSet.blnNotBrush = False
                                            Exit Function
                                         End Select
                                End If
'                           Else
'                                If MsgBox("在分时段页面中未设置『" & str限制项目 & "』的时段,你确定不设置时间段?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                                    If lngIndex = -1 Then
'                                        Call zlSaveTimePageSelected(str限制项目)
'                                        mTimeSet.blnNotBrush = True
'                                        tbSubPage.Item(lngPage).Selected = True
'                                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
'                                        mTimeSet.blnNotBrush = False
'                                    End If
'                                    Exit Function
'                                End If
                            End If
                        Else
                                If lngIndex = -1 Then
                                        If blnAllowNums = False Then
                                        
                                                strCommand = zlCommFun.ShowMsgbox("提醒", "    在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")与限号数(" & lng限号数 & ") 不等,你确定按当前设置的时段保存?" & vbCrLf & vbCrLf & _
                                                 "『是』:表示允许限号数与号数不一致" & vbCrLf & vbCrLf & _
                                                 "『忽略』:表示允许限号数与号数不一致，遇到类似的问题,不再提示。" & vbCrLf & vbCrLf & _
                                                 "『否』:表示不允许限号数与号数不一致,返回重新设置" & vbCrLf, "是(&O),忽略(&I),否(&C)", Me, vbQuestion)
                                                Select Case strCommand
                                                 Case "是"
                                                 Case "忽略"
                                                     blnAllowNums = True
                                                 Case Else
                                                    Call zlSaveTimePageSelected(str限制项目)
                                                    mTimeSet.blnNotBrush = True
                                                    tbSubPage.Item(lngPage).Selected = True
                                                    If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                                    mTimeSet.blnNotBrush = False
                                                     Exit Function
                                                 End Select
                                        End If
'                                   Else
'                                        If MsgBox("在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")与限号数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                End If
                        End If
                    ElseIf lng号数 > lng限号数 Then
                        Call MsgBox("在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")大于了限号数(" & lng限约数 & ") 你不能按当前设置的时段保存!", vbQuestion + vbOKOnly + vbDefaultButton2, gstrSysName)
                        If lngIndex = -1 Then
                            Call zlSaveTimePageSelected(str限制项目)
                            mTimeSet.blnNotBrush = True
                            tbSubPage.Item(lngPage).Selected = True
                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                            mTimeSet.blnNotBrush = False
                        End If
                        Exit Function
                    End If
                Else
                     '普通号检查限约数是否大于限号数
                    With vsTime(lngPage)
                        For i = 1 To .Rows - 1
                            For j = 1 To .Cols - 1 Step 2
                                If .TextMatrix(i, j) <> "" Then
                                    lng预约数 = lng预约数 + Val(.TextMatrix(i, j))
                                End If
                            Next
                        Next
                    End With
                End If
                If lng预约数 > lng限约数 Then
                   MsgBox "在分时段页面中的『" & str限制项目 & "』所设置的预约数(" & lng预约数 & ")大于了" & IIf(lng限号数 = lng限约数, "限号数(" & lng限约数 & ")", "限约数(" & lng限约数 & ")") & ",你不能按当前设置保存!", vbInformation, Me.Caption
                    If lngIndex = -1 Then
                        Call zlSaveTimePageSelected(str限制项目)
                        mTimeSet.blnNotBrush = True
                        tbSubPage.Item(lngPage).Selected = True
                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                        mTimeSet.blnNotBrush = False
                    End If
                   Exit Function
                End If
                If lng预约数 < lng限约数 And lng预约数 <> 0 Then
                    If lngIndex = -1 Then
                           If blnAllowYYNums = False Then
                                   strCommand = zlCommFun.ShowMsgbox("提醒", "    在分时段页面中的『" & str限制项目 & "』所设置的实际预约数(" & lng预约数 & ") 与限约数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?" & vbCrLf & vbCrLf & _
                                    "『是』:表示允许限约数与预约数不一致" & vbCrLf & vbCrLf & _
                                    "『忽略』:表示允许限约数与预约数不一致，遇到类似的问题,不再提示。" & vbCrLf & vbCrLf & _
                                    "『否』:表示不允许限约数与预约数不一致,返回重新设置" & vbCrLf, "是(&O),忽略(&I),否(&C)", Me, vbQuestion)
                                    Select Case strCommand
                                    Case "是"
                                    Case "忽略"
                                        blnAllowYYNums = True
                                    Case Else
                                       Call zlSaveTimePageSelected(str限制项目)
                                       mTimeSet.blnNotBrush = True
                                       tbSubPage.Item(lngPage).Selected = True
                                       If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                       mTimeSet.blnNotBrush = False
                                        Exit Function
                                    End Select
                           End If
'                      Else
'                            If MsgBox("在分时段页面中的『" & str限制项目 & "』所设置的实际预约数(" & lng预约数 & ") 与限约数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
        End If
    Next
    VsTimeValidate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckUsedCount() As Boolean
    '检查存在预约记录的安排 限号、限约，医技上班时段
    Dim var限制项目 As Variant, str限制项目 As String
    Dim lng已约数 As Long, lng最大序号 As Long
    Dim var原排班 As Variant, var星期 As Variant
    Dim i As Long, k As Long
    Dim lng限约数 As Long, lng限号数 As String, str上班时段 As String
    
    On Error GoTo ErrHandler
    Call LoadRegHistory
    If mrsRegHistory.RecordCount = 0 Then CheckUsedCount = True: Exit Function
    
    var限制项目 = Array("", "周日", "周一", "周二", "周三", "周四", "周五", "周六")
    lng限号数 = Val(txt限号.Text)
    lng限约数 = Val(txt限约.Text)
    str上班时段 = cbo天.Text

    For i = 1 To 7
        If opt天.Value = False Then
            lng限号数 = Val(vsPlan.TextMatrix(2, i))
            lng限约数 = Val(vsPlan.TextMatrix(3, i))
            str上班时段 = vsPlan.TextMatrix(1, i)
        End If
        lng已约数 = 0: lng最大序号 = 0
    
        str限制项目 = var限制项目(i)
        mrsRegHistory.Filter = "限制项目='" & str限制项目 & "'"
        If mrsRegHistory.RecordCount <> 0 Then
            lng已约数 = Val(Nvl(mrsRegHistory!统计))
            If lng已约数 > lng限号数 Then
               Call MsgBox(IIf(opt天.Value, "", str限制项目) & "限号数小于了" & IIf(opt天.Value, str限制项目, "") & "已经预约出去的数量[" & lng已约数 & "]，不能继续！", vbInformation, gstrSysName)
               Exit Function
            End If
            If lng已约数 > lng限约数 Then
               Call MsgBox(IIf(opt天.Value, "", str限制项目) & "限约数小于了" & IIf(opt天.Value, str限制项目, "") & "已经预约出去的数量[" & lng已约数 & "]，不能继续！", vbInformation, gstrSysName)
               Exit Function
            End If
            lng最大序号 = Val(Nvl(mrsRegHistory!最大序号))
            If lng最大序号 > lng限号数 Then
               Call MsgBox(IIf(opt天.Value, "", str限制项目) & "限号数小于了" & IIf(opt天.Value, str限制项目, "") & "已经预约出去的最大序号[" & lng最大序号 & "]，不能继续！", vbInformation, gstrSysName)
               Exit Function
            End If
            
            If lng已约数 > 0 Then
                If InStr(mstr原排班, ",") = 0 Then '"周一,上午;周二,下午;..."
                    '原来是“按天”
                    If str上班时段 <> mstr原排班 Then
                        Call MsgBox("当前计划生效时间内" & str限制项目 & "存在已经预约出去的挂号记录，不能修改排班！", vbInformation, gstrSysName)
                        Exit Function
                    End If
                Else
                    '原来是“按周”
                    var原排班 = Split(mstr原排班, ";")
                    For k = 0 To UBound(var原排班)
                        var星期 = Split(var原排班(k), ",")
                        If str限制项目 = var星期(0) Then
                            If var星期(1) <> "" And str上班时段 <> var星期(1) Then
                                Call MsgBox("当前计划生效时间内" & str限制项目 & "存在已经预约出去的挂号记录，不能修改排班！", vbInformation, gstrSysName)
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Next
    CheckUsedCount = True
    Exit Function
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function SaveVerify(ByVal lng上次计划ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:审核挂号安排计划
    '返回:审核成功,返回true, 否则返回False
    '编制:刘兴洪
    '日期:2009-09-14 17:11:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    Err = 0: On Error GoTo Errhand
    If CheckUsedCount() = False Then Exit Function
    
    'Zl_挂号安排计划_Verify(Id_In In 挂号安排计划.ID%Type,立即生效_in Number:=0)
    strSQL = "Zl_挂号安排计划_Verify(" & Val(mstr计划ID) & "," & chk立即生效.Value & "," & ZVal(lng上次计划ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveVerify = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function
Private Function SaveCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消审核挂号安排计划
    '返回:取消审核成功,返回true, 否则返回False
    '编制:刘兴洪
    '日期:2009-09-14 17:11:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsPlan As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "Select 1 From 挂号安排计划 Where 上次计划ID = [1]"
    Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, "检查调整计划", Val(mstr计划ID))
    If rsPlan.RecordCount > 0 Then
        MsgBox "当前计划有变更计划，不能取消审核。", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_挂号安排计划_Cancel(Id_In In 挂号安排计划.ID%Type) Is
    strSQL = "Zl_挂号安排计划_Cancel(" & Val(mstr计划ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveCancel = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
     SaveErrLog
End Function
Private Function SaveDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消审核挂号安排计划
    '返回:取消审核成功,返回true, 否则返回False
    '编制:刘兴洪
    '日期:2009-09-14 17:11:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo Errhand:
    'Zl_挂号安排计划_Delete(Id_In In 挂号安排计划.ID%Type) Is
    strSQL = "Zl_挂号安排计划_Delete(" & Val(mstr计划ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveDelete = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
     SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim lng上次计划ID As Long
    If mblnSaveMinorChange Then Call SaveMinorChange: Exit Sub
    If mEditType = ed_安排查阅 Then Unload Me: Exit Sub
    If mEditType = Ed_安排删除 Then
        If SaveDelete = False Then Exit Sub
        mblnSucces = True
        Unload Me: Exit Sub
    End If
    
    If mEditType = Ed_安排审核 Then
        If LongPlanIsValied(lng上次计划ID) = False Then Exit Sub
        If SaveVerify(lng上次计划ID) = False Then Exit Sub
        mblnSucces = True
        Unload Me: Exit Sub
    End If
    
    If mEditType = Ed_安排取消 Then
        If SaveCancel = False Then Exit Sub
        mblnSucces = True
        Unload Me: Exit Sub
    End If
    
    If IsValied(True) = False Then Exit Sub
    If LongPlanIsValied(lng上次计划ID) = False Then Exit Sub
    If SavePlan(lng上次计划ID) = False Then Exit Sub
    mblnSucces = True
    Unload Me
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With cmdOK
        .Left = ScaleWidth - .Width - 100
        cmdCancel.Left = .Left
        cmdHelp.Left = .Left
    End With
    With tbPage
        .Top = 50
        .Height = ScaleHeight - 100
        .Left = 50
        .Width = cmdOK.Left - .Left - 100
    End With
End Sub

Private Sub zlSaveTimePageSelected(ByVal str星期 As String)
    If tbPage.Selected Is Nothing Then Exit Sub
    If tbPage.Selected.index <> mPageIndex.EM_时段 Then
         tbPage.Item(mPageIndex.EM_时段).Selected = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ClearCustomData
End Sub

Private Sub mfrmOtherCalc_zlRefreshCon(ByVal varTimes As Variant)
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng限号 As Long
    Dim lng限约 As Long
    Dim dat开始时间 As Date
    Dim dat结束时间 As Date
    Dim lng序号 As Long
    Dim strTmp As String
    Dim str时段 As String
    Dim str限制时间 As String
    Dim lng默认间隔 As Long
    Dim lng分配个数 As Long
    Dim lng固定数量 As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim dat时点 As Date
    Dim str分段间隔 As String
    Dim str限制项目 As String
    Dim cllPro As Collection
    Dim varTemp As Variant
    Dim strStart As String
    Dim strEnd As String
    Dim int分钟 As Integer
    Dim str时点 As String
    Dim lng时间间隔 As Long
    Dim varData As Variant
    Dim str首时点 As String
    
    If Not mTimeSet.bln序号控制 Then Exit Sub
    If varTimes Is Nothing Then Exit Sub
    If varTimes("时间间隔") <> "" Then
        txtTimeOut.Text = Val(varTimes("时间间隔"))
        Call cmd设置时段_Click
        Exit Sub
    End If

    str分段间隔 = varTimes("分段间隔")
    If Trim(str分段间隔) = "" Then Exit Sub


    If mrs上班时间段 Is Nothing Then
        Call Init时间段
    End If
    str限制项目 = GetVsGridCaption(mTimeSet.lngSelIndex)


    If mrs上班时间段 Is Nothing Then Exit Sub
    mTimeSet.rsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
    If mTimeSet.rsRegPlan.RecordCount = 0 Then mTimeSet.rsRegPlan.Filter = 0: Exit Sub
    lng限号 = Nvl(mTimeSet.rsRegPlan!限号数, 0): lng限约 = Nvl(mTimeSet.rsRegPlan!限约数, 0)
    If lng限约 = 0 Then lng限约 = lng限号
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbInformation, Me.Caption
        Exit Sub
    End If


    str时段 = mTimeSet.rsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbInformation, Me.Caption
        Exit Sub

    End If

    Set cllPro = New Collection
    varData = Split(str分段间隔, ";")

    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int分钟 = Val(varTemp(1))
        varTemp = Split(varTemp(0), "～")
        strStart = varTemp(0)
        strEnd = varTemp(1)
        cllPro.Add int分钟, "K" & Replace(strStart, ":", "_")
        cllPro.Add strStart, "K" & Replace(strStart, ":", "_") & "_Start"
        cllPro.Add strEnd, "K" & Replace(strStart, ":", "_") & "_End"
    Next

    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "' And 已使用=0"
    Do While Not mTimeSet.rsAssign.EOF
        mTimeSet.rsAssign.Delete adAffectCurrent
        mTimeSet.rsAssign.MoveNext
    Loop
    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "'"
    If mTimeSet.rsAssign.RecordCount <> 0 Then
        lng固定数量 = mTimeSet.rsAssign.RecordCount
        lng默认间隔 = Val(Nvl(mTimeSet.rsAssign!时间间隔, lng时间间隔))
        lng时间间隔 = lng默认间隔
        Do While Not mTimeSet.rsAssign.EOF
            lng分配个数 = lng分配个数 + Val(Nvl(mTimeSet.rsAssign!限制数量))
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    mTimeSet.rsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        dat时点 = dat开始时间
        mrs上班时间段.MoveNext

        For i = j To lng限号
            If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                j = i
                Exit For
            End If
            If str时点 <> Format(dat时点, "HH:00") Then
                If str时点 = "" Or Format(str首时点, "HH:00") = Format(dat时点, "HH:00") Then
                    If str时点 = "" Then str首时点 = Format(dat时点, "HH:MM")
                    str时点 = Format(dat时点, "HH:MM")
    
                    If InStr("," & str分段间隔, str首时点 & "～") > 0 Then
                        lng时间间隔 = Val(cllPro("K" & Replace(str首时点, ":", "_")))
                    Else
                        lng时间间隔 = lng默认间隔
                    End If
                Else
                    str时点 = Format(dat时点, "HH:00")
                    '问题号:115865,焦博,2017/10/31,门诊挂号安排管理时段设置时修改时间间隔为"0"时报错
                    If InStr("," & str分段间隔, str时点 & "～") > 0 Then
                        lng时间间隔 = Val(cllPro("K" & Replace(str时点, ":", "_")))
                    Else
                        lng时间间隔 = lng默认间隔
                    End If
                End If
            End If
            If lng时间间隔 = 0 Then
                '问题号:119348,焦博,2018/1/9,新增安排，时段设置使用其他辅助计算，其中某个时间刻度的时间间隔为0，辅助计算后时段信息错误
                dat时点 = DateAdd("h", 1, dat时点)
                i = i - 1
            Else
                If i > lng固定数量 Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        If Format(DateAdd("n", lng时间间隔, dat时点), "yyyy-MM-dd hh:mm:00") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                            !结束时间 = Format(dat结束时间, "hh:mm:ss")
                            !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                        Else
                            !结束时间 = Format(DateAdd("n", lng时间间隔, dat时点), "hh:mm:00")
                            !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng时间间隔, dat时点), "hh:mm")
                        End If
                        !时间间隔 = lng时间间隔
                        !限制数量 = IIf(lng分配个数 >= lng限号, 0, 1)
                        !是否预约 = 0
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限号, 0, 1)
                    End With
                Else
                    mTimeSet.rsAssign.Filter = "序号=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mTimeSet.rsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng时间间隔
                    End If
                End If
                dat时点 = DateAdd("n", IIf(i > lng固定数量, lng时间间隔, lng默认间隔), dat时点)
            End If
        Next
        If i > lng限号 And mTimeSet.bln序号控制 Then
            blnExit = True
        End If
    Loop
    Call tbSubPage_SelectedChanged(tbSubPage(mTimeSet.lngSelIndex))
End Sub

Private Sub vsTime_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str时段() As String
     If mTimeSet.bln序号控制 Then
        str时段 = Split(vsTime(index).EditText, "-")
        If UBound(str时段) <> 1 Then
           MsgBox "输入的时间格式有误!请检查!", vbInformation, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(str时段(0)) Then
           MsgBox "输入的时间格式有误!请检查!", vbInformation, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(str时段(1)) Then
           MsgBox "输入的时间格式有误!请检查!", vbInformation, gstrSysName
           Cancel = True: Exit Sub
        End If
        If CDate(str时段(0)) >= CDate(str时段(1)) Then
           MsgBox "开始时间必须小于结束时间!请检查!", vbInformation, gstrSysName
           Cancel = True
        End If
     End If
    mTimeSet.blnChange = True
End Sub

Private Sub cmdOther_Click()
    Dim str安排 As String
    If Not mTimeSet.bln序号控制 Then Exit Sub
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    Call mfrmOtherCalc.zlShowMe(Me, Nvl(mTimeSet.rsRegPlan!排班), Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing '
End Sub

Private Sub vsTime_BeforeRowColChange(index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mTimeSet.bln序号控制 Then
        vsTime(index).Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
          cmd预约(index).Visible = False: Exit Sub
    End If
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    
    SetCtrlMove index, NewRow - (NewRow) Mod 2, NewCol
    If mTimeSet.bln序号控制 Then
        If vsTime(index).Cell(flexcpFontUnderline, NewRow, NewCol) = False And vsTime(index).Cell(flexcpBackColor, NewRow, NewCol) = 0 Then
            vsTime(index).Editable = flexEDKbdMouse
        Else
            vsTime(index).Editable = flexEDNone
        End If
        Exit Sub
    End If
    
    With vsTime(index)
        .Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
    End With
End Sub

Private Sub vsTime_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If Not mTimeSet.bln序号控制 Then Exit Sub
     
    With vsTime(index)
           
        If (.Row < 0 Or .Col < 1) Or (.Row > .Rows - 1 Or .Col > .Cols - 1) Then Exit Sub '没在有效单元格内
        If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
        If KeyCode = 13 Then
            Call cmd预约_Click(index)
            Exit Sub
        End If
        
        If KeyCode = 46 Then
            If cmd删除(index).Visible = False Then Exit Sub
            If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
            Call cmd删除_Click(index)
        End If
     End With
End Sub

Private Sub vsTime_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsTime_LostFocus(index As Integer)
 If Trim(vsTime(index).EditText) <> "" Then
    With vsTime(index)
        .TextMatrix(.Row, .Col) = .EditText
        mTimeSet.blnChange = True
    End With
 End If
End Sub

Private Sub cmdClearAll_Click()
    If Not mTimeSet.bln序号控制 Or mTimeSet.lngSelIndex < 0 Then Exit Sub
    With vsTime(mTimeSet.lngSelIndex)
        If .Rows = 0 Then Exit Sub
        .Cell(flexcpForeColor, 0, 1, .Rows - 1, .Cols - 1) = &H80000008
        .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = False
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub cmdSelAll_Click()
    If Not mTimeSet.bln序号控制 Or mTimeSet.lngSelIndex < 0 Then Exit Sub
    With vsTime(mTimeSet.lngSelIndex)
        If .Rows = 0 Then Exit Sub
        .Cell(flexcpForeColor, 0, 1, .Rows - 1, .Cols - 1) = vbBlue
        .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub SetCtrlMove(ByVal index As Integer, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnDel As Boolean
    With vsTime(index)
        If mTimeSet.bln序号控制 Then
            If Trim(.TextMatrix(NewRow, NewCol)) = "" Then
                cmd删除(index).Visible = False
                cmd预约(index).Visible = False
                Exit Sub
            End If
            cmd删除(index).Left = .Cell(flexcpLeft, NewRow, NewCol) + .Cell(flexcpWidth, NewRow, NewCol) - cmd删除(index).Width
            If .Row Mod 2 <> 0 Then
                cmd删除(index).Top = .Cell(flexcpTop, NewRow, NewCol)
            Else
                cmd删除(index).Top = .Cell(flexcpTop, NewRow, NewCol)
            End If
            cmd预约(index).Left = .Cell(flexcpLeft, NewRow, NewCol)
            cmd预约(index).Top = cmd删除(index).Top
            If NewCol < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(NewRow, NewCol + 1)) = ""
            Else
                blnDel = True
            End If
             
            blnDel = blnDel And Trim(.TextMatrix(NewRow, NewCol)) <> "" And Not .Cell(flexcpFontUnderline, NewRow, NewCol)
            cmd删除(index).Visible = blnDel And mTimeSet.bln序号控制
            cmd预约(index).Visible = True 'Val(txt限约.Text) <> 0
        Else
            cmd预约(index).Left = .Cell(flexcpTop, NewRow, NewCol)
            cmd预约(index).Top = .Cell(flexcpLeft, NewRow, NewCol)
            cmd预约(index).Visible = False
        End If
    End With
End Sub

Private Sub cmd预约_Click(index As Integer)
    Dim i As Integer, j As Integer
    Dim intStartRow As Integer, intEndRow As Integer, intStartCol As Integer, intEndCol As Integer
    If Not mTimeSet.bln序号控制 Or mTimeSet.lngSelIndex < 0 Then Exit Sub
    If chkAppoint.Value = 0 Then Exit Sub
    If mTimeSet.lngSelIndex <> index Then Exit Sub
    With vsTime(mTimeSet.lngSelIndex)
'        If .MouseRow < 0 Or .MouseCol < 0 Then Exit Sub
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If .Row > .RowSel Then
            intStartRow = .RowSel
            intEndRow = .Row
        Else
            intStartRow = .Row
            intEndRow = .RowSel
        End If
        If .Col > .ColSel Then
            intStartCol = .ColSel
            intEndCol = .Col
        Else
            intStartCol = .Col
            intEndCol = .ColSel
        End If
        For i = intStartRow To intEndRow Step 2
            For j = intStartCol To intEndCol
                If i <= .Rows - 1 And j <= .Cols - 1 Then
                    If .Cell(flexcpForeColor, i, j) = vbBlue Then
                       .Cell(flexcpForeColor, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = &H80000008
                        .Cell(flexcpFontBold, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = False
                    Else
                        .Cell(flexcpForeColor, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = vbBlue
                        .Cell(flexcpFontBold, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = True
                    End If
                End If
            Next j
        Next i
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub cmd删除_Click(index As Integer)
    Dim blnDel As Boolean
    Dim lngSelX As Long
    Dim lngSelY As Long
    Dim i As Long, j As Long
    Dim lngCurrSn As Long
    Dim lngStartCol As Long
    With vsTime(index)
        If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
        Else
                blnDel = True
        End If
        blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> "" And Not .Cell(flexcpFontUnderline, .Row, .Col)
        If Not blnDel Then Exit Sub
        If mTimeSet.bln序号控制 Then
          lngSelX = .Row - (.Row Mod 2): lngSelY = .Col
          lngCurrSn = Val(.TextMatrix(lngSelX, lngSelY))
          .TextMatrix(lngSelX, lngSelY) = ""
          .TextMatrix(lngSelX + 1, lngSelY) = ""
          
          For i = lngSelX To .Rows - 1 Step 2
            lngStartCol = 1
            If i = lngSelX Then lngStartCol = lngSelY
            For j = lngStartCol To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    .TextMatrix(i, j) = lngCurrSn
                     lngCurrSn = lngCurrSn + 1
                End If
            Next
         Next
        End If
        cmd删除(index).Visible = False
        cmd预约(index).Visible = False
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub picPage_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picPage(index)
        vsTime(index).Left = .ScaleLeft
        vsTime(index).Top = .ScaleTop
        vsTime(index).Width = .ScaleWidth
        vsTime(index).Height = .ScaleHeight
    End With
End Sub

Private Sub opt分诊_Click(index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    '71253 李南春 2014-04-15 14:23:10 将listView 替换为vsflexGrid
    If index <> 1 Then Exit Sub
    With vsDept
    For intCol = 0 To .Cols - 1
        For intRow = 0 To .Rows - 1
            If .Cell(flexcpChecked, intRow, intCol) = 1 Then
                Call ClearVsGridCheckValue
                .Row = intRow: .Col = intCol
                .Cell(flexcpChecked, intRow, intCol) = 1
                Exit Sub
            End If
        Next
    Next
    End With
End Sub

Private Sub vsPlan_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    If Row = 1 Then FinishEdit = True
End Sub

Private Sub opt生效时间_Click(index As Integer)
     dtpBegin.Enabled = opt生效时间(0).Value = False
     
     If opt生效时间(0).Value = True Then
        chk立即审核.Value = 1
     End If
End Sub
Private Sub opt天_Click()
    Dim i As Integer
    Dim strPlan As String
    Dim ctl As Control
    
    With vsPlan
        For i = 1 To .Cols - 1
            If Trim(.TextMatrix(1, i)) <> "" Then
                If strPlan = "" Then
                    strPlan = .TextMatrix(1, i)
                Else
                    If .TextMatrix(1, i) <> strPlan Then
                        strPlan = "": Exit For
                    End If
                End If
            End If
        Next
        For i = 1 To .Cols - 1
            .TextMatrix(1, i) = ""
            .TextMatrix(2, i) = ""
            .TextMatrix(3, i) = ""
        Next
        .Enabled = False: .TabStop = False
    End With
    opt天.Value = -True: txt限号.Enabled = True: txt限约.Enabled = (chkAppoint.Value = 1)
    cbo天.Enabled = True
    opt周.Value = False
    cbo天.ListIndex = cbo.FindIndex(cbo天, strPlan, True)
    cbo天.SetFocus

    '设置编辑背景色
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX", UCase("ComboBox")
            Call zlSetCtrolBackColor(ctl)
        Case UCase("ListView")
        Case UCase("DTPicker")
        Case Else
        End Select
    Next
End Sub

Private Sub opt周_Click()
    Dim i As Integer
    Dim ctl As Control
    
    If Trim(cbo天.Text) <> "" Then
        With vsPlan
            For i = 1 To .Cols - 1
                .TextMatrix(1, i) = cbo天.Text
                .TextMatrix(2, i) = txt限号.Text
                .TextMatrix(3, i) = txt限约.Text
            Next
            .Enabled = True: .TabStop = True
            .Col = 1: .SetFocus
        End With
    End If
    opt天.Value = False: txt限号.Enabled = False: txt限约.Enabled = False
    cbo天.Enabled = False: cbo天.ListIndex = -1
    opt周.Value = True: vsPlan.Enabled = True

    '设置编辑背景色
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX", UCase("ComboBox")
            Call zlSetCtrolBackColor(ctl)
        Case UCase("ListView")
        Case UCase("DTPicker")
        Case Else
        End Select
    Next
End Sub

Private Sub picTimeSet_Resize()
    Err = 0: On Error Resume Next
    With fra应用于
        .Top = picTimeSet.ScaleHeight - .Height - 50
        .Width = picTimeSet.ScaleWidth
        .Left = picTimeSet.ScaleLeft
        .Visible = True
    End With
    With tbSubPage
        .Top = txtTimeOut.Top + txtTimeOut.Height + 50
        .Left = picTimeSet.ScaleLeft
        .Width = picTimeSet.ScaleWidth
        .Height = fra应用于.Top - .Top - 100
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    PageChange Item
End Sub

Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    If Item.index = mPageIndex.EM_时段 Then
       mblnChangeByCode = True
       tbPage.Item(mPageIndex.EM_计划).Selected = True
        If IsValied() = False Then
            mblnChangeByCode = False
            Exit Sub
        End If
        tbPage.Item(mPageIndex.EM_时段).Selected = True
        mblnChangeByCode = False
        Call LoadTimePlan
        If mTimeSet.bln序号控制 = True Then
            cmdSelAll.Enabled = True
            cmdSelAll.Visible = True
            cmdClearAll.Enabled = True
            cmdClearAll.Visible = True
        Else
            cmdSelAll.Enabled = False
            cmdSelAll.Visible = False
            cmdClearAll.Enabled = False
            cmdClearAll.Visible = False
        End If
        mTimeSet.lngSelIndex = tbSubPage.Selected.index
    Else
        If mTimeSet.blnChange = False Then Exit Sub
        If zl_CheckMoveAssign() = False Then
             mblnChangeByCode = True
            tbPage.Item(mPageIndex.EM_时段).Selected = True
             mblnChangeByCode = False
        End If
    End If
End Sub



Private Sub LoadTimePlan()
    Dim i As Long
    Dim lng限号数 As Long
    Dim lng限约数 As Long
    Dim strTemp As String
    Dim str安排 As String
    Dim str排班 As String
    Dim str应诊时段 As String
    Dim str应诊     As String
    
    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing
    If mrsRegNewData Is Nothing Then
        Set mrsRegNewData = New ADODB.Recordset
        mrsRegNewData.Fields.Append "ID", adBigInt, 18
        mrsRegNewData.Fields.Append "限制项目", adVarChar, 20
        mrsRegNewData.Fields.Append "排班", adVarChar, 20
        mrsRegNewData.Fields.Append "限号数", adBigInt, 10
        mrsRegNewData.Fields.Append "限约数", adBigInt, 18
        mrsRegNewData.Fields.Append "序号控制", adBigInt, 18
        mrsRegNewData.CursorLocation = adUseClient
        mrsRegNewData.LockType = adLockOptimistic
        mrsRegNewData.CursorType = adOpenStatic
        mrsRegNewData.Open
     End If
     If opt天.Value = True Then
          lng限号数 = Val(txt限号.Text)
          lng限约数 = Val(txt限约.Text)
          str排班 = Me.cbo天.Text
          For i = 0 To 6
            strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
            '周一,限号数,限约数|周二,限号数,限约数|....
            str安排 = str安排 & "|" & strTemp & "," & lng限号数 & "," & lng限约数
             With mrsRegNewData
                .AddNew
                !ID = Val(mstr计划ID)
                !限制项目 = strTemp
                !排班 = str排班
                !限号数 = lng限号数
                !限约数 = lng限约数
                !序号控制 = Me.chk序号控制.Value
                .Update
            End With
            If InStr("|" & mPlanInfo.str应诊时段 & "|", "|" & strTemp & "-" & Trim(str排班) & "|") > 0 Then
                '如果没有改变当天的排班信息,则保持原来时段不变,
                str应诊时段 = str应诊时段 & "|" & strTemp
            End If
            str应诊 = str应诊 & "|" & strTemp & "-" & str排班
          Next
        Else
           With vsPlan
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTemp = Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                    lng限号数 = Val(Trim(vsPlan.TextMatrix(2, i)))
                    lng限约数 = Val(Trim(vsPlan.TextMatrix(3, i)))
                    str排班 = Trim(vsPlan.TextMatrix(1, i))
                    str安排 = str安排 & "|" & strTemp & "," & lng限号数 & "," & lng限约数
                    With mrsRegNewData
                        .AddNew
                        !ID = Val(mstr计划ID)
                        !限制项目 = strTemp
                        !排班 = str排班
                        !限号数 = lng限号数
                        !限约数 = lng限约数
                        !序号控制 = Me.chk序号控制.Value
                        .Update
                    End With
                    If InStr("|" & mPlanInfo.str应诊时段 & "|", "|" & strTemp & "-" & Trim(str排班) & "|") > 0 Then
                        '如果没有改变当天的排班信息,则保持原来时段不变,
                        str应诊时段 = str应诊时段 & "|" & strTemp
                    End If
                    str应诊 = str应诊 & "|" & strTemp & "-" & str排班
                End If
               
            Next
        End With
     End If
     If str安排 <> "" Then str安排 = Mid(str安排, 2)
     If str应诊时段 <> "" Then str应诊时段 = Mid(str应诊时段, 2)
     If str应诊 <> "" Then str应诊 = Mid(str应诊, 2)
     mPlanInfo.str应诊时段 = str应诊
'Public Enum mRegEditType
'Ed_计划安排 = 0
'Ed_安排修改 = 1
'Ed_安排删除 = 2
'Ed_安排审核 = 3
'Ed_安排取消 = 4
'Ed_安排查阅 = 5
'End Enum
     
     '增加计划,暂时不对已经预约部分mfrmTime.zlShowPagePlan str安排, mrsRegNewData, mrsRegHistory, chk序号控制.Value = 1, Switch(mEditType = ed_计划安排, EM_计划_增加, mEditType = Ed_安排修改, EM_计划_修改, True, EM_计划_查阅), mlng安排ID, Val(mstr计划ID)
     zlShowPagePlan str安排, mrsRegNewData, Nothing, chk序号控制.Value = 1, Switch(mEditType = ed_计划安排, EM_计划_增加, mEditType = Ed_安排修改, EM_计划_修改, True, EM_计划_查阅), mlng安排ID, Val(mstr计划ID), , str应诊时段
End Sub

Private Sub zlShowPagePlan(ByVal str安排类别 As String, ByVal rsRegPlan As ADODB.Recordset, ByRef rsHistory As ADODB.Recordset, _
                        ByVal bln序号控制 As Boolean, ByVal bytType As gPlanEditType, Optional ByVal lng安排ID As Long, _
                        Optional ByVal lng计划ID As Long, Optional ByVal blnBeforCheck As Boolean = False, Optional ByVal str应诊时段 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示页面
    '编制:刘兴洪
    '日期:2012-06-15 13:49:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    mTimeSet.str安排 = str安排类别
    mTimeSet.str应诊时段 = str应诊时段
                                                                          
    Set mTimeSet.rsRegPlan = rsRegPlan
    If bln序号控制 <> mTimeSet.bln序号控制 And Not mTimeSet.rsAssign Is Nothing Then
         mTimeSet.rsAssign.Filter = 0
         Do While Not mTimeSet.rsAssign.EOF
            mTimeSet.rsAssign.Delete
            mTimeSet.rsAssign.MoveNext
         Loop
         If blnBeforCheck Then Exit Sub
    End If
    mPlanEditType = bytType: mTimeSet.lng安排ID = lng安排ID: mTimeSet.lng计划ID = lng计划ID
    Set mTimeSet.rsHistory = rsHistory
    If Not blnBeforCheck Then Call ShowTimeSetPage
    If mTimeSet.blnIsInit Then
        Call AssignManage
    End If
    mTimeSet.blnIsInit = True
    Call InitRs(mTimeSet.bln序号控制 = bln序号控制)
    mTimeSet.bln序号控制 = bln序号控制
    If blnBeforCheck Then Exit Sub
    For i = 0 To 6
        Call tbSubPage_SelectedChanged(tbSubPage.Item(i))
    Next
 End Sub
 
 Private Sub tbSubPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim str限制项目 As String
   If Not mTimeSet.blnIsInit Then Exit Sub
   
   If Item.index <> mTimeSet.lngSelIndex And mTimeSet.lngSelIndex <> -1 Then '
'     If mTimeSet.lngSelIndex <> -1 And mTimeSet.blnChange Then
'        If VsTimeValidate(mTimeSet.lngSelIndex) = False Then
'            mTimeSet.blnOnChange = True
'            tbSubPage.Item(mTimeSet.lngSelIndex).Selected = True
'            mTimeSet.blnOnChange = False
'            Exit Sub
'        End If
'     End If
     
     str限制项目 = GetVsGridCaption(mTimeSet.lngSelIndex)
     If MoveAssign(str限制项目) = False Then
        If mTimeSet.lngSelIndex <> -1 Then tbSubPage.Item(mTimeSet.lngSelIndex).Selected = True
        Exit Sub
     End If
   End If
   
   If mTimeSet.blnOnChange Then Exit Sub
   mTimeSet.lngSelIndex = Item.index
   SetStyle mTimeSet.bln序号控制, Item.index
   
   LoadTimeSetPlan Item.Caption
   setVsGridSNStyle Item.index
End Sub
 
 Private Sub setVsGridSNStyle(ByVal lngIndex As Long)
 '如果分时段在vsFex表哥填充好数据后需要重新设置表哥样式
 '****************************************
'对表格样式进行设置
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
   
    If vsTime(lngIndex).Cols <= 1 Then Exit Sub
    If mTimeSet.bln序号控制 Then
        With vsTime(lngIndex)
            For i = 1 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
             Next
             .ColWidth(0) = 1200
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
             If .Rows > 0 Then
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
             End If
    '对时间段设置间隔背景
         End With
    Else
    
    End If
    With vsTime(lngIndex)
         If (mTimeSet.bln序号控制 And .Rows = 0) Or (mTimeSet.bln序号控制 = False And .Rows = 1) Then Exit Sub
         For i = IIf(mTimeSet.bln序号控制, 0, 1) To .Rows - 1 Step 2
             .Cell(flexcpBackColor, i, IIf(mTimeSet.bln序号控制, 1, 0), i, .Cols - 1) = &HE0E0D3
         Next
    End With

End Sub
 
Private Function LoadTimeSetPlan(ByVal str限制项目 As String) As Boolean
    Dim nIndex As Integer
    Dim i As Long, r As Long
    Dim strTime As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim str时点 As String
    Dim strData As String
    If mTimeSet.rsAssign Is Nothing Then Exit Function
    nIndex = GetVsGridIndex(str限制项目)
    cmd预约(nIndex).Visible = False
    cmd删除(nIndex).Visible = False
    If Not mTimeSet.bln序号控制 Then
        With vsTime(nIndex)
             
             mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "'"
            mTimeSet.rsAssign.Sort = "序号 asc "
               r = 1: i = -1
            Do While Not mTimeSet.rsAssign.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mTimeSet.rsAssign!限制数量))
                strTime = mTimeSet.rsAssign!时间段
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                If Val(Nvl(mTimeSet.rsAssign!已使用)) = 1 Then
                    .Cell(flexcpFontUnderline, r, i * 2, r, i * 2 + 1) = True
                Else
                   '不做颜色处理
                End If
                mTimeSet.rsAssign.MoveNext
            Loop
             mTimeSet.rsAssign.Filter = 0
        End With
        LoadTimeSetPlan = True
        Exit Function
    End If
    
    '-序号控制
    With vsTime(nIndex)
        .Cols = 1: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        .Cols = 2: .Clear
        lngRow = -1: lngCol = 0
        mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "'"
        If mTimeSet.rsAssign.RecordCount = 0 Then mTimeSet.rsAssign.Filter = 0: Exit Function
        i = 1
        mTimeSet.rsAssign.Sort = "序号 asc "
        Do While Not mTimeSet.rsAssign.EOF
             lngCol = lngCol + 1
             If str时点 <> Nvl(mTimeSet.rsAssign!时点) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                str时点 = Nvl(mTimeSet.rsAssign!时点)
                If lngRow > .Rows - 1 Then .Rows = .Rows + 2
                 .TextMatrix(lngRow - 1, 0) = Format(str时点, "hh:mm")
                 .TextMatrix(lngRow, 0) = Format(str时点, "hh:mm")
             End If
             strData = mTimeSet.rsAssign!序号
             strTime = mTimeSet.rsAssign!时间段
            If lngCol > .Cols - 1 Then .Cols = .Cols + 1
            If lngRow > .Rows - 1 Then .Rows = .Rows + 2
             .TextMatrix(lngRow - 1, lngCol) = strData
             .TextMatrix(lngRow, lngCol) = strTime
            If Val(Nvl(mTimeSet.rsAssign!是否预约)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            If Val(Nvl(mTimeSet.rsAssign!已使用)) = 1 Then
                    .Cell(flexcpFontUnderline, lngRow - 1, lngCol, lngRow, lngCol) = True
            Else
               '不做颜色处理
            End If
            mTimeSet.rsAssign.MoveNext
        Loop
        If .Rows = 0 Then .Rows = 1
    End With
End Function
 
Private Sub ClearCustomData()
     mTimeSet.str安排 = ""
     mTimeSet.bln序号控制 = False
     mTimeSet.lngSelIndex = 0
     mTimeSet.blnOnChange = False
     mTimeSet.lng安排ID = 0
     mTimeSet.lng计划ID = 0
     mTimeSet.blnIsInit = False
     Set mTimeSet.rsRegPlan = Nothing
     Set mTimeSet.rsAssign = Nothing
     mTimeSet.blnChange = False
     Set mTimeSet.rsHistory = Nothing
End Sub
 
Private Sub SetStyle(ByVal bln序号控制 As Boolean, ByVal lngIndex As Long)
    '设置
    Dim i As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    If lngIndex > vsTime.UBound Then Exit Sub
    If Not mTimeSet.blnIsInit Then Exit Sub
    With vsTime(lngIndex)
        If bln序号控制 Then
            If .Cols <= 1 Then Exit Sub
            .Rows = 0
            .FixedCols = 1
            .MergeCellsFixed = flexMergeFree
            .MergeCol(0) = True
            .FixedAlignment(0) = flexAlignRightTop
            .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
        Else
             .Clear
             .Cols = 8: .Rows = 1
             .MergeCol(0) = False
            .FixedCols = 0
            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedRows = 1
            
            .RowHeightMax = 400: .RowHeightMin = 400
            For i = 0 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "时间段"
            Next
            For i = 1 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "预约人数"
            Next
            For i = 0 To .Cols - 1
               .ColAlignment(i) = flexAlignCenterCenter
               .ColWidth(i) = 1200
            Next
        End If
    End With
End Sub
 
 Private Sub InitRs(Optional ByVal blnInitRs As Boolean = True)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If Not mTimeSet.rsAssign Is Nothing Then Exit Sub
    With mTimeSet.rsAssign
        Set mTimeSet.rsAssign = New ADODB.Recordset
        mTimeSet.rsAssign.Fields.Append "限制项目", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "开始时间", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "时点", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "结束时间", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "时间段", adVarChar, 50
        mTimeSet.rsAssign.Fields.Append "时间间隔", adBigInt, 4
        mTimeSet.rsAssign.Fields.Append "限制数量", adBigInt, 10
        mTimeSet.rsAssign.Fields.Append "是否预约", adBigInt, 18
        mTimeSet.rsAssign.Fields.Append "序号", adBigInt, 18
        mTimeSet.rsAssign.Fields.Append "已使用", adBigInt, 2
        mTimeSet.rsAssign.CursorLocation = adUseClient
        mTimeSet.rsAssign.LockType = adLockOptimistic
        mTimeSet.rsAssign.CursorType = adOpenStatic
        mTimeSet.rsAssign.Open
    End With
    If blnInitRs Then Call InitAssignRs
End Sub
 
 Private Function InitAssignRs() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng固定 As Long  '固定的序号不允许更改
    Dim i As Long
    '初始化已分配数据集合
    If mPlanEditType = EM_安排_增加 Then Exit Function
     On Error GoTo Hd
    If mPlanEditType = EM_安排_查阅 Or mPlanEditType = EM_安排_修改 Or mPlanEditType = EM_计划_增加 Then
        strSQL = "Select 序号, 星期 As 限制项目, To_Char(开始时间, 'hh24:mi:ss') As 开始时间, To_Char(结束时间, 'hh24:mi:ss') As 结束时间,"
        strSQL = strSQL & vbCrLf & "         是否预约 , 限制数量,To_Char(开始时间, 'hh24') || ':00:00' As 时点,To_Char(开始时间, 'hh24:mi') || '-' || To_Char(结束时间, 'hh24:mi') As 时间段"
        strSQL = strSQL & vbCrLf & " From 挂号安排时段 Where 安排ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By 星期"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTimeSet.lng安排ID)
    ElseIf mPlanEditType = EM_计划_查阅 Or mPlanEditType = EM_计划_修改 Then
        strSQL = "Select 序号, 星期 As 限制项目, To_Char(开始时间, 'hh24:mi:ss') As 开始时间, To_Char(结束时间, 'hh24:mi:ss') As 结束时间,"
        strSQL = strSQL & vbCrLf & "         是否预约 , 限制数量, To_Char(开始时间, 'hh24') || ':00:00' As 时点,To_Char(开始时间, 'hh24:mi') || '-' || To_Char(结束时间, 'hh24:mi') As 时间段"
        strSQL = strSQL & vbCrLf & " From 挂号计划时段 Where 计划ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By 星期"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTimeSet.lng计划ID)
    End If
    Do While Not rsTmp.EOF
            With mTimeSet.rsAssign
                .AddNew
                !限制项目 = Nvl(rsTmp!限制项目)
                !开始时间 = Nvl(rsTmp!开始时间, "00:00:00")
                !结束时间 = Nvl(rsTmp!结束时间, "00:00:00")
                !时间段 = Nvl(rsTmp!时间段, "__:__-__:__")
                !时间间隔 = DateDiff("n", CDate(!开始时间), CDate(!结束时间))
                !限制数量 = Val(Nvl(rsTmp!限制数量))
                !是否预约 = Val(Nvl(rsTmp!是否预约))
                !时点 = Nvl(rsTmp!时点, "00:00:00")
                !序号 = Val(Nvl(rsTmp!序号))
                lng固定 = 0
                If Not mTimeSet.rsHistory Is Nothing Then
                mTimeSet.rsHistory.Filter = "限制项目='" & Nvl(rsTmp!限制项目) & "'"
                    If mTimeSet.rsHistory.RecordCount > 0 Then
                        If CStr(mTimeSet.rsHistory!发生时间) >= CStr(Nvl(rsTmp!开始时间, "00:00:00")) Then
                            lng固定 = 1
                        End If
                    End If
                End If
                !已使用 = lng固定
                .Update
                
            End With
        rsTmp.MoveNext
    Loop
    Call AssignManage
    InitAssignRs = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub ShowTimeSetPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示页面
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-26 15:21:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    
    For j = 0 To tbSubPage.ItemCount - 1
         tbSubPage(j).Visible = False: tbSubPage(j).Enabled = False
         tbSubPage(j).Selected = False
    Next
    
    On Error GoTo errHandle
    varData = Split(mTimeSet.str安排, "|")
    lngIndex = -1: mTimeSet.lngSelIndex = -1
    For i = 0 To UBound(varData)
        ''周一,限号数,限约数|周二,限号数,限约数|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            For j = 0 To tbSubPage.ItemCount - 1
                If tbSubPage(j).Tag = varTemp(0) Then
                    If lngIndex < 0 Then lngIndex = j
                    tbSubPage(j).Visible = True: tbSubPage(j).Enabled = True
                    p = GetVsGridIndex(varTemp(0))
                    vsTime(p).Tag = varTemp(1) & "," & varTemp(2)
                    If mTimeSet.lngSelIndex = -1 Then mTimeSet.lngSelIndex = j: tbSubPage(j).Selected = True
                End If
            Next
        End If
    Next
    If mTimeSet.lngSelIndex = -1 Then mTimeSet.lngSelIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
 
Private Sub txt号别_GotFocus()
    zlControl.TxtSelAll txt号别
End Sub
Private Sub txt号别_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限号_GotFocus()
    zlControl.TxtSelAll txt限号
End Sub
Private Sub txt限号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限号_Validate(Cancel As Boolean)
    If Trim(txt限号.Text) = "" And Trim(txt限约.Text) <> "" Then
        MsgBox "限约必须限号!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    If Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
        MsgBox "限约数不能小于限号数!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub
Private Sub txt限约_GotFocus()
    zlControl.TxtSelAll txt限约
End Sub

Private Sub txt限约_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If Val(txt限号.Text) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限约_Validate(Cancel As Boolean)
    If Val(txt限号.Text) < Val(txt限约.Text) And _
        Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" Then
        MsgBox "限约数不能小于限号数!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub

Private Function zlCheckRegistPlanIsValied(ByRef blnMulitNumPlan As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前所输入的号码是否合法
    '出参:blnMulitNumPlan-返回是否有多个相同(同一项目,同一科室,同一人,不同号)的安排
    '返回:合法返回,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-29 10:26:45
    '检查规则（同一项目,同一科室,同一人,不同号）:
    '     1.同天内不能有交叉的安排
    '问题目:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str医生 As String
    Dim lng项目id As Long, lng科室ID As Long, lng医生ID As Long
    Dim str号别 As String, strTemp As String, strTemp1 As String
    Dim i As Long, bytCheckType As Byte '0-检查计划是否合法;1-检查安排中正在执行项目是否合法.
    Dim strTittle As String
    
    On Error GoTo errHandle
    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    lng项目id = cboItem.ItemData(cboItem.ListIndex)
    lng医生ID = 0: str医生 = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    
    '检查计划中是否存在重复
    bytCheckType = 0
goReCheck:
    If bytCheckType <> 0 Then

        strSQL = "" & _
        "   Select Distinct A.号码, A.周日 D0, A.周一 D1, A.周二 D2, A.周三 D3, A.周四 D4, A.周五 D5, A.周六 D6, " & _
        "                 Nvl(To_Char(a.开始时间, 'YYYY-MM-DD HH24:MI:SS'), '1901-01-01') 生效时间, " & _
        "                 Nvl(To_Char(a.终止时间, 'YYYY-MM-DD HH24:MI:SS'), '3000-01-01 00:00:00') 失效时间 " & _
        "   From 挂号安排 A,挂号安排 B " & _
        "   Where A.科室id = b.科室id And A.医生姓名 = b.医生姓名 And Nvl(A.医生id, 0) = nvl(b.医生id,0) " & _
        "               And a.ID + 0 <> [1]   And B.ID = [1]  " & _
        "   Order By 号码"
            strTittle = "安排"
    Else
        strSQL = "" & _
            "   Select  distinct A.号码,A.周日 D0,A.周一 D1,A.周二 D2,A.周三 D3,A.周四 D4,A.周五 D5,A.周六 D6," & _
            "           To_Char(A.生效时间,'YYYY-MM-DD HH24:MI:SS') 生效时间,To_Char(A.失效时间,'YYYY-MM-DD HH24:MI:SS') 失效时间" & _
            "   From 挂号安排计划 A, 挂号安排 B,挂号安排 C " & _
            "   Where A.安排ID=B.ID and B.科室ID=C.科室ID and B.医生姓名=C.医生姓名 and nvl(B.医生ID,0)=nvl(C.医生ID,0) " & _
            "           And B.ID+0<>[1] and C.ID=[1]  " & _
            "   Order by 号码"
            strTittle = "计划安排"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    blnMulitNumPlan = Not rsTemp.EOF
    If blnMulitNumPlan = False And bytCheckType = 0 Then
        bytCheckType = bytCheckType + 1
        GoTo goReCheck:
    End If
    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
    str号别 = ""
    Do While Not rsTemp.EOF
        str号别 = str号别 & "," & Nvl(rsTemp!号码)
        If (Nvl(rsTemp!生效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!生效时间) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
           (Nvl(rsTemp!失效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!失效时间) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
           (Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)) Or _
           (Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)) Then
           '时间内不能交叉
            If opt天.Value Then
                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & "  周日:" & Nvl(rsTemp!D0)
                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & "  周一:" & Nvl(rsTemp!D1)
                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & "  周二:" & Nvl(rsTemp!D2)
                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & "  周三:" & Nvl(rsTemp!D3)
                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & "  周四:" & Nvl(rsTemp!D4)
                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & "  周五:" & Nvl(rsTemp!D5)
                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & "  周六:" & Nvl(rsTemp!D6)
                If strTemp <> "" Then
                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & vbCrLf & "  生效时间:" & IIf(Nvl(rsTemp!生效时间) = "1901-01-01", "无限", Nvl(rsTemp!生效时间) & "-" & Nvl(rsTemp!失效时间)) & vbCrLf
                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号计划安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此计划安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckRegistPlanIsValied = False: Exit Function
                End If
            Else
                With vsPlan
                    For i = 0 To 6
                        strTemp1 = "  周" & Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                            '存在,肯定重复了
                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                        End If
                    Next
                End With
                If strTemp <> "" Then
                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  生效时间:" & IIf(Nvl(rsTemp!生效时间) = "1901-01-01", "无限", Nvl(rsTemp!生效时间) & "-" & Nvl(rsTemp!失效时间)) & vbCrLf
                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此计划安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckRegistPlanIsValied = False: Exit Function
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    If bytCheckType = 0 Then
        bytCheckType = bytCheckType + 1
        GoTo goReCheck:
    End If
    zlCheckRegistPlanIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     SaveErrLog
End Function

Private Sub Init时间段()
  '--------------------------------
  '功能:获取上下班时间段
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("上午上下班时间", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_时间.dat_上午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_上午上班 = CDate("08:00:00")
    End If
    
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_时间.dat_上午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_上午下班 = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("下午上下班时间", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_时间.dat_下午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_下午上班 = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_时间.dat_下午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_下午下班 = CDate("1900-01-01 18:00:00")
    End If
    With t_时间
         If .dat_上午上班 > .dat_上午下班 Then
            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
         End If
         If .dat_上午上班 > .dat_上午下班 Then
            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
         End If
    End With
    strSQL = _
    "       Select 时间段, 上班, 下班 " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select 时间段,To_Date('1900-01-01 ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 开始时间," & vbNewLine & _
    "                               To_Date(Decode(Sign(开始时间 - 终止时间), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 终止时间," & _
    "                               Sign(开始时间 - 终止时间) As 隔天, " & vbNewLine & _
    "                                To_Date('" & Format(t_时间.dat_上午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午上班时间, " & vbNewLine & _
    "                                To_Date('" & Format(t_时间.dat_上午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午下班时间, " & vbNewLine & _
    "                                 To_Date('" & Format(t_时间.dat_下午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午上班时间," & vbNewLine & _
    "                                 To_Date('" & Format(t_时间.dat_下午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午下班时间,出诊预留时间 As 预留时间 "
    strSQL = strSQL & vbNewLine & _
    "                       From 时间段 )" & vbNewLine & _
    "           Select 时间段, '无' As 标签, 0 As 标志, 开始时间 As 上班, 终止时间 - Nvl(预留时间, 0) / 24 / 60 As 下班, 开始时间, 终止时间," & _
    "                  上午上班时间 As 上班时间, 上午下班时间 As 下班时间" & vbNewLine & _
    "            From Tb  Where (开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) And " & _
    "                      (开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select 时间段, '有-上午' As 标签, 1 As 标志, Decode(Sign(上午上班时间 - 开始时间), 1, 上午上班时间, 开始时间) As 上班, " & vbNewLine & _
    "                        Decode(Sign(终止时间 - 上午下班时间), 1, 上午下班时间, 终止时间) - Nvl(预留时间, 0) / 24 / 60 As 下班, 开始时间, 终止时间, " & _
    "                        上午上班时间 As 上班时间, 上午下班时间 As 下班时间 " & vbNewLine & _
    "           From Tb a Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select 时间段, '有-下午' As 标签, 1 As 标志, Decode(Sign(下午上班时间 - 开始时间), 1, 下午上班时间, 开始时间) As 上班, " & _
    "                   Decode(Sign(终止时间 - 下午下班时间), 1, 下午下班时间, 终止时间) - Nvl(预留时间, 0) / 24 / 60 As 下班, 开始时间, 终止时间, 下午上班时间 As 上班时间, 下午下班时间 As 下班时间 " & vbNewLine & _
    "         From Tb a   Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By 时间段,上班"
     Set mrs上班时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Function AssignReapportion(ByVal lng间隔时间 As Long, ByVal str限制项目 As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng限号 As Long
    Dim lng限约 As Long
    Dim dat开始时间 As Date
    Dim dat结束时间 As Date
    Dim lng序号 As Long
    Dim strTmp As String
    Dim str时段 As String
    Dim str限制时间 As String
    Dim lng默认间隔 As Long
    Dim lng分配个数 As Long
    Dim lng固定数量 As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim dat时点 As Date
    If mrs上班时间段 Is Nothing Then
        Call Init时间段
    End If

    If mrs上班时间段 Is Nothing Then Exit Function
    mTimeSet.rsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
    If mTimeSet.rsRegPlan.RecordCount = 0 Then mTimeSet.rsRegPlan.Filter = 0: Exit Function
    lng限号 = Nvl(mTimeSet.rsRegPlan!限号数, 0): lng限约 = Nvl(mTimeSet.rsRegPlan!限约数, 0)
    If lng限约 = 0 Then lng限约 = lng限号
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbInformation, Me.Caption
        Exit Function
    End If


    str时段 = mTimeSet.rsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbInformation, Me.Caption
        Exit Function

    End If

    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "' And 已使用=0"
    Do While Not mTimeSet.rsAssign.EOF
        mTimeSet.rsAssign.Delete adAffectCurrent
        mTimeSet.rsAssign.MoveNext
    Loop
    mTimeSet.rsAssign.Filter = "限制项目='" & str限制项目 & "'"
    If mTimeSet.rsAssign.RecordCount <> 0 Then
        lng固定数量 = mTimeSet.rsAssign.RecordCount
        lng默认间隔 = Val(Nvl(mTimeSet.rsAssign!时间间隔, lng间隔时间))
        Do While Not mTimeSet.rsAssign.EOF
            lng分配个数 = lng分配个数 + Val(Nvl(mTimeSet.rsAssign!限制数量))
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    mTimeSet.rsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        dat时点 = dat开始时间
        mrs上班时间段.MoveNext

        If mTimeSet.bln序号控制 Then
            For i = j To lng限号
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                If i > lng固定数量 Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        If Format(DateAdd("n", lng间隔时间, dat时点), "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                            !结束时间 = Format(dat结束时间, "hh:mm:ss")
                            !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                        Else
                            !结束时间 = Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm:00")
                            !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm")
                        End If
                        !时间间隔 = lng间隔时间
                        !限制数量 = IIf(lng分配个数 >= lng限号, 0, 1)
                        !是否预约 = 0
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限号, 0, 1)
                    End With
                Else
                    mTimeSet.rsAssign.Filter = "序号=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mTimeSet.rsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng间隔时间
                    End If
                End If
                dat时点 = DateAdd("n", IIf(i > lng固定数量, lng间隔时间, lng默认间隔), dat时点)
            Next

        Else    '非序号控制

            Do While Not Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss")

                ' If lngStart > lng限约 Then blnExit = True: Exit For
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then Exit Do

                If i > lng固定数量 Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        If Format(DateAdd("n", lng间隔时间, dat时点), "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                            !结束时间 = Format(dat结束时间, "hh:mm:ss")
                            !时间段 = Format(dat时点, "hh:mm") & "-" & Format(dat结束时间, "hh:mm")
                        Else
                            !结束时间 = Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm:00")
                            !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm")
                        End If
                        !时间间隔 = lng间隔时间
                        !限制数量 = IIf(lng分配个数 >= lng限约, 0, 1)
                        !是否预约 = 1
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限约, 0, 1)
                    End With
                Else
                    mTimeSet.rsAssign.Filter = "序号=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mTimeSet.rsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng间隔时间
                    End If
                End If
                dat时点 = DateAdd("n", IIf(i > lng固定数量, lng间隔时间, lng默认间隔), dat时点)
                i = i + 1
            Loop
        End If
        If i > lng限号 And mTimeSet.bln序号控制 Then
            blnExit = True
        End If
    Loop
    AssignReapportion = True
End Function

Private Sub txtTimeOut_Change()
    If Val(txtTimeOut.Text) > 1440 Then txtTimeOut.Text = 1440
End Sub

Private Sub cmd设置时段_Click()
    If AssignReapportion(Val(txtTimeOut.Text), tbSubPage.Item(mTimeSet.lngSelIndex).Caption) = False Then Exit Sub
    Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
End Sub

Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If mEditType <> ed_计划安排 And mEditType <> Ed_安排修改 Then Cancel = True: Exit Sub
        If Not opt周.Value = True Then Cancel = True: Exit Sub
        If Row = 3 And chkAppoint.Value = 0 Then Cancel = True
    End With
End Sub
Private Sub vsPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
      '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关的格式
    '编制:刘兴洪
    '日期:2011-11-11 11:33:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPlan
        If Row = 1 Then
              If Trim(.EditText) = "" Then
               .TextMatrix(2, Col) = ""
               .TextMatrix(3, Col) = ""
            End If
            Exit Sub
        End If
        If Val(.TextMatrix(Row, Col)) <> 0 Then
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###;;;")
        End If
    End With
    Exit Sub
End Sub
Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
   Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
    vsPlan.ColComboList(NewCol) = ""
    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
        vsPlan.TextMatrix(2, OldCol) = ""
        vsPlan.TextMatrix(3, OldCol) = ""
    End If
    If OldRow = 2 And Trim(vsPlan.TextMatrix(3, OldCol)) = "" And mbln自动默认限约数 Then
        vsPlan.TextMatrix(3, OldCol) = vsPlan.TextMatrix(2, OldCol)
    End If
    If NewRow <> 1 Then Exit Sub
    vsPlan.ColComboList(NewCol) = vsPlan.Tag
End Sub
Private Sub vsPlan_GotFocus()
    Call zl_VsGridGotFocus(vsPlan)
End Sub
Private Sub vsPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    With vsPlan
        If KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsPlan
        If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub

Private Sub vsPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPlan
            If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub
Private Sub vsPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub vsPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPlan
        If Row <= 1 Then Exit Sub
        VsFlxGridCheckKeyPress vsPlan, Row, Col, KeyAscii, m数字式
    End With
End Sub
Private Sub vsPlan_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsPlan)
End Sub

Private Sub vsPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String
    Dim str限制项目 As String
    Dim lng已约数  As Long
    '数据验证
    With vsPlan
        str限制项目 = Switch(Col = 1, "周日", Col = 2, "周一", Col = 3, "周二", Col = 4, "周三", Col = 5, "周四", Col = 6, "周五", True, "周六")
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If .Row <= 1 Then Exit Sub
        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        If Val(strKey) <> 0 Then
            strKey = Format(Abs(Val(strKey)), "####;;;")
        End If
        If Row = 2 Then
            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
                If MsgBox("限号数小于了限约数,是否清空限约数?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
                .TextMatrix(3, Col) = ""
            End If
        ElseIf Row = 3 Then
            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
                Call MsgBox("限号数小于了限约数,不能继续", vbInformation, gstrSysName)
                Cancel = True: Exit Sub
            End If
        End If
        .EditText = strKey
    End With
End Sub




Private Sub cboDoctor_Validate(Cancel As Boolean)
       
    '指定医生时不能指定多个科室
    If Trim(cboDoctor.Text) <> "" Then
        opt分诊(2).Enabled = False
        opt分诊(3).Enabled = False
        If opt分诊(2).Value Or opt分诊(3).Value Then opt分诊(0).Value = True
    Else
        opt分诊(2).Enabled = True
        opt分诊(3).Enabled = True
    End If
End Sub

Private Sub LoadDoctor()
    Set mrsDoctor = GetDoctor(Val(cbo科室.ItemData(cbo科室.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!姓名
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
        mrsDoctor.MoveNext
    Loop
End Sub

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    If KeyAscii <> 13 Then Exit Sub
    If cboDoctor.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsDoctor Is Nothing Then Exit Sub
    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text, True, "") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Function Check时段() As Boolean
    '新增加计划时 获取原有的安排是否具有时段
    '修改计划时 获取原计划是否具有时段
   Dim strSQL           As String
   Dim rsTmp            As ADODB.Recordset
   If mEditType <> Ed_安排修改 And mEditType <> ed_计划安排 Then Exit Function
    On Error GoTo Hd
    If mEditType = ed_计划安排 Then
        strSQL = " Select 1 As Hdata From 挂号安排时段 Where 安排id =[1] And Rownum=1"
    Else
        strSQL = "Select 1  as haveData From 挂号计划时段 Where 计划ID=[2] and Rownum=1"
    End If
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
     Check时段 = Not rsTmp.EOF
    Set rsTmp = Nothing
   
   Exit Function
Hd:
   If ErrCenter() = 1 Then
        Resume
   End If
   SaveErrLog
End Function


Private Function IsValidation() As Boolean
    '检查 生效时间是否合法
     If mbln限制修改 Then
      Select Case mEditType
        
                     ' dtpBegin.MinDate = mdtMinCustom
          Case Ed_安排审核
             If Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss") < Format(mdtMinCustom, "yyyy-mm-dd hh:mm:ss") Then
                If MsgBox("该计划在生效日期后已存在预约号,是否继续?", vbYesNo + vbDefaultButton1 + vbInformation, Me.Caption) = vbNo Then
                    Exit Function
                End If
            End If
          Case Ed_安排修改
                'dtpBegin.MinDate = mdtMinCustom
      End Select
    End If
    IsValidation = True
 End Function
 Private Function CheckExistsBooking(str号别 As String, Optional dtCustom As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定号别是否存在预约挂号单
    '入参:str号别-号别
    '返回:存在,返回true,否则返回False
    '编制:
    '日期:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select /*+ Rule*/ Max(发生时间) 时间" & vbNewLine & _
            "From 门诊费用记录" & vbNewLine & _
            "Where 记录性质 = 4 And 记录状态 In (0, 1) And 计算单位 = [1] And 发生时间 > 登记时间"
    If gint预约天数 = 0 Then
        strSQL = strSQL & " And 发生时间 > Sysdate"
    Else
        strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别)
    
    CheckExistsBooking = Not IsNull(rsTmp!时间)
    dtCustom = IIf(CheckExistsBooking, rsTmp!时间, zlDatabase.Currentdate)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadRegHistory() As Boolean
    Dim strSQL As String
    strSQL = "Select 限制项目, Max(最大序号) As 最大序号, Max(统计) As 统计, Max(发生时间) As 发生时间" & vbNewLine & _
            " From (Select Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六') As 限制项目," & vbNewLine & _
            "              Max(Nvl(a.号序, 0)) As 最大序号, Count(1) As 统计, To_Char(Max(发生时间), 'hh24:mi:ss') As 发生时间," & vbNewLine & _
            "              To_Char(发生时间, 'YYYY-MM-DD') As 发生日期" & vbNewLine & _
            "       From 病人挂号记录 A, 挂号安排 B" & vbNewLine & _
            "       Where a.记录状态 = 1 And a.发生时间 Between [2] And [3] And a.号别 = b.号码 And b.Id = [1] " & vbNewLine & _
            "       Group By Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六')," & vbNewLine & _
            "                To_Char(发生时间, 'YYYY-MM-DD'))" & vbNewLine & _
            " Group By 限制项目"

    On Error GoTo Hd:
    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, dtpBegin, dtpEndDate)
    LoadRegHistory = True
    
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
 
Public Property Let 自动默认限约数(ByVal vNewValue As Boolean)
    mbln自动默认限约数 = vNewValue
End Property

Private Sub LoadvsDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:vsDept属性设置
    '编制:李南春
    '日期:2014-04-14 16:05:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsDept
        .Cols = 1
        .Rows = 11
        .FixedCols = 0
        .FixedRows = 0
        .RowHeight(-1) = 300
        .AllowSelection = False
        .BackColorSel = &HE0E0E0
        .BackColorBkg = &H80000005
        .SheetBorder = &H80000005 '边线颜色
        .GridColor = &H80000005 '网格线颜色
        .ColWidthMin = 1200
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearVsGridCheckValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除网格控件的复选框值
    '编制:李南春
    '日期:2014-04-14 18:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer
    Dim i  As Integer
    Dim intRow As Integer
    On Error GoTo errHandle

    With vsDept
        .Redraw = flexRDNone
        intRow = -1
        For i = 0 To .Rows - 1
            If .Cell(flexcpChecked, i, .Cols - 1) = 0 Then intRow = i: Exit For
        Next
        .Cell(flexcpChecked, 0, 0, .Rows - 1, .Cols - 1) = 2
        If intRow <> -1 Then .Cell(flexcpChecked, intRow, .Cols - 1, .Rows - 1, .Cols - 1) = 0
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsDept.Cell(flexcpChecked, Row, Col) = 0 Then Cancel = True
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim intType As Integer
    If opt分诊(1).Value Then
        intType = vsDept.Cell(flexcpChecked, Row, Col)
        Call ClearVsGridCheckValue
        vsDept.Cell(flexcpChecked, Row, Col) = intType
    End If
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsDept_GotFocus()
    Dim intRow As Integer
    Dim intCol As Integer
    On Error GoTo errHandle
    
    With vsDept
        If .Row >= 0 And .Col >= 0 Then Exit Sub
        For intCol = 0 To .Cols - 1
            For intRow = 0 To .Rows - 1
                If .Cell(flexcpChecked, intRow, intCol) = 1 Then
                    .Row = intRow: .Col = intCol
                    Exit Sub
                End If
            Next
        Next
        If .Rows >= 0 And .Cols >= 0 Then .Row = 0: .Col = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
