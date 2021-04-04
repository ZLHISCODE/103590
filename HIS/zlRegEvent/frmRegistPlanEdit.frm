VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmRegistPlanEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "挂号安排编辑"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   Icon            =   "frmRegistPlanEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9840
      TabIndex        =   32
      Top             =   1065
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9840
      TabIndex        =   31
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   9840
      TabIndex        =   37
      Top             =   1590
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   780
      Left            =   9240
      TabIndex        =   38
      Top             =   3720
      Width           =   1575
      _Version        =   589884
      _ExtentX        =   2778
      _ExtentY        =   1376
      _StockProps     =   64
   End
   Begin VB.PictureBox picTimeSet 
      BorderStyle     =   0  'None
      Height          =   6900
      Left            =   1380
      ScaleHeight     =   6900
      ScaleWidth      =   9525
      TabIndex        =   39
      Top             =   2235
      Width           =   9525
      Begin VB.CommandButton cmdAuto 
         Caption         =   "自动计算(&A)"
         Height          =   350
         Left            =   5505
         TabIndex        =   57
         ToolTipText     =   "通过输入的限号数,自动分配时间间隔进行计算"
         Top             =   45
         Width           =   1150
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "全清(&D)"
         Height          =   350
         Left            =   8370
         TabIndex        =   56
         ToolTipText     =   "点击重新计算时段"
         Top             =   45
         Width           =   1150
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   6930
         TabIndex        =   55
         ToolTipText     =   "点击重新计算时段"
         Top             =   45
         Width           =   1150
      End
      Begin VB.Frame fra应用于 
         Caption         =   "应用于…"
         Height          =   615
         Left            =   675
         TabIndex        =   50
         Top             =   6825
         Width           =   7755
         Begin VB.OptionButton opt应用于 
            Caption         =   "本医生(张三)"
            Height          =   255
            Index           =   1
            Left            =   2100
            TabIndex        =   54
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt应用于 
            Caption         =   "本号码"
            Height          =   255
            Index           =   0
            Left            =   795
            TabIndex        =   53
            Top             =   255
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "本科室(内科)"
            Height          =   255
            Left            =   3870
            TabIndex        =   52
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt所有 
            Caption         =   "所有号别"
            Height          =   255
            Left            =   5685
            TabIndex        =   51
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "其他辅助计算(&T)"
         Height          =   350
         Left            =   3690
         TabIndex        =   43
         ToolTipText     =   "点击重新计算时段"
         Top             =   45
         Width           =   1515
      End
      Begin VB.CommandButton cmd设置时段 
         Caption         =   "辅助计算(&F)"
         Height          =   350
         Left            =   2235
         TabIndex        =   42
         ToolTipText     =   "点击重新计算时段"
         Top             =   45
         Width           =   1150
      End
      Begin VB.TextBox txtTimeOut 
         Height          =   300
         Left            =   1185
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "10"
         Top             =   75
         Width           =   465
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   3540
         Index           =   0
         Left            =   690
         ScaleHeight     =   3540
         ScaleWidth      =   2535
         TabIndex        =   40
         Top             =   990
         Width           =   2535
      End
      Begin MSComCtl2.UpDown udTime 
         Height          =   300
         Left            =   1650
         TabIndex        =   44
         Top             =   75
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "Frame3"
         BuddyDispid     =   196630
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
         Left            =   225
         TabIndex        =   45
         Top             =   1380
         Width           =   2535
         _Version        =   589884
         _ExtentX        =   4471
         _ExtentY        =   8599
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTime 
         Height          =   5475
         Index           =   0
         Left            =   1005
         TabIndex        =   46
         Top             =   1245
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
         FormatString    =   $"frmRegistPlanEdit.frx":000C
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
         Begin VB.CommandButton cmd删除 
            Caption         =   "删"
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   48
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmd预约 
            Caption         =   "预"
            Height          =   255
            Index           =   0
            Left            =   2685
            TabIndex        =   47
            Top             =   2535
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "时间间隔(分)"
         Height          =   180
         Left            =   75
         TabIndex        =   49
         Top             =   135
         Width           =   1080
      End
   End
   Begin VB.PictureBox picBaseBack 
      BorderStyle     =   0  'None
      Height          =   8865
      Left            =   120
      ScaleHeight     =   8865
      ScaleWidth      =   10125
      TabIndex        =   33
      Top             =   120
      Width           =   10125
      Begin VB.Frame Frame4 
         Caption         =   "应诊诊室:"
         Height          =   3980
         Left            =   240
         TabIndex        =   36
         Top             =   4560
         Width           =   8895
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   3480
            Left            =   150
            TabIndex        =   30
            Top             =   300
            Width           =   8595
            _cx             =   15161
            _cy             =   6138
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
            GridColor       =   -2147483628
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483634
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
            Cols            =   2
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
            TabIndex        =   29
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "动态分诊"
            Height          =   180
            Index           =   2
            Left            =   3180
            TabIndex        =   28
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "指定诊室"
            Height          =   180
            Index           =   1
            Left            =   2010
            TabIndex        =   27
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "不分诊"
            Height          =   180
            Index           =   0
            Left            =   1020
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "应诊时间"
         Height          =   2550
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   8925
         Begin VB.TextBox txt限号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   17
            Top             =   292
            Width           =   1215
         End
         Begin VB.TextBox txt限约 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4980
            MaxLength       =   5
            TabIndex        =   19
            Top             =   292
            Width           =   1215
         End
         Begin VB.CheckBox chk有效期 
            Caption         =   "有效期"
            Height          =   195
            Left            =   255
            TabIndex        =   22
            Top             =   2115
            Width           =   855
         End
         Begin VB.ComboBox cbo天 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   292
            Width           =   1110
         End
         Begin VB.OptionButton opt周 
            Caption         =   "每周(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   20
            Top             =   630
            Width           =   930
         End
         Begin VB.OptionButton opt天 
            Caption         =   "每天(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   14
            Top             =   285
            Width           =   960
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   1170
            TabIndex        =   23
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   192544771
            CurrentDate     =   38091
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   3555
            TabIndex        =   25
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   192544771
            CurrentDate     =   38091
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   1275
            Left            =   1140
            TabIndex        =   21
            Top             =   675
            Width           =   7650
            _cx             =   13494
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
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmRegistPlanEdit.frx":0081
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "限号"
            Height          =   180
            Left            =   2610
            TabIndex        =   16
            Top             =   352
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "限约"
            Height          =   180
            Left            =   4545
            TabIndex        =   18
            Top             =   345
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            Height          =   180
            Left            =   3315
            TabIndex        =   24
            Top             =   2115
            Width           =   180
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "基本信息"
         Height          =   1500
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   8970
         Begin VB.TextBox txtAppLimit 
            Height          =   315
            Left            =   7125
            TabIndex        =   13
            Top             =   1058
            Width           =   765
         End
         Begin VB.CheckBox chkAppoint 
            Caption         =   "可预约          天"
            Height          =   300
            Left            =   6240
            TabIndex        =   12
            Top             =   1065
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chk序号控制 
            Caption         =   "序号控制"
            Height          =   255
            Left            =   2130
            TabIndex        =   2
            Top             =   285
            Width           =   1095
         End
         Begin VB.ComboBox cbo号类 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   2115
         End
         Begin VB.CheckBox chk病案 
            Caption         =   "挂号时必须建病案"
            Height          =   195
            Left            =   3615
            TabIndex        =   11
            Top             =   1118
            Width           =   1845
         End
         Begin VB.ComboBox cbo科室 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            TabIndex        =   6
            Text            =   "cbo科室"
            Top             =   660
            Width           =   2115
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   1050
            TabIndex        =   10
            Top             =   1065
            Width           =   2115
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   660
            Width           =   2115
         End
         Begin VB.TextBox txt号别 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            MaxLength       =   5
            TabIndex        =   1
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "号类"
            Height          =   180
            Left            =   3600
            TabIndex        =   3
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lbl医生 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "院内医生↓"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "项目"
            Height          =   180
            Left            =   3615
            TabIndex        =   7
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "科室"
            Height          =   180
            Left            =   645
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
            Left            =   615
            TabIndex        =   0
            Top             =   330
            Width           =   390
         End
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "院内医生"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "含外援医生"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlngModule As Long, mstrPrivs As String, mlngID As Long, mfrmMain As Form, mblnChange As Boolean
Private mrs科室 As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnSucces As Boolean
Private mlng缺省挂号科室ID  As Long '在挂号安排时，根据主界面中选择的科室进行缺省
Private mrs时间段 As ADODB.Recordset
Private mstr限制修改 As String '在某一天或者多天的安排限制更改
Private mbln自动默认限约数 As Boolean '45519 自动默认限约数
Private mblnMinorChange As Boolean
Public Enum RegistEditType
    edt_新增 = 0
    edt_修改 = 1
    edt_查阅 = 2
End Enum
Private mEditType As RegistEditType
'对外上班时间
Private Type t_上班时间
  dat_上午上班 As Date
  dat_上午下班 As Date
  dat_下午上班 As Date
  dat_下午下班 As Date
End Type
Private t_时间 As t_上班时间
Private mrs上班时间段 As ADODB.Recordset
Private mrs限号          As ADODB.Recordset

Private mPlanEditType As gPlanEditType

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
    strKey As String
    str限制修改 As String
End Type

Private mTimeSet As TimeSet
Private mintSysAppLimit As Integer
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther
Attribute mfrmOtherCalc.VB_VarHelpID = -1

Private mstr科室ID As String
Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
Private mblnOnly院内医生 As Boolean '仅只能输院内医生
Private Type PlanInfo               '安排改变需要对比的信息
    str应诊时段      As String
    str排班         As String       '排班信息
    str限号         As String       '限号信息
    bln序号         As Boolean      '是否序号控制
    bln时间段       As Boolean      '是否设置了时间段
End Type

Private mPlanInfo     As PlanInfo '原始的安排信息  主要用于安排修改时 相应信息的比较

Private Enum mPageIndex
    EM_安排 = 0
    EM_时段 = 1
End Enum

Private Enum mPgIndex
    Pg_计划安排 = 1
    Pg_计划时段 = 2
End Enum

Private mblnChangeByCode As Boolean '是否是代码控制改变了tabelpage的显示页
Private mrsRegOldData As ADODB.Recordset '本地数据集保存,原始挂号安排
Private mrsRegNewData As ADODB.Recordset '本地数据集保存 重新设置后的安排
Private mrsRegHistory As ADODB.Recordset '历次挂号的数据集
Private mcll预约信息  As Collection '保存已经预约出去的预约信息 K星期_数量 /K星期_日期
Private mblnChangeDist As Boolean


Private Function LoadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '返回:加载成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 12:14:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL          As String
    Dim rsTemp          As New ADODB.Recordset
    Dim i               As Long
    Dim j               As Long
    Dim strTemp         As String
    Dim rs限号          As ADODB.Recordset
    Dim bln每周         As Boolean
    Dim bln限号         As Boolean
    Dim str限号         As String
    Dim bln限约         As Boolean
    Dim str限约         As String
    Dim blnExitFor      As Boolean
    Dim rsTmp           As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:

    
    If mEditType = edt_新增 Then
        txt号别.Text = GetNext号别
        txt限号.Text = ""
        txt限约.Text = ""
        chk病案.Value = 0

        If cbo科室.ListIndex >= 0 Then
            If mlng缺省挂号科室ID <> cbo科室.ItemData(cbo科室.ListIndex) Then
                cbo科室.ListIndex = -1
                cboItem.ListIndex = -1
                cboDoctor.Text = ""
            End If
        Else
            cbo科室.ListIndex = -1
            cboItem.ListIndex = -1
            cboDoctor.Text = ""
        End If
        dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = CDate("3000-01-01")

        opt天.Value = True
        cbo天.Enabled = True
        cbo天.ListIndex = cbo.FindIndex(cbo天, "全日", True)
        If cbo天.ListIndex = -1 Then cbo天.ListIndex = 0
        opt周.Value = False
        vsPlan.Enabled = False
        LoadCard = True
        opt分诊(0).Value = True
        mTimeSet.bln序号控制 = False
        
        '清空门诊诊室的选项
        For i = 0 To vsDept.Cols - 1
            For j = 0 To vsDept.Rows - 1
                If vsDept.Cell(flexcpChecked, j, i) <> 0 Then vsDept.Cell(flexcpChecked, j, i) = 2
            Next
        Next
        Exit Function
    End If
    '修改或查看
    strSQL = " " & _
    "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
    "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,A.默认时段间隔, " & _
    "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室,A.预约天数 " & _
    "   From 挂号安排 A,收费项目目录 B,部门表 D " & _
    "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
    "         And A.Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)

    If rsTemp.EOF Then
        ShowMsgbox "未找到指定的号别,请检查!"
        Exit Function
    End If
    strSQL = "Select 限制项目,限号数,限约数 From  挂号安排限制 where 安排ID=[1]       "
    Set rs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    
    chkAppoint.Value = 0
    Do While Not rs限号.EOF
        If IsNull(rs限号!限约数) Then
            chkAppoint.Value = 1
            Exit Do
        Else
            If Val(Nvl(rs限号!限约数)) <> 0 Then
                chkAppoint.Value = 1
                Exit Do
            End If
        End If
        rs限号.MoveNext
    Loop
    If rs限号.RecordCount <> 0 Then rs限号.MoveFirst
    
    cbo号类.ListIndex = cbo.FindIndex(cbo号类, Nvl(rsTemp!号类), True)
    txt号别.Text = Nvl(rsTemp!号码)

    cbo科室.ListIndex = cbo.FindIndex(cbo科室, Nvl(rsTemp!科室), True)
    cboItem.ListIndex = cbo.FindIndex(cboItem, Nvl(rsTemp!项目), True)

    cboDoctor.ListIndex = cbo.FindIndex(cboDoctor, Nvl(rsTemp!医生姓名), True)
    If cboDoctor.ListIndex = -1 Then cboDoctor.Text = Nvl(rsTemp!医生姓名)


    chk病案.Value = IIf(Val(Nvl(rsTemp!病案必须)) = 1, 1, 0)

    chk序号控制.Value = IIf(Val(Nvl(rsTemp!序号控制)) = 1, 1, 0):     chk序号控制.Tag = chk序号控制.Value
    mTimeSet.bln序号控制 = Val(rsTemp!序号控制) = 1
    
    '获取修改前的安排是否序号控制
    mPlanInfo.bln序号 = IIf(Val(Nvl(rsTemp!序号控制)) = 1, True, False)
    '有效时间范围
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
    If Not IsNull(rsTemp!开始时间) Then
        chk有效期.Value = 1
        dtpBegin.Value = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM:SS"))
        If Not IsNull(rsTemp!终止时间) Then
            dtpEnd.Value = CDate(Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS"))
        End If
    End If

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
                !id = mlngID
                !限制项目 = Nvl(rs限号!限制项目)
                !限号数 = Val(Nvl(rs限号!限号数))
                !限约数 = Val(Nvl(rs限号!限约数))
                !序号控制 = Val(Nvl(rsTemp!序号控制))
                .Update
            End With
            rs限号.MoveNext
        Loop
    End With

    Call LoadRegHistory

    '---------------------------------------------------
    '判断 每日安排 限号数 限约数 等是否一致
    '---------------------------------------------------
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

   If bln每周 Or mrsRegHistory.RecordCount > 0 Then
        '每周
        opt周.Value = True
        With vsPlan
            For i = 1 To .Cols - 1
                strTemp = Switch(i - 1 = 0, "日", i - 1 = 1, "一", i - 1 = 2, "二", i - 1 = 3, "三", i - 1 = 4, "四", i - 1 = 5, "五", True, "六")
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("周" & strTemp))
                rs限号.Filter = "限制项目='" & "周" & strTemp & "'"
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
                If InStr(mstr限制修改, ";周" & strTemp & ";") > 0 Then
                    .Cell(flexcpForeColor, 2, i, 3, i) = vbBlue
                End If
            Next
        End With
        opt天.Value = False: cbo天.Enabled = False: txt限号.Enabled = False: txt限约.Enabled = False
        vsPlan.Enabled = True: chk序号控制.Enabled = mstr限制修改 = ""
    Else
        '每天
        opt天.Value = True:  cbo天.ListIndex = cbo.FindIndex(cbo天, Nvl(rsTemp!周日), True)
        If cbo天.ListIndex = -1 Then cbo天.ListIndex = 0:
        opt周.Value = False: vsPlan.Enabled = False
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
    '获取修改前的 时间段和 限号数
    '用于在保存时 对比限号限约以及时间段是否发生了变化
    '如果发生了变化则需要提示  操作员重新设置时段信息
    '------------------------------
   mPlanInfo.str排班 = ""
   mPlanInfo.str限号 = ""
   mPlanInfo.str应诊时段 = ""
    If bln每周 Or mrsRegHistory.RecordCount > 0 Then
         For i = 1 To vsPlan.Cols - 1
            mPlanInfo.str排班 = mPlanInfo.str排班 & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"
            mPlanInfo.str应诊时段 = mPlanInfo.str应诊时段 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六") & "-" & Trim(vsPlan.TextMatrix(1, i))
                mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                     mPlanInfo.str限号 = mPlanInfo.str限号 & ",0,0"
                Else
                     mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Trim(vsPlan.TextMatrix(3, i))
                End If
        Next
    Else
         For i = 1 To 7
             mPlanInfo.str排班 = mPlanInfo.str排班 & "'" & Trim(cbo天.Text) & "',"
             mPlanInfo.str应诊时段 = mPlanInfo.str应诊时段 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六") & "-" & Trim(cbo天.Text)
             mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
             mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(txt限号.Text) & "," & txt限约.Text
        Next
    End If
    If mPlanInfo.str限号 <> "" Then mPlanInfo.str限号 = Mid(mPlanInfo.str限号, 2)
    If mPlanInfo.str应诊时段 <> "" Then mPlanInfo.str应诊时段 = Mid(mPlanInfo.str应诊时段, 2)
    '-------------------------------

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

    strSQL = "Select 号表ID,门诊诊室　From 挂号安排诊室 Where 号表ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    '71253 李南春 2014-04-15 11:30:10 将listView 替换为vsflexGrid
    
    With vsDept
        blnExitFor = False
        Do While Not rsTmp.EOF
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If Nvl(rsTmp!门诊诊室) = .TextMatrix(j, i) Then
                        .Cell(flexcpChecked, j, i) = 1
                        blnExitFor = True
                        Exit For
                    End If
                Next
                If blnExitFor Then blnExitFor = False: Exit For
            Next
            rsTmp.MoveNext
        Loop
    End With
    rsTmp.Close
    
    If mstr限制修改 <> "" Then opt天.Enabled = False
    '如果是修改时 获取原来的安排是否已经安排了时段
    If mEditType = edt_修改 Then mPlanInfo.bln时间段 = Check时段
    If mrsRegHistory.RecordCount > 0 Then opt天.Enabled = False
    If chkAppoint.Value = 1 Then
        txtAppLimit.Enabled = True
        txtAppLimit.Text = Nvl(rsTemp!预约天数, mintSysAppLimit)
    Else
        txtAppLimit.Enabled = False
        txtAppLimit.Text = Nvl(rsTemp!预约天数, mintSysAppLimit)
    End If
    LoadCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub chkAppoint_Click()
    Dim i As Integer
    If chkAppoint.Value = 0 Then
        If opt天.Value = True Then
            txt限约.Enabled = False
            txt限约.BackColor = &H8000000F
        End If
        txt限约.Text = ""
        txtAppLimit.Enabled = False
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(3, i) = ""
        Next i
    Else
        If opt天.Value = True Then
            txt限约.Enabled = True
            txt限约.BackColor = vbWhite
        End If
        txtAppLimit.Enabled = True
        If Val(txt限约.Text) = 0 Then txt限约.Text = ""
        For i = 1 To vsPlan.Cols - 1
            If Val(vsPlan.TextMatrix(3, i)) = 0 Then vsPlan.TextMatrix(3, i) = ""
        Next i
    End If
End Sub

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
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Function
    End If

    str时段 = mTimeSet.rsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbOKOnly, Me.Caption
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

Public Function ShowEdit(ByVal frmMain As Form, ByVal EditType As RegistEditType, _
    ByVal lngModule As Long, ByVal strPrivs As String, Optional lngID As Long = 0, _
    Optional lng缺省科室ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗体
    '     EditType-编辑类型
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 10:25:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs: mlngID = lngID: mlng缺省挂号科室ID = lng缺省科室ID
    mEditType = EditType: mblnSucces = False
    mblnChange = False
    If EditType = edt_修改 Then
        mstr限制修改 = zl_Get预约信息(lngID)
    End If
    Me.Show 1, frmMain
    ShowEdit = mblnSucces
    
End Function

Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        cboDoctor.ListIndex = GetCboIndex(cboDoctor, cboDoctor)
'    End If
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
        If mblnOnly院内医生 = False Then
                zlCommFun.PressKey vbKeyTab
        End If
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
      If mblnOnly院内医生 Then
           If cboDoctor.ListIndex < 0 Then cboDoctor.Text = ""
      End If
      
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

Private Sub cbo科室_Click()
    mblnCboClick = True
    If cbo科室.ListIndex = -1 Then Exit Sub
    Call LoadDoctor
End Sub

Private Sub LoadDoctor()
    Set mrsDoctor = GetDoctor(Val(cbo科室.ItemData(cbo科室.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!姓名
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!id
        mrsDoctor.MoveNext
    Loop
End Sub

Private Sub cbo科室_GotFocus()
    zlControl.TxtSelAll cbo科室
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo科室.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If cbo科室.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        mblnCboClick = True
        If Select科室(Me, mlngModule, mrs科室, cbo科室, cbo科室.Text) = True Then
            mblnCboClick = False
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If cbo科室.Enabled Then cbo科室.SetFocus
        mblnCboClick = False
        zlControl.TxtSelAll cbo科室
    Else
       ' Call zlControl.CboSetIndex(cbo科室.hWnd, zlControl.CboMatchIndex(cbo科室.hWnd, KeyAscii))
    End If
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
 '如果在cbo的keypress事件中用了弹出列表的的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
    If Not mblnCboClick Then cbo科室_Click
    mblnCboClick = False
End Sub

Private Sub chk有效期_Click()
    dtpBegin.Enabled = chk有效期.Value = 1
    dtpEnd.Enabled = chk有效期.Value = 1
    
    If Visible And dtpBegin.Enabled Then
        dtpBegin.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Function GetDoctorPlan(lng医生ID As Long, str医生姓名 As String) As ADODB.Recordset
'功能:返回指定医生ID或姓名的已有号别的时间信息
'   用于检查新增或修改的号别是否与现有的号别在时间上重复
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 号码,周日 D0,周一 D1,周二 D2,周三 D3,周四 D4,周五 D5,周六 D6," & _
            " To_Char(开始时间,'YYYY-MM-DD HH24:MI:SS') 开始时间,To_Char(终止时间,'YYYY-MM-DD HH24:MI:SS') 终止时间" & _
            " From 挂号安排 Where (终止时间 is null or 终止时间>sysdate) And " & IIf(lng医生ID <> 0, " 医生ID=[1]", " 医生姓名=[1]") & _
            IIf(mEditType = edt_新增, "", " And ID<>[2]")
    Set GetDoctorPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(lng医生ID <> 0, lng医生ID, str医生姓名), mlngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckExistsBooking() As Boolean
'功能:检查当前时间段之外是否存在预约挂号单
    Dim rsTemp As ADODB.Recordset, rsBooking As ADODB.Recordset, strSQL As String
    Dim i As Long, str时间段 As String
        
    On Error GoTo errH
    If opt天.Value Then
        str时间段 = _
               "Select 1 From 时间段 b Where b.时间段 = [2] And (" & _
               " ('3000-01-10 '||To_Char(a.发生时间,'HH24:MI:SS')" & _
               " Between" & _
               " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-09 '||To_Char(b.开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS'))" & _
               " And" & _
               " '3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))" & _
               " Or" & _
               " ('3000-01-10 '||To_Char(a.发生时间,'HH24:MI:SS')" & _
               " Between" & _
               " '3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS')" & _
               " And" & _
               " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-11 '||To_Char(b.终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))))"
        
        strSQL = "Select  /*+ Rule*/ Min(发生时间) 时间" & vbNewLine & _
            "From 门诊费用记录 a" & vbNewLine & _
            "Where 记录性质 = 4 And 记录状态 In (0, 1) And 计算单位 = [1] And 发生时间 > 登记时间"
        If gint预约天数 = 0 Then
            strSQL = strSQL & " And 发生时间 > Sysdate"
        Else
            strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
        End If
        strSQL = strSQL & " And Not Exists (" & str时间段 & ")"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt号别.Text, Trim(cbo天.Text))
        CheckExistsBooking = Not IsNull(rsTemp!时间)
    Else
        strSQL = "Select /*+ Rule*/ 发生时间,To_Char(发生时间,'D') 星期 From 门诊费用记录 a Where 记录性质 = 4 and 记录状态 In(0,1) And 计算单位 = [1] And 发生时间 > 登记时间"
        If gint预约天数 = 0 Then
            strSQL = strSQL & " And 发生时间 > Sysdate"
        Else
            strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
        End If
        
        Set rsBooking = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt号别.Text)
        For i = 1 To rsBooking.RecordCount
            str时间段 = Trim(vsPlan.TextMatrix(1, rsBooking!星期 - 1))
            If str时间段 = "" Then
                CheckExistsBooking = True
            Else
               strSQL = _
                    "Select Count(*) cnt From 时间段 b Where b.时间段 = [2] And (" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-09 '||To_Char(b.开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS'))" & _
                    " And" & _
                    " '3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))" & _
                    " Or" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " '3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS')" & _
                    " And" & _
                    " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-11 '||To_Char(b.终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))))"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(rsBooking!发生时间), str时间段)
                CheckExistsBooking = rsTemp!cnt = 0
            End If
            
            If CheckExistsBooking Then Exit Function
            rsBooking.MoveNext
        Next
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveMinorChange()
    Dim strSQL As String, intCount As Integer
    Dim str诊室 As String, lng预约天数 As Long
    Dim i As Long
    Dim j As Long
    Dim rsTemp As ADODB.Recordset
    Dim intSync As Integer
    If CheckRegistDays(Trim(txt号别.Text)) = False Then Exit Sub
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
    
    If chkAppoint.Value = 1 Then
        lng预约天数 = IIf(txtAppLimit.Text = "", gint预约天数, txtAppLimit.Text)
    Else
        lng预约天数 = 0
    End If
    
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str诊室 = str诊室 & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str诊室 = Mid(str诊室, 2)
    
    strSQL = "Select 1 From 挂号安排计划 Where 安排ID=[1] And 失效时间 > Sysdate"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If Not rsTemp.EOF And mblnChangeDist Then
        If MsgBox("修改的安排存在正生效和未生效的计划,是否同步更改计划的诊室设置?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            intSync = 1
        Else
            intSync = 0
        End If
    End If
    
    strSQL = "Zl_挂号安排_Modify("
    strSQL = strSQL & mlngID & ",'"
    strSQL = strSQL & str诊室 & "',"
    strSQL = strSQL & lng预约天数 & ","
    strSQL = strSQL & intCount & ","
    strSQL = strSQL & intSync & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Unload Me
End Sub


Private Function CheckRegistDays(ByVal str号别 As String) As Boolean
'功能:检查预约天数
    Dim lng检查天数 As Long
    On Error GoTo errH
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If chkAppoint.Value = 1 Then
        lng检查天数 = Val(Nvl(txtAppLimit.Text, gint预约天数))
    Else
        lng检查天数 = 0
    End If
    strSQL = "Select Max(发生时间) As 时间 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 = 1 And 发生时间 > Sysdate + [1] And 号别 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng检查天数, str号别)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!时间) <> "" Then
            MsgBox "在" & Format(rsTmp!时间, "YYYY-MM-DD") & "存在超过当前预约天数的预约记录,不能继续,请将可预约天数调大!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If chkAppoint.Value = 1 And txtAppLimit.Text <> "" And Val(txtAppLimit.Text) <= 0 Then
        MsgBox "允许预约时,预约天数不能小于等于0!", vbInformation, gstrSysName
        If txtAppLimit.Visible And txtAppLimit.Enabled Then txtAppLimit.SetFocus
        Exit Function
    End If
    
    CheckRegistDays = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim i As Integer, intCount As Integer, j As Integer
    Dim str时间段 As String, str诊室 As String, str限号 As String
    Dim lngNextID As Long, lng医生ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim cllPro As Collection, lng预约天数 As Long
    Dim str号别 As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '是否多次安排
    Dim blnChange       As Boolean '是否改变了 时间安排
    Dim strMsg          As String
    
    If mblnMinorChange Then Call SaveMinorChange: Exit Sub
    If mEditType = edt_查阅 Then Unload Me: Exit Sub
    If Me.tbPage.Item(mPageIndex.EM_安排).Selected = False Then
        mblnChangeByCode = True
        tbPage.Item(mPageIndex.EM_安排).Selected = True
        mblnChangeByCode = False
    End If
    If CheckRegistDays(Trim(txt号别.Text)) = False Then Exit Sub
    If mblnOnly院内医生 Then
        If cboDoctor.ListIndex < 0 And cboDoctor.Text <> "" Then
                MsgBox "你选择的医生不存在,请重新输入医生!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                If cboDoctor.Enabled Then cboDoctor.SetFocus
                Exit Sub
        End If
    End If
    '完整性检查
    If Trim(txt号别) = "" Then
        MsgBox "号别不能为空！", vbInformation, gstrSysName
        txt号别.SetFocus: Exit Sub
    End If
    If cbo科室.ListIndex = -1 Then
        MsgBox "未设置号别所对应的科室！", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Sub
    End If
    If cboItem.ListIndex = -1 Then
        MsgBox "未设置号别所对应的挂号项目！", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    If dtpBegin.Enabled And dtpEnd.Enabled Then
        If dtpBegin.Value >= dtpEnd.Value Then
            MsgBox "开始时间应该小于结束时间。", vbInformation, gstrSysName
            dtpBegin.SetFocus: Exit Sub
        End If
    End If

    If opt天.Value Then
        If cbo天.ListIndex = -1 Then
            MsgBox "该号别每天的应诊时间未设置！", vbInformation, gstrSysName
            cbo天.SetFocus: Exit Sub
        End If
        If chk序号控制.Value = 1 Then
            If Val(txt限号.Text) = 0 And Val(txt限约.Text) = 0 Then
                MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
                txt限号.SetFocus: Exit Sub
            End If
        End If
        '限号限约规则
        If Trim(txt限号.Text) <> "" Then
            If Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                txt限约.SetFocus: Exit Sub
            End If
        ElseIf Trim(txt限约.Text) <> "" Then
            MsgBox "限约必须限号！", vbInformation, gstrSysName
            txt限号.SetFocus: Exit Sub
        End If
    Else
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
                    If chk序号控制.Value = 1 Then
                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                              MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
                              .Row = 2: .Col = i
                              .SetFocus: Exit Sub
                          End If
                      End If
                        '限号限约规则
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Sub
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            MsgBox "限约必须限号！", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Sub
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "该号别每周的应诊时间未设置！", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Sub
            End If
        End With
    End If
    
    If CheckRegistDays(Trim(txt号别.Text)) = False Then Exit Sub
    
    '诊室判断
    If opt分诊(1).Value Or opt分诊(2).Value Or opt分诊(3).Value Then
        '71253 李南春 2014-04-15 11:30:10 将listView 替换为vsflexGrid
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

    '项目价格判断
    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
        MsgBox "项目""" & cboItem.Text & """未设置有效价格,请先到收费项目管理中设置！", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    '取医生ID
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    If lng医生ID = 0 And cboDoctor.Text <> "" Then
        strSQL = "Select 1 From 人员表 Where 姓名 = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDoctor.Text)
        If Not rsTemp.EOF Then
            MsgBox "医生""" & cboDoctor.Text & """不属于科室""" & cbo科室.Text & """,请重新设置该号别的科室与医生信息！", vbInformation, gstrSysName
            cboDoctor.SetFocus: Exit Sub
        End If
    End If
    
'    '问题:现在一个医生可以加入重复号了
'    If zlCheckPlanArrageIsValied = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
'
'    If zlCheckRegistPlanIsValied(blnMulitNumPlan) = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
    '是否同一医生的安排时间段是否重复或交叉
    If Trim(cboDoctor.Text) <> "" Then
        Set rsDoctorPlan = GetDoctorPlan(lng医生ID, cboDoctor.Text)
        If rsDoctorPlan.RecordCount > 0 Then
            strSQL = "Select 时间段, 开始时间, Decode(Sign(终止时间 - 开始时间), 1, 终止时间 , 终止时间+ 1) 终止时间 From 时间段"
            Set rsNewDate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            Set rsOldDate = rsNewDate.Clone
        End If

        strInfo = ""
        For j = 1 To rsDoctorPlan.RecordCount
            strTmp = ""
            For i = 0 To IIf(opt天.Value, 6, vsPlan.Cols - 2)
               strOld = "" & rsDoctorPlan.Fields("D" & i).Value
               If opt天.Value Then
                   strNew = cbo天.Text
               Else
                   strNew = Trim(vsPlan.TextMatrix(1, i + 1))
               End If

               rsNewDate.Filter = "时间段='" & strNew & "'"
               rsOldDate.Filter = "时间段='" & strOld & "'"
               If rsNewDate.RecordCount > 0 And rsOldDate.RecordCount > 0 Then
                    If rsNewDate!开始时间 >= rsOldDate!开始时间 And rsNewDate!开始时间 <= rsOldDate!终止时间 Or rsNewDate!终止时间 >= rsOldDate!开始时间 And rsNewDate!终止时间 <= rsOldDate!终止时间 Or rsNewDate!开始时间 <= rsOldDate!开始时间 And rsNewDate!终止时间 >= rsOldDate!终止时间 Then
                    '时间交叉,再判断效期是否交叉
                         If chk有效期.Value = 0 Then
                             strTmp = strTmp & "," & "星期" & Choose(i + 1, "日", "一", "二", "三", "四", "五", "六") & ":" & strOld
                         Else
                             '为简化判断,假定数据按规范保存,开始时间和结束时间,要么都有,要么都没有,所以仅以开始时间来判断有无
                             If IsNull(rsDoctorPlan!开始时间) Then
                                 strTmp = strTmp & "," & "星期" & Choose(i + 1, "日", "一", "二", "三", "四", "五", "六") & ":" & strOld
                             Else
                                 If dtpBegin.Value >= CDate(rsDoctorPlan!开始时间) And dtpBegin.Value <= CDate(Nvl(rsDoctorPlan!终止时间, "3000-01-01")) Or dtpEnd.Value >= CDate(rsDoctorPlan!开始时间) And dtpEnd.Value <= CDate(Nvl(rsDoctorPlan!终止时间, "3000-01-01")) Or dtpBegin.Value <= CDate(rsDoctorPlan!开始时间) And dtpEnd.Value >= CDate(Nvl(rsDoctorPlan!终止时间, "3000-01-01")) Then
                                    strTmp = strTmp & "," & "星期" & Choose(i + 1, "日", "一", "二", "三", "四", "五", "六") & ":" & strOld
                                 End If
                             End If
                         End If
                    End If
               End If
            Next
            If strTmp <> "" Then
                strInfo = strInfo & vbCrLf & "在号别 [" & rsDoctorPlan!号码 & "] 中已有如下安排:" & vbCrLf & "        " & Mid(strTmp, 2)
                If Not IsNull(rsDoctorPlan!开始时间) Then
                    strInfo = strInfo & vbCrLf & "        有效期:" & rsDoctorPlan!开始时间 & "~" & rsDoctorPlan!终止时间
                Else
                    strInfo = strInfo & vbCrLf & "        有效期:不限"
                End If
            End If
            rsDoctorPlan.MoveNext
        Next
        If strInfo <> "" Then
            If blnMulitNumPlan Then
                '多次安排时,不能存在交叉
                Call MsgBox("发现" & cboDoctor.Text & "医生存在与当前号别重复或交叉的挂号安排" & vbCrLf & strInfo & vbCrLf & vbCrLf & "不能安排!", vbInformation + vbOKOnly, gstrSysName)
                Exit Sub
            Else
                If MsgBox("发现" & cboDoctor.Text & "医生存在与当前号别重复或交叉的挂号安排" & vbCrLf & strInfo & vbCrLf & vbCrLf & "确实要保存当前号别吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If

    If Not mEditType = edt_新增 Then
        If CheckExistsBooking() Then
            If MsgBox("该号别当前安排的时间段之外存在预约挂号单,是否要继续?", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    '先检查
    '取时间段
    str限号 = ""
    If opt天.Value Then '每天
        For i = 1 To 7
            str时间段 = str时间段 & "'" & Trim(cbo天.Text) & "',"
            str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
            str限号 = str限号 & "," & Val(txt限号.Text) & "," & IIf(chkAppoint.Value = 0, "0", txt限约.Text)
        Next
    Else
        For i = 1 To vsPlan.Cols - 1
            str时间段 = str时间段 & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"
            If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
                str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                    str限号 = str限号 & ",0,0"
                Else
                    str限号 = str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & IIf(chkAppoint.Value = 0, "0", Trim(vsPlan.TextMatrix(3, i)))
                End If
            End If
        Next
    End If
    If str限号 <> "" Then str限号 = Mid(str限号, 2)
    
    If chkAppoint.Value = 1 Then
        If txtAppLimit.Text <> "" And Val(txtAppLimit.Text) <= 0 Then
            MsgBox "允许预约的情况下，预约天数至少需要有1天！", vbInformation, gstrSysName
            txtAppLimit.SetFocus: Exit Sub
        End If
        lng预约天数 = Val(IIf(txtAppLimit.Text = "", gint预约天数, Val(txtAppLimit.Text)))
    Else
        lng预约天数 = 0
    End If

    '取挂号诊室
    '71253 李南春 2014-04-15 11:30:10 将listView 替换为vsflexGrid
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str诊室 = str诊室 & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str诊室 = Mid(str诊室, 2)
    
    '取分诊方式
    intCount = 0
    For i = 0 To opt分诊.UBound
        If opt分诊(i).Value Then intCount = i: Exit For
    Next

    '取开始时间范围
    strBegin = "NULL": strEnd = "NULL"
    If chk有效期.Value = 1 Then
        strBegin = "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strEnd = "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    End If

      '查看是否改变了排班或者 改变了 限号数 限约数 或者序号控制
    blnChange = (str限号 <> mPlanInfo.str限号) Or (str时间段 <> mPlanInfo.str排班)
    blnChange = blnChange Or (chk序号控制.Value <> IIf(mPlanInfo.bln序号, 1, 0))
    str限号 = "'" & str限号 & "',"
    Set cllPro = New Collection
    '取ID
    If mEditType = edt_新增 Then

        '新增
        lngNextID = zlDatabase.GetNextId("挂号安排")

        strSQL = "zl_挂号安排_INSERT(" & _
            lngNextID & ",'" & Trim(txt号别.Text) & "','" & cbo号类.Text & "'," & _
            cbo科室.ItemData(cbo科室.ListIndex) & "," & _
            cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "'," & _
            lng医生ID & "," & _
            chk病案.Value & "," & str时间段 & str限号 & intCount & "," & _
            "'" & str诊室 & "'," & strBegin & "," & strEnd & ",1," & chk序号控制.Value & ",0," & _
            5 & "," & lng预约天数 & ")"
    Else
'
' Zl_挂号安排_Insert
'(
'  Id_In       挂号安排.ID%Type,
'  号码_In     挂号安排.号码%Type,
'  号类_In     挂号安排.号类%Type,
'  科室id_In   挂号安排.科室id%Type,
'  项目id_In   挂号安排.项目id%Type,
'  医生_In     挂号安排.医生姓名%Type,
'  医生id_In   挂号安排.医生id%Type,
'  病案必须_In 挂号安排.病案必须%Type,
'  周日_In     挂号安排.周日%Type,
'  周一_In     挂号安排.周一%Type,
'  周二_In     挂号安排.周二%Type,
'  周三_In     挂号安排.周三%Type,
'  周四_In     挂号安排.周四%Type,
'  周五_In     挂号安排.周五%Type,
'  周六_In     挂号安排.周六%Type,
'  限号控制_In Varchar2,
'  分诊方式_In 挂号安排.分诊方式%Type,
'  诊室_In     Varchar2,
'  开始时间_In 挂号安排.开始时间%Type,
'  终止时间_In 挂号安排.终止时间%Type,
'  新增_In     Number,
'  序号控制_In 挂号安排.序号控制%Type,
'  处理类型_In Number:=0,
'  默认时段间隔_In 挂号安排.默认时段间隔%Type
') As
'  -----------------------------------------------------------
'  --参数：
'  --  诊室_IN=以';'号分隔的多个诊室名称
'  --  限号控制_IN:|周一,22(限号),13(限约)|周二,20(限号),11(限约)....
'  --  处理类型_IN:修改安排时 对时段数据的处理 0--不处理 1--删除时段信息
        '修改

        lngNextID = mlngID
        strSQL = "    " & vbNewLine & "zl_挂号安排_INSERT("
        strSQL = strSQL & vbNewLine & lngNextID
        strSQL = strSQL & vbNewLine & ",'" & (txt号别.Text) & "','" & cbo号类.Text & "',"
        strSQL = strSQL & vbNewLine & cbo科室.ItemData(cbo科室.ListIndex) & ","
        strSQL = strSQL & vbNewLine & cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "',"
        strSQL = strSQL & vbNewLine & lng医生ID & "," & chk病案.Value & ","
        strSQL = strSQL & vbNewLine & str时间段 & str限号 & intCount & ","
        strSQL = strSQL & vbNewLine & "'" & str诊室 & "'," & strBegin & "," & strEnd & ",0," & chk序号控制.Value & ","
        strSQL = strSQL & vbNewLine & IIf(chk序号控制.Value <> IIf(mPlanInfo.bln序号, 1, 0), 1, 0) & ","
        strSQL = strSQL & vbNewLine & 5 & "," & lng预约天数 & ")"

    End If

    On Error GoTo errH
    zlAddArray cllPro, strSQL

    LoadTimePlan True
    If SaveTimeSetData(lngNextID, cllPro) = False Then Exit Sub
    
    On Error GoTo Errhand
    zlExecuteProcedureArrAy cllPro, Me.Caption
    On Error GoTo 0
    mblnSucces = True

    If mEditType <> edt_新增 Then Unload Me: Exit Sub
    Call LoadCard
    mblnChangeByCode = True
    tbPage.Item(mPageIndex.EM_安排).Selected = True
    mblnChangeByCode = False
    Call ClearCustomData
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearCustomData()
     mTimeSet.str安排 = ""
     mTimeSet.bln序号控制 = False
     mTimeSet.lngSelIndex = 0
     mTimeSet.blnOnChange = False
     mTimeSet.lng安排ID = 0
     mTimeSet.lng计划ID = 0
     mTimeSet.blnIsInit = False
     Set mrs限号 = Nothing
     Set mTimeSet.rsRegPlan = Nothing
     Set mTimeSet.rsAssign = Nothing
     mTimeSet.strKey = ""
     mTimeSet.blnChange = False
     mTimeSet.str限制修改 = ""
     Set mTimeSet.rsHistory = Nothing
End Sub

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
                           bln时段 = True
                           Exit For
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
'                                Else
'                                     If MsgBox("在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")与限号数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
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
                   MsgBox "在分时段页面中的『" & str限制项目 & "』所设置的预约数(" & lng预约数 & ")大于了" & IIf(lng限号数 = lng限约数, "限号数(" & lng限约数 & ")", "限约数(" & lng限约数 & ")") & ",你不能按当前设置保存!", vbOKOnly, Me.Caption
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

Private Sub zlSaveTimePageSelected(ByVal str星期 As String)
    If tbPage.Selected Is Nothing Then Exit Sub
    If tbPage.Selected.index <> mPageIndex.EM_时段 Then
         tbPage.Item(mPageIndex.EM_时段).Selected = True
    End If
End Sub


Private Sub txtAppLimit_Validate(Cancel As Boolean)
    If chkAppoint.Value = 1 And txtAppLimit.Text <> "" And Val(txtAppLimit.Text) <= 0 Then
        MsgBox "允许预约时,预约天数不能小于等于0!", vbInformation, gstrSysName
        If txtAppLimit.Visible And txtAppLimit.Enabled Then txtAppLimit.SetFocus
        Cancel = True
    End If
End Sub

Private Sub vsTime_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str时段() As String
     If mTimeSet.bln序号控制 Then
        str时段 = Split(vsTime(index).EditText, "-")
        If UBound(str时段) <> 1 Then
           MsgBox "输入的时间格式有误!请检查!", vbOKOnly, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(str时段(0)) Then
           MsgBox "输入的时间格式有误!请检查!", vbOKOnly, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(str时段(1)) Then
           MsgBox "输入的时间格式有误!请检查!", vbOKOnly, gstrSysName
           Cancel = True: Exit Sub
        End If
        If CDate(str时段(0)) >= CDate(str时段(1)) Then
           MsgBox "开始时间必须小于结束时间!请检查!", vbOKOnly, gstrSysName
           Cancel = True
        End If
     End If
    mTimeSet.blnChange = True
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
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Function
    End If


    str时段 = mTimeSet.rsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbOKOnly, Me.Caption
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

Private Sub cmd设置时段_Click()
    If AssignReapportion(Val(txtTimeOut.Text), tbSubPage.Item(mTimeSet.lngSelIndex).Caption) = False Then Exit Sub
    Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
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
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Sub
    End If


    str时段 = mTimeSet.rsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbOKOnly, Me.Caption
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
                dat时点 = DateAdd("h", 1, dat时点)
                i = i - 1
            Else
                If i > lng固定数量 Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        If Format(DateAdd("n", lng时间间隔, dat时点), "hh:mm:00") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
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
    If VsTimeValidate(-1) = False Then
        Exit Function
    End If
    
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

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '返回:成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2009-09-15 13:14:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim bln所属部门 As Boolean
    Dim lngColsWidth As Long
    Dim intRow As Integer
    
    Err = 0: On Error GoTo Errhand:
    gint号长 = GetMaxLen
    mintSysAppLimit = Val(zlDatabase.GetPara("挂号允许预约天数", glngSys))
    
    strSQL = "" & _
    "   Select '    ' 时间段 From dual Union All  " & _
    "   Select 时间段 From 时间段"
    Set mrs时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsPlan
        .Clear 1
        .Tag = .BuildComboList(mrs时间段, "时间段")
        
        .ColComboList(1) = .BuildComboList(mrs时间段, "时间段")
        For i = 2 To .Cols - 1
            .ColComboList(i) = .ColComboList(0)
        Next
    End With
    With cbo天
        Do While Not mrs时间段.EOF
            cbo天.AddItem Nvl(mrs时间段!时间段)
            mrs时间段.MoveNext
        Loop
        .ListIndex = 0
    End With
    
   '取出门诊临床科室
    Set mrs科室 = GetDepartments("'临床'", "1,3", Not zlStr.IsHavePrivs(mstrPrivs, "所有科室"))
    If mrs科室.RecordCount = 0 Then
        MsgBox "你不具备可用的临床科室信息或你权限不足,请先到部门管理中进行设置或找系统管理员分配权限！", vbInformation, gstrSysName
        Exit Function
    End If
    
    cbo科室.Clear
    Do While Not mrs科室.EOF
        cbo科室.AddItem mrs科室!名称
        cbo科室.ItemData(cbo科室.NewIndex) = Val(Nvl(mrs科室!id))
        If mlng缺省挂号科室ID = Val(Nvl(mrs科室!id)) Then cbo科室.ListIndex = cbo科室.NewIndex  '刘兴洪:增加从主界面中传入的科室
        mrs科室.MoveNext
    Loop
        
    '挂号项目
    strSQL = "Select ID as 序号,名称 From 收费项目目录 " & _
        " Where 类别='1' And (Sysdate Between 建档时间 And 撤档时间 Or 建档时间<Sysdate And 撤档时间 Is Null)" & _
        " And (站点='" & gstrNodeNo & "' Or 站点 is Null) " & _
        " Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用的挂号项目信息,请先到挂号项目设置中初始！", vbInformation, gstrSysName
        Exit Function
    End If
    cboItem.Clear
    Do While Not rsTemp.EOF
        cboItem.AddItem rsTemp!名称
        cboItem.ItemData(cboItem.NewIndex) = rsTemp!序号
        rsTemp.MoveNext
    Loop
    
    '号类
    strSQL = "Select 编码,名称,缺省标志 From 号类 Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbo号类.Clear
    Do While Not rsTemp.EOF
        cbo号类.AddItem rsTemp!名称
        If IIf(IsNull(rsTemp!缺省标志), 0, rsTemp!缺省标志) = 1 Then
            cbo号类.ListIndex = cbo号类.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    
    '门诊诊室
    strSQL = "Select 编码,名称　From 门诊诊室 Where (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '71253 李南春 2014-04-14 16:05:10 诊室名称显示不全
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
            If lngColsWidth > .ClientWidth Then .Height = .Height + 130: Frame4.Height = Frame4.Height + 30
            .Editable = flexEDKbdMouse
        End With
    End If
    
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub Form_Load()
    Dim intType As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    mblnFirst = True
    Call InitPage
    mblnOnly院内医生 = Val(zlDatabase.GetPara("只允许选院内医生", glngSys, mlngModule, "0", , InStr(1, mstrPrivs, ";参数设置;") > 0, intType)) = 1
    If mblnOnly院内医生 Then
        mnuViewDoctor(0).Checked = True
        mnuViewDoctor(1).Checked = False
    Else
        mnuViewDoctor(0).Checked = False
        mnuViewDoctor(1).Checked = True
    End If
    Call LoadTimeSetControl
    Call LoadvsDept
    lbl医生.Tag = IIf(mblnOnly院内医生, "0", "1")
    lbl医生.Caption = IIf(mblnOnly院内医生, "院内医生", "医生") & IIf(lbl医生.Tag = "1", "↓", "")
    lbl医生.ToolTipText = IIf(mblnOnly院内医生, "只能选院内建档医生", "含外援医生(除了可以选择院内医生外，还可以输入外援医生)")
    
End Sub

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

Private Sub Form_Activate()
    Dim i As Integer, intIndex As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitData = False Then Unload Me: Exit Sub
    If LoadCard = False Then Unload Me: Exit Sub
    Call cboDoctor_Validate(False)
    For i = 0 To opt分诊.UBound
        If opt分诊(i).Value Then Call opt分诊_Click(i): Exit For
    Next
    If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
    strSQL = "Select 1 From 挂号安排计划 Where 安排id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If Not rsTmp.EOF Then
        mblnMinorChange = True
        If txtAppLimit.Enabled And txtAppLimit.Visible Then
            txtAppLimit.SetFocus
            zlControl.TxtSelAll txtAppLimit
        End If
        tbPage.Item(1).Visible = False
        txt号别.Enabled = False
        chkAppoint.Enabled = False
        chk序号控制.Enabled = False
        chk病案.Enabled = False
        Frame3.Enabled = False
        cbo号类.Enabled = False
        cboDoctor.Enabled = False
        cbo科室.Enabled = False
        cboItem.Enabled = False
        opt分诊.Item(0).Enabled = False
        opt分诊.Item(1).Enabled = False
        opt分诊.Item(2).Enabled = False
        opt分诊.Item(3).Enabled = False
        vsPlan.HighLight = flexHighlightNever
        If zlStr.IsHavePrivs(mstrPrivs, "修改诊室") = False Then
            For i = 0 To 3
                If opt分诊(i).Value = True Then intIndex = i
            Next i
            opt分诊(0).Enabled = False
            opt分诊(1).Enabled = False
            opt分诊(2).Enabled = False
            opt分诊(3).Enabled = False
            opt分诊(intIndex).Value = True
            vsDept.Enabled = False
        End If
    Else
        mblnMinorChange = False
    End If
    mblnChangeDist = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cbo科室 Then Exit Sub
    If Me.ActiveControl Is cboDoctor Then Exit Sub
    If Me.ActiveControl Is vsPlan Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mstr限制修改 = ""
    Set mcll预约信息 = Nothing
    If Not mrs上班时间段 Is Nothing Then
        Set mrs上班时间段 = Nothing
    End If
    If Not mrs限号 Is Nothing Then
        Set mrs限号 = Nothing
    End If
    If Not mrsRegHistory Is Nothing Then
        Set mrsRegHistory = Nothing
    End If
    If Not mrsRegNewData Is Nothing Then
        Set mrsRegNewData = Nothing
    End If
    If Not mrsRegOldData Is Nothing Then
        Set mrsRegOldData = Nothing
    End If
    '72729:刘尔旋,2014-05-06,第一次修改点击取消或者右上角的X按钮后，再次修改时间段显示不正确的问题
    Call ClearCustomData
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
 

Private Sub lbl医生_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 0 Then Exit Sub
        If Val(lbl医生.Tag) = 0 Then Exit Sub
        
        PopupMenu mnuPopu, 2
End Sub

Private Sub ClearVsGridCheckValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除网格控件的复选框值
    '编制:李南春
    '日期:2014-04-14 18:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
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

Private Sub mfrmTime_zlSaveTimePageSelected(ByVal str星期 As String)
       If tbPage.Selected Is Nothing Then Exit Sub
       If tbPage.Selected.index <> mPageIndex.EM_时段 Then
            tbPage.Item(mPageIndex.EM_时段).Selected = True
       End If
End Sub
Private Sub mnuViewDoctor_Click(index As Integer)
        mnuViewDoctor(index).Checked = True
        If index = 0 Then
            mnuViewDoctor(1).Checked = False: mblnOnly院内医生 = True
        Else
            mnuViewDoctor(0).Checked = False: mblnOnly院内医生 = False
        End If
 
        lbl医生.Caption = IIf(mblnOnly院内医生, "院内医生", "医生") & "↓"
        lbl医生.ToolTipText = IIf(mblnOnly院内医生, "只能选择院内建档医生", "含外援医生(除了可以选择院内医生外，还可以输入外援医生)")
End Sub
Private Sub opt分诊_Click(index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    
    '71253 李南春 2014-04-15 11:30:10 将listView 替换为vsflexGrid
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
    mblnChangeDist = True
End Sub

Private Sub opt天_Click()
    Dim i As Integer
    Dim strPlan As String
    
    For i = 0 To vsPlan.Cols - 1
        If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
            If strPlan = "" Then
                strPlan = vsPlan.TextMatrix(1, i)
            Else
                If vsPlan.TextMatrix(1, i) <> strPlan Then
                    strPlan = "": Exit For
                End If
            End If
        End If
    Next
    
    opt天.Value = -True: txt限号.Enabled = True: txt限约.Enabled = (chkAppoint.Value = 1)
    cbo天.Enabled = True
    
    opt周.Value = False
    With vsPlan
        .Enabled = False: .TabStop = False
        For i = 1 To 7
             .TextMatrix(1, i) = ""
             .TextMatrix(2, i) = ""
             .TextMatrix(3, i) = ""
        Next
    End With
    
    cbo天.ListIndex = cbo.FindIndex(cbo天, strPlan, True)
    cbo天.SetFocus
End Sub

Private Sub opt周_Click()
    Dim i As Integer
    
    If Trim(cbo天.Text) <> "" Then
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(1, i) = cbo天.Text
            vsPlan.TextMatrix(2, i) = txt限号.Text
            vsPlan.TextMatrix(3, i) = txt限约.Text
        Next
    End If
    
    opt天.Value = False
    cbo天.Enabled = False: txt限号.Enabled = False: txt限约.Enabled = False
    cbo天.ListIndex = -1

    opt周.Value = True
    vsPlan.Enabled = True: vsPlan.TabStop = True
    vsPlan.Col = 1: vsPlan.SetFocus
End Sub

Private Sub txt号别_GotFocus()
    Call zlControl.TxtSelAll(txt号别)
End Sub

Private Sub txt号别_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限号_GotFocus()
    Call zlControl.TxtSelAll(txt限号)
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
        MsgBox "限号数应大于限约数!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txt限约_GotFocus()
    Call zlControl.TxtSelAll(txt限约)
End Sub

Private Sub txt限约_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If Val(txt限号.Text) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限约_Validate(Cancel As Boolean)
    If Val(txt限号.Text) < Val(txt限约.Text) And _
        Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" Then
        MsgBox "限约数应小于限号数!", vbInformation, gstrSysName
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
    Dim i As Long
    On Error GoTo errHandle
    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    lng项目id = cboItem.ItemData(cboItem.ListIndex)
    lng医生ID = 0: str医生 = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    strSQL = "" & _
        "   Select 号码,序号,周日 D0,周一 D1,周二 D2,周三 D3,周四 D4,周五 D5,周六 D6," & _
        "           To_Char(开始时间,'YYYY-MM-DD HH24:MI:SS') 开始时间,To_Char(终止时间,'YYYY-MM-DD HH24:MI:SS') 终止时间" & _
        "   From 挂号安排  "

    If lng医生ID = 0 Then
        strSQL = strSQL & _
            "   Where 科室id=[1] and  项目ID =[2] and 医生姓名=[3] and nvl(医生ID,0)=0 and ID<>" & mlngID & " Order by 序号"
    Else
        strSQL = strSQL & _
        "   Where 科室id=[1] and  项目ID =[2] and  医生ID=[4] and ID<>" & mlngID & " Order by 序号"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, lng项目id, str医生, lng医生ID)
    blnMulitNumPlan = Not rsTemp.EOF
    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
    str号别 = ""
    Do While Not rsTemp.EOF
        str号别 = str号别 & "," & Nvl(rsTemp!号码)
        If opt天.Value Then
            If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " 周日:" & Nvl(rsTemp!D0)
            If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " 周一:" & Nvl(rsTemp!D1)
            If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " 周二:" & Nvl(rsTemp!D2)
            If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " 周三:" & Nvl(rsTemp!D3)
            If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " 周四:" & Nvl(rsTemp!D4)
            If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " 周五:" & Nvl(rsTemp!D5)
            If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " 周六:" & Nvl(rsTemp!D6)
            If strTemp <> "" Then
                strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下安排:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        Else
            With vsPlan
                For i = 0 To 6
                    strTemp1 = "周" & Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                    If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                        '存在,肯定重复了
                        strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                    End If
                Next
            End With
            If strTemp <> "" Then
                strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下安排:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        End If
        rsTemp.MoveNext
    Loop
    If str号别 <> "" Then str号别 = Mid(str号别, 2)
    If MsgBox("注意:" & vbCrLf & "   发现『" & cboDoctor.Text & "』医生已经存在如下安排:" & vbCrLf & "    " & str号别 & vbCrLf & "   是否继续对该医生进行安排?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        zlCheckRegistPlanIsValied = True: Exit Function
    End If
    zlCheckRegistPlanIsValied = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function zlCheckPlanArrageIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查计划安排是否有效
    '返回:检查计划安排是否存在相关的安排,如果有相关的安排,则返回False,否则返回true
    '编制:刘兴洪
    '日期:2010-12-29 19:53:56
    '问题目:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str医生 As String
    Dim lng项目id As Long, lng科室ID As Long, lng医生ID As Long
    Dim str号别 As String, strTemp As String, strTemp1 As String
    Dim blnCheck As Boolean
    Dim i As Long
    On Error GoTo errHandle
    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    lng项目id = cboItem.ItemData(cboItem.ListIndex)
    lng医生ID = 0: str医生 = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  distinct A.号码,A.周日 D0,A.周一 D1,A.周二 D2,A.周三 D3,A.周四 D4,A.周五 D5,A.周六 D6," & _
    "           To_Char(生效时间,'YYYY-MM-DD HH24:MI:SS') 生效时间,To_Char(失效时间,'YYYY-MM-DD HH24:MI:SS') 失效时间" & _
    "   From 挂号安排计划 A, 挂号安排 B " & _
    "   Where A.安排ID=B.ID    " & _
    "      and   B.科室id=[1] and  B.项目ID =[2] and B.医生姓名=[3] and nvl(B.医生ID,0)=[4] and B.ID<>" & mlngID & _
    "   Order by 号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, lng项目id, str医生, lng医生ID)
    If rsTemp.EOF Then
        zlCheckPlanArrageIsValied = True: Exit Function
    End If
    Do While Not rsTemp.EOF
        str号别 = str号别 & "," & Nvl(rsTemp!号码)
        blnCheck = chk有效期.Value = 0
        If chk有效期.Value = 1 Then
            blnCheck = Nvl(rsTemp!生效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!生效时间) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Nvl(rsTemp!失效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!失效时间) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)
            blnCheck = blnCheck Or Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)
             
        End If
        If blnCheck Then
            If opt天.Value Then
                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " 周日:" & Nvl(rsTemp!D0)
                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " 周一:" & Nvl(rsTemp!D1)
                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " 周二:" & Nvl(rsTemp!D2)
                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " 周三:" & Nvl(rsTemp!D3)
                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " 周四:" & Nvl(rsTemp!D4)
                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " 周五:" & Nvl(rsTemp!D5)
                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " 周六:" & Nvl(rsTemp!D6)
                If strTemp <> "" Then
                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下计划安排:" & vbCrLf & "        " & Mid(strTemp, 2)
                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            Else
                With vsPlan
                    For i = 0 To 6
                        strTemp1 = "周" & Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                            '存在,肯定重复了
                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                        End If
                    Next
                End With
                If strTemp <> "" Then
                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下计划安排:" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  生效时间:" & IIf(Nvl(rsTemp!生效时间) = "1901-01-01", "无限", Nvl(rsTemp!生效时间) & "-" & Nvl(rsTemp!失效时间)) & vbCrLf
                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    zlCheckPlanArrageIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If mEditType = edt_查阅 Then Cancel = True: Exit Sub
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
    Dim strTmp As String
    Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
    vsPlan.ColComboList(NewCol) = ""
     
    If mstr限制修改 <> "" Then
        strTmp = ";周" & vsPlan.TextMatrix(0, NewCol) & ";"
        vsPlan.Editable = flexEDKbdMouse
        If InStr(mstr限制修改, strTmp) > 0 And NewRow = 1 Then vsPlan.Editable = flexEDNone
    End If
    If OldRow = 2 And Trim(vsPlan.TextMatrix(3, OldCol)) = "" And mbln自动默认限约数 Then
        vsPlan.TextMatrix(3, OldCol) = vsPlan.TextMatrix(2, OldCol)
    End If
    
    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
        vsPlan.TextMatrix(2, OldCol) = ""
        vsPlan.TextMatrix(3, OldCol) = ""
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
    Dim strKey As String, intCol As Integer, strTemp As String, strTmp As String
    '数据验证
    With vsPlan
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If .Row <= 1 Then Exit Sub
        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        If Val(strKey) <> 0 Then
            strKey = Format(Abs(Val(strKey)), "####;;;")
        End If
         If mstr限制修改 <> "" Then
               strTmp = "周" & vsPlan.TextMatrix(0, Col)
               'vsPlan.Editable = flexEDKbdMouse
               If InStr(mstr限制修改, ";" & strTmp & ";") > 0 Then
                   If mcll预约信息 Is Nothing Then
                        Cancel = Val(strKey) < Val(.TextMatrix(Row, Col))
                   Else
                        Cancel = Val(mcll预约信息("K" & strTmp & "_数量")) > Val(strKey)
                        If Cancel Then Exit Sub
                        If chk序号控制.Value = 1 Then
                            Cancel = Val(mcll预约信息("K" & strTmp & "_序号")) > Val(strKey)
                        End If
                   End If
               End If
        End If
        If Cancel Then Exit Sub
        If Row = 2 Then
            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
                If MsgBox("限号数小于了限约数,是否清空限约数?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
                .TextMatrix(3, Col) = ""
            End If
        ElseIf Row = 3 Then
            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
                Call MsgBox("限号数小于了限约数,不能继续", vbOKOnly, gstrSysName)
                Cancel = True: Exit Sub
            End If
        End If

        .EditText = strKey
    End With
End Sub


Private Function Check时段() As Boolean
    '----------------------------------
    '判断是否分时段
    '----------------------------------
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset

    If mEditType = edt_查阅 Or mEditType = edt_新增 Then Exit Function

    On Error GoTo Hd
    strSQL = _
    "   Select 1 As Hdata From 挂号安排时段 Where 安排id =[1] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
     Check时段 = Not rsTmp.EOF
    Set rsTmp = Nothing
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function zl_Get预约信息(ByVal lng安排ID As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim cll预约信息 As Collection
    strSQL = "    " & vbCrLf & " Select 日期, Max(预约日期) As 预约日期, Max(数量) As 数量,序号"
    strSQL = strSQL & vbCrLf & " From ("
    strSQL = strSQL & vbCrLf & "    Select Decode(To_Char(A.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',"
    strSQL = strSQL & vbCrLf & "                    '7', '周六') As 日期, To_Char(A.发生时间, 'yyyy-mm-dd') As 预约日期, Count(Rownum) As 数量, B.ID,Max(Nvl(A.号序,0)) as 序号 "
    strSQL = strSQL & vbCrLf & "    From 病人挂号记录 A, 挂号安排　b"
    strSQL = strSQL & vbCrLf & "    Where A.号别 = B.号码 And A.记录状态 = 1 And b.ID = [1] And"
    strSQL = strSQL & vbCrLf & "          A.发生时间 > A.登记时间"
'    If gint预约天数 = 0 Then
    strSQL = strSQL & " And A.发生时间 > Sysdate "
'    Else
'        strSQL = strSQL & " And A.发生时间 Between Sysdate And Sysdate+" & gint预约天数
'    End If
    strSQL = strSQL & vbCrLf & "    Group By To_Char(A.发生时间, 'yyyy-mm-dd'),"
    strSQL = strSQL & vbCrLf & "              Decode(To_Char(A.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',"
    strSQL = strSQL & vbCrLf & "                      '周五', '7', '周六'), B.ID)"
    strSQL = strSQL & vbCrLf & " Group By 日期,序号"
  On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    If rsTmp.EOF Then Exit Function
    Set cll预约信息 = New Collection
    Do While Not rsTmp.EOF
        If InStr(strTmp, Nvl(rsTmp!日期)) <= 0 Or strTmp = "" Then
            strTmp = strTmp & ";" & Nvl(rsTmp!日期)
            cll预约信息.Add Nvl(rsTmp!数量), "K" & Nvl(rsTmp!日期) & "_数量"
            cll预约信息.Add Nvl(rsTmp!预约日期), "K" & Nvl(rsTmp!日期) & "_日期"
            cll预约信息.Add Nvl(rsTmp!序号), "K" & Nvl(rsTmp!日期) & "_序号"
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then strTmp = strTmp & ";"
    Set mcll预约信息 = cll预约信息
    zl_Get预约信息 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Property Let 自动默认限约数(ByVal vNewValue As Boolean)
    mbln自动默认限约数 = vNewValue
End Property


Private Sub InitPage()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_计划安排, "计划安排", picBaseBack.Hwnd, 0)
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

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    PageChange Item
End Sub

Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    If Item.index = mPageIndex.EM_时段 Then
       mblnChangeByCode = True
       tbPage.Item(mPageIndex.EM_安排).Selected = True
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

Private Sub LoadTimePlan(Optional ByVal blnSaveBeforCheck As Boolean = False)
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
                !id = 0
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
                        !id = Val(mlngID)
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

     zlShowPagePlan str安排, mrsRegNewData, mrsRegHistory, chk序号控制.Value = 1, Switch(mEditType = ed_计划安排, EM_安排_增加, mEditType = Ed_安排修改, EM_安排_修改, True, EM_安排_查阅), mlngID, Val(0), blnSaveBeforCheck, str应诊时段
End Sub

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
       If tbSubPage.Item(i).Selected Then
            Call tbSubPage_SelectedChanged(tbSubPage.Item(i))
            Exit For
       End If
    Next
 End Sub

Private Function LoadRegHistory() As Boolean
    Dim strSQL As String
    strSQL = "Select 限制项目, Max(最大序号) As 最大序号, Max(统计) As 统计, Max(发生时间) As 发生时间" & vbNewLine & _
            " From (Select Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六') As 限制项目," & vbNewLine & _
            "              Max(Nvl(a.号序, 0)) As 最大序号, Count(1) As 统计, To_Char(Max(发生时间), 'hh24:mi:ss') As 发生时间," & vbNewLine & _
            "              To_Char(发生时间, 'YYYY-MM-DD') As 发生日期" & vbNewLine & _
            "       From 病人挂号记录 A, 挂号安排 B" & vbNewLine & _
            "       Where a.记录状态 = 1 And a.发生时间 Between Sysdate And Sysdate + Nvl(b.预约天数, " & IIf(gint预约天数 = 0, 15, gint预约天数) & ") And a.号别 = b.号码 And b.Id = [1] " & vbNewLine & _
            "       Group By Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六')," & vbNewLine & _
            "                To_Char(发生时间, 'YYYY-MM-DD'))" & vbNewLine & _
            " Group By 限制项目"
                    
    On Error GoTo Hd:
    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    LoadRegHistory = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function IsValied() As Boolean
     Dim i As Integer, intCount As Integer, j As Integer
    Dim str时间段 As String, str诊室 As String, str限号 As String
    Dim lngNextID As Long, lng医生ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim str号别 As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '是否多次安排
    Dim blnChange       As Boolean '是否改变了 时间安排
    Dim strMsg          As String

    If opt天.Value Then
        If cbo天.ListIndex = -1 Then
            MsgBox "该号别每天的应诊时间未设置！", vbInformation, gstrSysName
            cbo天.SetFocus: Exit Function
        End If

        If Val(txt限号.Text) = 0 And Val(txt限约.Text) = 0 Then
            MsgBox "安排设置时段时,必须设置限号或限约数！", vbInformation, gstrSysName
            If txt限号.Visible And txt限号.Enabled Then txt限号.SetFocus
            Exit Function
        End If
        If (chkAppoint.Value = 0 And chk序号控制.Value = 0) Or (chkAppoint.Value = 1 And txt限约.Text <> "" And Val(txt限约.Text) = 0 And chk序号控制.Value = 0) Then
            MsgBox "非序号控制的安排设置时段时,必须是可预约的安排！", vbInformation, gstrSysName
            If txt限号.Visible And txt限号.Enabled Then txt限号.SetFocus
            Exit Function
        End If
        '限号限约规则
        If Trim(txt限号.Text) <> "" Then
            If Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                txt限约.SetFocus: Exit Function
            End If
        ElseIf Trim(txt限约.Text) <> "" Then
            MsgBox "限约必须限号！", vbInformation, gstrSysName
            txt限号.SetFocus: Exit Function
        End If
    Else
        If chkAppoint.Value = 0 And chk序号控制.Value = 0 Then
            MsgBox "非序号控制的安排设置时段时,必须是可预约的安排！", vbInformation, gstrSysName
            Exit Function
        End If
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))

                        If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                            MsgBox "安排设置时段时,必须设置限号或限约数！", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
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
    IsValied = True
End Function

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
    mblnChangeDist = True
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
