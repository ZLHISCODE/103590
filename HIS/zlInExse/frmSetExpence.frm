VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab stab 
      Height          =   5175
      Left            =   105
      TabIndex        =   33
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "记帐参数"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTY"
      Tab(0).Control(1)=   "txt转出"
      Tab(0).Control(2)=   "txtOutDay0"
      Tab(0).Control(3)=   "fraDoctor"
      Tab(0).Control(4)=   "lst收费类别"
      Tab(0).Control(5)=   "UDOutDay(0)"
      Tab(0).Control(6)=   "chk转出"
      Tab(0).Control(7)=   "fra药房"
      Tab(0).Control(8)=   "lblOutDate(0)"
      Tab(0).Control(9)=   "Label1"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "结帐参数(&1)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk缺省金额"
      Tab(1).Control(1)=   "chkRefundStyle"
      Tab(1).Control(2)=   "chk(14)"
      Tab(1).Control(3)=   "UDOutDay(1)"
      Tab(1).Control(4)=   "txtOutDay1"
      Tab(1).Control(5)=   "lblOutDate(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "结帐票据控制(&2)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblInUse"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblOutUse"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraTitle"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdPrintSetup"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdListPrintSet"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmd退款收据"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdBillZY"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cboInvoiceKindZY"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "fra票据格式"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdOwnFee"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboInvoiceKindMZ"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdBillMZ"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "其他票据控制(&3)"
      TabPicture(3)   =   "frmSetExpence.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdRed"
      Tab(3).Control(1)=   "fraRed"
      Tab(3).Control(2)=   "cmdPrepayPrintSet"
      Tab(3).Control(3)=   "fraPrepay"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdBillMZ 
         Caption         =   "结帐票据设置(&P)"
         Height          =   350
         Left            =   5055
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1560
      End
      Begin VB.ComboBox cboInvoiceKindMZ 
         Height          =   300
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1950
         Width           =   3270
      End
      Begin VB.CheckBox chk缺省金额 
         Caption         =   "结帐退款时选择现金结算缺省退款金额"
         Height          =   255
         Left            =   -74655
         TabIndex        =   26
         Top             =   1635
         Width           =   3525
      End
      Begin VB.CommandButton cmdRed 
         Caption         =   "结帐红票打印设置(&S)"
         Height          =   350
         Left            =   -71235
         TabIndex        =   50
         Top             =   4170
         Width           =   1965
      End
      Begin VB.Frame fraRed 
         Caption         =   "作废红票格式"
         Height          =   1515
         Left            =   -74970
         TabIndex        =   48
         Top             =   2550
         Width           =   6600
         Begin VSFlex8Ctl.VSFlexGrid vsRedFormat 
            Height          =   1155
            Left            =   60
            TabIndex        =   49
            Top             =   225
            Width           =   6375
            _cx             =   11245
            _cy             =   2037
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
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":007C
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
      Begin VB.CommandButton cmdPrepayPrintSet 
         Caption         =   "预交票据打印设置(&S)"
         Height          =   350
         Left            =   -74190
         TabIndex        =   43
         Top             =   4170
         Width           =   1965
      End
      Begin VB.CommandButton cmdOwnFee 
         Caption         =   "自费清单设置(&4)"
         Height          =   350
         Left            =   165
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1860
      End
      Begin VB.Frame fra票据格式 
         Caption         =   "收费票据格式"
         Height          =   1620
         Left            =   120
         TabIndex        =   32
         Top             =   2655
         Width           =   6540
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1320
            Left            =   60
            TabIndex        =   39
            Top             =   225
            Width           =   6330
            _cx             =   11165
            _cy             =   2328
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
            FormatString    =   $"frmSetExpence.frx":010E
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
      Begin VB.ComboBox cboInvoiceKindZY 
         Height          =   300
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2325
         Width           =   3270
      End
      Begin VB.Frame fraTY 
         Height          =   1170
         Left            =   -74760
         TabIndex        =   0
         Top             =   690
         Width           =   2400
         Begin VB.CheckBox chk 
            Caption         =   "住院留观病人记帐"
            Height          =   195
            Index           =   5
            Left            =   195
            TabIndex        =   3
            Top             =   840
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊留观病人记帐"
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   2
            Top             =   570
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "开单人定开单科室"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   1
            Top             =   300
            Width           =   1740
         End
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用预交票据"
         Height          =   1995
         Left            =   -74970
         TabIndex        =   41
         Top             =   450
         Width           =   6600
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1455
            Left            =   75
            TabIndex        =   42
            Top             =   270
            Width           =   6375
            _cx             =   11245
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
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":01B4
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
      Begin VB.CommandButton cmdBillZY 
         Caption         =   "结帐票据设置(&P)"
         Height          =   350
         Left            =   5055
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2295
         Width           =   1560
      End
      Begin VB.CommandButton cmd退款收据 
         Caption         =   "退款收据设置(&3)"
         Height          =   350
         Left            =   4050
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4305
         Width           =   1620
      End
      Begin VB.CommandButton cmdListPrintSet 
         Caption         =   "打印费用明细设置(1)"
         Height          =   350
         Left            =   165
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4305
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "回单票据打印设置(&2)"
         Height          =   350
         Left            =   2115
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4305
         Width           =   1860
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用收费票据"
         Height          =   1470
         Left            =   90
         TabIndex        =   30
         Top             =   435
         Width           =   6540
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1125
            Left            =   75
            TabIndex        =   31
            Top             =   255
            Width           =   6330
            _cx             =   11165
            _cy             =   1984
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
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":0294
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
      Begin VB.CheckBox chkRefundStyle 
         Caption         =   "结帐退款缺省按预交缴款方式"
         Height          =   255
         Left            =   -74655
         TabIndex        =   25
         Top             =   1335
         Width           =   3525
      End
      Begin VB.TextBox txt转出 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73725
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "3"
         Top             =   3690
         Width           =   255
      End
      Begin VB.TextBox txtOutDay0 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73920
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "设置为 0 表示只能选择在院病人"
         Top             =   3315
         Width           =   450
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "显示开单人"
         Height          =   1170
         Left            =   -72045
         TabIndex        =   4
         Top             =   720
         Width           =   1755
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按简码"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   5
            Top             =   435
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按编码"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   6
            Top             =   735
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Index           =   14
         Left            =   -74655
         TabIndex        =   24
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   1065
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   1
         Left            =   -73410
         TabIndex        =   34
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay1"
         BuddyDispid     =   196637
         OrigLeft        =   1486
         OrigTop         =   3375
         OrigRight       =   1726
         OrigBottom      =   3645
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOutDay1 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73860
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "0"
         ToolTipText     =   "设置为 0 表示只能选择在院病人"
         Top             =   690
         Width           =   450
      End
      Begin VB.ListBox lst收费类别 
         Height          =   3000
         Left            =   -70095
         Style           =   1  'Checkbox
         TabIndex        =   16
         ToolTipText     =   "请复选允许使用的收费类别"
         Top             =   930
         Width           =   1545
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   0
         Left            =   -73470
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay0"
         BuddyDispid     =   196633
         OrigLeft        =   1486
         OrigTop         =   2760
         OrigRight       =   1726
         OrigBottom      =   3030
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chk转出 
         Caption         =   "显示最近   天转出的病人"
         Height          =   195
         Left            =   -74715
         TabIndex        =   20
         Top             =   3720
         Width           =   2370
      End
      Begin VB.Frame fra药房 
         Caption         =   " 药房与发料部门设置 "
         Height          =   1185
         Left            =   -74745
         TabIndex        =   7
         Top             =   1965
         Width           =   4470
         Begin VB.ComboBox cbo卫材 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   1305
         End
         Begin VB.ComboBox cbo中药 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   1305
         End
         Begin VB.ComboBox cbo西药 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   1305
         End
         Begin VB.ComboBox cbo成药 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发料部门"
            Height          =   180
            Left            =   2100
            TabIndex        =   14
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中草药"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "西成药"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中成药"
            Height          =   180
            Left            =   2280
            TabIndex        =   10
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Label lblOutUse 
         AutoSize        =   -1  'True
         Caption         =   "门诊结帐票据使用"
         Height          =   180
         Left            =   150
         TabIndex        =   51
         Top             =   2010
         Width           =   1440
      End
      Begin VB.Label lblInUse 
         AutoSize        =   -1  'True
         Caption         =   "住院结帐票据使用"
         Height          =   180
         Left            =   150
         TabIndex        =   27
         Top             =   2385
         Width           =   1440
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许选择         天内出院的病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74700
         TabIndex        =   17
         Top             =   3360
         Width           =   2790
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许选择         天内出院的病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   -74655
         TabIndex        =   22
         Top             =   750
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入类别:"
         Height          =   180
         Left            =   -70125
         TabIndex        =   35
         Top             =   660
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   7035
      TabIndex        =   46
      Top             =   1710
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7035
      TabIndex        =   44
      Top             =   645
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7035
      TabIndex        =   45
      Top             =   1185
      Width           =   1110
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7035
      TabIndex        =   47
      Top             =   4275
      Width           =   1110
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytInFun As Byte '0=记帐,1=结帐
Public mbytUseType As Byte '0:普通记帐,1-科室分散记帐,2-医技科室记帐
Public mstrPrivs As String
Public mlngModul As Long
Public mblnOnlyDrugStock As Boolean  '仅显示药房设置
Private Enum chkBPS
    C0记帐 = 0
    C1划价 = 1
    C2审核 = 2
End Enum
Private Enum chks
    C03开单人定科室 = 3
    C04门诊留观记帐 = 4
    C05住院留观记帐 = 5
    C09医保结帐不打 = 9
    C14LED欢迎信息 = 14
End Enum
Private Enum InvoiceKind
    C1收费收据 = 1
    C3结帐收据 = 3
    C4多种收据 = 10
End Enum
Private Const CModule As Long = 1150    '住院记帐操作
Private mstrOptPrivs As String


Private Sub zlOnlyDrugStrock()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:仅显示药房的相关设置
    '编制:刘兴洪
    '日期:2010-01-25 15:24:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    Err = 0: On Error GoTo ErrHand:
    If Not (mblnOnlyDrugStock And mbytInFun = 0) Then Exit Sub
    
    For Each ctl In Me.Controls
       Select Case UCase(TypeName(ctl))
       Case UCase("ImageList")
       Case UCase("sstab")
            ctl.Visible = True
       Case Else
            If ctl Is fra药房 Or ctl.Container Is fra药房 Or ctl Is cmdOK Or ctl Is cmdCancel Then
                ctl.Visible = True
            Else
                 ctl.Visible = False
            End If
       End Select
    Next
    
    fra药房.Top = fraTY.Top
    Me.Height = 3525: Me.Width = 5470
    cmdCancel.Top = ScaleHeight - cmdCancel.Height - 100
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 100
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    stab.Height = cmdOK.Top - stab.Top - 100
    stab.Width = ScaleWidth - stab.Left * 2
    stab.TabCaption(0) = "药房设置"
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboInvoiceKindZY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo成药_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub cbo卫材_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo西药_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub cbo中药_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 

Private Sub chkRefundStyle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk缺省金额_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'问题:27380
Private Sub chk转出_Click()
    txt转出.Enabled = chk转出.Value = 1
    If txt转出.Visible And txt转出.Enabled Then txt转出.SetFocus
End Sub
Private Sub chk转出_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdListPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me)
End Sub

Private Sub cmdOwnFee_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_4", Me)
End Sub

Private Sub cmdPrintSetup_Click()
     Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me)
End Sub
Private Sub cmd退款收据_Click()
    '刘兴洪 问题:27776 日期:2010-02-04 16:44:39
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me)
End Sub

 
 
Private Sub lst收费类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDoctorKind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtOutDay0_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtOutDay1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt转出_GotFocus()
   zlControl.TxtSelAll txt转出
End Sub

Private Sub txt转出_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cboInvoiceKindZY_Click()
    Dim bytKind As Byte
    If Visible Then '启动时强制调用
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3结帐收据
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1收费收据
        Else
            bytKind = InvoiceKind.C4多种收据
        End If
        Call InitShareInvoice(bytKind)
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3结帐收据, InvoiceKind.C1收费收据))
        'Call SetFactBillFormat
    End If
End Sub

Private Sub cboInvoiceKindMZ_Click()
    Dim bytKind As Byte
    If Visible Then '启动时强制调用
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3结帐收据
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1收费收据
        Else
            bytKind = InvoiceKind.C4多种收据
        End If
        Call InitShareInvoice(bytKind)
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3结帐收据, InvoiceKind.C1收费收据))
        'Call SetFactBillFormat
    End If
End Sub

Private Sub cmdBillZY_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdBillMZ_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdRed_Click()
    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6"), Me)
End Sub

Private Sub cmdCancel_Click()
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1137)
End Sub

Private Sub cmdHelp_Click()
    Select Case stab.Tab
        Case 0
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence1"
        Case 1
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence2"
    End Select
End Sub

Private Sub cmdPrepayPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdOK_Click()
    Dim strValue As String, i As Long, lngShareID As Long
    Dim blnHavePrivs As Boolean, strTemp As String
    Dim blnBillOptSet As Boolean
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If mbytInFun = 0 And cbo西药.Visible Then
        If cbo西药.ListIndex = -1 And cbo西药.ListCount > 0 And cbo西药.Enabled Then
            MsgBox "请选择西药房.", vbInformation, gstrSysName
            stab.Tab = 0: cbo西药.SetFocus: Exit Sub
        End If
        If cbo成药.ListIndex = -1 And cbo成药.ListCount > 0 And cbo成药.Enabled Then
            MsgBox "请选择成药房.", vbInformation, gstrSysName
            stab.Tab = 0: cbo成药.SetFocus: Exit Sub
        End If
        If cbo中药.ListIndex = -1 And cbo中药.ListCount > 0 And cbo中药.Enabled Then
            MsgBox "请选择中药房.", vbInformation, gstrSysName
            stab.Tab = 0: cbo中药.SetFocus: Exit Sub
        End If
        If cbo卫材.ListIndex = -1 And cbo卫材.ListCount > 0 And cbo卫材.Enabled Then
            MsgBox "请选择卫材发料部门.", vbInformation, gstrSysName
            stab.Tab = 0: cbo卫材.SetFocus: Exit Sub
        End If
    End If
    '保存参数注册信息
    '当不使用门诊留观记帐时,检查如果不显示门诊科室是否有其它可用记帐科室
    If mbytInFun = 0 And (mbytUseType = 0 Or mbytUseType = 1) And chk(chks.C04门诊留观记帐).Value = 0 Then
        If Not CheckUnits Then
            MsgBox "当不使用门诊留观记帐时,你没有可以记帐的科室,参数无法被设置！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    If mbytInFun = 0 Then
        blnBillOptSet = InStr(1, mstrOptPrivs, ";记帐选项设置;") > 0
    
        '药房
        zlDatabase.SetPara "缺省中药房", IIf(cbo中药.ListIndex = 0, "0", cbo中药.ItemData(cbo中药.ListIndex)), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "缺省西药房", IIf(cbo西药.ListIndex = 0, "0", cbo西药.ItemData(cbo西药.ListIndex)), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "缺省成药房", IIf(cbo成药.ListIndex = 0, "0", cbo成药.ItemData(cbo成药.ListIndex)), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "缺省发料部门", IIf(cbo卫材.ListIndex = 0, "0", cbo卫材.ItemData(cbo卫材.ListIndex)), glngSys, CModule, blnBillOptSet
        If mblnOnlyDrugStock Then GoTo GoOver:
        
        '1150的参数
        '--------------------------------------------------------------------------------
        '收费类别
        For i = lst收费类别.ListCount - 1 To 0 Step -1
            If lst收费类别.Selected(i) Then strValue = strValue & "'" & Chr(lst收费类别.ItemData(i)) & "',"
        Next
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "收费类别", strValue, glngSys, CModule, blnBillOptSet
    
           
        '留观病人记帐
        zlDatabase.SetPara "门诊留观病人记帐", chk(chks.C04门诊留观记帐).Value, glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "住院留观病人记帐", chk(chks.C05住院留观记帐).Value, glngSys, CModule, blnBillOptSet
        
        zlDatabase.SetPara "出院病人天数", Val(txtOutDay0.Text), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "开单人显示方式", IIf(optDoctorKind(0).Value, 1, 2), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "科室医生", IIf(chk(chks.C03开单人定科室).Value = 1, 0, 1), glngSys, CModule, blnBillOptSet
        
        If mbytUseType = 1 Then
            '刘兴洪 问题:27380 日期:2010-01-22 14:45:32
            zlDatabase.SetPara "最近转出天数", IIf(chk转出.Value = 1, "1", "0") & "|" & Val(txt转出.Text), glngSys, mlngModul, blnHavePrivs
        End If
    Else
        '本地共用结帐票据
        zlDatabase.SetPara "住院结帐票据类型", cboInvoiceKindZY.ListIndex, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "门诊结帐票据类型", cboInvoiceKindMZ.ListIndex, glngSys, mlngModul, blnHavePrivs
        Call SaveInvoice
        
'        lngShareID = 0
'        For i = 1 To lvwBill.ListItems.Count
'            If lvwBill.ListItems(i).Checked Then lngShareID = Val(Mid(lvwBill.ListItems(i).Key, 2))
'        Next
'        zlDatabase.SetPara "共用结帐票据批次", lngShareID, glngSys, mlngModul, blnHavePrivs
        
        'LED设备
        zlDatabase.SetPara "LED显示欢迎信息", chk(chks.C14LED欢迎信息).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "出院病人天数", Val(txtOutDay1.Text), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "结帐退款缺省方式", chkRefundStyle.Value, glngSys, mlngModul, blnHavePrivs  '30036
        zlDatabase.SetPara "退款现金结算缺省金额", chk缺省金额.Value, glngSys, mlngModul, blnHavePrivs
    End If
GoOver:
    If mblnOnlyDrugStock Then
        Call zlInit药房
    Else
        Call InitLocPar(mlngModul)
    End If
    gblnOK = True
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If stab.TabVisible(0) Then
        If chk(chks.C03开单人定科室).Visible And chk(chks.C03开单人定科室).Enabled Then chk(chks.C03开单人定科室).SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub


Private Sub Load药房()
    Dim rsTmp As ADODB.Recordset
        
    On Error GoTo errH
    Set rsTmp = GetDepartments("'中药房','西药房','成药房','发料部门'", "2,3")
        
    cbo中药.AddItem "人工选择"
    cbo西药.AddItem "人工选择"
    cbo成药.AddItem "人工选择"
    cbo卫材.AddItem "人工选择"
    
    If Not rsTmp.EOF Then
        rsTmp.Filter = "工作性质='中药房'"
        Do While Not rsTmp.EOF
            cbo中药.AddItem rsTmp!名称
            cbo中药.ItemData(cbo中药.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "工作性质='西药房'"
        Do While Not rsTmp.EOF
            cbo西药.AddItem rsTmp!名称
            cbo西药.ItemData(cbo西药.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "工作性质='成药房'"
        Do While Not rsTmp.EOF
            cbo成药.AddItem rsTmp!名称
            cbo成药.ItemData(cbo成药.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "工作性质='发料部门'"
        Do While Not rsTmp.EOF
            cbo卫材.AddItem rsTmp!名称
            cbo卫材.ItemData(cbo卫材.ListCount - 1) = rsTmp!ID
                            
            rsTmp.MoveNext
        Loop
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim i As Long, strValue As String, blnParSet As Boolean, blnBillOptSet As Boolean
    Dim strDefault As String
    Dim varData As Variant
    Dim bytKind As Byte
    
    gblnOK = False
    On Error GoTo errH
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If mbytInFun = 0 Then
        mstrOptPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p记帐操作) & ";"
        blnBillOptSet = InStr(1, mstrOptPrivs, "记帐选项设置") > 0
        '不是1150的参数
        '--------------------------------------------------------------------------------------
        
        '1150的参数
        '------------------------------------------------------------------
        '收费类别(挂号除外)
        strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 序号"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        Do While Not rsTmp.EOF
            lst收费类别.AddItem rsTmp!类别
            lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
            rsTmp.MoveNext
        Loop
        strValue = zlDatabase.GetPara("收费类别", glngSys, CModule, , Array(lst收费类别), blnBillOptSet)
        If strValue = "" Then
            For i = 0 To lst收费类别.ListCount - 1
                lst收费类别.Selected(i) = True
            Next
        Else
            For i = 0 To lst收费类别.ListCount - 1
                If InStr(strValue, Chr(lst收费类别.ItemData(i))) Then lst收费类别.Selected(i) = True
            Next
        End If
        If lst收费类别.ListCount > 0 Then lst收费类别.TopIndex = 0: lst收费类别.ListIndex = 0
        
        '留观病人记帐
        chk(chks.C04门诊留观记帐).Value = IIf(zlDatabase.GetPara("门诊留观病人记帐", glngSys, CModule, , Array(chk(chks.C04门诊留观记帐)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C05住院留观记帐).Value = IIf(zlDatabase.GetPara("住院留观病人记帐", glngSys, CModule, , Array(chk(chks.C05住院留观记帐)), blnBillOptSet) = "1", 1, 0)
                      
        txtOutDay0.Text = Val(zlDatabase.GetPara("出院病人天数", glngSys, CModule, 0, Array(txtOutDay0, lblOutDate(0), UDOutDay(0)), blnBillOptSet))
        If Val(zlDatabase.GetPara("开单人显示方式", glngSys, CModule, 0, Array(optDoctorKind(0), optDoctorKind(1)), blnBillOptSet)) = 1 Then
            optDoctorKind(0).Value = True
        Else
            optDoctorKind(1).Value = True
        End If
        
        
        chk(chks.C03开单人定科室).Value = IIf(zlDatabase.GetPara("科室医生", glngSys, CModule, , Array(chk(chks.C03开单人定科室)), blnBillOptSet) = "1", 0, 1)
        
                
        
       
        '--------------------------
        Call Load药房
        
        strValue = zlDatabase.GetPara("缺省中药房", glngSys, CModule, , Array(cbo中药), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo中药, strValue, True)
        If cbo中药.ListIndex = -1 And Val(strValue) = 0 Then cbo中药.ListIndex = 0
        
        strValue = zlDatabase.GetPara("缺省西药房", glngSys, CModule, , Array(cbo西药), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo西药, strValue, True)
        If cbo西药.ListIndex = -1 And Val(strValue) = 0 Then cbo西药.ListIndex = 0
        
        strValue = zlDatabase.GetPara("缺省成药房", glngSys, CModule, , Array(cbo成药), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo成药, strValue, True)
        If cbo成药.ListIndex = -1 And Val(strValue) = 0 Then cbo成药.ListIndex = 0
        
        strValue = zlDatabase.GetPara("缺省发料部门", glngSys, CModule, , Array(cbo卫材), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo卫材, strValue, True)
        If cbo卫材.ListIndex = -1 And Val(strValue) = 0 Then cbo卫材.ListIndex = 0
        
        
        chk转出.Visible = False: txt转出.Visible = False
        If mbytUseType = 1 Then
            '刘兴洪 问题:27380 日期:2010-01-22 14:45:32
            chk转出.Visible = True: txt转出.Visible = True
            Dim str转出 As String
            'CModule
            str转出 = zlDatabase.GetPara("最近转出天数", glngSys, mlngModul, "0|3", Array(chk转出, txt转出), InStr(1, mstrPrivs, ";参数设置;") > 0)
            txt转出.Text = Val(Split(str转出 & "|", "|")(1))
            chk转出.Value = IIf(Val(Split(str转出 & "|", "|")(0)) = 1, 1, 0)
        End If
        
    ElseIf mbytInFun = 1 Then
        chkRefundStyle.Value = IIf(Val(zlDatabase.GetPara("结帐退款缺省方式", glngSys, mlngModul, , Array(chkRefundStyle), blnParSet)) = 1, 1, 0)
        chk缺省金额.Value = IIf(Val(zlDatabase.GetPara("退款现金结算缺省金额", glngSys, mlngModul, , Array(chk缺省金额), blnParSet)) = 1, 1, 0)
        
        cboInvoiceKindZY.AddItem "住院医疗费收据"
        cboInvoiceKindZY.AddItem "门诊医疗费收据"
        i = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, 0, Array(cboInvoiceKindZY), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindZY.ListIndex = i
        
        cboInvoiceKindMZ.AddItem "住院医疗费收据"
        cboInvoiceKindMZ.AddItem "门诊医疗费收据"
        i = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, 0, Array(cboInvoiceKindMZ), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindMZ.ListIndex = i
        
        If InStr(1, mstrPrivs, ";门诊费用结帐;") = 0 Then '不允许对门诊费用结帐时,只能使用住院医疗费收据
            cboInvoiceKindZY.ListIndex = 0
            cboInvoiceKindZY.Enabled = False
            cboInvoiceKindMZ.Enabled = False
        End If
        
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3结帐收据
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1收费收据
        Else
            bytKind = InvoiceKind.C4多种收据
        End If
        Call InitShareInvoice(bytKind)
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3结帐收据, InvoiceKind.C1收费收据))
        '问题:35142
        'Call SetFactBillFormat '设置普通和医保病人结帐发票格式
        'LED设备
        chk(chks.C14LED欢迎信息).Value = IIf(zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, "1", Array(chk(chks.C14LED欢迎信息)), blnParSet) = "1", 1, 0)
        txtOutDay1.Text = Val(zlDatabase.GetPara("出院病人天数", glngSys, mlngModul, 0, Array(txtOutDay1, lblOutDate(1), UDOutDay(1)), blnParSet))
    End If
    If mbytInFun = 0 Then
        stab.TabVisible(1) = False
        stab.TabVisible(2) = False
        stab.TabVisible(3) = False
        '问题:27380
        txt转出.Visible = mbytUseType = 1 '科室分散记帐
        chk转出.Visible = mbytUseType = 1 '科室分散记帐

    ElseIf mbytInFun = 1 Then
        stab.TabVisible(0) = False
    End If
    Call zlOnlyDrugStrock
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'
'Private Sub SetShareInvoice(ByVal bytKind As Byte)
'    Dim rstmp As New ADODB.Recordset, strSQL As String
'    Dim i As Long, lngShareID As Long
'    Dim objItem As ListItem
'
'    '读取可用公用结帐领用
'    Set rstmp = GetShareInvoiceGroupID(bytKind)
'    lngShareID = Val(zlDatabase.GetPara("共用结帐票据批次", glngSys, mlngModul, 0, Array(lvwBill), InStr(1, mstrPrivs, ";参数设置;") > 0))
'    lvwBill.ListItems.Clear
'    For i = 1 To rstmp.RecordCount
'        Set objItem = lvwBill.ListItems.Add(, "_" & rstmp!ID, rstmp!领用人, , 1)
'        objItem.SubItems(1) = Format(rstmp!登记时间, "yyyy-MM-dd")
'        objItem.SubItems(2) = rstmp!开始号码 & "," & rstmp!终止号码
'        objItem.SubItems(3) = rstmp!剩余数量
'        If rstmp!ID = lngShareID Then
'            objItem.Checked = True
'            objItem.Selected = True
'            lngShareID = 0
'        End If
'        rstmp.MoveNext
'    Next
'    If lngShareID <> 0 Then zlDatabase.SetPara "共用结帐票据批次", 0, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
'
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    mbytUseType = 0
    zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("预交类型"))) = Trim(.Cell(flexcpData, i, .ColIndex("预交类型"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub

Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("选择")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub
'
'Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim i As Long
'    For i = 1 To lvwBill.ListItems.Count
'        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
'    Next
'    Item.Selected = True
'End Sub

Private Sub txtOutDay0_GotFocus()
    zlControl.TxtSelAll txtOutDay0
End Sub

Private Sub txtOutDay0_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOutDay1_GotFocus()
    zlControl.TxtSelAll txtOutDay1
End Sub

Private Sub txtOutDay1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function CheckUnits() As Boolean
'功能：检查按参数设置之后,是否有可用记帐临床科室
'说明：当不使用门诊留观记帐之后,将不显示门诊临床科室
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lng病区ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    '有权则显示门诊观察室对应的临床科室,住院留观与住院相同
    If InStr(mstrPrivs, ";门诊留观记帐;") And (chk(chks.C04门诊留观记帐).Value = 1) Then
        strSQL = "1,2,3"
    Else
        strSQL = "2,3"
    End If
    If InStr(";" & mstrPrivs, ";所有病区;") > 0 Then
        strSQL = _
             " Select Distinct A.ID,A.编码,A.名称" & _
             " From 部门表 A,部门性质说明 B" & _
             " Where B.部门ID = A.ID And B.服务对象 IN(" & strSQL & ") And B.工作性质 IN('临床','手术')" & _
             " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
             " Order by A.编码"
    Else
        '求有权限的科室：本身所在科室+所属病区包含的科室
        '#当操作员属于门诊观察室时，即使没有门诊留观记帐的权限,也显示对应的门诊临床科室,但无法记帐
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And B.服务对象 IN(" & strSQL & ") And B.工作性质 IN('临床','手术')" & _
            " Order by A.编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    CheckUnits = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'
'Private Sub SetFactBillFormat()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:设置发票格式
'    '编制:刘兴洪
'    '日期:2010-12-31 19:29:48
'    '问题:35142
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strRptName As String, rstmp As ADODB.Recordset, i As Long, blnParSet As Boolean, strSQL As String
'    blnParSet = zlStr.IsHavePrivs(mstrPrivs, ";参数设置;")
'    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
'    cboFactNormal.Clear: cboFactMediCare.Clear
'
'    cboFactNormal.AddItem "使用本地缺省格式"
'    cboFactMediCare.AddItem "使用本地缺省格式"
'    '    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
'    strSQL = "" & _
'    "   Select B.说明,B.序号 From zlReports A,zlRptFmts B" & _
'    "    Where A.ID=B.报表ID And A.编号=[1] " & _
'    "   Order by b.序号"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
'    For i = 1 To rstmp.RecordCount
'        cboFactNormal.AddItem rstmp!说明
'        cboFactNormal.ItemData(cboFactNormal.NewIndex) = rstmp!序号
'        cboFactMediCare.AddItem rstmp!说明
'        cboFactMediCare.ItemData(cboFactMediCare.NewIndex) = rstmp!序号
'        rstmp.MoveNext
'    Next
'    cboFactNormal.ListIndex = 0: cboFactMediCare.ListIndex = 0
'    i = Val(zlDatabase.GetPara("普通发票格式", glngSys, mlngModul, , Array(lblFactNormal, cboFactNormal), blnParSet))
'    Call zlControl.CboLocate(cboFactNormal, i, True)
'    i = Val(zlDatabase.GetPara("医保发票格式", glngSys, mlngModul, , Array(lblFactMediCare, cboFactMediCare), blnParSet))
'    Call zlControl.CboLocate(cboFactMediCare, i, True)
'End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(Row, .ColIndex("使用类别"))) = Trim(.TextMatrix(i, .ColIndex("使用类别"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
 
Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("选择")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub vsBillFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "结帐票据格式", False, False
End Sub
Private Sub vsBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "结帐票据格式", False, False
End Sub

Private Sub vsBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillFormat
        Select Case Col
        Case .ColIndex("住院结帐票据格式")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case .ColIndex("门诊结帐票据格式")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsRedFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "结帐红票格式", False, False
End Sub
Private Sub vsRedFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "结帐红票格式", False, False
End Sub

Private Sub vsRedFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRedFormat
        Select Case Col
        Case .ColIndex("票据格式")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存发票相关票据
    '编制:刘兴洪
    '日期:2011-04-28 18:16:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '保存共享票据
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("使用类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用结帐票据批次", strValue, glngSys, mlngModul, blnHavePrivs
    
    '保存预交票据
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("预交类型")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用预交票据批次", strValue, glngSys, mlngModul, blnHavePrivs
    
    Dim strPrintMode As String
    '保存门诊格式
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("门诊结帐票据格式")))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "门诊结帐发票格式", strValue, glngSys, mlngModul, blnHavePrivs
    End With
    
    '保存住院格式
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("住院结帐票据格式")))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "住院结帐发票格式", strValue, glngSys, mlngModul, blnHavePrivs
    End With
    
    strValue = "": strPrintMode = ""
    With vsRedFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "作废发票格式", strValue, glngSys, mlngModul, blnHavePrivs
    End With
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-28 18:24:16
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
     If mbytInFun <> 0 Then isValied = True: Exit Function
     
    isValied = False
    On Error GoTo errHandle
    '检查每种使用种式只能一个选择
    With vsBill
        str类别 = "-"
        For i = 1 To vsBill.Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("使用类别"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("使用类别")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("使用类别"))) = Trim(.TextMatrix(j, .ColIndex("使用类别"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    使用类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    
      '检查每种使用预交只能一个选择
    With vsPrepay
        str类别 = "-"
        For i = 1 To .Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("预交类型"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("预交类型")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("预交类型"))) = Trim(.TextMatrix(j, .ColIndex("预交类型"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    预交类型为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    

    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitShareInvoice(ByVal intKind As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '     intKind:
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant
    Dim varType As Variant, varTemp1 As Variant
    Dim intTYPE As Integer, intType1 As Integer, intType2 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strRptName As String, blnHavePrivs As Boolean
    Dim strPrintMode As String, varDataMZ As Variant
    Dim str合约单位结帐 As String, strShareInvoiceMZ As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    On Error GoTo errHandle
    
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
    zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "结帐票据格式", False, False
    zl_vsGrid_Para_Restore mlngModul, vsRedFormat, Me.Name, "结帐红票格式", False, False
    strShareInvoice = zlDatabase.GetPara("共用结帐票据批次", glngSys, mlngModul, , , True, intTYPE)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And Not blnHavePrivs Then .Editable = flexEDNone
    End With
    
    
    '格式:领用ID1,使用类别1|领用IDn,使用类别n|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(intKind)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!使用类别, " ")
            .TextMatrix(lngRow, .ColIndex("领用人")) = NVL(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(NVL(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    '住院票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("住院结帐票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    
    strRptName = IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    '门诊票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
    With vsBillFormat
        .ColComboList(.ColIndex("门诊结帐票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    
    '读取参数值
    strShareInvoice = zlDatabase.GetPara("住院结帐发票格式", glngSys, mlngModul, , , True, intTYPE)
    strShareInvoiceMZ = zlDatabase.GetPara("门诊结帐发票格式", glngSys, mlngModul, , , True, intTYPE)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat
         .ColData(.ColIndex("住院结帐票据格式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intTYPE
        Case 1, 3, 5, 15
             .ColData(.ColIndex("住院结帐票据格式")) = IIf(intTYPE = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("住院结帐票据格式"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("结帐后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    With vsBillFormat
         .ColData(.ColIndex("门诊结帐票据格式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intTYPE
        Case 1, 3, 5, 15
             .ColData(.ColIndex("门诊结帐票据格式")) = IIf(intTYPE = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("门诊结帐票据格式"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("结帐后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    varData = Split(strShareInvoice, "|")
    varDataMZ = Split(strShareInvoiceMZ, "|")
    strSQL = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("住院结帐票据格式")) = "0"
            .TextMatrix(lngRow, .ColIndex("门诊结帐票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("住院结帐票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varDataMZ)
                varTemp = Split(varDataMZ(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("门诊结帐票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("住院结帐票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("住院结帐票据格式"), .Rows - 1, .ColIndex("住院结帐票据格式")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("门诊结帐票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("门诊结帐票据格式"), .Rows - 1, .ColIndex("门诊结帐票据格式")) = vbBlue
        End If
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6")
    '票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
    With vsRedFormat
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    
    '读取参数值
    strShareInvoice = zlDatabase.GetPara("作废发票格式", glngSys, mlngModul, , , True, intTYPE)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsRedFormat
         .ColData(.ColIndex("票据格式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intTYPE
        Case 1, 3, 5, 15
             .ColData(.ColIndex("票据格式")) = IIf(intTYPE = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("结帐后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    
    varData = Split(strShareInvoice, "|")
    strSQL = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsRedFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("票据格式"), .Rows - 1, .ColIndex("票据格式")) = vbBlue
        End If
    End With
    
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModul, , , True, intTYPE)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,预交类别ID1|领用IDn,预交类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!使用类别, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "门诊和住院共用"
            ElseIf Val(NVL(rsTemp!使用类别, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
            Else
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = Val(NVL(rsTemp!使用类别))
            
            .TextMatrix(lngRow, .ColIndex("领用人")) = NVL(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(NVL(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("预交类型"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


