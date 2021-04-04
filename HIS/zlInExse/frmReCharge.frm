VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmReCharge 
   Caption         =   "病人费用销帐申请"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   14805
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboRemark 
      Height          =   330
      Left            =   11520
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame fraTop 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   0
      Left            =   45
      TabIndex        =   27
      Top             =   45
      Visible         =   0   'False
      Width           =   12195
      Begin VB.CheckBox chk项目 
         Caption         =   "已执行项目"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   683
         Width           =   1365
      End
      Begin VB.CheckBox chk项目 
         Caption         =   "未执行项目"
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   9
         Top             =   683
         Width           =   1350
      End
      Begin VB.Frame fraPatiInfor 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   135
         TabIndex        =   43
         Tag             =   "2700"
         Top             =   255
         Width           =   2865
         Begin VB.TextBox txtPatient 
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1155
            MaxLength       =   100
            TabIndex        =   1
            Tag             =   "1580"
            Top             =   0
            Width           =   1605
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   495
            TabIndex        =   44
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmReCharge.frx":038A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "宋体"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            DefaultCardType =   "0"
            BackColor       =   -2147483633
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            Caption         =   "病人"
            ForeColor       =   &H80000007&
            Height          =   210
            Index           =   7
            Left            =   0
            TabIndex        =   45
            Top             =   75
            Width           =   420
         End
      End
      Begin VB.ComboBox cbo次数 
         Height          =   330
         Left            =   10215
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   285
         Width           =   1845
      End
      Begin VB.CheckBox chkShowOthers 
         Caption         =   "显示他科执行费用"
         Height          =   315
         Left            =   6885
         TabIndex        =   11
         Top             =   683
         Width           =   2070
      End
      Begin VB.ComboBox cboBaby 
         Height          =   330
         Left            =   10215
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   675
         Width           =   1845
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "忽略期间"
         Height          =   255
         Left            =   9000
         TabIndex        =   6
         Top             =   323
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpApplyE 
         Height          =   360
         Left            =   6885
         TabIndex        =   5
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   93323267
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpApplyB 
         Height          =   360
         Left            =   3915
         TabIndex        =   3
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   93323267
         CurrentDate     =   36257
      End
      Begin zl9InExse.ComboxExpend cboKind 
         Height          =   360
         Left            =   3915
         TabIndex        =   10
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   9
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收费类别"
         Height          =   210
         Left            =   3015
         TabIndex        =   51
         Top             =   735
         Width           =   840
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "费用期间"
         Height          =   210
         Left            =   3015
         TabIndex        =   49
         Top             =   315
         Width           =   855
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   210
         Left            =   6360
         TabIndex        =   48
         Top             =   315
         Width           =   255
      End
      Begin VB.Label lblPatiInfo 
         Caption         =   "性别：     年龄：        住院号：             床号：         科室：       病区：      付款方式："
         Height          =   210
         Left            =   120
         TabIndex        =   47
         Top             =   1065
         Width           =   11895
      End
      Begin VB.Label lblShowBabyFee 
         AutoSize        =   -1  'True
         Caption         =   "婴儿费显示"
         Height          =   210
         Left            =   9060
         TabIndex        =   46
         Top             =   735
         Width           =   1050
      End
   End
   Begin VB.Frame fraTop 
      Height          =   1140
      Index           =   1
      Left            =   210
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   10470
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   7545
         TabIndex        =   14
         Top             =   665
         Width           =   1380
      End
      Begin VB.CheckBox chkDateAudit 
         Caption         =   "忽略期间"
         Height          =   255
         Left            =   6090
         TabIndex        =   36
         Top             =   705
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpAuditE 
         Height          =   360
         Left            =   3915
         TabIndex        =   13
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   93323267
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpAuditB 
         Height          =   360
         Left            =   1305
         TabIndex        =   2
         Top             =   660
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   93323267
         CurrentDate     =   36257
      End
      Begin VB.Label lblAuditDate 
         BackStyle       =   0  'Transparent
         Caption         =   "申请期间"
         Height          =   210
         Left            =   420
         TabIndex        =   0
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   210
         Left            =   3555
         TabIndex        =   4
         Top             =   735
         Width           =   255
      End
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   435
      ScaleWidth      =   11295
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5070
      Width           =   11295
   End
   Begin VB.CheckBox chkVerfy 
      Caption         =   "销帐申请同时完成审核"
      Height          =   420
      Left            =   10650
      TabIndex        =   40
      Top             =   1590
      Width           =   2505
   End
   Begin VB.CommandButton cmdAudit 
      Caption         =   "审核(&A)"
      Height          =   350
      Left            =   12945
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   34
      Tag             =   "已处理"
      Top             =   4440
      Width           =   10905
      _cx             =   19235
      _cy             =   1508
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":046D
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
      OwnerDraw       =   1
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
   Begin VB.Frame fraCmd 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdCancelRefuse 
         Caption         =   "取消拒绝(&C)"
         Height          =   350
         Left            =   4440
         TabIndex        =   42
         Top             =   30
         Width           =   1350
      End
      Begin VB.CommandButton cmdOKAudit 
         Caption         =   "确认审核(&S)"
         Height          =   350
         Left            =   2940
         TabIndex        =   21
         Top             =   30
         Width           =   1350
      End
   End
   Begin VB.Frame fraCmd 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   26
      Top             =   2160
      Visible         =   0   'False
      Width           =   6570
      Begin VB.CheckBox chkOtherOperator 
         Caption         =   "显示他人销帐申请"
         Height          =   315
         Left            =   3015
         TabIndex        =   41
         Top             =   30
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.CommandButton cmdCancelApply 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消申请(&C)"
         Height          =   350
         Left            =   5130
         TabIndex        =   20
         ToolTipText     =   "热键：F2"
         Top             =   0
         Width           =   1350
      End
      Begin VB.ComboBox cboState 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   15
         Width           =   1815
      End
      Begin VB.Label lblState 
         Caption         =   "审核状态"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.Frame fraCmd 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2535
      TabIndex        =   24
      Top             =   1620
      Visible         =   0   'False
      Width           =   8070
      Begin VB.CommandButton cmdSeleItem 
         Caption         =   "…"
         Height          =   300
         Left            =   4605
         TabIndex        =   38
         Top             =   15
         Width           =   300
      End
      Begin VB.TextBox txtFeeItem 
         Height          =   350
         Left            =   1080
         TabIndex        =   16
         Top             =   0
         Width           =   3855
      End
      Begin VB.CommandButton cmdAllDetail 
         Caption         =   "所有费用(&A)"
         Height          =   350
         Left            =   5175
         TabIndex        =   17
         Top             =   0
         Width           =   1350
      End
      Begin VB.CommandButton cmdOKApply 
         Caption         =   "确认申请(&S)"
         Height          =   350
         Left            =   6585
         TabIndex        =   18
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label lblItem 
         Caption         =   "销帐项目"
         Height          =   255
         Left            =   195
         TabIndex        =   25
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   12945
      TabIndex        =   23
      Top             =   600
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   12945
      TabIndex        =   22
      Top             =   120
      Width           =   1350
   End
   Begin MSComctlLib.TabStrip tbsType 
      Height          =   375
      Left            =   45
      TabIndex        =   15
      Top             =   1620
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "销帐申请"
            Key             =   "T1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "已申请明细"
            Key             =   "T2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   1695
      Left            =   0
      TabIndex        =   32
      Tag             =   "明细"
      Top             =   5610
      Width           =   7245
      _cx             =   12779
      _cy             =   2990
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":04E2
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
      OwnerDraw       =   1
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
      Height          =   3495
      Index           =   0
      Left            =   0
      TabIndex        =   31
      Tag             =   "待处理"
      Top             =   2085
      Width           =   10905
      _cx             =   19235
      _cy             =   6165
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":0557
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
      OwnerDraw       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfTogether 
      Height          =   1695
      Left            =   7320
      TabIndex        =   37
      Tag             =   "明细"
      ToolTipText     =   "一并给药药品"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3570
      _cx             =   6297
      _cy             =   2990
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReCharge.frx":05CC
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
      OwnerDraw       =   1
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   39
      Top             =   7875
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReCharge.frx":05F6
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmReCharge.frx":0E8A
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmReCharge.frx":1064
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
End
Attribute VB_Name = "frmReCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mbytUseType As Byte  '0-病区调用,1-医技科室调用,2-医生站调用(只能申请药品，且无审核功能)
Public mbytFun As Byte      '0-申请,1-审核
Public mlngDeptID As Long   '病区调用时传入当前操作的病人病区ID,医技科室调用时传入医技科室ID
Public mstrPrivs As String
Public mlngPatientID As Long '传入病人ID
Public mstrInNO As String
Public mlngAdviceID As Long
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private Const mlngModul = 1150
Private Const HeadApply = "类别,4,850|项目名称,1,3500|规格,1,2500|产地,1,2500|药品来源,1,1200|单位,1,550|数量,7,850|销帐数量,7,1000|销帐金额,7,1000|原始销帐数量,7,0|原始销帐金额,7,0"
Private Const HeadApplied = "选择,4,850|姓名,1,850|性别,1,550|类别,1,850|项目名称,1,2500|规格,1,2000|产地,1,2500|药品来源,1,1200|单位,1,550|销帐数量,7,1000|销帐金额,7,1000|申请人,1,850|申请时间,1,2100"
Private Const HeadAudit = "审核,4,550|姓名,1,850|性别,1,550|病人病区,1,1100|床号,1,650|类别,1,850|项目名称,1,2500|规格,1,2000|产地,1,2500|药品来源,1,1200|单位,1,550|销帐数量,7,1000|销帐金额,7,1000|申请人,1,850|申请时间,1,2100"
Private Const HeadAudited = "状态,4,550|姓名,1,850|性别,1,550|病人病区,1,1200|床号,1,650|类别,1,850|项目名称,1,2500|规格,1,2000|产地,1,2500|药品来源,1,1200|单位,1,550|销帐数量,7,1000|申请人,1,850|申请时间,1,2100"
Private Const HeadApplyDetail = "执行状态,4,1000|婴儿费,4,600|NO,4,1000|发生时间,1,2100|执行科室,1,1200|开单科室,1,1200|单价,7,1250|付数,7,850|数次,7,850|应收金额,7,1050|实收金额,7,1050|销帐数量,7,1000|销帐金额,7,1000|销帐原因,1,2500|原始销帐数量,7,0|原始销帐金额,7,0|原始销帐原因,1,0"
Private Const HeadAppliedDetail = "NO,4,1000|发生时间,1,2100|执行科室,1,1200|开单科室,1,1200|销帐数量,7,1000|销帐金额,7,1000|销帐原因,1,2500"
Private Const HeadAuditDetail = "NO,4,1000|发生时间,1,2100|开单科室,1,1200|销帐数量,7,1000|销帐原因,1,2500"
Private Const HeadAuditedDetail = "NO,4,1000|发生时间,1,2100|开单科室,1,1200|销帐数量,7,1000|销帐原因,1,2500"
Private mblnInit As Boolean
Private mrsApplyDept As ADODB.Recordset
Private mblnOperatorICU As Boolean  '当前操作员是ICU科室的
Private mblnPatiDeptICU As Boolean '病人当前科室是否为ICU病人
Private mrsOperatorDept As ADODB.Recordset '操作员部门ID
Private mblnOperatorNurse As Boolean '当前操作员是否护士
Private mstrOperatorDeptIDs As String  '操作员所属科室ID(性质为"护士"的)
Private mrs停嘱原因 As ADODB.Recordset
Private mlngPrevRow As Long
'控制变量
Private Enum EFun
    E申请 = 0
    E审核 = 1
End Enum
Private Enum ESTATE
    E全部 = 0
    E未审核 = 1
    E审核通过 = 2
    E审核未通过 = 3
End Enum
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
    部分冲销明细 As Boolean
    冲销已结帐单据 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private ColApply As Collection
Private ColApplied As Collection
Private ColAudit As Collection
Private ColAudited As Collection

Private mbonNotEnter As Boolean
Private mlngPreFeeItemID As Long '排序时记录当前行
Private mstrUnitIDs As String   '操作员有权限的病区或部门ID集
Private mblnUnChange As Boolean
Private mblnNotClick As Boolean

'数据变量
Private mrsInfo As ADODB.Recordset
Private mrsApply As ADODB.Recordset     '申请明细
Private mrsApplied As ADODB.Recordset   '已申请明细
Private mrsAudit As ADODB.Recordset     '待审核明细
Private mrsAudited As ADODB.Recordset   '已审核明细
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mlngOldY As Long
'消息相关对象变量
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private Sub cbo次数_Click()
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If cbo次数.ListIndex < 0 Then Exit Sub
    If cbo次数.ItemData(cbo次数.ListIndex) = 0 Then Exit Sub
    If zlIsAllowFeeChange(Nvl(Val(mrsInfo!病人ID)), cbo次数.ItemData(cbo次数.ListIndex)) = False Then Exit Sub
End Sub

Private Sub chkDateAudit_Click()
    dtpAuditB.Enabled = chkDateAudit.Value = 0
    dtpAuditE.Enabled = dtpAuditB.Enabled
End Sub

Private Sub chkOtherOperator_Click()
    Call cboState_Click
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk项目_Click(Index As Integer)
    Dim i As Integer
    i = IIf(Index = 0, 1, 0)
    If chk项目(Index).Value = 0 Then    '至少选一种
        If chk项目(i).Value = 0 Then chk项目(i).Value = 1
    End If
End Sub

Private Sub cbo次数_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAudit_Click()
    Dim frmTmp As New frmReCharge
    
    With frmTmp
        .mlngDeptID = mlngDeptID
        .mbytUseType = 0
        .mbytFun = 1
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
End Sub

Private Sub chkShowOthers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub chk项目_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpApplyB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpApplyE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdSeleItem_Click()
    If zlSelectItem("") = False Then Exit Sub
End Sub

Private Sub Form_Activate()
    Call InitInput
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF4
        If Shift = vbCtrlMask Then
            Dim intIndex As Integer
            intIndex = IDKIND.GetKindIndex("IC卡号")
            If intIndex < 0 Then Exit Sub
            IDKIND.IDKIND = intIndex: Call IDKind_Click(IDKIND.GetCurCard)
        ElseIf Me.ActiveControl Is txtPatient Then
            If IDKIND.Enabled Then
                If Shift = vbShiftMask Then
                    IDKIND.IDKIND = IIf(IDKIND.IDKIND = 0, UBound(Split(IDKIND.IDKindStr, ";")), IDKIND.IDKIND - 1)
                Else
                    IDKIND.IDKIND = IIf(IDKIND.IDKIND = UBound(Split(IDKIND.IDKindStr, ";")), 0, IDKIND.IDKIND + 1)
                End If
            End If
        End If
    Case vbKeyF5
        If cmdRefresh.Visible Then Call cmdRefresh_Click
        If cmdAllDetail.Visible And cmdAllDetail.Enabled Then Call cmdAllDetail_Click
    Case vbKeyF6  '定位到病人输入框
        txtPatient.SetFocus
        Call zlControl.TxtSelAll(txtPatient)
    Case vbKeyF7    '切换输入法
        If gbln简码切换 Then
            If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
                If stbThis.Panels("WB").Bevel = sbrRaised Then
                    Call stbThis_PanelClick(stbThis.Panels("WB"))
                Else
                    Call stbThis_PanelClick(stbThis.Panels("PY"))
                End If
            End If
        End If
    End Select
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
        Exit Sub
    End If
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID = 0 Then Exit Sub
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    txtPatient.Text = strID
    Dim objCard  As Card
    Set objCard = IDKIND.GetIDKindCard("身份证")
    If objCard Is Nothing Then Exit Sub
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub chkDate_KeyPress(KeyAscii As Integer)
    SendKeys "{Tab}"
End Sub


Private Sub cmdAllDetail_Click()
    If mrsApply.State = 1 Then
        If mrsApply.RecordCount > 0 Then
            mrsApply.Filter = "销帐数量<>0"
            If mrsApply.RecordCount > 0 Then
                If MsgBox("重新读取记录后,当前已输入的信息将丢失,你确认要继续吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    If dtpApplyB.Value > dtpApplyE.Value Then
        MsgBox "开始时间不能大于结束时间.", vbInformation, gstrSysName
        If dtpApplyB.Visible And dtpApplyB.Enabled Then dtpApplyB.SetFocus
        Exit Sub
    End If
    Call LoadMainData(0)
    vsfMain(0).SetFocus
    Call ShowSumMoney
End Sub

Private Sub cmdRefresh_Click()
    If dtpAuditB.Value > dtpAuditE.Value Then
        MsgBox "开始时间不能大于结束时间.", vbInformation, gstrSysName
        If dtpAuditB.Visible And dtpAuditB.Enabled Then dtpAuditB.SetFocus
        Exit Sub
    End If
    If mbytFun = E申请 Then
        Call cboState_Click
    Else
        Call LoadMainData(0)
    End If
End Sub
Private Sub cboState_Click()
    Dim strFirstCol As String, lngWidth As Long
    Dim intState As Integer
    
    If Not Visible Or cboState.ListIndex = -1 Then Exit Sub
    
    intState = Val(cboState.ItemData(cboState.ListIndex))
    
    cmdCancelApply.Visible = intState = E未审核
        
    Call LoadMainData(0)
    
    strFirstCol = "状态"
    chkOtherOperator.Visible = False
    Select Case intState
    Case ESTATE.E全部
        lngWidth = 550
    Case ESTATE.E未审核
        strFirstCol = "选择"
        lngWidth = 550
        chkOtherOperator.Visible = InStr(1, mstrPrivsOpt, ";取消他人申请;") > 0
    Case ESTATE.E审核通过
        lngWidth = 0
    Case ESTATE.E审核未通过
        lngWidth = 0
    End Select
    vsfMain(1).TextMatrix(0, ColApplied("选择")) = strFirstCol
    vsfMain(1).ColWidth(ColApplied("选择")) = lngWidth
    Call ShowSumMoney
End Sub

Private Sub chkDate_Click()
    dtpApplyB.Enabled = chkDate.Value = 0
    dtpApplyE.Enabled = chkDate.Value = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub AdjustFace()
    On Error Resume Next
    
    fraTop(1).Top = fraTop(0).Top
    fraTop(1).Left = fraTop(0).Left
    
    If mbytFun = E申请 Then
        Me.Caption = "病人费用销帐申请"
        If tbsType.SelectedItem.Key = "T1" Then
            fraTop(0).Visible = True
            fraTop(1).Visible = False
            tbsType.Top = fraTop(0).Top + fraTop(0).Height + 100
            Set fraPatiInfor.Container = fraTop(0)
            fraPatiInfor.Width = Val(fraPatiInfor.Tag)
            txtPatient.Width = Val(txtPatient.Tag)
           ' fraPatiInfor.Top = dtpApplyE.Top - 10
        Else
            fraTop(0).Visible = False
            fraTop(1).Visible = True
            tbsType.Top = fraTop(1).Top + fraTop(1).Height + 100
            Set fraPatiInfor.Container = fraTop(1)
            fraPatiInfor.Width = Val(fraPatiInfor.Tag) + 520
            txtPatient.Width = Val(txtPatient.Tag) + 520
            chkDateAudit.Visible = True
        End If
        fraCmd(0).Left = tbsType.Left + tbsType.Width + 50
    Else
        Me.Caption = "病人费用销帐审核"
        fraTop(0).Visible = False
        fraTop(1).Visible = True
        tbsType.Top = fraTop(1).Top + fraTop(1).Height + 100
        Set fraPatiInfor.Container = fraTop(1)
        fraPatiInfor.Width = Val(fraPatiInfor.Tag) + 520
        txtPatient.Width = Val(txtPatient.Tag) + 520
    End If
    
    fraCmd(0).Top = tbsType.Top + (fraCmd(0).Height - tbsType.Height) \ 2
    fraCmd(0).Left = tbsType.Left + tbsType.Width + 100
    fraCmd(1).Top = fraCmd(0).Top: fraCmd(1).Left = fraCmd(0).Left
    fraCmd(2).Top = fraCmd(0).Top: fraCmd(2).Left = fraCmd(0).Left
            
    vsfMain(0).Top = tbsType.Top + tbsType.Height + 100
    If picHsc.Top - vsfMain(0).Top - 20 < 500 Then
        vsfMain(0).Height = 500
        picHsc.Top = vsfMain(0).Top + vsfMain(0).Height + 10
        vsfDetail.Top = picHsc.Top + picHsc.Height + 10
        vsfDetail.Height = stbThis.Top - vsfDetail.Top - 20
    Else
        vsfMain(0).Height = picHsc.Top - vsfMain(0).Top - 20
    End If
    vsfMain(1).Top = vsfMain(0).Top
    vsfMain(1).Height = vsfMain(0).Height
    vsfMain(1).Left = vsfMain(0).Left
    vsfMain(1).Width = vsfMain(0).Width
End Sub

Private Sub InitFace()
    Dim i As Integer
    
    Call AdjustFace
        
    tbsType.Tabs("T1").Selected = True
    
    Call InitMainHead(True)
    Call InitDetailHead(True)
    If mbytFun = E申请 Then
        chkDateAudit.Visible = False
        txtPatient.ToolTipText = "定位快捷键F6"
        tbsType.Tabs("T1").Caption = "销帐申请"
        tbsType.Tabs("T2").Caption = "已申请明细"
        
        Set ColApply = New Collection
        Set ColApplied = New Collection
        For i = 0 To vsfMain(0).Cols - 1
            ColApply.Add i, vsfMain(0).TextMatrix(0, i)
        Next
        For i = 0 To vsfMain(1).Cols - 1
            ColApplied.Add i, vsfMain(1).TextMatrix(0, i)
        Next
        
        chkVerfy.Visible = InStr(1, mstrPrivsOpt, ";销帐审核;") > 0  '34994
        chkVerfy.Value = IIf(zlDatabase.GetPara("销帐申请同时审核", glngSys, Enum_Inside_Program.p记帐操作, "0", Array(chkVerfy), InStr(1, mstrPrivsOpt, ";记帐选项设置;") > 0) = "1", 1, 0)
        chkShowOthers.Value = IIf(zlDatabase.GetPara("显示他科执行费用", glngSys, Enum_Inside_Program.p记帐操作, "1", Array(chkShowOthers), InStr(1, mstrPrivsOpt, ";记帐选项设置;") > 0) = "1", 1, 0)
    Else
        chkVerfy.Visible = False  '34994
        chkDateAudit.Value = 1
        tbsType.Tabs("T1").Caption = "销帐审核"
        tbsType.Tabs("T2").Caption = "已审核明细"
        
        Set ColAudit = New Collection
        Set ColAudited = New Collection
        For i = 0 To vsfMain(0).Cols - 1
            ColAudit.Add i, vsfMain(0).TextMatrix(0, i)
        Next
        For i = 0 To vsfMain(1).Cols - 1
            ColAudited.Add i, vsfMain(1).TextMatrix(0, i)
        Next
    End If
    Call InitInput
End Sub

Private Sub InitInput()
    '初始化输入法
    stbThis.Panels("PY").Visible = gbln简码切换
    stbThis.Panels("WB").Visible = gbln简码切换
    If gbytCode = 0 Then
        stbThis.Panels("PY").Bevel = sbrInset
        stbThis.Panels("WB").Bevel = sbrRaised
    ElseIf gbytCode = 1 Then
        stbThis.Panels("PY").Bevel = sbrRaised
        stbThis.Panels("WB").Bevel = sbrInset
    Else
        stbThis.Panels("PY").Bevel = sbrInset
        stbThis.Panels("WB").Bevel = sbrInset
    End If
End Sub

Private Sub InitData()
    Dim DatSys As Date
    Dim rsOperator As ADODB.Recordset
    Dim i As Long, strTmp As String, arrTmp As Variant
    
    Set mrsInfo = New ADODB.Recordset
            
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    '60679
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    strSQL = "" & _
    "   Select 1 From 人员表 a,人员性质说明 b" & _
    "   Where a.ID = b.人员ID And b.人员性质='护士'  and A.id=[1] " & _
    "           And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
    "           And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    mblnOperatorNurse = rsTemp.RecordCount <> 0
    rsTemp.Close
    Set rsTemp = Nothing
    Set rsOperator = GetOperatorDept
    With rsOperator
        If .RecordCount <> 0 Then .MoveFirst
        mstrOperatorDeptIDs = ""
        Do While Not .EOF
            mstrOperatorDeptIDs = mstrOperatorDeptIDs & "," & Nvl(!ID)
            .MoveNext
        Loop
        mstrOperatorDeptIDs = mstrOperatorDeptIDs & ","
    End With
    
    cboKind.Clear: cboKind.AddItem "0", "所有收费类别", True, True, True
    strSQL = "Select 编码,类别 From 收费类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTemp.EOF
        cboKind.AddItem "" & rsTemp!编码, "" & rsTemp!类别, False, True, True
        rsTemp.MoveNext
    Loop
    
    strSQL = "Select 名称 From 停嘱原因"
    Set mrs停嘱原因 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    mstrUnitIDs = GetUserUnits
    DatSys = zlDatabase.Currentdate
    If mbytFun = E申请 Then
        Set mrsApply = New ADODB.Recordset
        Set mrsApplied = New ADODB.Recordset
            
        dtpApplyB.Value = DateAdd("D", -5, DatSys)
        dtpApplyE.Value = DatSys
    
        strTmp = "0-全部,1-未审核,2-审核通过,3-审核未通过"
        arrTmp = Split(strTmp, ",")
        cboState.Clear
        For i = 0 To UBound(arrTmp)
            cboState.AddItem arrTmp(i)
            cboState.ItemData(cboState.NewIndex) = i
        Next
        cboState.ListIndex = 0
        cmdCancelApply.Visible = False
        
        If InStr(mstrPrivsOpt, "销帐审核") > 0 And mbytUseType <> 2 Then
            cmdAudit.Visible = True
            cmdAudit.Top = cmdHelp.Top
            cmdHelp.Top = cmdHelp.Top + cmdHelp.Height + 100
        End If
    Else
        Set mrsAudit = New ADODB.Recordset
        Set mrsAudited = New ADODB.Recordset
    End If
    
    dtpAuditB.Value = DateAdd("D", -5, DatSys)
    dtpAuditE.Value = CDate(Format(DatSys, "yyyy-MM-dd 23:59:59"))
    
    
    On Error Resume Next
    If mbytFun = E申请 Then
        dtpApplyB.Value = CDate(zlDatabase.GetPara("费用开始时间", glngSys, mlngModul, Format(dtpApplyB.Value, "YYYY-MM-DD HH:MM:SS"), Array(dtpApplyB, dtpApplyE), zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")))
        chkDate.Value = IIf(zlDatabase.GetPara("忽略期间", glngSys, mlngModul, "0", Array(chkDate), zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")) = "0", 0, 1)
        i = Val(zlDatabase.GetPara("项目显示方式", glngSys, mlngModul, "0", Array(chk项目(0), chk项目(1)), InStr(mstrPrivsOpt, "记帐选项设置")))
        Select Case i
        Case 1
            chk项目(0).Value = 1: chk项目(1).Value = 0
        Case 2
            chk项目(0).Value = 0: chk项目(1).Value = 1
        Case Else
            chk项目(0).Value = 1: chk项目(1).Value = 1
        End Select
        
        fraCmd(0).Enabled = False
        txtFeeItem.Enabled = False
        cmdAllDetail.Enabled = False
        cmdOKApply.Enabled = False
        '59051
        chkDateAudit.Value = IIf(zlDatabase.GetPara("申请明细忽略期间", glngSys, Enum_Inside_Program.p记帐操作, "0", Array(chkDateAudit), zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")) = "0", 0, 1)
    Else
        dtpAuditB.Value = zlDatabase.GetPara("审核开始时间", glngSys, mlngModul, Format(dtpAuditB.Value, "YYYY-MM-DD HH:MM:SS"), Array(dtpAuditB, dtpApplyE), zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置"))
        cmdOKAudit.Enabled = False
    End If
    
    If mlngPatientID <> 0 Then      ' And mbytFun = E申请
        txtPatient.Text = "-" & mlngPatientID
        Call txtPatient_KeyPress(13)
    End If
End Sub

Private Sub InitMainHead(Optional blnSetWidth As Boolean, Optional bytScope As Byte)
'参数:
'   bytScope=0-初始化两张表,1-初始化第一张表,2-初始化第二张表
    Dim i As Long, ArrTmp0 As Variant, ArrTmp1 As Variant, arrTmp As Variant
    
    If mbytFun = E申请 Then
        ArrTmp0 = Split(HeadApply, "|")
        ArrTmp1 = Split(HeadApplied, "|")
    Else
        ArrTmp0 = Split(HeadAudit, "|")
        ArrTmp1 = Split(HeadAudited, "|")
    End If
    If bytScope = 0 Or bytScope = 1 Then
        With vsfMain(0)
            .Redraw = flexRDNone
            .Clear
            .RowHeightMin = 320: .Rows = 2
            .Cols = UBound(ArrTmp0) + 1
            For i = 0 To .Cols - 1
                arrTmp = Split(ArrTmp0(i), ",")
                .TextMatrix(0, i) = arrTmp(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
                If blnSetWidth Then
                    .FixedAlignment(i) = flexAlignCenterCenter
                    .ColAlignment(i) = arrTmp(1)
                    .ColWidth(i) = arrTmp(2)
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
    
    If bytScope = 0 Or bytScope = 2 Then
        With vsfMain(1)
            .Redraw = flexRDNone
            .Clear
            .RowHeightMin = 320
            .Rows = 2
            .Cols = UBound(ArrTmp1) + 1
            For i = 0 To .Cols - 1
                arrTmp = Split(ArrTmp1(i), ",")
                .TextMatrix(0, i) = arrTmp(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
                If blnSetWidth Then
                    .FixedAlignment(i) = flexAlignCenterCenter
                    .ColAlignment(i) = arrTmp(1)
                    .ColWidth(i) = arrTmp(2)
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
End Sub

Private Sub InitDetailHead(Optional blnSetWidth As Boolean)
    Dim ArrTmpDetail As Variant, arrTmp As Variant
    Dim i As Long
    Dim strHead As String
    
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            strHead = HeadApplyDetail
        Else
            strHead = HeadAppliedDetail
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            strHead = HeadAuditDetail
        Else
            strHead = HeadAuditedDetail
        End If
    End If
    
    vsfDetail.Clear
    vsfDetail.Rows = 2
    vsfDetail.RowHeightMin = 320
    ArrTmpDetail = Split(strHead, "|")
    vsfDetail.Cols = UBound(ArrTmpDetail) + 1
     
    With vsfDetail
        For i = 0 To .Cols - 1
            arrTmp = Split(ArrTmpDetail(i), ",")
            .TextMatrix(0, i) = arrTmp(0)
            .ColKey(i) = .TextMatrix(0, i)
            
            If blnSetWidth Then
                .FixedAlignment(i) = flexAlignCenterCenter
                .ColAlignment(i) = arrTmp(1)
                .ColWidth(i) = arrTmp(2)
            End If
        Next
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    gblnOK = False
    '55368
    Call LoadBabyCombox
    mblnOperatorICU = zlisCheckOperatorICU
    Call initCardSquareData
    
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作) & ";" & mstrPrivs
    Call RestoreWinState(Me, App.ProductName)
    Call InitFace
     '问题:39373
     '55368
    Call RestoreFlexState(vsfMain(0), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call RestoreFlexState(vsfMain(1), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Me.WindowState = vbMaximized
     '问题:47798
    Call GetRegisterItem(g私有模块, Me.Name, "idkind", strTmp)
    Err = 0: On Error Resume Next
    IDKIND.IDKIND = Val(strTmp)
    Err = 0: On Error GoTo 0
    
    Call InitData
    Call zlMsgModule_Init
    If mstrInNO <> "" Or mlngAdviceID <> 0 Then
        mblnInit = True
        Call LoadMainData(0, mstrInNO, mlngAdviceID)
        mblnInit = False
        mstrInNO = ""
        mlngAdviceID = 0
    End If
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long
    
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
        
    vsfMain(0).Left = Me.ScaleLeft + 20
    vsfMain(0).Width = Me.ScaleLeft + Me.ScaleWidth - vsfMain(0).Left - 20
    vsfMain(1).Left = vsfMain(0).Left
    vsfMain(1).Width = vsfMain(0).Width
    vsfDetail.Left = vsfMain(0).Left
    vsfDetail.Width = vsfMain(0).Width - IIf(vsfTogether.Visible, vsfTogether.Width + 50, 0)
    picHsc.Width = vsfMain(0).Width
    
    If vsfMain(0).Visible Then
        lngTmp = Me.ScaleTop + Me.ScaleHeight - (picHsc.Height + vsfDetail.Height + stbThis.Height + 30) - vsfMain(0).Top
        If lngTmp > 500 Then
            vsfMain(0).Height = lngTmp
            picHsc.Top = vsfMain(0).Top + vsfMain(0).Height + 10
            vsfDetail.Top = picHsc.Top + picHsc.Height + 10
        End If
    ElseIf vsfMain(1).Visible Then
        lngTmp = Me.ScaleTop + Me.ScaleHeight - (picHsc.Height + vsfDetail.Height + stbThis.Height + 30) - vsfMain(1).Top
        If lngTmp > 500 Then
            vsfMain(1).Height = lngTmp
            picHsc.Top = vsfMain(0).Top + vsfMain(1).Height + 10
            vsfDetail.Top = picHsc.Top + picHsc.Height + 10
        End If
    End If
    
    If mbytFun = EFun.E申请 Then
        If vsfTogether.Visible Then
            vsfTogether.Top = vsfDetail.Top
            vsfTogether.Height = vsfDetail.Height
            vsfTogether.Left = vsfDetail.Left + vsfDetail.Width + 50
        End If
        chkVerfy.Top = fraCmd(0).Top + 15
        chkVerfy.Width = IIf(Me.ScaleWidth - chkVerfy.Left > 2555, 2505, Me.ScaleWidth - chkVerfy.Left)
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveFlexState(vsfMain(0), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call SaveFlexState(vsfMain(1), App.ProductName & "\" & Me.Name & "-" & mbytFun)
    Call SaveFlexState(vsfDetail, App.ProductName & "\" & Me.Name & "-" & mbytFun)
    '55368
    Call zlDatabase.SetPara("销帐申请婴儿费显示规则", cboBaby.ItemData(cboBaby.ListIndex), glngSys, Enum_Inside_Program.p记帐操作, InStr(mstrPrivsOpt, ";记帐选项设置;") > 0)
    If mbytFun = E申请 Then
        zlDatabase.SetPara "费用开始时间", Format(dtpApplyB.Value, "YYYY-MM-DD HH:MM:SS"), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")
        zlDatabase.SetPara "忽略期间", chkDate.Value, glngSys, mlngModul
        zlDatabase.SetPara "项目显示方式", IIf(chk项目(0).Value = 1 And chk项目(1).Value = 0, 1, IIf(chk项目(0).Value = 0 And chk项目(1).Value = 1, 2, 0)), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")
        zlDatabase.SetPara "销帐申请同时审核", IIf(chkVerfy.Value = 1, 1, 0), glngSys, Enum_Inside_Program.p记帐操作, zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")
        zlDatabase.SetPara "申请明细忽略期间", chkDateAudit.Value, glngSys, Enum_Inside_Program.p记帐操作, zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")
        zlDatabase.SetPara "显示他科执行费用", IIf(chkShowOthers.Value = 1, 1, 0), glngSys, Enum_Inside_Program.p记帐操作, zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")
    Else
        zlDatabase.SetPara "审核开始时间", Format(dtpAuditB.Value, "YYYY-MM-DD HH:MM:SS"), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivsOpt, "记帐选项设置")
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mrsApplyDept = Nothing
    Set mrs停嘱原因 = Nothing
    Set ColApply = Nothing
    Set ColApplied = Nothing
    Set ColAudit = Nothing
    Set ColAudited = Nothing
    Set mrsInfo = Nothing
    Set mrsApply = Nothing
    Set mrsApplied = Nothing
    Set mrsAudit = Nothing
    Set mrsAudited = Nothing

    mlngDeptID = 0
    mlngPatientID = 0
     '问题:47798
    Call SaveRegisterItem(g私有模块, Me.Name, "idkind", IDKIND.IDKIND)
    Set mrsOperatorDept = Nothing
    mblnOperatorNurse = False
    mstrOperatorDeptIDs = ""
    
    '消息拆卸
    zlMsgModule_Unload
End Sub

Private Sub picHsc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngOldY = Y
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsfMain(0).Height + Y - mlngOldY <= 500 Or vsfDetail.Height - Y + mlngOldY <= 500 Then Exit Sub
        
        picHsc.Top = picHsc.Top + Y - mlngOldY
        If vsfMain(0).Visible Then
            vsfMain(0).Height = picHsc.Top - vsfMain(0).Top ' vsfMain(0).vsfMain(0).Height + Y
            vsfMain(1).Height = vsfMain(0).Height
        Else
            vsfMain(1).Height = picHsc.Top - vsfMain(1).Top ' vsfMain(1).Height + Y
            vsfMain(0).Height = vsfMain(1).Height
        End If
        
        vsfDetail.Top = picHsc.Top + picHsc.Height ' vsfDetail.Top + Y
        vsfDetail.Height = IIf(ScaleHeight - vsfDetail.Top - stbThis.Height < 0, 0, ScaleHeight - vsfDetail.Top - stbThis.Height) ' vsfDetail.Height - Y
        
        If cboRemark.Visible Then
            cboRemark.Top = vsfDetail.Top + vsfDetail.RowPos(vsfDetail.Row)
            cboRemark.Left = vsfDetail.Left + vsfDetail.ColPos(vsfDetail.ColIndex("销帐原因"))
        End If
        
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Key
    Case "PY", "WB"
        If Panel.Bevel = sbrRaised And gbln简码切换 Then
            '切换并保存简码匹配方式
            Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            If Panel.Key = "PY" Then
                stbThis.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Else
                stbThis.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            End If
            zlDatabase.SetPara "简码方式", IIf(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIf(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
            gbytCode = Val(zlDatabase.GetPara("简码方式", , , 0))
        End If
    End Select
End Sub

Private Sub txtFeeItem_Change()
    txtFeeItem.Tag = ""
End Sub

Private Sub txtFeeItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtFeeItem.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txtFeeItem.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zlSelectItem(Trim(txtFeeItem.Text)) = False Then Exit Sub
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKIND.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    Call IDKIND.SetAutoReadCard(False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If Trim(txtPatient.Text) = "" Then
         If mbytFun = E申请 Then
            If tbsType.SelectedItem.Key <> "T1" Then
                Call ClearPatientInfo
            End If
        Else
            Call ClearPatientInfo
        End If
    End If
    
    If mrsInfo.State = 0 And Trim(txtPatient.Text) <> "" Then txtPatient.Text = ""
    If mrsInfo.State = 1 Then
        If txtPatient.Text <> mrsInfo!姓名 Then txtPatient.Text = mrsInfo!姓名
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            .mlngUnitID = mlngDeptID
            .mbytUseType = 4
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
        End With
    Else
        If IDKIND.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.GetCurCard.名称 = "门诊号" Or IDKIND.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKIND.GetCurCard, blnCard, txtPatient.Text)
      End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOutMsg As Boolean
    Call ClearPatientInfo
    If Not GetPatient(objCard, strInput, blnCard, blnOutMsg) Then
        If Not blnOutMsg Then stbThis.Panels(2).Text = "没有找到该病人,请检查输入内容!"
        Call zlControl.TxtSelAll(txtPatient)
        Exit Sub
    End If
    If Not IsNull(mrsInfo!险类) Then
        MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(mrsInfo!险类))
        MCPAR.部分冲销明细 = gclsInsure.GetCapability(support允许部分冲销明细, , Val(mrsInfo!险类))
        MCPAR.冲销已结帐单据 = gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , Val(mrsInfo!险类))
        If MCPAR.记帐作废上传 Then
            If Not gclsInsure.GetCapability(support允许部份冲销单据, , Val(mrsInfo!险类)) Then  '不能部分销帐
                MsgBox "当前医保不允许部分冲销单据，不支持采用申请审核模式销帐！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    stbThis.Panels(2).Text = ""
    Call LoadPatientInfo
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub ClearPatientInfo()
    Set mrsInfo = New ADODB.Recordset
    
    txtPatient.ForeColor = Me.ForeColor
    lblPatiInfo.Caption = "性别：     年龄：        住院号：             床号：         科室：       病区：        付款方式： "
    
    fraCmd(0).Enabled = False
    txtFeeItem.Enabled = False
    cmdAllDetail.Enabled = False
    cmdOKApply.Enabled = False
    
    If vsfMain(0).Rows >= 2 Then
        If Val(vsfMain(0).RowData(1)) <> 0 Then
            Call InitMainHead(False, 1)
            Call InitDetailHead(False)
            Set mrsApply = New ADODB.Recordset
        End If
    End If
End Sub

Private Sub LoadPatientInfo()
    txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
    mblnNotClick = True
    txtPatient.Text = mrsInfo!姓名
    lblPatiInfo.Caption = "性别：" & mrsInfo!性别 & "   年龄：" & mrsInfo!年龄 & "   住院号：" & mrsInfo!住院号 & _
                          "   床号：" & mrsInfo!床号 & "   科室：" & mrsInfo!科室 & "   病区：" & mrsInfo!病区 & "   付款方式：" & mrsInfo!医疗付款方式
    fraCmd(0).Enabled = True
    fraCmd(0).Enabled = True
    txtFeeItem.Enabled = True
    cmdAllDetail.Enabled = True
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    mblnNotClick = False
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参:blnOutMsg-true返回已经提示,否则未提示
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 16:53:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strIF As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    '是否具有强制记帐权限
    If InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 And InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)=0)"
    Else
        strIF = " And B.出院日期 is NULL And Nvl(B.状态,0)<>3"
    End If
    
    strSQL = _
    "   Select A.病人ID,B.主页ID,B.出院科室ID," & _
    "          Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, A.年龄, A.住院号, B.出院病床 床号, " & _
    "          C.名称 科室, D.名称 病区, A.医疗付款方式, B.险类,B.病人类型,a.门诊号,A.身份证号" & vbNewLine & _
    "   From 病人信息 A, 病案主页 B, 部门表 C, 部门表 D,病人余额 X" & vbNewLine & _
    "   Where A.病人id = B.病人id And A.主页ID = B.主页ID And B.出院科室ID = C.ID And B.当前病区id = D.ID" & _
    "       And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2  And A.停用时间 Is Null" & strIF
        
        '问题:38332:取消站点限制,因为可能存在对转出病人的处理
'        " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _

    If blnCard = True And objCard.名称 Like "姓名*" Then   '刷卡
    
        If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKIND.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strInput = Mid(strInput, 2)
        If strInput = "" Then strInput = "0"
        strSQL = strSQL & " And A.门诊号=[2]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strInput = Mid(strInput, 2)
        If strInput = "" Then strInput = "0"
        strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
    ElseIf Left(strInput, 1) = "/" And mbytUseType <> 1 And mlngDeptID <> 0 Then   '床位号,医技科室调用时不使用床号,病区调用进入时选所有病区时不使用床号
        '41654 And IsNumeric(Mid(strInput, 2))
        strSQL = strSQL & " And B.当前病区ID=[3] And B.出院病床=[1]"
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                strPati = "" & _
                "   Select A.病人ID as ID,A.病人ID,A.住院号, A.门诊号, Nvl(b.性别, a.性别) As 性别, A.年龄, A.住院次数, A.家庭地址, A.工作单位," & vbNewLine & _
                "       To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,  To_Char(B.入院日期,'YYYY-MM-DD') as 入院日期, To_Char(B.出院日期,'YYYY-MM-DD') as 出院日期" & vbNewLine & _
                "   From 病人信息 A, 病案主页 B,病人余额 X" & vbNewLine & _
                "   Where A.病人id = B.病人id(+) And A.主页ID = B.主页id(+) And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2 And A.停用时间 Is Null And A.姓名 = [1]" & strIF & vbNewLine & _
                "   Order By Decode(住院号, Null, 1, 0), 入院日期 Desc"
                        
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!病人ID)
                    strSQL = strSQL & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset:  Exit Function
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlngDeptID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: Exit Function
    If zlPatiIS病案已编目(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = True Then    '问题:28725
        Set mrsInfo = New ADODB.Recordset
        blnOutMsg = True
        Exit Function
    End If
    mblnPatiDeptICU = zlisCheckDeptICU(Val(Nvl(mrsInfo!出院科室ID)))
    Call Load住院次数
    
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub Load住院次数()
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    
    With cbo次数
        .Clear
        .AddItem "所有住院"
        .ListIndex = 0
    End With
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    gstrSQL = "select 主页ID From 病案主页 where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Nvl(mrsInfo!病人ID)))
    With cbo次数
        Do While Not rsTemp.EOF
            .AddItem "第" & Val(Nvl(rsTemp!主页ID)) & "次住院"
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!主页ID))
            If Val(Nvl(mrsInfo!主页ID)) = Val(Nvl(rsTemp!主页ID)) Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetWindowsTittle()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体标题
    '编制:刘兴洪
    '日期:2009-10-26 15:21:22
    '问题:25850
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case mbytFun
    Case E申请
        
        'mbytFun As Byte      '0-申请,1-审核
        'mlngDeptID As Long   '病区调用时传入当前操作的病人病区ID,医技科室调用时传入医技科室ID
        If mlngDeptID = 0 Then
            fraTop(0).ForeColor = vbRed
            If tbsType.SelectedItem.Key = "T1" Then
               fraTop(0).Caption = "申请部门：" & "申请部门未选择!"
            Else
                fraTop(0).Caption = ""
            End If
        Else
            fraTop(0).ForeColor = vbRed
            If tbsType.SelectedItem.Key = "T1" Then
                fraTop(0).Caption = "申请部门：" & "申请部门未选择!"
                If mrsApplyDept Is Nothing Then
                    GoTo GetApplyDept:
                ElseIf mrsApplyDept.State <> 1 Then
                    GoTo GetApplyDept:
                Else
                    fraTop(0).Caption = "申请部门：" & "申请部门未选择!"
                    If mrsApplyDept.EOF = False Then
                        fraTop(0).Caption = "申请部门：" & Nvl(mrsApplyDept!名称)
                        fraTop(0).ForeColor = &H80000012
                    End If
                End If
            Else
                fraTop(0).Caption = ""
            End If
        End If
    Case Else
        fraTop(0).Caption = ""
        fraTop(0).ForeColor = &H80000012
    End Select
    Exit Sub
GetApplyDept:

    On Error GoTo errHandle
    gstrSQL = "Select 名称  From 部门表 where id=[1]"
    Set mrsApplyDept = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptID)
    fraTop(0).Caption = "申请部门：" & "申请部门未选择!"
    If mrsApplyDept.EOF = False Then
        fraTop(0).Caption = "申请部门：" & Nvl(mrsApplyDept!名称)
        fraTop(0).ForeColor = &H80000012
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub tbsType_Click()
    Dim lngFeeItemID As Long
    
    Me.AutoRedraw = False
    Call AdjustFace
    
    Call SetWindowsTittle
    
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            vsfMain(0).Visible = True
            vsfMain(1).Visible = False
            fraCmd(0).Visible = True
            fraCmd(1).Visible = False
            chkVerfy.Visible = InStr(1, mstrPrivsOpt, ";销帐审核;") > 0  '34994
            If Visible Then
                vsfMain(0).SetFocus
                lngFeeItemID = vsfMain(0).RowData(vsfMain(0).Row)
                Call ShowDetail(lngFeeItemID)
            End If
        Else
            chkVerfy.Visible = False '34994
            vsfMain(0).Visible = False
            vsfMain(1).Visible = True
            fraCmd(0).Visible = False
            fraCmd(1).Visible = True
            If Visible Then
                vsfMain(1).SetFocus
                Call cmdRefresh_Click
            End If
        End If
        Call Form_Resize
        Call ShowSumMoney
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            lblAuditDate.Caption = "申请期间"
            vsfMain(0).Visible = True
            vsfMain(1).Visible = False
            fraCmd(2).Visible = True
            cmdOKAudit.Caption = "确认审核(&S)"
            cmdCancelRefuse.Visible = False
            chkDateAudit.Visible = True
            Call chkDateAudit_Click
            
            If Visible Then
                vsfMain(0).SetFocus
                lngFeeItemID = vsfMain(0).RowData(vsfMain(0).Row)
                cmdOKAudit.Enabled = lngFeeItemID > 0
                Call ShowDetail(lngFeeItemID)
            End If
        Else
            lblAuditDate.Caption = "审核期间"
            vsfMain(0).Visible = False
            vsfMain(1).Visible = True
            fraCmd(2).Visible = True
            cmdOKAudit.Caption = "重审拒绝(&S)"
            cmdCancelRefuse.Visible = True
            
            chkDateAudit.Visible = False
            dtpAuditB.Enabled = True
            dtpAuditE.Enabled = dtpAuditB.Enabled
            
            
            If Visible Then
                vsfMain(1).SetFocus
                Call cmdRefresh_Click
            End If
        End If
        Call ShowSumMoney
    End If
    Me.AutoRedraw = True
End Sub

Private Sub txtFeeItem_GotFocus()
    zlControl.TxtSelAll txtFeeItem
End Sub

Private Function zlSelectItem(ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的销帐项目
    '入参:strKey-搜索条件
    '出参:
    '返回:选择成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2009-09-21 14:23:25
    '问题:25182
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strIF As String, strSQL As String, DatBegin As Date, DatEnd As Date, strWhere As String
    Dim strSearch As String, vRect As RECT, blnCancel As Boolean
    Dim strDosage As String '配药中心条件
    Dim lng主页ID As Long
    Dim intBaby As Integer, strWhereICU As String
    
    '59220
    On Error GoTo errHandler
    strIF = " And A.病人id = [1] And A.记录状态 > 0"
    '问题:39373
    '55368
    intBaby = cboBaby.ItemData(cboBaby.ListIndex)
    Select Case intBaby
    Case 0  '不含婴儿费
        strIF = strIF & " And nvl(A.婴儿费,0)= 0"
    Case 1  '含婴儿费
    Case Else '显示第几个婴儿费
        strIF = strIF & " And nvl(A.婴儿费,0)= [9]"
    End Select
    '问题:40304
    lng主页ID = 0
    If cbo次数.ListIndex >= 0 Then
         lng主页ID = cbo次数.ItemData(cbo次数.ListIndex)
    End If
    strIF = strIF & IIf(lng主页ID = 0, "", " And nvl(A.主页ID,0)= [8]")
        
    If mlngDeptID <> 0 Then
        If mbytUseType <> 1 Then
            If Not mblnOperatorICU Then
                strWhereICU = " And Instr(','||[6]||',',','||A.病人病区id||',')>0"
                '问题:43940:由于会诊医生也存在开单科室<>病人科室的情况,因此, _
                '       经与周韬讨论,直接以开单科室ID是否为临床性质判断, '
                '       不再用病人科室ID=开单科室ID来判断是否为临床开的单了
                 'exists(select 1 From 部门性质说明 where A.开单部门id=部门ID And 工作性质='临床')
                
                '问题:36462
                strWhereICU = strWhereICU & _
                    " And (Exists(select 1 From 部门性质说明 where A.开单部门id=部门ID And 工作性质='临床') " & _
                    "           And (Instr(',5,6,7,', ',' || A.收费类别 || ',') > 0 Or (A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 1)) " & _
                    "      Or (Instr(',5,6,7,', ',' || A.收费类别 || ',') = 0 Or A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 0))"
            ElseIf Not mblnPatiDeptICU Then
                '以当时病人科室是否为ICU部分:42526
                strWhereICU = _
                    " And (Exists(Select 1 From  部门性质说明 J1  Where A.病人科室ID=J1.部门ID And J1.工作性质='ICU') " & _
                    "      Or (Exists(select 1 From 部门性质说明 where A.开单部门id=部门ID And 工作性质='临床') " & _
                    "          And (Instr(',5,6,7,', ',' || A.收费类别 || ',') > 0 Or (A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 1)) " & _
                    "      Or (Instr(',5,6,7,', ',' || A.收费类别 || ',') = 0 Or A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 0)) )"
            End If
        Else
            strIF = strIF & " And A.开单部门id+0 = [2]"
        End If
    End If
      If chkDate.Value = 0 Then
        If dtpApplyB.Value <= dtpApplyE.Value Then
            DatBegin = dtpApplyB.Value
            DatEnd = dtpApplyE.Value
        Else
            DatBegin = dtpApplyE.Value
            DatEnd = dtpApplyB.Value
        End If
        '59220
        strIF = strIF & " And A.发生时间+0 Between [4] And [5]"
    End If
    '36391:将1替换为RowNum,避免Oracle视图自动合并:42333
    '77686,李南春,2014/9/18,单据类别限制
    strDosage = " And Not Exists (Select  Rownum as 序号" & _
        " From 住院费用记录 J, 药品收发记录 B1, 输液配药内容 C1" & _
        " Where j.NO = a.NO And a.记录性质 = j.记录性质 And  nvl(A.价格父号, A.序号) = Nvl(J.价格父号, J.序号)" & _
        "       And B1.费用id = j.ID And B1.ID = C1.收发id And instr( ',8,9,10,21,24,25,26,',','||B1.单据||',')>0)  "
    
    '问题:29887,55380
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    blnYP = zlStr.IsHavePrivs(mstrPrivsOpt, "药品销帐申请")
    blnZL = zlStr.IsHavePrivs(mstrPrivsOpt, "诊疗销帐申请")
    blnWC = zlStr.IsHavePrivs(mstrPrivsOpt, "卫材销帐申请")
    
  If blnYP And blnWC And blnZL Then
        '全部,不限制
    ElseIf blnYP And blnWC And blnZL = False Then
        strIF = strIF & "  And  A.收费类别 In('4','5','6','7')"
    ElseIf blnYP And blnWC = False And blnZL Then
        strIF = strIF & "  And  A.收费类别 <>'4'"
    ElseIf blnYP And blnWC = False And blnZL = False Then
        strIF = strIF & "  And  A.收费类别 In('5','6','7')"
    ElseIf blnYP = False And blnWC And blnZL = False Then
        strIF = strIF & "  And  A.收费类别 ='4'"
    ElseIf blnYP = False And blnWC And blnZL Then
        strIF = strIF & "  And instr( '5,6,7',  A.收费类别)=0 "
    ElseIf blnYP = False And blnWC = False And blnZL Then
        strIF = strIF & "  And instr( '4,5,6,7',  A.收费类别)=0 "
    Else
        MsgBox "注意:" & vbCrLf & "  你不具备药品、卫材及诊疗销帐申请的权限,请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strWhere = ""
    '问题:30523
    '由于以前对未执行的药品或卫材,只能对病区本身发的药品或卫材进行申请,而由药房发药或卫材的,则不能够申请处理,因此出现程序在此环节上的流程处理漏洞.
    '所以现在取消了该限制,现在的处理方式是如果药房存在未执行的,则申请时,审核部门只能为病区(如果在病区审核前,该药品被药品执行,则禁止审核),执行了的,则为执行部门.
    If chk项目(0).Value = 1 And chk项目(1).Value = 0 Then '只显示已执行的
        strWhere = _
            " And Exists(Select 1 From 住院费用记录 B " & _
            "            Where A.NO = B.NO And A.记录性质 = B.记录性质" & _
            "                  And Nvl(A.价格父号, A.序号) = Nvl(B.价格父号, B.序号)  And B.执行状态 <> 0 )" & vbNewLine
    ElseIf chk项目(0).Value = 0 And chk项目(1).Value = 1 Then '只显示未执行的
        strWhere = _
            " And Exists(Select 1 From 住院费用记录 B" & _
            "            Where A.NO = B.NO And A.记录性质 = B.记录性质" & _
            "                  And Nvl(A.价格父号, A.序号) = Nvl(B.价格父号, B.序号)  And B.执行状态 = 0 )" & vbNewLine
    ElseIf chk项目(0).Value = 0 And chk项目(1).Value = 0 Then '未选择执行项目的,缺省为全选
    Else
    End If
    
    If strKey <> "" Then
        strSearch = IIf(Len(strKey) < 3, "", gstrLike) & strKey & "%"
        If zlCommFun.IsNumOrChar(strKey) Then
            strIF = strIF & vbCrLf & _
                " And Exists(Select 1 From  收费项目目录 Q1,收费项目别名 Q2" & _
                "            Where Q1.ID=Q2.收费细目ID and A.收费细目id=Q1.id" & _
                "                  And (Q1.编码 like upper([7]) or ( Q2.简码 like upper([7]) and Q2.码类 in (3," & gbytCode + 1 & "))))"
        Else
            strIF = strIF & vbCrLf & _
                " And Exists(Select 1 From 收费项目目录 Q1,收费项目别名 Q2" & _
                "            where Q1.ID=Q2.收费细目ID and A.收费细目id=Q1.id And Q2.名称 like upper([7]))"
        End If
    End If

    '未结帐的(结帐并作废当未结账)
    strIF = strIF & _
            " And (A.NO, Nvl(A.价格父号, A.序号)) In (" & vbNewLine & _
            "       Select A.No ,Nvl(A.价格父号, A.序号)" & _
            "       From 住院费用记录 A" & vbNewLine & _
            "       Where Mod(A.记录性质, 10) = 2 " & strIF & vbNewLine & _
            "       Group By A.NO, Mod(A.记录性质, 10), Nvl(A.价格父号, A.序号)" & vbNewLine & _
            "       Having Nvl(Sum(结帐金额),0) = 0)"
    
    strSQL = _
        " Select a.ID, a.NO, a.执行状态, a.价格父号, a.序号, a.发生时间, a.执行部门id, a.开单部门id, a.收费类别," & _
        "        a.收费细目id, a.标准单价, a.付数, a.数次, a.应收金额, a.实收金额, a.结帐ID, a.医嘱序号, a.病人科室id, a.病人病区ID" & _
        " From 住院费用记录 A" & _
        " Where a.记录性质 = 2" & strDosage & strWhere & strIF
    strSQL = strSQL & " Union All " & Replace(strSQL, "住院费用记录", "门诊费用记录")

    '未退数量不等于零的
    '退过药后,因为退的时候只输数次,所以付数不准,都取1
    '如果是药品,没有发药的不允许申请,可能发药后又退药了,所以要用Exists子查询判断,不能直接用执行状态<>0
    strSQL = "Select Max(ID) ID, NO, 发生时间, 序号, 执行部门id,开单部门id, 收费类别, 收费细目id, Avg(单价) 单价," & vbNewLine & _
            "       Decode(Sign(Min(执行状态)), -1, 1, Sum(付数)) 付数," & vbNewLine & _
            "       Decode(Sign(Min(执行状态)), -1, Sum(付数 * 数次), Sum(数次)) 数次, Sum(应收金额) 应收金额, Sum(实收金额) 实收金额, 结帐ID, 医嘱序号" & vbNewLine & _
            "From (Select Max(Decode(Sign(A.执行状态), -1, 0, Decode(A.价格父号, Null, A.ID, 0))) ID, A.执行状态, A.发生时间, A.NO," & vbNewLine & _
            "              Nvl(A.价格父号, A.序号) As 序号, A.执行部门id,A.开单部门id, A.收费类别, A.收费细目id, Avg(A.标准单价) 单价," & vbNewLine & _
            "              Avg(A.付数) 付数, Avg(A.数次) 数次, Sum(A.应收金额) 应收金额, Sum(A.实收金额) 实收金额, A.结帐ID, A.医嘱序号" & vbNewLine & _
            "       From (" & strSQL & ") A, 材料特性 C" & vbNewLine & _
            "       Where A.收费细目id = C.材料id(+)  " & strWhereICU & _
            "       Group By A.NO, A.执行状态, Nvl(A.价格父号, A.序号), A.发生时间, A.执行部门id,A.开单部门id, A.收费类别, A.收费细目id, A.结帐ID, A.医嘱序号)" & vbNewLine & _
            "Group By NO, 发生时间, 序号, 执行部门id,开单部门id, 收费类别, 收费细目id, 结帐ID, 医嘱序号" & vbNewLine & _
            "Having Sum(付数 * 数次) <> 0 "
    
    '可申请销帐的明细
    strSQL = _
        " Select Distinct c.ID,d.名称 As 类别, c.编码,c.名称,c.规格 " & _
        " From (" & strSQL & ") A, 病人费用销帐 B,收费项目目录 C,收费项目类别 D" & vbNewLine & _
        " Where a.收费细目ID=c.ID And c.类别=d.编码 And a.ID = b.费用id(+) And b.状态(+) = 0" & _
        " Order By 类别,名称,规格"
    
    vRect = zlControl.GetControlRect(txtFeeItem.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "销帐费用项目", 1, " ", "请选择", _
        False, False, True, vRect.Left, vRect.Top, txtFeeItem.Height, blnCancel, False, True, _
        Val(mrsInfo!病人ID), mlngDeptID, 0, DatBegin, DatEnd, mstrUnitIDs, strSearch, lng主页ID, intBaby - 1)
    If blnCancel Then
        zlControl.TxtSelAll txtFeeItem
        If txtFeeItem.Enabled And txtFeeItem.Visible Then txtFeeItem.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "未找到项目,可能此病人未发生此费用，请检查!", vbInformation + vbDefaultButton1, gstrSysName
        zlControl.TxtSelAll txtFeeItem
        If txtFeeItem.Enabled And txtFeeItem.Visible Then txtFeeItem.SetFocus
        Exit Function
    End If
        
    '加载相关费用信息数据
    txtFeeItem.Text = Nvl(rsTemp!名称): txtFeeItem.Tag = Nvl(rsTemp!ID)
    Call LoadMainData(rsTemp!ID)
    stbThis.Panels(2).Text = ""
    zlControl.TxtSelAll txtFeeItem
    If txtFeeItem.Enabled And txtFeeItem.Visible Then txtFeeItem.SetFocus
    zlSelectItem = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadMainData(ByVal lngFeeItemID As Long, Optional ByVal strNO As String, Optional ByVal lngAdviceID As Long)
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            If mrsInfo.State = 0 Then Exit Sub
            Call LoadApplyData(lngFeeItemID, lngAdviceID, strNO)
        Else
            Call LoadAppliedData
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            Call LoadAuditData(0)
        Else
            Call LoadAuditData(1)
        End If
    End If
End Sub

Private Function zlGetVarBoundSQL(ByVal strVars As String, ByVal lngStep As Long, ByRef strSQL As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取绑定变量的SQL,主要是受Oracle限制
    '入参:strVars -分离串(用逗号分离)
    '       lngStep-步长(即绑定变量从好多开始)
    '出参:strSQL-返回的SQL
    '返回:返回各绑定变量,主要是10个数组
    '编制:刘兴洪
    '日期:2010-12-27 15:37:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intR As Long, strItems As String, strSubTable As String
    Dim varData As Variant, i As Long, strValues(0 To 10) As String
    strItems = "": strSubTable = ""
    intR = 0:
    varData = Split(strVars, ",")
    For i = 0 To UBound(varData)
        If Len(strItems) > 2000 Then
            If intR <= 10 Then
                strValues(intR) = Mid(strItems, 2)
                strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list([" & intR + lngStep & "]) As ZLTOOLS.t_numlist))"
            Else
                strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list('" & Mid(strItems, 2) & "')  As ZLTOOLS.t_numlist))"
            End If
            strItems = "": intR = intR + 1
        End If
        strItems = strItems & "," & varData(i)
    Next
    
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list([" & intR + lngStep & "]) As ZLTOOLS.t_numlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
                "  Select  Column_Value  As ID From Table(Cast(f_num2list('" & Mid(strItems, 2) & "')  As ZLTOOLS.t_numlist))"
        End If
    End If
    If strSubTable <> "" Then strSubTable = Mid(strSubTable, 11)
    strSQL = strSubTable: zlGetVarBoundSQL = strValues
End Function
Private Function zlApplyToVerify(ByRef str费用ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐审核
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-27 14:53:02
    '问题:34994
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strMCNO As String, strTable As String
    Dim cllPro As Collection, varValue As Variant, strSQL As String
    Dim strIF As String, strNos As String
    
    On Error GoTo errHandle
    varValue = zlGetVarBoundSQL(str费用ID, 3, strTable)
    strTable = _
        " With C1 As (" & strTable & ")" & _
        "  Select 0 As 费用来源, a.ID, a.NO, a.记录性质, a.序号, a.病人id, a.主页id, a.姓名, a.性别, a.操作员姓名, a.登记时间" & _
        "  From 住院费用记录 A,C1" & _
        "  Where a.ID = c1.ID" & _
        "  Union All" & _
        "  Select 1 As 费用来源, a.ID, a.NO, a.记录性质, a.序号, a.病人id, a.主页id, a.姓名, a.性别, a.操作员姓名, a.登记时间" & _
        "  From 门诊费用记录 A,C1" & _
        "  Where a.ID = c1.ID"
    
    strIF = " And Instr(','||[1]||',',','||A.审核部门ID||',')>0 And A.审核部门ID=A.申请部门ID and a.申请部门ID=[2] and A.状态 = 0"
    '是否具有强制记帐权限
    If Not (InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 And InStr(mstrPrivsOpt, "出院结清强制记帐") > 0) Then
        If InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 Then
            strIF = strIF & " And ((G.出院日期 is NULL And Nvl(G.状态,0)<>3) Or Nvl(Y.费用余额,0)<>0)"
        ElseIf InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
            strIF = strIF & " And ((G.出院日期 is NULL And Nvl(G.状态,0)<>3) Or Nvl(Y.费用余额,0)=0)"
        Else
            strIF = strIF & " And G.出院日期 is NULL And Nvl(G.状态,0)<>3"
        End If
    End If
    
    strSQL = _
        " Select /*+RULE*/b.费用来源, a.费用id ID, a.审核部门id, a.申请类别," & _
        "        To_Char(a.申请时间, 'YYYY-MM-DD HH24:MI:SS') 申请时间, a.数量," & _
        "        b.No, b.序号, b.记录性质, g.险类, a.状态, b.姓名, b.性别, b.操作员姓名, b.登记时间" & _
        " From (" & strTable & ") B, 病案主页 G, 病人余额 Y, 病人费用销帐 A" & _
        " Where a.费用id = b.Id And b.病人id = g.病人id And b.主页id = g.主页id" & _
        "       And b.病人id = y.病人id(+) And y.性质(+) = 1 And y.类型(+) = 2" & strIF
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnitIDs, mlngDeptID, _
        CStr(varValue(0)), CStr(varValue(1)), CStr(varValue(2)), CStr(varValue(3)), CStr(varValue(4)), CStr(varValue(5)), _
        CStr(varValue(6)), CStr(varValue(7)), CStr(varValue(8)), CStr(varValue(9)), CStr(varValue(10)))
    If rsTemp.EOF Then zlApplyToVerify = True: Exit Function
    
    Set cllPro = New Collection
    Do While Not rsTemp.EOF
        If zlCheckFeeIsValied(Val(Nvl(rsTemp!费用来源)), Val(Nvl(rsTemp!ID)), _
            Val(Nvl(rsTemp!审核部门id)), Val(Nvl(rsTemp!申请类别))) = False Then Exit Function
        
        'Zl_病人费用销帐_Audit
        strSQL = "Zl_病人费用销帐_Audit("
        '  Id_In       病人费用销帐.费用id%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!ID)) & ","
        '  申请时间_In 病人费用销帐.申请时间%Type,
        strSQL = strSQL & "To_Date('" & Format(Nvl(rsTemp!申请时间), "yyyy-mm-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  审核人_In   病人费用销帐.审核人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  审核时间_In 病人费用销帐.审核时间%Type,
        strSQL = strSQL & "To_Date('" & Format(Nvl(rsTemp!申请时间), "yyyy-mm-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  状态_In     病人费用销帐.状态%Type,--1-审核通过,2-审核未通过
        strSQL = strSQL & "" & "1" & ","
        '  Int自动退料 Integer := 1,
        strSQL = strSQL & "" & "1" & ","
        '  申请类别_In 病人费用销帐.申请类别%Type := 1--对药品和卫材有效,缺省为已执行的药品或卫材
        strSQL = strSQL & "" & Val(Nvl(rsTemp!申请类别)) & ")"
        zlAddArray cllPro, strSQL
        
        If Val(Nvl(rsTemp!费用来源)) = 0 Then
            'Zl_住院记帐记录_Delete
            strSQL = "ZL_住院记帐记录_Delete("
            '  No_In           住院费用记录.No%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!NO) & "',"
            '  序号_In         Varchar2,
            strSQL = strSQL & "'" & Val(Nvl(rsTemp!序号)) & ":" & Val(Nvl(rsTemp!数量)) & "',"
            '  操作员编号_In   住院费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In   住院费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  记录性质_In     住院费用记录.记录性质%Type := 2,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!记录性质)) & ","
            '  操作状态_In     Number := 0,--0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
            strSQL = strSQL & "" & "1" & ")"
            zlAddArray cllPro, strSQL
        Else
            'Zl_门诊记帐记录_Delete
            strSQL = "Zl_门诊记帐记录_Delete("
            '  No_In         门诊费用记录.No%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!NO) & "',"
            '  序号_In       Varchar2,
            strSQL = strSQL & "'" & Val(Nvl(rsTemp!序号)) & ":" & Val(Nvl(rsTemp!数量)) & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type
            strSQL = strSQL & "'" & UserInfo.姓名 & "')"
            zlAddArray cllPro, strSQL
        End If
        
        If Not IsNull(rsTemp!险类) And InStr("," & strMCNO & ",", "," & rsTemp!NO & ",") = 0 Then
            MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val("" & rsTemp!险类))
            MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val("" & rsTemp!险类))
            strMCNO = "|" & Nvl(rsTemp!NO) & "," & Val(Nvl(rsTemp!险类)) & "," & _
                IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
        End If
        
        If InStr("," & strNos & ",", "," & Nvl(rsTemp!NO) & ",") = 0 Then
            '单据操作时间限制检查
            If Not BillOperCheck(IIf(Val(Nvl(rsTemp!费用来源)) = 0, 5, 4), _
                Nvl(rsTemp!操作员姓名), Format(Nvl(rsTemp!登记时间), "YYYY-MM-DD HH:MM:SS"), _
                "销帐审核", Nvl(rsTemp!NO), , 2, , False, False) Then Exit Function
            strNos = strNos & "," & Nvl(rsTemp!NO)
        End If
        rsTemp.MoveNext
    Loop
    If strMCNO <> "" Then strMCNO = Mid(strMCNO, 2)
    
    If ExecuteDataSave(cllPro, strMCNO) = False Then Exit Function
    
    stbThis.Panels(2).Text = "数据销帐审核成功!"
    zlApplyToVerify = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LoadAuditData(ByVal bytType As Byte)
'参数:bytType=0-待审核记录,1-已审核记录
    Dim strSQL As String, strDetail As String, strDosage As String
    Dim rsTmp As ADODB.Recordset
    Dim strIF As String, strFirstCol As String
    Dim DatBegin As Date, DatEnd As Date
    Dim lng病人ID As Long, strForceAccount As String
    
    On Error GoTo errHandle
    If dtpAuditB.Value <= dtpAuditE.Value Then
        DatBegin = dtpAuditB.Value
        DatEnd = dtpAuditE.Value
    Else
        DatBegin = dtpAuditE.Value
        DatEnd = dtpAuditB.Value
    End If
        
    strIF = " And Instr(','||[3]||',',','||A.审核部门ID||',')>0"
    
    If bytType = 0 Then
        If chkDateAudit.Value = 0 Then
            strIF = strIF & " And A.申请时间 Between [1] And [2]"
        End If
        strIF = strIF & " And A.状态 = 0"
        strFirstCol = "' ' 审核, "
    Else
        strIF = strIF & " And A.审核时间 Between [1] And [2]"
        strIF = strIF & " And A.状态 IN(1,2)"
        strFirstCol = "Decode(状态,1,'√','×') 状态, "
    End If
    
    If bytType = 0 Then
        '是否具有强制记帐权限
        If Not (InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 And InStr(mstrPrivsOpt, "出院结清强制记帐") > 0) Then
            If InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 Then
                strForceAccount = " And ((G.出院日期 is NULL And Nvl(G.状态,0)<>3) Or Nvl(Y.费用余额,0)<>0)"
            ElseIf InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
                strForceAccount = " And ((G.出院日期 is NULL And Nvl(G.状态,0)<>3) Or Nvl(Y.费用余额,0)=0)"
            Else
                strForceAccount = " And G.出院日期 is NULL And Nvl(G.状态,0)<>3"
            End If
        End If
    End If
    '问题:42827,42837
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    If lng病人ID <> 0 Then
        strIF = strIF & " And B.病人id+0 = [4] "
    End If
    '问题59958,刘尔旋:显示的申请信息,应该排除进入输液配药中心的药品
    '77686,李南春,2014/9/18,单据类别限制
    strDosage = _
        " And Not Exists(Select RowNum as 序号" & _
        "                From 药品收发记录 B1, 输液配药内容 C1" & _
        "                Where B1.费用id = B.ID And B1.ID = C1.收发id And instr( ',8,9,10,21,24,25,26,',','||B1.单据||',')>0) "
    
    strDetail = _
        " Select 0 As 费用来源,b.Id, b.病人id, b.主页id, b.姓名, b.性别, b.No, b.记录性质, b.序号, b.收费类别," & _
        "        b.实收金额, b.付数, b.数次, b.执行状态, b.执行部门id, b.收费细目id, b.病人病区id," & _
        "        b.开单部门id, b.医嘱序号, b.发生时间, b.登记时间, b.操作员姓名, b.结帐id, a.数量, a.状态," & _
        "        a.申请人, a.申请时间, a.审核部门ID, a.申请类别, a.销帐原因" & _
        " From 病人费用销帐 A,住院费用记录 B" & _
        " Where A.费用id = B.ID " & strDosage & strIF
    strDetail = strDetail & " Union All " & _
        Replace(Replace(strDetail, "住院费用记录", "门诊费用记录"), "0 As 费用来源", "1 As 费用来源")
    
    '明细记录
    strDetail = _
        " Select a.费用来源,a.ID, a.序号, a.记录性质, a.状态, a.姓名, a.性别, G.险类, F.名称 类别, D.名称 病人病区," & _
        "        G.出院病床 床号, E.名称 开单科室, a.收费细目id,C.名称 As 项目名称, " & vbNewLine & _
        "        C.规格, Nvl(X.住院单位,C.计算单位) as 单位,a.NO, a.发生时间, a.登记时间, a.操作员姓名," & _
        "        a.数量/Nvl(X.住院包装,1) As 销帐数量,a.数量 As 售价销帐数量,a.数量*Nvl(a.实收金额,0)/a.付数/a.数次 As 销帐金额 ," & vbNewLine & _
        "        a.申请人, To_Char(a.申请时间,'YYYY-MM-DD HH24:MI:SS') As 申请时间, C.产地, X.药品来源,a.结帐ID,a.执行状态," & _
        "        a.执行部门ID,a.审核部门ID,a.申请类别,a.销帐原因" & vbNewLine & _
        " From (" & strDetail & ") A, 病案主页 G, 病人余额 Y, 收费项目目录 C," & _
        "       药品规格 X, 部门表 D, 部门表 E, 收费项目类别 F" & vbNewLine & _
        " Where a.收费细目id = C.ID And a.病人病区id = D.ID(+) And a.开单部门ID = E.ID " & vbNewLine & _
        "       And a.收费类别 = F.编码 And a.病人id = G.病人id And a.主页id = G.主页id And a.收费细目ID=X.药品ID(+)" & _
        "       And a.病人ID=Y.病人ID(+) And Y.性质(+)=1 And Y.类型(+)=2" & strForceAccount
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strDetail, Me.Caption, DatBegin, DatEnd, mstrUnitIDs, lng病人ID)
    If bytType = 0 Then
        Set mrsAudit = rsTmp
    Else
        Set mrsAudited = rsTmp
    End If
     
    strSQL = _
        " Select " & strFirstCol & "姓名, 性别, 病人病区, 床号, 类别, 项目名称,收费细目ID, 规格, 单位,申请类别," & _
        "       Sum(销帐数量) 销帐数量, Sum(销帐金额) 销帐金额, 申请人, 申请时间, 产地, 药品来源" & vbNewLine & _
        " From (" & strDetail & ")" & vbNewLine & _
        " Group by 收费细目id,申请类别, 状态, 姓名, 性别, 病人病区, 床号, 类别, 项目名称, 规格, 单位, 申请人," & _
        "       申请时间, 产地, 药品来源" & vbNewLine & _
        " Order by 申请时间 Desc,姓名,类别,项目名称,规格"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, DatBegin, DatEnd, mstrUnitIDs, lng病人ID)
    Call ShowMainData(rsTmp)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadAppliedData()
'功能：读取已申请的销帐单
    Dim strSQL As String, strDetail As String
    Dim rsTmp As ADODB.Recordset
    Dim strIF As String, strFirstCol As String, strDosage As String
    Dim DatBegin As Date, DatEnd As Date
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    If Not chkDateAudit.Value = 1 Then
        If dtpAuditB.Value <= dtpAuditE.Value Then
            DatBegin = dtpAuditB.Value
            DatEnd = dtpAuditE.Value
        Else
            DatBegin = dtpAuditE.Value
            DatEnd = dtpAuditB.Value
        End If
        strIF = " And A.申请时间 Between [1] And [2]"
    End If
    If mlngDeptID <> 0 Then
        If mbytUseType <> 1 Then
            strIF = strIF & " And Instr(','||[4]||',',','||A.申请部门ID||',')>0"
        Else
            strIF = strIF & " And A.申请部门ID = [3]"
        End If
    End If
    '问题:42827,42837
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    If lng病人ID <> 0 Or chkDateAudit.Value = 1 Then
        strIF = Replace(strIF, "A.申请时间", "A.申请时间+0") & "  And B.病人id  = [6] "
    End If
    
    '0-全部,1-未审核,2-审核通过,3-审核未通过
    Select Case Val(cboState.ItemData(cboState.ListIndex))
        Case ESTATE.E全部
            strFirstCol = "Decode(状态,0,' ',1,'√','×') 状态,"
        Case ESTATE.E未审核
            strFirstCol = "' ' 状态,"
            '问题:42716
            strIF = strIF & " And A.状态 = 0 "
            If Not (chkOtherOperator.Value = 1 And chkOtherOperator.Visible) Then
               strIF = strIF & " And A.申请人 = [5]"
            End If
        Case ESTATE.E审核通过
            strFirstCol = "'√' 状态,"
            strIF = strIF & " And A.状态 = 1"
        Case ESTATE.E审核未通过
            strFirstCol = "' ' 状态,"
            strIF = strIF & " And A.状态 = 2 And A.申请人 = [5]"
    End Select
    '问题59958,刘尔旋:显示的申请信息,应该排除进入输液配药中心的药品
    '77686,李南春,2014/9/18,单据类别限制
    strDosage = _
        " And Not Exists(Select RowNum as 序号" & _
        "                From 药品收发记录 B1, 输液配药内容 C1" & _
        "                Where B1.费用id = B.ID And B1.ID = C1.收发id And instr( ',8,9,10,21,24,25,26,',','||B1.单据||',')>0)  "
    
    strDetail = _
        " Select b.Id, b.No, b.姓名, b.性别, b.收费类别, b.收费细目ID, b.发生时间, b.标准单价, a.数量," & _
        "        b.执行部门id, b.开单部门id, b.医嘱序号, a.申请人, a.申请时间, a.销帐原因, a.状态" & _
        " From 病人费用销帐 A,住院费用记录 B" & _
        " Where A.费用id = B.ID " & strDosage & strIF
    strDetail = strDetail & " Union All " & Replace(strDetail, "住院费用记录", "门诊费用记录")
    
    '明细记录
    strDetail = _
        "Select A.ID, A.状态, a.姓名, a.性别, F.名称 类别,a.收费类别, A.收费细目id,C.名称 As 项目名称, C.规格," & vbNewLine & _
        "       Nvl(X.住院单位,C.计算单位) as 单位,a.NO, a.发生时间, D.名称 执行科室,E.名称 开单科室," & vbNewLine & _
        "       a.数量/Nvl(X.住院包装,1) As 销帐数量,a.数量*nvl(a.标准单价,0) As 销帐金额, a.申请人," & _
        "       To_Char(a.申请时间,'YYYY-MM-DD HH24:MI:SS') 申请时间,C.产地, X.药品来源,a.医嘱序号,a.销帐原因" & vbNewLine & _
        "From (" & strDetail & ") A, 收费项目目录 C, 收费项目类别 F, 药品规格 X, 部门表 D, 部门表 E" & vbNewLine & _
        "Where A.收费细目id = C.ID And a.执行部门id = D.ID And a.开单部门id = E.ID" & _
        "      And a.收费类别 = F.编码 And A.收费细目ID=X.药品ID(+) "
    Set mrsApplied = zlDatabase.OpenSQLRecord(strDetail, Me.Caption, DatBegin, DatEnd, mlngDeptID, mstrUnitIDs, UserInfo.姓名, lng病人ID)
     
    strSQL = _
        " Select " & strFirstCol & " 姓名, 性别, 类别, 项目名称,收费细目ID, 规格, 单位," & _
        "       Sum(销帐数量) 销帐数量,sum(销帐金额) as 销帐金额, 申请人, 申请时间, 产地, 药品来源" & vbNewLine & _
        " From (" & strDetail & ")" & vbNewLine & _
        " Group by 收费细目id, 状态, 姓名, 性别, 类别, 项目名称, 规格, 单位, 申请人, 申请时间, 产地, 药品来源" & vbNewLine & _
        " Order by 申请时间 Desc,姓名,类别,项目名称,规格"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, DatBegin, DatEnd, mlngDeptID, mstrUnitIDs, UserInfo.姓名, lng病人ID)
    Call ShowMainData(rsTmp)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadApplyData(ByVal lngFeeItemID As Long, Optional ByVal lngAdviceID As Long, _
                        Optional ByVal strNO As String, Optional lngSerial As Long)
    '功能:读取申请销帐记录
    Dim strSQL As String, strSQLDetail As String
    Dim rsTmp As ADODB.Recordset
    Dim strIF As String, blnAppend As Boolean, blnVsfEmpt As Boolean
    Dim DatBegin As Date, DatEnd As Date
    Dim strWhere As String, strWhereExists As String
    Dim strTable As String
    Dim strDosage As String '配药中心配药条件
    Dim strWhereOthers As String
    Dim lng主页ID As Long, str收费类别 As String
    Dim intBaby As Integer, strWhereICU As String
    
    On Error GoTo errHandle
    If lngFeeItemID <> 0 Then
        If vsfMain(0).Rows > 1 Then
            blnVsfEmpt = Val(vsfMain(0).RowData(1)) = 0
        Else
            blnVsfEmpt = True
        End If
        
        If Not blnVsfEmpt Then
            If CheckExistFeeItem(lngFeeItemID) Then
                If MsgBox("输入的销帐项目已存在于列表中,你要清除列表中的内容,只显示该项目吗?", _
                    vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                blnAppend = False
            Else
                blnAppend = True
            End If
        End If
    End If
    
    '全部走病人索引:问题:29176
    strIF = " And A.病人id = [1] And A.记录状态 > 0  "
    '问题:39373
    '55368
    intBaby = cboBaby.ItemData(cboBaby.ListIndex)
    Select Case intBaby
    Case 0  '不含婴儿费
        strIF = strIF & " And nvl(A.婴儿费,0)= 0"
    Case 1  '含婴儿费
    Case Else '显示第几个婴儿费
        strIF = strIF & " And nvl(A.婴儿费,0)= [8]"
    End Select

    '问题:40304
    lng主页ID = 0
    If cbo次数.ListIndex >= 0 Then
         lng主页ID = cbo次数.ItemData(cbo次数.ListIndex)
    End If
    strIF = strIF & IIf(lng主页ID = 0, "", " And nvl(A.主页ID,0)= [7]")
    
    str收费类别 = cboKind.GetNodesCheckedDatas
    If str收费类别 = "" And cboKind.GetNodesCheckedDatas(False) = "" Then
        MsgBox "请选择一项收费类别!", vbInformation, gstrSysName
        Exit Sub
    End If
    strIF = strIF & IIf(Replace(str收费类别, ",", "") = "", "", " And Instr('," & str收费类别 & ",',',' || A.收费类别 || ',') > 0")
    
    If mlngDeptID <> 0 Then
        '0-病区调用,1-医技科室调用,2-医生站调用(只能申请药品，且无审核功能)
        If mbytUseType <> 1 Then
            '38463
            If Not mblnOperatorICU Then
                strWhereICU = " And Instr(','||[6]||',',','||A.病人病区id||',')>0"
                ' 问题:36462
                '如果是医技科室开单, 则护士和医生站(临床)不允许显示药品和卫材
                '如果是医技站申请,也不能看到临床开的单
                '问题:43940:由于会诊医生也存在开单科室<>病人科室的情况,因此, _
                '       经与周韬讨论,直接以开单科室ID是否为临床性质判断, '
                '       不再用病人科室ID=开单科室ID来判断是否为临床开的单了
                 'exists(select 1 From 部门性质说明 where A.开单部门id=部门ID And 工作性质='临床')
                strWhereICU = strWhereICU & _
                    " And (Exists(select 1 From 部门性质说明 where A.开单部门id=部门ID And 工作性质='临床')" & _
                    "       And (Instr(',5,6,7,', ',' || A.收费类别 || ',') > 0 Or (A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 1))  " & _
                    "      Or (Instr(',5,6,7,', ',' || A.收费类别 || ',') = 0 Or A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 0))"
            ElseIf Not mblnPatiDeptICU Then
                '以当时病人科室是否为ICU部分:42526
                strWhereICU = _
                    " And (Exists(Select 1 From  部门性质说明 J1  Where A.病人科室ID=J1.部门ID And J1.工作性质='ICU') " & _
                    "      Or (Exists(select 1 From 部门性质说明 where A.开单部门id=部门ID And 工作性质='临床') " & _
                    "          And (Instr(',5,6,7,', ',' || A.收费类别 || ',') > 0 Or (A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 1)) " & _
                    "      Or (Instr(',5,6,7,', ',' || A.收费类别 || ',') = 0 Or A.收费类别 = '4' And Nvl(C.跟踪在用, 0) = 0)) )"
            End If
        Else
            strIF = strIF & " And A.开单部门id+0 = [2]"
        End If
    End If
    If lngFeeItemID <> 0 Then
        strIF = strIF & " And A.收费细目ID+0 = [3]"
    End If
    If lngAdviceID <> 0 Then
        strIF = strIF & _
            " And A.NO In (Select Distinct a.No" & _
            "              From 病人医嘱发送 A, 病人医嘱记录 B" & _
            "              Where a.医嘱id = b.Id And (b.Id = [10] Or b.相关id = [10]))"
    End If
    If strNO <> "" Then
        strIF = strIF & " And A.NO = [11]"
    End If
    If lngSerial <> 0 Then
        strIF = strIF & " And A.序号 = [12]"
    End If
    If chkDate.Value = 0 Then
        If dtpApplyB.Value <= dtpApplyE.Value Then
            DatBegin = dtpApplyB.Value
            DatEnd = dtpApplyE.Value
        Else
            DatBegin = dtpApplyE.Value
            DatEnd = dtpApplyB.Value
        End If
        strIF = strIF & " And A.发生时间+0 Between [4] And [5]"
    End If
    
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    blnYP = zlStr.IsHavePrivs(mstrPrivsOpt, "药品销帐申请")
    blnZL = zlStr.IsHavePrivs(mstrPrivsOpt, "诊疗销帐申请")
    blnWC = zlStr.IsHavePrivs(mstrPrivsOpt, "卫材销帐申请")
    
    If blnYP And blnWC And blnZL Then
        '全部,不限制
    ElseIf blnYP And blnWC And blnZL = False Then
        strIF = strIF & "  And  收费类别 In('4','5','6','7')"
    ElseIf blnYP And blnWC = False And blnZL Then
        strIF = strIF & "  And  收费类别 <>'4'"
    ElseIf blnYP And blnWC = False And blnZL = False Then
        strIF = strIF & "  And  收费类别 In('5','6','7')"
    ElseIf blnYP = False And blnWC And blnZL = False Then
        strIF = strIF & "  And  收费类别 ='4'"
    ElseIf blnYP = False And blnWC And blnZL Then
        strIF = strIF & "  And instr( '5,6,7',  收费类别)=0 "
    ElseIf blnYP = False And blnWC = False And blnZL Then
        strIF = strIF & "  And instr( '4,5,6,7',  收费类别)=0 "
    Else
        MsgBox "注意:" & vbCrLf & "  你不具备药品、卫材及诊疗销帐申请的权限,请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '69899:刘尔旋,2014-02-09,过滤其他科室的执行费用
    strWhereOthers = IIf(chkShowOthers.Value = 0, " And Exists(Select 1 From 部门人员 Where A.执行部门ID =部门ID And 人员ID= [9])", " ")

    '36391:将1替换为RowNum,避免Oracle视图自动合并:42333
    '77686,李南春,2014/9/18,单据类别限制
    strDosage = _
        " And Not Exists(Select RowNum as 序号" & _
        "                From 住院费用记录 J, 药品收发记录 B1, 输液配药内容 C1" & _
        "                Where j.NO = a.NO And a.记录性质 = j.记录性质 and nvl(A.价格父号, A.序号) = Nvl(J.价格父号, J.序号)" & _
        "                      And B1.费用id = j.ID And B1.ID = C1.收发id And instr( ',8,9,10,21,24,25,26,',','||B1.单据||',')>0)  "
    
    strWhere = ""
    '问题:30523
    '由于以前对未执行的药品或卫材,只能对病区本身发的药品或卫材进行申请,而由药房发药或卫材的,则不能够申请处理,因此出现程序在此环节上的流程处理漏洞.
    '所以现在取消了该限制,现在的处理方式是如果药房存在未执行的,则申请时,审核部门只能为病区(如果在病区审核前,该药品被药品执行,则禁止审核),执行了的,则为执行部门.
    If chk项目(0).Value = 1 And chk项目(1).Value = 0 Then '只显示已执行的
        strWhere = _
            " And Exists(Select 1 From 住院费用记录 B" & _
            "            Where A.NO = B.NO And A.记录性质 = B.记录性质" & _
            "                  And Nvl(A.价格父号, A.序号) = Nvl(B.价格父号, B.序号) And B.执行状态 <> 0 )" & vbNewLine
    ElseIf chk项目(0).Value = 0 And chk项目(1).Value = 1 Then '只显示未执行的
        strWhere = _
            " And Exists(Select 1 From 住院费用记录 B" & _
            "            Where A.NO = B.NO And A.记录性质 = B.记录性质" & _
            "                  And Nvl(A.价格父号, A.序号) = Nvl(B.价格父号, B.序号) And B.执行状态 = 0 )" & vbNewLine
    ElseIf chk项目(0).Value = 0 And chk项目(1).Value = 0 Then '未选择执行项目的,缺省为全选
    Else
    End If
    
    '未结帐的(结帐并作废当未结账)
    strWhereExists = "" & _
    "   And exists(Select 1 From 住院费用记录 A1" & vbNewLine & _
    "              Where Mod(A1.记录性质, 10) = 2 And A.NO=A1.NO and Nvl(A.价格父号, A.序号)=Nvl(A1.价格父号, A1.序号)" & vbNewLine & _
    "              Group By A1.NO, Mod(A1.记录性质, 10), Nvl(A1.价格父号, A1.序号)" & vbNewLine & _
    "              Having Nvl(Sum(A1.结帐金额),0) = 0) "
    
    strTable = _
        " Select 0 As 费用来源,a.Id, a.No, a.序号, a.价格父号, a.收费类别, a.收费细目id, a.标准单价," & _
        "        a.付数, a.数次, a.应收金额, a.实收金额, a.发生时间, a.登记时间, a.操作员姓名," & _
        "        a.执行状态, a.婴儿费, a.执行部门id, a.病人科室ID, a.开单部门id, a.结帐id, a.医嘱序号, a.病人病区ID" & _
        " From 住院费用记录 A" & _
        " Where a.记录性质 = 2 And Nvl(a.病人病区ID,0)<>0 " & _
        "       And Exists(Select 1 From 住院费用记录" & _
        "                  Where A.NO = NO And A.记录性质 = 记录性质" & _
        "                        And Nvl(A.价格父号, A.序号) = Nvl(价格父号, 序号) And 执行状态 <> 0 )" & vbNewLine & _
                strWhereExists & strIF & strWhereOthers & strDosage
    strTable = strTable & " Union All " & _
        Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "0 As 费用来源", "1 As 费用来源")

    strTable = _
    " Select Max(a.费用来源) As 费用来源,Max(Decode(Sign(A.执行状态), -1, 0, Decode(A.价格父号, Null, A.ID, 0))) ID, " & _
    "        A.执行状态,nvl(A.婴儿费,0) as 婴儿费, A.发生时间, Max(A.登记时间) 登记时间, " & vbNewLine & _
    "        Max(Decode(Sign(A.执行状态), -1, Null,A.操作员姓名)) 操作员姓名, A.NO," & vbNewLine & _
    "        Nvl(A.价格父号, A.序号) As 序号, A.执行部门id,A.开单部门id, A.收费类别, A.收费细目id, Avg(A.标准单价) 单价," & vbNewLine & _
    "        Avg(A.付数) 付数, Avg(A.数次) 数次, Sum(A.应收金额) 应收金额, Sum(A.实收金额) 实收金额, A.结帐ID, A.医嘱序号  " & vbNewLine & _
    " From (" & strTable & ") A, 材料特性 C" & vbNewLine & _
    " Where A.收费细目id = C.材料id(+)" & strWhereICU & _
    "       And (A.收费类别 in ('5','6','7') or (A.收费类别='4' and  nvl(C.跟踪在用,0) = 1)) " & _
    " Group By A.NO,A.执行状态,nvl(A.婴儿费,0),  Nvl(A.价格父号, A.序号),A.发生时间, A.执行部门id," & _
    "       A.开单部门id, A.收费类别, A.收费细目id,A.结帐ID,  A.医嘱序号 "
    
    strTable = "" & _
    " Select Max(费用来源) As 费用来源,Max(ID) ID, NO, 发生时间, Max(登记时间) 登记时间," & _
    "        Max(操作员姓名) as 操作员姓名,max(婴儿费) as 婴儿费, 序号,  " & _
    "        执行部门id,开单部门id, 收费类别, 收费细目id, Avg(单价) 单价," & vbNewLine & _
    "        Sum(付数 * 数次)  数次, Sum(应收金额) 应收金额, Sum(实收金额) 实收金额, Max(结帐ID) as 结帐ID, 医嘱序号" & vbNewLine & _
    " From (" & strTable & ") " & _
    " Group By NO, 发生时间, 序号, 执行部门id,开单部门id, 收费类别, 收费细目id, 医嘱序号" & vbNewLine & _
    " Having Sum(付数 * 数次) <> 0 "
    
    '问题:38388
    strTable = " With 费用  as ( " & strTable & ") "
    strSQL = ""
    If chk项目(0).Value = 1 Or (chk项目(0).Value = 0 And chk项目(1).Value = 0) Then
        '已执行项目,需要处理已退药部分
        strSQL = strSQL & " UNION ALL " & _
        "      Select Max(c1.费用来源) As 费用来源,C1.ID,-1 as 执行状态,max(C1.婴儿费) as 婴儿费,C1.发生时间,C1.登记时间,C1.操作员姓名," & _
        "             C1.NO,C1.序号,C1.执行部门id,C1.开单部门id, C1.收费类别, C1.收费细目id, max(C1.单价) 单价," & _
        "             1 as 付数, -1* Sum(B.实际数量)  as 数次 ," & _
        "             -1*Sum(C1.应收金额)*Round(Sum(Nvl(B.付数,1)*B.实际数量) /  sum(C1.数次),5) as 应收金额," & _
        "             -1*Sum(C1.实收金额)*Round(Sum(Nvl(B.付数,1)*B.实际数量) / sum(C1.数次),5) as 实收金额," & _
        "             C1.结帐ID, C1.医嘱序号,1 as 现执行状态,1 as 药品卫材" & _
        "      From  费用 C1,药品收发记录  B " & _
        "      Where C1.ID=B.费用ID And MOD(B.记录状态,3)=1 And B.单据 In(24,25,26,8,9,10) And B.审核人 is NULL " & _
        "      Group By C1.ID,C1.发生时间,C1.登记时间,C1.操作员姓名,C1.NO,C1.序号,C1.执行部门id,C1.开单部门id," & _
        "               C1.收费类别, C1.收费细目id,C1.结帐ID, C1.医嘱序号"
    End If
    
    If chk项目(1).Value = 1 Or (chk项目(0).Value = 0 And chk项目(1).Value = 0) Then
        strSQL = strSQL & " Union ALL " & _
        "      Select Max(c1.费用来源) As 费用来源,C1.ID,0 as 执行状态,max(C1.婴儿费) as 婴儿费 ,C1.发生时间,C1.登记时间,C1.操作员姓名," & _
        "             C1.NO,C1.序号,C1.执行部门id,C1.开单部门id, C1.收费类别, C1.收费细目id, max(C1.单价) 单价," & _
        "             1 as 付数,  Sum(Nvl(B.付数,1)*B.实际数量)  as 数次 ," & _
        "             Sum(C1.应收金额)*Round(Sum(Nvl(B.付数,1)*B.实际数量) /  sum(C1.数次),5) as 应收金额," & _
        "             Sum(C1.实收金额)*Round(Sum(Nvl(B.付数,1)*B.实际数量) / sum(C1.数次),5) as 实收金额," & _
        "             C1.结帐ID, C1.医嘱序号,0 as 现执行状态,1 as 药品卫材" & _
        "      From  费用 C1,药品收发记录  B " & _
        "      Where C1.ID=B.费用ID And MOD(B.记录状态,3)=1 And B.单据  In(24,25,26,8,9,10) And B.审核人 is NULL " & _
        "      Group By C1.ID,C1.发生时间,C1.登记时间,C1.操作员姓名,C1.NO,C1.序号,C1.执行部门id,C1.开单部门id," & _
        "               C1.收费类别, C1.收费细目id,C1.结帐ID, C1.医嘱序号"
    End If
    
    strSQLDetail = _
        " Select 0 As 费用来源,a.Id, a.No, a.记录状态, a.序号, a.价格父号, a.收费类别, a.收费细目id, a.标准单价," & _
        "        a.付数, a.数次, a.应收金额, a.实收金额, a.发生时间, a.登记时间, a.操作员姓名," & _
        "        a.执行状态, a.婴儿费, a.执行部门id, a.病人科室ID, a.开单部门id, a.结帐id, a.医嘱序号, a.病人病区ID" & _
        " From 住院费用记录 A" & _
        " Where a.记录性质 = 2 And Nvl(a.病人病区ID,0)<>0 " & _
                strWhereExists & strIF & strWhereOthers & strDosage
    strSQLDetail = strSQLDetail & " Union All " & _
        Replace(Replace(strSQLDetail, "住院费用记录", "门诊费用记录"), "0 As 费用来源", "1 As 费用来源")
    
    '未退数量不等于零的
    '退过药后,因为退的时候只输数次,所以付数不准,都取1
    '如果是药品, 可能发药后又退药了,所以要用Exists子查询判断,不能直接用执行状态<>0
    '31313:Max(结帐ID):主要是解决先结帐后,再对记帐单进行销帐的情况
    strSQLDetail = "" & _
        " Select Max(费用来源) As 费用来源,Max(ID) ID,现执行状态 as 执行状态,max(婴儿费) as 婴儿费,药品卫材, NO, 发生时间, Max(登记时间) 登记时间," & _
        "       Max(操作员姓名) as 操作员姓名, 序号, 执行部门id,开单部门id, 收费类别, 收费细目id, Avg(单价) 单价," & vbNewLine & _
        "       Decode(Sign(Min(执行状态)), -1, 1, Sum(付数)) 付数,Decode(Sign(Min(执行状态)), -1, Sum(付数 * 数次), Sum(数次)) 数次, " & vbNewLine & _
        "       Sum(应收金额) 应收金额, Sum(实收金额) 实收金额, Max(结帐ID) as 结帐ID, 医嘱序号" & vbNewLine & _
        " From (Select Max(a.费用来源) As 费用来源,Max(Decode(Sign(A.执行状态), -1, 0, Decode(A.价格父号, Null, A.ID, 0))) ID," & _
        "              A.执行状态 ,max(nvl(A.婴儿费,0)) as 婴儿费, A.发生时间, Max(A.登记时间) 登记时间, " & vbNewLine & _
        "              Decode(Sign(A.执行状态), -1, Null,A.操作员姓名) 操作员姓名, A.NO," & vbNewLine & _
        "              Nvl(A.价格父号, A.序号) As 序号, A.执行部门id,A.开单部门id, A.收费类别, A.收费细目id, Avg(A.标准单价) 单价," & vbNewLine & _
        "              Avg(A.付数) 付数, Avg(A.数次) 数次, Sum(A.应收金额) 应收金额, Sum(A.实收金额) 实收金额, A.结帐ID, A.医嘱序号, " & vbNewLine & _
        "              Max(Decode(Sign(A.执行状态), -1, 1, Decode(A.价格父号, Null, decode(A.执行状态,2,1,1,1,decode(A.记录状态,1,0,1)), 1))) 现执行状态," & _
        "             decode(A.收费类别,'5',1,'6',1,'7',1,'4',decode(Max(nvl(C.跟踪在用,0)),1,1,0),0) as 药品卫材  " & _
        "        From (" & strSQLDetail & ") A, 材料特性 C" & vbNewLine & _
        "        Where A.收费细目id = C.材料id(+)" & strWhereICU & _
        "        Group By A.NO, A.执行状态,Decode(Sign(A.执行状态), -1, Null,A.操作员姓名), Nvl(A.价格父号, A.序号), " & _
        "           A.发生时间, A.执行部门id,A.开单部门id, A.收费类别, A.收费细目id, A.结帐ID, A.医嘱序号 " & _
                 strSQL & _
        "       )" & vbNewLine & _
        " Group By NO, 发生时间, 序号, 现执行状态,执行部门id,开单部门id, 收费类别,药品卫材, 收费细目id, 医嘱序号" & vbNewLine & _
        " Having Sum(付数 * 数次) <> 0 "
    
    strSQLDetail = strTable & vbCrLf & strSQLDetail
    
    '30523:屏蔽
    '"            And   (A.收费类别 Not In ('4', '5', '6', '7') Or A.收费类别 = '4' And C.跟踪在用 = 0 Or" & vbNewLine & _
    '"                     (A.收费类别 In ('5', '6', '7') Or A.收费类别 = '4' And C.跟踪在用 = 1) And Exists" & vbNewLine & _
    '"                        (Select 1" & vbNewLine & _
    '"              From 住院费用记录 B" & vbNewLine & _
    '"              Where A.NO = B.NO And A.记录性质 = B.记录性质 And Nvl(A.价格父号, A.序号) = Nvl(B.价格父号, B.序号)  " & vbNewLine & _
    '"                         And (B.执行状态 <> 0 " & strWhere & ")))" & vbNewLine
    
    '可申请销帐的明细
    'A.单价*Nvl(X.住院包装,1) as 单价,:问题:42823
    strSQL = _
        " Select a.费用来源,A.ID,A.执行状态,婴儿费, A.NO, A.序号, A.发生时间, A.登记时间, A.操作员姓名,  " & _
        "        B.名称 执行科室,A.执行部门ID,D.名称 开单科室,A.开单部门ID, A.收费类别, A.收费细目ID," & _
        "        A.单价*Nvl(X.住院包装,1) as 单价, A.付数, A.数次 as 售价数次,A.数次/Nvl(X.住院包装,1) 数次," & vbNewLine & _
        "        A.应收金额, A.实收金额, Nvl(C.数量, 0)/Nvl(X.住院包装,1) 销帐数量,nvl(C.数量,0)*A.单价 as 销帐金额," & _
        "        Nvl(X.住院包装,1) 住院包装, A.结帐ID, A.医嘱序号, Nvl(c.销帐原因,e.操作说明) As 销帐原因" & vbNewLine & _
        " From (" & strSQLDetail & ") A, 病人费用销帐 C,药品规格 X, 部门表 B, 部门表 D, 病人医嘱状态 E" & vbNewLine & _
        " Where A.执行部门id = B.ID And A.医嘱序号 = E.医嘱ID(+) And E.操作类型(+) = 8 And A.开单部门id = D.ID" & _
        "       And A.ID = C.费用id(+) and decode(A.药品卫材,1,A.执行状态,0)=C.申请类别(+) And A.收费细目ID=X.药品ID(+)" & vbNewLine & _
        "       And C.状态(+) = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID), mlngDeptID, lngFeeItemID, DatBegin, DatEnd, _
        mstrUnitIDs, lng主页ID, intBaby - 1, UserInfo.ID, lngAdviceID, strNO, lngSerial)
    Call MakeApplyRecordSet(rsTmp, blnAppend) '为了修改数量,转为可修改的记录集
    
    '明细按收费细目汇总
        strSQL = _
        " Select A.收费细目ID, C.名称 as 类别, C.编码 as 收费类别, max(B.名称) as 项目名称," & _
        "        B.规格, Nvl(X.住院单位,B.计算单位) as 单位,B.产地, X.药品来源,Sum(A.付数 * A.数次/Nvl(X.住院包装,1)) 数量," & vbNewLine & _
        "       Sum(Nvl(D.数量/Nvl(X.住院包装,1), 0)) 销帐数量,sum(Nvl(D.数量,0)*nvl(A.单价,0)) as 销帐金额 " & vbNewLine & _
        " From (" & strSQLDetail & ") A, 收费项目目录 B, 收费项目类别 C, 病人费用销帐 D, 药品规格 X" & vbNewLine & _
        " Where A.收费细目ID = B.ID And A.收费类别 = C.编码 And A.ID = D.费用id(+) And D.状态(+) = 0" & _
        "       And decode(A.药品卫材,1,A.执行状态,0)=D.申请类别(+)  And A.收费细目ID=X.药品ID(+)" & vbNewLine & _
        " Group By A.收费细目ID,A.收费类别,C.名称, C.编码, B.名称, B.规格, Nvl(X.住院单位,B.计算单位),B.产地, X.药品来源" & _
        " Order By 类别,项目名称,规格"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID), mlngDeptID, lngFeeItemID, DatBegin, DatEnd, mstrUnitIDs, lng主页ID, intBaby - 1, UserInfo.ID, lngAdviceID, strNO, lngSerial)
    
    Call ShowMainData(rsTmp, blnAppend)
    If rsTmp.RecordCount = 0 And mblnInit = True Then
        MsgBox "无法找到销账申请的单据，请调整条件后重试。", vbInformation, gstrSysName
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckExistFeeItem(ByRef lngFeeItemID As Long) As Boolean
    Dim i As Long
    
    For i = 1 To vsfMain(0).Rows - 1
        If lngFeeItemID = Val(vsfMain(0).RowData(i)) Then
            CheckExistFeeItem = True
            Exit For
        End If
    Next
End Function

Private Sub MakeApplyRecordSet(ByRef rsDetail As ADODB.Recordset, ByVal blnAppend As Boolean)
'功能：将申请销帐的记录集转换为可修改的记录集
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    
    If Not blnAppend Then
        rsTmp.Fields.Append "ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "优先", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "结帐ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "执行状态", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "医嘱序号", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "收费类别", adVarChar, 20, adFldIsNullable
        rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
        rsTmp.Fields.Append "序号", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "发生时间", adDBTimeStamp, , adFldIsNullable
        rsTmp.Fields.Append "登记时间", adDBTimeStamp, , adFldIsNullable
        rsTmp.Fields.Append "操作员姓名", adVarChar, 100, adFldIsNullable
        rsTmp.Fields.Append "执行科室", adVarChar, 100, adFldIsNullable
        rsTmp.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
        rsTmp.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "付数", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "售价数次", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "数次", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "应收金额", adCurrency, , adFldIsNullable
        rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
        rsTmp.Fields.Append "销帐数量", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "销帐金额", adDouble, , adFldIsNullable     '问题:35595
        rsTmp.Fields.Append "原始销帐金额", adDouble, , adFldIsNullable '问题:35595
        rsTmp.Fields.Append "原始销帐数量", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "住院包装", adDouble, , adFldIsNullable
        rsTmp.Fields.Append "婴儿费", adBigInt, , adFldIsNullable
        rsTmp.Fields.Append "销帐原因", adVarChar, 200, adFldIsNullable
        rsTmp.Fields.Append "原始销帐原因", adVarChar, 200, adFldIsNullable
        rsTmp.Fields.Append "费用来源", adBigInt, , adFldIsNullable
        
        rsTmp.CursorLocation = adUseClient
        rsTmp.LockType = adLockOptimistic
        rsTmp.CursorType = adOpenStatic
        rsTmp.Open
        
        Set mrsApply = rsTmp
    End If

    With mrsApply
        For i = 1 To rsDetail.RecordCount
            .AddNew
            !优先 = 0
            If mblnOperatorNurse Then
                '60679
                '如果是护士,则在输入申请数量后,先按操作员所属科室开单进行分配
                If InStr(1, mstrOperatorDeptIDs, "," & Val(Nvl(rsDetail!开单部门ID)) & ",") > 0 Then
                    !优先 = 1
                End If
            End If
            !ID = rsDetail!ID
            !结帐ID = rsDetail!结帐ID
            !执行状态 = Val(Nvl(rsDetail!执行状态))
            !医嘱序号 = rsDetail!医嘱序号
            !收费类别 = rsDetail!收费类别
            !收费细目ID = rsDetail!收费细目ID
            !NO = rsDetail!NO
            !序号 = rsDetail!序号
            !发生时间 = rsDetail!发生时间
            !登记时间 = rsDetail!登记时间
            !操作员姓名 = rsDetail!操作员姓名
            !执行科室 = rsDetail!执行科室
            !开单科室 = rsDetail!开单科室
            !执行部门ID = rsDetail!执行部门ID
            !开单部门ID = rsDetail!开单部门ID
            !单价 = rsDetail!单价
            !付数 = rsDetail!付数
            !售价数次 = rsDetail!售价数次
            !数次 = rsDetail!数次
            !应收金额 = rsDetail!应收金额
            !实收金额 = rsDetail!实收金额
            !销帐数量 = rsDetail!销帐数量
            !销帐金额 = rsDetail!销帐金额 '问题:35595
            !原始销帐金额 = rsDetail!销帐金额 '问题:35595
            !原始销帐数量 = rsDetail!销帐数量
            !住院包装 = rsDetail!住院包装
            !婴儿费 = rsDetail!婴儿费 '39374
            !销帐原因 = rsDetail!销帐原因
            !原始销帐原因 = rsDetail!销帐原因
            !费用来源 = rsDetail!费用来源
            .Update
            rsDetail.MoveNext
        Next
        If .RecordCount > 0 Then .MoveFirst
    End With
End Sub

Private Sub ShowMainData(ByRef rsTmp As ADODB.Recordset, Optional ByVal blnAppend As Boolean)
'参数:blnAppend=True-追加,False-重新加载
    Dim i As Long, j As Long, lngInitRows As Long
    Dim intState As Integer
    Dim vsfCurrent As VSFlexGrid
    
    If tbsType.SelectedItem.Key = "T1" Then
        Set vsfCurrent = vsfMain(0)
        If mbytFun = E申请 Then
            cmdOKApply.Enabled = rsTmp.RecordCount > 0
        Else
            cmdOKAudit.Enabled = rsTmp.RecordCount > 0
        End If
    Else
        Set vsfCurrent = vsfMain(1)
        If mbytFun = E申请 Then
            intState = Val(cboState.ItemData(cboState.ListIndex))
            cmdCancelApply.Enabled = rsTmp.RecordCount > 0
        Else
            cmdCancelRefuse.Enabled = False
            cmdOKAudit.Enabled = False
        End If
    End If
    
    If blnAppend Then   'And mbytFun = E申请 And tbsType.SelectedItem.Key = "T1"
        lngInitRows = vsfCurrent.Rows
        If vsfCurrent.Rows = 2 Then
            If Val(vsfCurrent.RowData(1)) = 0 Then lngInitRows = 1
        End If
    Else
        Call InitMainHead(False, IIf(tbsType.SelectedItem.Key = "T1", 1, 2))
        lngInitRows = 1
    End If
    
    
    With vsfCurrent
        If rsTmp.RecordCount <> 0 Then
            .Redraw = flexRDNone
            .Rows = rsTmp.RecordCount + lngInitRows
            For i = lngInitRows To .Rows - 1
                If mbytFun = E申请 Then
                    If tbsType.SelectedItem.Key = "T1" Then
                        .TextMatrix(i, ColApply("类别")) = rsTmp!类别
                        .TextMatrix(i, ColApply("项目名称")) = rsTmp!项目名称
                        .TextMatrix(i, ColApply("规格")) = "" & rsTmp!规格
                        .TextMatrix(i, ColApply("单位")) = "" & rsTmp!单位
                        .TextMatrix(i, ColApply("产地")) = "" & rsTmp!产地
                        '.TextMatrix(i, ColApply("婴儿费")) = IIf(Val(Nvl(rsTmp!婴儿费)) <> 0, "√", "")
                        .TextMatrix(i, ColApply("药品来源")) = "" & rsTmp!药品来源
                        .TextMatrix(i, ColApply("数量")) = FormatEx(rsTmp!数量, 5)
                        .TextMatrix(i, ColApply("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                        .TextMatrix(i, ColApply("销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                        .TextMatrix(i, ColApply("原始销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                        .TextMatrix(i, ColApply("原始销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                        .RowData(i) = Val(rsTmp!收费细目ID)
                                    
                        '设置可修改列的颜色
                        mbonNotEnter = True
                        .Row = i
                        .Col = ColApply("销帐数量")
                        .CellBackColor = &HE7CFBA    '蓝色
                        mbonNotEnter = False
                    Else
                        .TextMatrix(i, ColApplied("选择")) = rsTmp!状态
                        .TextMatrix(i, ColApplied("姓名")) = rsTmp!姓名
                        .TextMatrix(i, ColApplied("性别")) = "" & rsTmp!性别
                        .TextMatrix(i, ColApplied("类别")) = rsTmp!类别
                        .TextMatrix(i, ColApplied("项目名称")) = rsTmp!项目名称
                        .TextMatrix(i, ColApplied("规格")) = "" & rsTmp!规格
                        .TextMatrix(i, ColApplied("单位")) = "" & rsTmp!单位
                        .TextMatrix(i, ColApplied("产地")) = "" & rsTmp!产地
                        .TextMatrix(i, ColApplied("药品来源")) = "" & rsTmp!药品来源
                        .TextMatrix(i, ColApplied("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                        .TextMatrix(i, ColApplied("销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                        .TextMatrix(i, ColApplied("申请人")) = rsTmp!申请人
                        .TextMatrix(i, ColApplied("申请时间")) = rsTmp!申请时间
                        .RowData(i) = Val(rsTmp!收费细目ID)
                        
                        mbonNotEnter = True
                        .Row = i
                        If intState = ESTATE.E未审核 Then
                            .Col = ColApplied("选择")
                            .CellBackColor = &HE7CFBA    '蓝色
                        ElseIf intState = ESTATE.E全部 Then
                            For j = 0 To .Cols - 1
                                .Col = j
                                If rsTmp!状态 = "√" Then
                                    .CellForeColor = &HC00000
                                ElseIf rsTmp!状态 = "×" Then
                                    .CellForeColor = &HC0&
                                End If
                            Next
                        ElseIf intState = ESTATE.E审核通过 Then
                            For j = 0 To .Cols - 1
                                .Col = j
                                .CellForeColor = &HC00000
                            Next
                        ElseIf intState = ESTATE.E审核未通过 Then
                            For j = 0 To .Cols - 1
                                .Col = j
                                .CellForeColor = &HC0&
                            Next
                        End If
                        mbonNotEnter = False
                    End If
                Else
                    If tbsType.SelectedItem.Key = "T1" Then
                        .TextMatrix(i, ColAudit("审核")) = rsTmp!审核
                        .Cell(flexcpData, i, ColAudit("审核")) = Val(Nvl(rsTmp!申请类别))
                        
                        .TextMatrix(i, ColAudit("姓名")) = rsTmp!姓名
                        .TextMatrix(i, ColAudit("性别")) = "" & rsTmp!性别
                        .TextMatrix(i, ColAudit("病人病区")) = "" & rsTmp!病人病区
                        .TextMatrix(i, ColAudit("床号")) = "" & rsTmp!床号
                        .TextMatrix(i, ColAudit("类别")) = rsTmp!类别
                        .TextMatrix(i, ColAudit("项目名称")) = rsTmp!项目名称
                        .TextMatrix(i, ColAudit("规格")) = "" & rsTmp!规格
                        .TextMatrix(i, ColAudit("产地")) = "" & rsTmp!产地
                        .TextMatrix(i, ColAudit("药品来源")) = "" & rsTmp!药品来源
                        .TextMatrix(i, ColAudit("单位")) = "" & rsTmp!单位
                        .TextMatrix(i, ColAudit("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                        .TextMatrix(i, ColAudit("销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                        .TextMatrix(i, ColAudit("申请人")) = rsTmp!申请人
                        .TextMatrix(i, ColAudit("申请时间")) = rsTmp!申请时间
                        .RowData(i) = Val(rsTmp!收费细目ID)
                        
                        mbonNotEnter = True
                        .Row = i
                        .Col = ColAudit("审核")
                        .CellBackColor = &HE7CFBA    '蓝色
                        mbonNotEnter = False
                    Else
                        .Cell(flexcpData, i, ColAudited("状态")) = Val(Nvl(rsTmp!申请类别))
                        .TextMatrix(i, ColAudited("状态")) = rsTmp!状态
                        .TextMatrix(i, ColAudited("姓名")) = rsTmp!姓名
                        .TextMatrix(i, ColAudited("性别")) = "" & rsTmp!性别
                        .TextMatrix(i, ColAudited("病人病区")) = "" & rsTmp!病人病区
                        .TextMatrix(i, ColAudited("床号")) = "" & rsTmp!床号
                        .TextMatrix(i, ColAudited("类别")) = rsTmp!类别
                        .TextMatrix(i, ColAudited("项目名称")) = rsTmp!项目名称
                        .TextMatrix(i, ColAudited("规格")) = "" & rsTmp!规格
                        .TextMatrix(i, ColAudited("产地")) = "" & rsTmp!产地
                        .TextMatrix(i, ColAudited("药品来源")) = "" & rsTmp!药品来源
                        .TextMatrix(i, ColAudited("单位")) = "" & rsTmp!单位
                        .TextMatrix(i, ColAudited("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                        .TextMatrix(i, ColAudited("申请人")) = rsTmp!申请人
                        .TextMatrix(i, ColAudited("申请时间")) = rsTmp!申请时间
                        .RowData(i) = Val(rsTmp!收费细目ID)
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            .Row = 1: .Col = 0
            If mbytFun = E申请 Then
                If tbsType.SelectedItem.Key = "T1" Then
                    .Col = ColApply("销帐数量") '调用事件AfterRowColChange
                End If
            Else
                If tbsType.SelectedItem.Key = "T1" Then
                    .Row = 0: .Col = ColAudit("审核")
                    .CellBackColor = &HE7CFBA    '蓝色
                    .Row = 1
                End If
            End If
            
            .Redraw = flexRDDirect
        End If
        Call ShowDetail(.RowData(.Row))
    End With
End Sub


Private Sub ShowDetail(ByVal lngFeeItem As Long)
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            Set rsTmp = mrsApply
            rsTmp.Filter = "收费细目ID=" & lngFeeItem   '注意,它会改变原记录集的Filter
        Else
            Set rsTmp = mrsApplied
            With vsfMain(1)
                rsTmp.Filter = "收费细目ID=" & lngFeeItem & " And 申请人='" & .TextMatrix(.Row, ColApplied("申请人")) & _
                            "' And 申请时间='" & .TextMatrix(.Row, ColApplied("申请时间")) & "'"
            End With
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            Set rsTmp = mrsAudit
            With vsfMain(0)
                rsTmp.Filter = "收费细目ID=" & lngFeeItem & " And 申请类别=" & Val(.Cell(flexcpData, .Row, ColAudit("审核"))) & " And 申请人='" & .TextMatrix(.Row, ColAudit("申请人")) & _
                            "' And 申请时间='" & .TextMatrix(.Row, ColAudit("申请时间")) & "'"
            End With
        Else
            Set rsTmp = mrsAudited
            With vsfMain(1)
                rsTmp.Filter = "收费细目ID=" & lngFeeItem & " And 申请人='" & .TextMatrix(.Row, ColAudited("申请人")) & _
                            "' And 申请时间='" & .TextMatrix(.Row, ColAudited("申请时间")) & "'"
            End With
        End If
    End If
    
    Call InitDetailHead(True)   '因显示的列不同,要重设宽度
       
    If rsTmp.State = 0 Then Exit Sub
    If rsTmp.RecordCount = 0 Then Exit Sub
    rsTmp.Sort = IIf(tbsType.SelectedItem.Key = "T1", "执行状态,", "") & "发生时间 Desc,NO Desc"
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    With vsfDetail
        .Redraw = flexRDNone
        .Rows = rsTmp.RecordCount + 1
        mblnUnChange = True
        For i = 1 To .Rows - 1
            If mbytFun = E申请 Then
                If tbsType.SelectedItem.Key = "T1" Then
                    If InStr(1, "5,6,7", Nvl(rsTmp!收费类别)) > 0 Then
                        .TextMatrix(i, .ColIndex("执行状态")) = IIf(Val(Nvl(rsTmp!执行状态)) = 0, "未发药", "已发药")
                    ElseIf Nvl(rsTmp!收费类别) = "4" Then
                        .TextMatrix(i, .ColIndex("执行状态")) = IIf(Val(Nvl(rsTmp!执行状态)) = 0, "未发料", "已发料")
                    Else
                         .TextMatrix(i, .ColIndex("执行状态")) = IIf(Val(Nvl(rsTmp!执行状态)) = 0, "未执行", "已执行")
                    End If
                    
                    .Cell(flexcpData, i, .ColIndex("执行状态")) = Nvl(rsTmp!执行状态)
                    .TextMatrix(i, .ColIndex("NO")) = rsTmp!NO
                    .TextMatrix(i, .ColIndex("发生时间")) = Format(rsTmp!发生时间, "YYYY-MM-DD HH:MM:SS")
                    .TextMatrix(i, .ColIndex("婴儿费")) = IIf(Val(Nvl(rsTmp!婴儿费)) >= 1, "√", "")
                    .TextMatrix(i, .ColIndex("执行科室")) = rsTmp!执行科室
                    .TextMatrix(i, .ColIndex("开单科室")) = rsTmp!开单科室
                    .TextMatrix(i, .ColIndex("单价")) = Format(rsTmp!单价, "######" & gstrFeePrecisionFmt)
                    .TextMatrix(i, .ColIndex("付数")) = rsTmp!付数
                    .TextMatrix(i, .ColIndex("销帐原因")) = Nvl(rsTmp!销帐原因)
                    .TextMatrix(i, .ColIndex("原始销帐原因")) = Nvl(rsTmp!销帐原因)
                    .TextMatrix(i, .ColIndex("数次")) = FormatEx(rsTmp!数次, 5)
                    .TextMatrix(i, .ColIndex("应收金额")) = Format(rsTmp!应收金额, "#######" & gstrDec)
                    .TextMatrix(i, .ColIndex("实收金额")) = Format(rsTmp!实收金额, "#######" & gstrDec)
                    .TextMatrix(i, .ColIndex("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                    .TextMatrix(i, .ColIndex("销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                    .TextMatrix(i, .ColIndex("原始销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                    .TextMatrix(i, .ColIndex("原始销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                    .RowData(i) = Val(rsTmp!ID)
                    .Cell(flexcpBackColor, i, .ColIndex("销帐数量")) = &HE7CFBA    '蓝色
                    .Cell(flexcpBackColor, i, .ColIndex("销帐原因")) = &HE7CFBA    '蓝色
                    .Cell(flexcpBackColor, i, .ColIndex("执行状态")) = Me.BackColor     '灰色
                    
                Else
                    .TextMatrix(i, .ColIndex("NO")) = rsTmp!NO
                    .TextMatrix(i, .ColIndex("发生时间")) = Format(rsTmp!发生时间, "YYYY-MM-DD HH:MM:SS")
                    .TextMatrix(i, .ColIndex("执行科室")) = rsTmp!执行科室
                    .TextMatrix(i, .ColIndex("开单科室")) = rsTmp!开单科室
                    .TextMatrix(i, .ColIndex("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                    .TextMatrix(i, .ColIndex("销帐金额")) = FormatEx(rsTmp!销帐金额, 5)
                    .TextMatrix(i, .ColIndex("销帐原因")) = Nvl(rsTmp!销帐原因)
                    .RowData(i) = Val(rsTmp!ID)
                End If
            Else
                .TextMatrix(i, .ColIndex("NO")) = rsTmp!NO
                .TextMatrix(i, .ColIndex("发生时间")) = Format(rsTmp!发生时间, "YYYY-MM-DD HH:MM:SS")
                .TextMatrix(i, .ColIndex("开单科室")) = rsTmp!开单科室
                .TextMatrix(i, .ColIndex("销帐数量")) = FormatEx(rsTmp!销帐数量, 5)
                .TextMatrix(i, .ColIndex("销帐原因")) = Nvl(rsTmp!销帐原因)
                .RowData(i) = Val(rsTmp!ID)
            End If
            rsTmp.MoveNext
        Next
        mblnUnChange = False
        .Row = 0: .Col = 0
        .Row = 1: .Col = 0
        If mbytFun = E申请 And tbsType.SelectedItem.Key = "T1" Then
            .Col = .ColIndex("销帐数量") '调用事件AfterRowColChange
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cboRemark_Click()
    Dim lngCol As Long
    If zlCommFun.ActualLen(cboRemark.Text) > 200 Then
        MsgBox "录入的销账原因超过100个汉字或者200个字符,请重新录入!", vbInformation, gstrSysName
        cboRemark.SetFocus
        cboRemark.SelStart = 0
        cboRemark.SelLength = Len(cboRemark.Text)
        Exit Sub
    End If
    vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("销帐原因")) = cboRemark.Text
    lngCol = vsfDetail.ColIndex("执行状态")
    mrsApply.Filter = "ID=" & vsfDetail.RowData(vsfDetail.Row) & IIf(lngCol >= 0, " And 执行状态=" & Val(vsfDetail.Cell(flexcpData, vsfDetail.Row, lngCol)), "")
    If mrsApply.RecordCount > 0 Then
        mrsApply!销帐原因 = cboRemark.Text
        mrsApply.Update
    End If
End Sub

Private Sub cboRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If KeyCode = 13 Then
        If zlCommFun.ActualLen(cboRemark.Text) > 200 Then
            MsgBox "录入的销账原因超过100个汉字或者200个字符,请重新录入!", vbInformation, gstrSysName
            cboRemark.SetFocus
            cboRemark.SelStart = 0
            cboRemark.SelLength = Len(cboRemark.Text)
            Exit Sub
        End If
        vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("销帐原因")) = cboRemark.Text
        lngCol = vsfDetail.ColIndex("执行状态")
        mrsApply.Filter = "ID=" & vsfDetail.RowData(vsfDetail.Row) & IIf(lngCol >= 0, " And 执行状态=" & Val(vsfDetail.Cell(flexcpData, vsfDetail.Row, lngCol)), "")
        If mrsApply.RecordCount > 0 Then
            mrsApply!销帐原因 = cboRemark.Text
            mrsApply.Update
        End If
        zlControl.ControlSetFocus vsfDetail
        cboRemark.Visible = False: cboRemark.Tag = ""
        vsfDetail.Select vsfDetail.Row, vsfDetail.ColIndex("销帐原因") - 2
    End If
End Sub

Private Sub cboRemark_LostFocus()
    Dim lngCol As Long
    If mlngPrevRow > vsfDetail.Rows - 1 Then
        cboRemark.Visible = False: cboRemark.Tag = "": Exit Sub
    End If
    If Val(cboRemark.Tag) = Val(vsfDetail.RowData(mlngPrevRow)) Then
        If zlCommFun.ActualLen(cboRemark.Text) > 200 Then
            MsgBox "录入的销账原因超过100个汉字或者200个字符,请重新录入!", vbInformation, gstrSysName
            cboRemark.SetFocus
            cboRemark.SelStart = 0
            cboRemark.SelLength = Len(cboRemark.Text)
            Exit Sub
        End If
        vsfDetail.TextMatrix(mlngPrevRow, vsfDetail.ColIndex("销帐原因")) = cboRemark.Text
        lngCol = vsfDetail.ColIndex("执行状态")
        mrsApply.Filter = "ID=" & vsfDetail.RowData(mlngPrevRow) & IIf(lngCol >= 0, " And 执行状态=" & Val(vsfDetail.Cell(flexcpData, mlngPrevRow, lngCol)), "")
        If mrsApply.RecordCount > 0 Then
            mrsApply!销帐原因 = cboRemark.Text
            mrsApply.Update
        End If
    End If
    If Me.ActiveControl Is cboRemark Then zlControl.ControlSetFocus vsfDetail
    cboRemark.Visible = False: cboRemark.Tag = ""
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errHandler
    If vsfDetail.Col = vsfDetail.ColIndex("销帐原因") Then
        With vsfDetail
            If Val(.RowData(.Row)) = 0 Or Not (mbytFun = E申请 And tbsType.SelectedItem.Key = "T1") Then Exit Sub
        
            cboRemark.Top = vsfDetail.Top + vsfDetail.RowPos(vsfDetail.Row) + 15
            cboRemark.Left = vsfDetail.Left + vsfDetail.ColPos(vsfDetail.ColIndex("销帐原因")) + 15
            cboRemark.Width = vsfDetail.ColWidth(vsfDetail.ColIndex("销帐原因"))
            If mrs停嘱原因.RecordCount <> 0 Then
                mrs停嘱原因.MoveFirst
                cboRemark.Clear
                Do While Not mrs停嘱原因.EOF
                    cboRemark.AddItem Nvl(mrs停嘱原因!名称)
                    mrs停嘱原因.MoveNext
                Loop
            End If
            cboRemark.Text = .TextMatrix(.Row, .ColIndex("销帐原因"))
            mlngPrevRow = .Row
            cboRemark.ZOrder: cboRemark.Visible = True: cboRemark.Tag = vsfDetail.RowData(.Row)
            cboRemark.SetFocus
        End With
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfMain_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim dblTotalNum As Double, i As Long, lngFeeItem As Long, blnDo As Boolean
    Dim dbl销帐金额 As Double
    
    On Error GoTo errHandler
    dblTotalNum = Val(vsfMain(Index).EditText)
    lngFeeItem = Val(vsfMain(Index).RowData(Row))
    
    '按后进先出分配明细
    With mrsApply
        .Filter = "收费细目ID=" & lngFeeItem
        dbl销帐金额 = 0
        If .RecordCount = 0 Then
            MsgBox "数据异常,未能修改明细记录的数量!", vbInformation, gstrSysName
            Exit Sub
        End If
        .Sort = "执行状态,优先 Desc,发生时间 Desc"
        For i = 1 To .RecordCount
            If dblTotalNum = 0 Then
                !销帐数量 = 0
                !销帐金额 = 0
                .Update
            Else
                If Not MCPAR.部分冲销明细 And Not IsNull(mrsInfo!险类) And dblTotalNum < !付数 * !数次 Then
                    If Val(vsfMain(Index).EditText) = dblTotalNum Then
                        MsgBox "不允许对医保病人进行部分冲销明细", vbInformation, gstrSysName
                        vsfMain(Index).TextMatrix(Row, Col) = 0
                        dblTotalNum = 0 '要继续循环,把其它行的冲销明细设为0
                    Else
                        MsgBox "不允许对医保病人进行部分冲销明细,单据[" & !NO & "]不能冲销.", vbInformation, gstrSysName
                        '当前单据不能冲销,但后面的单据可能可以完全冲销.
                    End If
                    !销帐数量 = 0
                    !销帐金额 = 0
                    .Update
                Else
                    blnDo = True
                    If Not IsNull(!结帐ID) Then
                        If CheckBalance(!费用来源, !NO, !序号) Then   '目前未结帐的没有提取出来,下面的程序暂时没有使用
                            If Not IsNull(mrsInfo!险类) Then
                                If Not MCPAR.冲销已结帐单据 Then blnDo = False
                            Else
                                Select Case gbytBillOpt
                                Case 1
                                    If MsgBox("单据[" & !NO & "]中的当前销帐项目已经结帐,确定要申请销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnDo = False
                                Case 2
                                    MsgBox "单据[" & !NO & "]中的当前销帐项目已经结帐,不能销帐！", vbExclamation, gstrSysName
                                    blnDo = False
                                End Select
                            End If
                        End If
                    End If
                    '检查输液配药中心是否处理了未发药部分
                    '问题:?????
                    If InStr(1, "4,5,6,7", Nvl(!收费类别)) > 0 And Val(Nvl(!医嘱序号)) <> 0 Then
                        If Val(Nvl(!执行状态)) = 0 Then  '只有未执行部分才会存在检查
                            If 药品存在配药中心(Val(Nvl(mrsApply!医嘱序号))) Then
                                MsgBox "单据[" & !NO & "]中的当前销帐项目在输液配药中心已经使用了该药品或卫材,不能销帐", vbExclamation, gstrSysName
                               blnDo = False
                            End If
                        End If
                    End If
                    If blnDo Then
                        !销帐数量 = IIf(dblTotalNum <= !付数 * !数次, dblTotalNum, !付数 * !数次)
                        !销帐金额 = Nvl(!销帐数量, 0) * Nvl(!单价, 0)
                        .Update
                        dblTotalNum = dblTotalNum - !销帐数量
                        dbl销帐金额 = dbl销帐金额 + Nvl(!销帐数量, 0) * Nvl(!单价, 0)
                    Else
                        !销帐数量 = 0
                        !销帐金额 = 0
                        .Update
                    End If
                End If
            End If
            .MoveNext
        Next
        
        If dblTotalNum <> 0 Then
            vsfMain(Index).TextMatrix(Row, Col) = Val(vsfMain(Index).EditText) - dblTotalNum
        End If
        vsfMain(Index).TextMatrix(Row, ColApply("销帐金额")) = FormatEx(dbl销帐金额, 5)
    End With
    Call ShowDetail(lngFeeItem)
    Call ShowSumMoney
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckBalance(ByVal byt费用来源 As Byte, ByVal strNO As String, ByVal lngRow As Long) As Boolean
    '功能:检查有结帐ID的某条单据明细,是否已结帐(结帐作废要当成没有结帐)
    '入参:
    '   byt费用来源 0-住院费用记录,1-门诊费用记录
    '返回:True:已结帐,False:未结帐
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTable As String
    
    On Error GoTo errHandler
    strTable = IIf(byt费用来源 = 1, "门诊费用记录", "住院费用记录")
    strSQL = _
        " Select 1" & vbNewLine & _
        " From " & strTable & " A" & vbNewLine & _
        " Where Mod(A.记录性质, 10) = 2 And NO = [1] And Nvl(A.价格父号, A.序号) = [2]" & vbNewLine & _
        " Group By A.NO, Mod(A.记录性质, 10), Nvl(A.价格父号, A.序号)" & vbNewLine & _
        " Having Nvl(Sum(结帐金额),0) = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lngRow)
    CheckBalance = rsTmp.RecordCount = 0
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfMain_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngFeeItem As Long
    
    If mbonNotEnter Or NewRow = 0 Then Exit Sub
    
   With vsfMain(Index)
        If OldRow <> NewRow Then
            lngFeeItem = Val(.RowData(NewRow))
            If lngFeeItem = 0 Then Exit Sub '异常
            Call ShowDetail(lngFeeItem)
        End If
            
        If OldCol <> NewCol Then
            If mbytFun = E申请 And tbsType.SelectedItem.Key = "T1" Then
                If NewCol = ColApply("销帐数量") And Val(.RowData(NewRow)) <> 0 Then
                    .Editable = flexEDKbdMouse
                Else
                    .Editable = flexEDNone
                End If
            End If
        End If
        If mbytFun = E审核 And tbsType.SelectedItem.Key = "T2" Then
            If .TextMatrix(NewRow, ColAudited("状态")) = "×" Then
                cmdOKAudit.Enabled = True
                cmdCancelRefuse.Enabled = True
            Else
                cmdOKAudit.Enabled = False
                cmdCancelRefuse.Enabled = False
            End If
        End If
    End With
End Sub

Private Function SaveRefuse(ByVal blnCancel As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:执行取消拒绝或者审核拒绝操作
    '入参:blnCancel-True表示取消拒绝 False表示审核拒绝
    '返回:成功返回True,失败返回False
    '编制:刘尔旋
    '日期:2014-4-15
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strDate As String, strNos As String
    Dim intRow As Integer, cllPro As Collection
    Dim strMCNO As String
     
    '84026:李南春,2015/4/20，容错处理
    On Error GoTo ErrHand
    With mrsAudited
        intRow = vsfMain(1).Row
        .Filter = "收费细目ID=" & vsfMain(1).RowData(intRow) & _
                " And 申请类别=" & Val(vsfMain(1).Cell(flexcpData, intRow, ColAudited("状态"))) & _
                " And 申请人='" & vsfMain(1).TextMatrix(intRow, ColAudited("申请人")) & "'" & _
                " And 申请时间='" & vsfMain(1).TextMatrix(intRow, ColAudited("申请时间")) & "'"
        If .RecordCount = 0 Then Exit Function
        
        Set cllPro = New Collection
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Do While Not .EOF
            'Zl_病人费用销帐_Cancel
            strSQL = "Zl_病人费用销帐_Cancel("
            '  Id_In       病人费用销帐.费用id%Type,
            strSQL = strSQL & "" & !ID & ","
            '  申请时间_In 病人费用销帐.申请时间%Type,
            strSQL = strSQL & "To_Date('" & !申请时间 & "','YYYY-MM-DD HH24:MI:SS'),"
            '  审核人_In   病人费用销帐.审核人%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  审核时间_In 病人费用销帐.审核时间%Type,
            strSQL = strSQL & "" & strDate & ","
            '  操作类型_In Number, --操作类型_IN:0-审核拒绝 1-取消拒绝
            strSQL = strSQL & "" & IIf(blnCancel, "1", "0") & ","
            '  Int自动退料 Integer := 1,
            strSQL = strSQL & "" & "1" & ","
            '  申请类别_In 病人费用销帐.申请类别%Type := 1
            strSQL = strSQL & "" & Val(vsfMain(1).Cell(flexcpData, intRow, ColAudited("状态"))) & ")"
            zlAddArray cllPro, strSQL
        
            If Not blnCancel Then
                If Val(Nvl(!费用来源)) = 0 Then
                    'Zl_住院记帐记录_Delete
                    strSQL = "ZL_住院记帐记录_Delete("
                    '  No_In           住院费用记录.No%Type,
                    strSQL = strSQL & "'" & Nvl(!NO) & "',"
                    '  序号_In         Varchar2,
                    strSQL = strSQL & "'" & Val(Nvl(!序号)) & ":" & Val(Nvl(!售价销帐数量)) & "',"
                    '  操作员编号_In   住院费用记录.操作员编号%Type,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '  操作员姓名_In   住院费用记录.操作员姓名%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '  记录性质_In     住院费用记录.记录性质%Type := 2,
                    strSQL = strSQL & "" & Val(Nvl(!记录性质)) & ","
                    '  操作状态_In     Number := 0,--0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
                    strSQL = strSQL & "" & "1" & ")"
                    zlAddArray cllPro, strSQL
                Else
                    'Zl_门诊记帐记录_Delete
                    strSQL = "Zl_门诊记帐记录_Delete("
                    '  No_In         门诊费用记录.No%Type,
                    strSQL = strSQL & "'" & Nvl(!NO) & "',"
                    '  序号_In       Varchar2,
                    strSQL = strSQL & "'" & Val(Nvl(!序号)) & ":" & Val(Nvl(!售价销帐数量)) & "',"
                    '  操作员编号_In 门诊费用记录.操作员编号%Type,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '  操作员姓名_In 门诊费用记录.操作员姓名%Type
                    strSQL = strSQL & "'" & UserInfo.姓名 & "')"
                    zlAddArray cllPro, strSQL
                End If
                
                If Not IsNull(!险类) And InStr("," & strMCNO & ",", "," & !NO & ",") = 0 Then
                    MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                    MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                    strMCNO = "|" & !NO & "," & !险类 & "," & _
                        IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                End If
                
                If InStr("," & strNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                    '单据操作时间限制检查
                    If Not BillOperCheck(IIf(Val(Nvl(!费用来源)) = 0, 5, 4), _
                        Nvl(!操作员姓名), Format(Nvl(!登记时间), "YYYY-MM-DD HH:MM:SS"), _
                        "销帐审核", Nvl(!NO), , 2, , False, False) Then Exit Function
                    strNos = strNos & "," & Nvl(!NO)
                End If
            End If
            
            .MoveNext
        Loop
        If strMCNO <> "" Then strMCNO = Mid(strMCNO, 2)
            
        If ExecuteDataSave(cllPro, strMCNO) = False Then Exit Function
    End With
    
    Call cmdRefresh_Click
    SaveRefuse = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 药品存在配药中心(ByVal lng医嘱ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查医嘱开的药品是否已经在配置中心使用了
    '返回：成在返回True,否则返回False
    '编制：刘兴洪
    '日期：2010-07-29 14:55:19
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = _
        " Select 1 " & _
        " From 病人医嘱记录 A, 病人医嘱发送 B, 输液配药记录 D " & _
        " Where A.相关id = B.医嘱id And B.医嘱id = D.医嘱id And B.发送号 = D.发送号 And A.ID = [1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    药品存在配药中心 = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub vsfMain_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    For i = 1 To vsfMain(Index).Rows - 1
        If mlngPreFeeItemID = vsfMain(Index).RowData(i) Then vsfMain(Index).Row = i
    Next
End Sub

Private Sub vsfMain_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mlngPreFeeItemID = vsfMain(Index).RowData(vsfMain(Index).Row)
End Sub

Private Sub vsfMain_DblClick(Index As Integer)
    Dim i As Long, strResult As String, intState As Integer
    
    If mbytFun = E申请 And tbsType.SelectedItem.Key = "T2" Then
        With vsfMain(Index)
            If .Col = 0 Then
                intState = Val(cboState.ItemData(cboState.ListIndex))
                If intState = ESTATE.E未审核 Then
                    If .MouseRow = 0 Then
                        If .ColData(ColApplied("选择")) = "" Then
                            .ColData(ColApplied("选择")) = "√"
                        Else
                            .ColData(ColApplied("选择")) = ""
                        End If
                        strResult = .ColData(ColApplied("选择"))
                        For i = 1 To .Rows - 1
                            .TextMatrix(i, ColApplied("选择")) = strResult
                        Next
                    Else
                        If .TextMatrix(.Row, ColApplied("选择")) = "√" Then
                            .TextMatrix(.Row, ColApplied("选择")) = ""
                        Else
                            .TextMatrix(.Row, ColApplied("选择")) = "√"
                        End If
                    End If
                End If
            End If
        End With
        
    ElseIf mbytFun = E审核 And tbsType.SelectedItem.Key = "T1" Then
        With vsfMain(Index)
            If .Col = 0 Then
                If .MouseRow = 0 Then
                    If .ColData(ColAudit("审核")) = "" Then
                        .ColData(ColAudit("审核")) = "√"
                    Else
                        .ColData(ColAudit("审核")) = ""
                    End If
                    strResult = .ColData(ColAudit("审核"))
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, ColAudit("审核")) = strResult
                        If strResult = "√" Then
                            If Not CheckCanAudit(.RowData(i), .TextMatrix(i, ColAudited("姓名")), .TextMatrix(i, ColAudited("申请人")), .TextMatrix(i, ColAudited("申请时间"))) Then .TextMatrix(i, ColAudit("审核")) = ""
                        End If
                    Next
                Else
                    Select Case Trim(.TextMatrix(.Row, ColAudit("审核")))
                        Case "√"
                            .TextMatrix(.Row, ColAudit("审核")) = "×"
                        Case "×"
                            .TextMatrix(.Row, ColAudit("审核")) = ""
                        Case ""
                            If CheckCanAudit(.RowData(.Row), .TextMatrix(.Row, ColAudited("姓名")), .TextMatrix(.Row, ColAudited("申请人")), .TextMatrix(.Row, ColAudited("申请时间"))) Then
                                .TextMatrix(.Row, ColAudit("审核")) = "√"
                            Else
                                .TextMatrix(.Row, ColAudit("审核")) = ""
                            End If
                    End Select
                End If
            End If
        End With
    End If
End Sub

Private Function CheckCanAudit(ByVal lngFeeItemID As Long, ByVal strPatient As String, ByVal strOperater As String, ByVal strDate As String) As Boolean
'功能:检查待审核的费用项目的单据明细行是否已结帐,已结帐的不允许审核(冲销)
    Dim i As Long
    
    '问题:29613
    If mrsAudit Is Nothing Then Exit Function
    If mrsAudit.State <> 1 Then Exit Function
    
    CheckCanAudit = True
    With mrsAudit
        .Filter = "收费细目id=" & lngFeeItemID & " And 姓名='" & strPatient & "'" & _
                " And 申请人='" & strOperater & "' And 申请时间='" & strDate & "'"
        For i = 1 To .RecordCount
            If Not IsNull(!结帐ID) Then
                If CheckBalance(!费用来源, !NO, !序号) Then
                    If Not IsNull(!险类) Then
                        If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , Val(!险类)) Then
                            MsgBox "不允许冲销医保病人已结帐的单据[" & !NO & "].", vbInformation, gstrSysName
                            CheckCanAudit = False
                            Exit For
                        End If
                    Else
                        Select Case gbytBillOpt
                        Case 1
                            If MsgBox("单据[" & !NO & "]中的当前销帐项目已经结帐,确定要申请销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then CheckCanAudit = False: Exit For
                        Case 2
                            MsgBox "单据[" & !NO & "]中的当前销帐项目已经结帐,不能销帐！", vbInformation, gstrSysName
                            CheckCanAudit = False
                            Exit For
                        End Select
                    End If
                End If
            End If
            .MoveNext
        Next
    End With
End Function

Private Sub vsfMain_EnterCell(Index As Integer)
    With vsfMain(Index)
        .BackColorSel = .CellBackColor
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub vsfMain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsfMain(Index)
            If Val(.RowData(.Row)) = 0 Or Not (mbytFun = E申请 And tbsType.SelectedItem.Key = "T1") Then Exit Sub
                        
            If .Col = ColApply("销帐数量") Then
                .TextMatrix(.Row, .Col) = 0
            End If
        End With
    End If
End Sub

Private Sub vsfMain_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsfMain(Index)
            KeyAscii = 0
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If cmdOKApply.Visible And cmdOKApply.Enabled Then cmdOKApply.SetFocus
            End If
        End With
    End If
End Sub

Private Sub vsfMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsfMain(Index)
        If mbytFun = E审核 And tbsType.SelectedItem.Key = "T1" Then
            If .MouseCol = 0 And .MouseRow = 0 Then
                .ToolTipText = "双击全选,再次双击全部取消."
            Else
                .ToolTipText = ""
            End If
        End If
    End With
End Sub

Private Sub vsfMain_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsfMain(Index).EditSelStart = 0
    vsfMain(Index).EditSelLength = Len(vsfMain(Index).EditText)
End Sub

Private Sub vsfMain_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsfMain(Index)
        If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
        If Val(.EditText) > Val(.TextMatrix(Row, ColApply("数量"))) Then
            stbThis.Panels(2).Text = "申请数量不能大于可销帐数量!"
            Cancel = True
        Else
            stbThis.Panels(2).Text = ""
        End If
    End With
End Sub

Private Sub vsfMain_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub


Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset
    
    If mblnUnChange Then Exit Sub
    With vsfDetail
        If OldCol <> NewCol Then
            .Editable = flexEDNone
            If Val(.RowData(NewRow)) = 0 Or Not (mbytFun = E申请 And tbsType.SelectedItem.Key = "T1") Then Exit Sub
            
            If NewCol = .ColIndex("销帐数量") Then .Editable = flexEDKbdMouse
            If NewCol = .ColIndex("销帐原因") Then .Editable = flexEDKbdMouse
        End If
        If OldRow <> NewRow Then
            If cboRemark.Visible And OldRow < vsfDetail.Rows Then
                If Val(cboRemark.Tag) = Val(vsfDetail.RowData(OldRow)) Then
                    Dim lngCol As Long
                    If zlCommFun.ActualLen(cboRemark.Text) > 200 Then
                        MsgBox "录入的销账原因超过100个汉字或者200个字符,请重新录入!", vbInformation, gstrSysName
                        cboRemark.SetFocus
                        cboRemark.SelStart = 0
                        cboRemark.SelLength = Len(cboRemark.Text)
                        Exit Sub
                    End If
                    vsfDetail.TextMatrix(OldRow, vsfDetail.ColIndex("销帐原因")) = cboRemark.Text
                    lngCol = vsfDetail.ColIndex("执行状态")
                    mrsApply.Filter = "ID=" & vsfDetail.RowData(OldRow) & IIf(lngCol >= 0, " And 执行状态=" & Val(vsfDetail.Cell(flexcpData, OldRow, lngCol)), "")
                    If mrsApply.RecordCount > 0 Then
                        mrsApply!销帐原因 = cboRemark.Text
                        mrsApply.Update
                    End If
                End If
                cboRemark.Visible = False: cboRemark.Tag = ""
            End If
            vsfTogether.Visible = False
            If Val(.RowData(NewRow)) <> 0 And mbytFun = E申请 Then
                If tbsType.SelectedItem.Key = "T1" Then
                    Set rsTmp = mrsApply
                Else
                    Set rsTmp = mrsApplied
                End If
                rsTmp.Filter = "ID=" & Val(.RowData(NewRow))    '注意,它会改变原记录集的Filter
                If InStr(1, ",5,6,7,", "," & rsTmp!收费类别 & ",") > 0 And Not IsNull(rsTmp!医嘱序号) Then
                    '显示一并给药情况
                    Call ShowTogetherMedi(Val(rsTmp!医嘱序号), Val(.RowData(NewRow)))
                End If
            End If
            Call Form_Resize
        End If
    End With
End Sub

Private Sub ShowTogetherMedi(ByVal lngAdviceID As Long, ByVal lngFeeItemID As Long)
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    vsfTogether.Clear
    vsfTogether.Rows = 1
    vsfTogether.TextMatrix(0, 0) = "一并给药药品"
 
    strSQL = "Select 1" & vbNewLine & _
            "From 住院费用记录 A, 住院费用记录 B" & vbNewLine & _
            "Where A.ID = [1] And A.医嘱序号 is Not Null And A.NO = B.NO And A.记录性质 = B.记录性质" & _
            "      And A.记录状态 = B.记录状态 And A.执行状态 = B.执行状态 " & vbNewLine & _
            "      And A.收费细目id = B.收费细目id And A.登记时间 = B.登记时间 Having Count(A.ID) > 1"
    strSQL = strSQL & " Union All " & Replace(strSQL, "住院费用记录", "门诊费用记录")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFeeItemID)
    If rsTmp.RecordCount > 0 Then
        strSQL = "Select B.医嘱内容 From 病人医嘱记录 A, 病人医嘱记录 B" & vbNewLine & _
                "Where A.ID = [1] And A.相关id = B.相关id And A.ID <> B.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
        If rsTmp.RecordCount > 0 Then
            Set vsfTogether.DataSource = rsTmp
            vsfTogether.TextMatrix(0, 0) = "一并给药药品"
        End If
    End If
    vsfTogether.Visible = vsfTogether.Rows > 1
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetail_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsfDetail.EditSelStart = 0
    vsfDetail.EditSelLength = Len(vsfDetail.EditText)
End Sub

Private Sub vsfDetail_EnterCell()
    With vsfDetail
        .BackColorSel = .CellBackColor
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub vsfDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblTotal As Double
    
    With vsfDetail
        If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
        dblTotal = Val(.TextMatrix(Row, .ColIndex("付数")) * .TextMatrix(Row, .ColIndex("数次")))
        If Val(.EditText) > dblTotal Then
            stbThis.Panels(2).Text = "申请数量不能大于可销帐数量!"
            Cancel = True
        Else
            stbThis.Panels(2).Text = ""
            If Val(.EditText) < dblTotal And Val(.EditText) <> 0 Then
                If Not MCPAR.部分冲销明细 And Not IsNull(mrsInfo!险类) Then
                    stbThis.Panels(2).Text = "不允许对医保病人进行部分冲销明细."
                    Cancel = True
                    Exit Sub
                End If
            End If
            If .ColIndex("执行状态") < 0 Then
                mrsApply.Filter = "ID=" & .RowData(Row)
            Else
                mrsApply.Filter = "ID=" & .RowData(Row) & " And 执行状态=" & Val(.Cell(flexcpData, .Row, .ColIndex("执行状态")))
            End If
            If mrsApply.RecordCount > 0 Then
                '检查输液配药中心是否处理了未发药部分
                '问题:?????
                If InStr(1, "4,5,6,7", Nvl(mrsApply!收费类别)) > 0 And Val(Nvl(mrsApply!医嘱序号)) <> 0 And .ColIndex("执行状态") >= 0 Then
                    If Val(.Cell(flexcpData, .Row, .ColIndex("执行状态"))) = 0 Then '只有未执行部分才会存在检查
                        If 药品存在配药中心(Val(Nvl(mrsApply!医嘱序号))) Then
                            stbThis.Panels(2).Text = "输液配药中心已经使用了该药品或卫材."
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                If Not IsNull(mrsApply!结帐ID) Then '目前未结帐的没有提取出来,下面的程序暂时没有使用
                    If CheckBalance(mrsApply!费用来源, mrsApply!NO, mrsApply!序号) Then
                        If Not IsNull(mrsInfo!险类) Then
                            If Not MCPAR.冲销已结帐单据 Then
                                stbThis.Panels(2).Text = "不允许冲销医保病人已结帐的单据."
                                Cancel = True
                            End If
                        Else
                            Select Case gbytBillOpt
                            Case 1
                                If MsgBox("单据[" & mrsApply!NO & "]中的当前销帐项目已经结帐,确定要申请销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
                            Case 2
                                stbThis.Panels(2).Text = "单据[" & mrsApply!NO & "]中的当前销帐项目已经结帐,不能销帐！"
                                Cancel = True
                            End Select
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       With vsfDetail
            If .Row < .Rows - 1 Then KeyAscii = 0: .Row = .Row + 1
       End With
    End If
End Sub


Private Sub vsfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsfDetail
            If Val(.RowData(.Row)) = 0 Or Not (mbytFun = E申请 And tbsType.SelectedItem.Key = "T1") Then Exit Sub
            
            If .Col = .ColIndex("销帐数量") Then
                .TextMatrix(.Row, .Col) = "0"
            End If
        End With
    End If
End Sub

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblLack As Double
    Dim lngCol As Long
    Dim dblMny As Double
    
    lngCol = vsfDetail.ColIndex("执行状态")
    mrsApply.Filter = "ID=" & vsfDetail.RowData(Row) & IIf(lngCol >= 0, " And 执行状态=" & Val(vsfDetail.Cell(flexcpData, Row, lngCol)), "")
    If mrsApply.RecordCount > 0 Then
        dblLack = Val(vsfDetail.EditText) - mrsApply!销帐数量
        dblMny = (Val(vsfDetail.EditText) - mrsApply!销帐数量) * mrsApply!单价
        
        mrsApply!销帐数量 = vsfDetail.EditText
        mrsApply!销帐金额 = mrsApply!销帐数量 * mrsApply!单价
        mrsApply.Update
        vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("销帐金额")) = FormatEx(mrsApply!销帐金额, 5)
        vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("销帐数量")) = vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("销帐数量")) + dblLack
        vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("销帐金额")) = vsfMain(0).TextMatrix(vsfMain(0).Row, ColApply("销帐金额")) + dblMny
        Call ShowSumMoney
    End If
End Sub
Private Sub cmdCancelApply_Click()
    Call SaveData
    gblnOK = True
End Sub

Private Sub cmdOKApply_Click()
    If mlngDeptID = 0 Then
        MsgBox "没有选择申请部门, 不能确认申请!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If SaveData = False Then Exit Sub
    '问题:26551
    'gblnOK = True
End Sub

Private Sub cmdOKAudit_Click()
    If mbytFun = E审核 And tbsType.SelectedItem.Key = "T2" Then
        If SaveRefuse(False) = False Then Exit Sub
    Else
        If SaveData = False Then Exit Sub
        gblnOK = True
    End If
End Sub

Private Sub cmdCancelRefuse_Click()
    If SaveRefuse(True) = False Then Exit Sub
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, cllPro As Collection
    Dim i As Long, lngTmp As Long, strMCNO As String
    Dim strDate As String, str费用IDs As String, str费用ID As String, strTmp As String
    Dim dbl数量 As Double, strNos As String
    Dim str审核费用ID As String, strMsgDate As String
    Dim strUserDeptIDs As String, str收费细目IDs As String, strKey费用IDs As String
    
    On Error GoTo errHandler
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            strUserDeptIDs = "," & GetUserDeptIDs & ","
            With mrsApply
                If .State = 0 Then Exit Function
                .Filter = ""
                For i = 1 To .RecordCount
                    If !销帐数量 <> !原始销帐数量 Or Nvl(!销帐原因) <> Nvl(!原始销帐原因) Then
                        str费用IDs = str费用IDs & "," & !ID
                        str审核费用ID = str审核费用ID & "," & !ID
                        If InStr(strUserDeptIDs, "," & !开单部门ID & ",") = 0 Then
                            If InStr(str收费细目IDs & ",", "," & !收费细目ID & ",") = 0 Then
                                If MsgBox("单据号:" & Nvl(!NO) & "中的收费项目为:" & GetItemName(!收费细目ID) & "不是自己所属的开单部门,是否还要进行申请销帐?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
                                str收费细目IDs = str收费细目IDs & "," & !收费细目ID
                            End If
                        End If
                    End If
                    .MoveNext
                Next
                If str费用IDs <> "" Then str费用IDs = Mid(str费用IDs, 2)
                If str审核费用ID <> "" Then str审核费用ID = Mid(str审核费用ID, 2)
                
                If str费用IDs = "" Then
                    stbThis.Panels(2).Text = "所有记录都没有填写销帐数量或者原因!"
                    Exit Function
                End If
            End With
        Else
            For i = 1 To vsfMain(1).Rows - 1
                If vsfMain(1).TextMatrix(i, ColApplied("选择")) = "√" And Val(vsfMain(1).RowData(i)) <> 0 Then Exit For
            Next
            If i > vsfMain(1).Rows - 1 Then
                stbThis.Panels(2).Text = "没有选择取消申请的记录!请在要取消申请的记录的""选择""列上双击。"
                Exit Function
            End If
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            For i = 1 To vsfMain(0).Rows - 1
                If Val(vsfMain(0).RowData(i)) <> 0 Then
                    If vsfMain(0).TextMatrix(i, ColAudit("审核")) = "√" Or vsfMain(0).TextMatrix(i, ColAudit("审核")) = "×" Then Exit For
                End If
            Next
            If i > vsfMain(0).Rows - 1 Then
                stbThis.Panels(2).Text = "没有选择审核的记录!请在要审核的记录的""审核""列上双击。"
                Exit Function
            End If
        End If
    End If
    
    Set cllPro = New Collection
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            strMsgDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            strDate = "To_Date('" & strMsgDate & "','YYYY-MM-DD HH24:MI:SS')"
            
            With mrsApply
                .Filter = ""
                strKey费用IDs = ""
                Do While Not .EOF
                    If InStr(1, "," & str费用IDs & ",", "," & !ID & ",") > 0 Then
                        dbl数量 = !销帐数量 * !住院包装
                        If !销帐数量 = !数次 Then dbl数量 = !售价数次
                        
                        'Zl_病人费用销帐_Insert
                        strSQL = "Zl_病人费用销帐_Insert("
                        '  Id_In         In 病人费用销帐.费用id%Type,
                        strSQL = strSQL & "" & !ID & ","
                        '  收费细目id_In In 病人费用销帐.收费细目id%Type,
                        strSQL = strSQL & "" & !收费细目ID & ","
                        '  申请部门id_In In 病人费用销帐.申请部门id%Type,
                        strSQL = strSQL & "" & mlngDeptID & ","
                        '  数量_In       In 病人费用销帐.数量%Type,
                        strSQL = strSQL & "" & dbl数量 & ","
                        '  申请人_In     In 病人费用销帐.申请人%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '  申请时间_In   In 病人费用销帐.申请时间%Type,
                        strSQL = strSQL & "" & strDate & ","
                        '  申请类别_In   In 病人费用销帐.申请类别%Type,--对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
                        strSQL = strSQL & "" & Val(Nvl(!执行状态)) & ","
                        '  删除标志_In   In Integer := 0,--删除病人费用销帐时的条件:1-删除时不管申请类别,0-删除时,根据申请类别来进行删除(因为可能出现在申请销帐时,存在已执行和未执行两种状态)
                        strSQL = strSQL & "" & IIf(InStr(1, "," & strKey费用IDs & ",", "," & Nvl(!ID) & ",") > 0, 0, 1) & ","
                        '  配药id_In     In Integer := 0,
                        strSQL = strSQL & "" & "0" & ","
                        '  销帐原因_In   In 病人费用销帐.销帐原因%Type := Null,
                        strSQL = strSQL & "'" & Nvl(!销帐原因) & "')"
                        '  配液更新_In   In Number := 1--是否 输液配药记录 状态字段。1-要更新，0-不更新
                        zlAddArray cllPro, strSQL
                        
                        strKey费用IDs = strKey费用IDs & "," & !ID
                        If InStr("," & strNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                            '单据操作时间限制检查
                            If Not BillOperCheck(IIf(Val(Nvl(!费用来源)) = 0, 5, 4), _
                                Nvl(!操作员姓名), Format(Nvl(!登记时间), "YYYY-MM-DD HH:MM:SS"), _
                                "销帐申请", Nvl(!NO), , 2, , False, False) Then Exit Function
                            strNos = strNos & "," & Nvl(!NO)
                        End If
                    End If
                    .MoveNext
                Loop
            End With
        Else
            With mrsApplied
                For i = 1 To vsfMain(1).Rows - 1
                    If vsfMain(1).TextMatrix(i, ColApplied("选择")) = "√" Then
                        .Filter = "收费细目ID=" & vsfMain(1).RowData(i) & _
                                " And 申请人='" & vsfMain(1).TextMatrix(i, ColApplied("申请人")) & "'" & _
                                " And 申请时间='" & vsfMain(1).TextMatrix(i, ColApplied("申请时间")) & "'"
                        Do While Not .EOF
                            str费用IDs = str费用IDs & "," & !ID
                            .MoveNext
                        Loop
                    End If
                Next
                If str费用IDs <> "" Then str费用IDs = Mid(str费用IDs, 2)
            End With
            
            While str费用IDs <> ""
                str费用IDs = str费用IDs & ","
                If Len(str费用IDs) > 3998 Then
                    lngTmp = InStrRev(Mid(str费用IDs, 1, 3998), ",")
                    str费用ID = Mid(str费用IDs, 1, lngTmp - 1)
                    str费用IDs = Mid(str费用IDs, lngTmp + 1)
                Else
                    str费用ID = Mid(str费用IDs, 1, Len(str费用IDs) - 1)
                    str费用IDs = ""
                End If
                
                strSQL = "zl_病人费用销帐_Delete('" & str费用ID & "')"
                zlAddArray cllPro, strSQL
            Wend
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            With mrsAudit
                For i = 1 To vsfMain(0).Rows - 1
                    strTmp = vsfMain(0).TextMatrix(i, ColAudit("审核"))
                    If strTmp = "√" Or strTmp = "×" Then
                        .Filter = "收费细目ID=" & vsfMain(0).RowData(i) & _
                                " And 申请类别=" & Val(vsfMain(0).Cell(flexcpData, i, ColAudit("审核"))) & _
                                " And 申请人='" & vsfMain(0).TextMatrix(i, ColAudit("申请人")) & "'" & _
                                " And 申请时间='" & vsfMain(0).TextMatrix(i, ColAudit("申请时间")) & "'"
                        
                        Do While Not .EOF
                            If zlCheckFeeIsValied(Val(Nvl(!费用来源)), Val(Nvl(!ID)), _
                                Val(Nvl(!审核部门id)), Val(vsfMain(0).Cell(flexcpData, i, ColAudit("审核")))) = False Then Exit Function
                            
                            'Zl_病人费用销帐_Audit
                            strSQL = "Zl_病人费用销帐_Audit("
                            '  Id_In       病人费用销帐.费用id%Type,
                            strSQL = strSQL & "" & Val(Nvl(!ID)) & ","
                            '  申请时间_In 病人费用销帐.申请时间%Type,
                            strSQL = strSQL & "To_Date('" & Format(Nvl(!申请时间), "yyyy-mm-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                            '  审核人_In   病人费用销帐.审核人%Type,
                            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                            '  审核时间_In 病人费用销帐.审核时间%Type,
                            strSQL = strSQL & "" & strDate & ","
                            '  状态_In     病人费用销帐.状态%Type,--1-审核通过,2-审核未通过
                            strSQL = strSQL & "" & IIf(strTmp = "√", "1", "2") & ","
                            '  Int自动退料 Integer := 1,
                            strSQL = strSQL & "" & "1" & ","
                            '  申请类别_In 病人费用销帐.申请类别%Type := 1--对药品和卫材有效,缺省为已执行的药品或卫材
                            strSQL = strSQL & "" & Val(vsfMain(0).Cell(flexcpData, i, ColAudit("审核"))) & ")"
                            zlAddArray cllPro, strSQL
                                    
                            If strTmp = "√" Then
                                If Val(Nvl(!费用来源)) = 0 Then
                                    'Zl_住院记帐记录_Delete
                                    strSQL = "ZL_住院记帐记录_Delete("
                                    '  No_In           住院费用记录.No%Type,
                                    strSQL = strSQL & "'" & Nvl(!NO) & "',"
                                    '  序号_In         Varchar2,
                                    strSQL = strSQL & "'" & Val(Nvl(!序号)) & ":" & Val(Nvl(!售价销帐数量)) & "',"
                                    '  操作员编号_In   住院费用记录.操作员编号%Type,
                                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                                    '  操作员姓名_In   住院费用记录.操作员姓名%Type,
                                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                                    '  记录性质_In     住院费用记录.记录性质%Type := 2,
                                    strSQL = strSQL & "" & Val(Nvl(!记录性质)) & ","
                                    '  操作状态_In     Number := 0,--0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
                                    strSQL = strSQL & "" & "1" & ")"
                                    zlAddArray cllPro, strSQL
                                Else
                                    'Zl_门诊记帐记录_Delete
                                    strSQL = "Zl_门诊记帐记录_Delete("
                                    '  No_In         门诊费用记录.No%Type,
                                    strSQL = strSQL & "'" & Nvl(!NO) & "',"
                                    '  序号_In       Varchar2,
                                    strSQL = strSQL & "'" & Val(Nvl(!序号)) & ":" & Val(Nvl(!售价销帐数量)) & "',"
                                    '  操作员编号_In 门诊费用记录.操作员编号%Type,
                                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                                    '  操作员姓名_In 门诊费用记录.操作员姓名%Type
                                    strSQL = strSQL & "'" & UserInfo.姓名 & "')"
                                    zlAddArray cllPro, strSQL
                                End If
                            End If
                                    
                            If Not IsNull(!险类) And InStr("," & strMCNO & ",", "," & !NO & ",") = 0 Then
                                MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                                MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                                strMCNO = "|" & !NO & "," & !险类 & "," & _
                                    IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                            End If
                            
                            If InStr("," & strNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                                '单据操作时间限制检查
                                If Not BillOperCheck(IIf(Val(Nvl(!费用来源)) = 0, 5, 4), _
                                    Nvl(!操作员姓名), Format(Nvl(!登记时间), "YYYY-MM-DD HH:MM:SS"), _
                                    "销帐审核", Nvl(!NO), , 2, , False, False) Then Exit Function
                                strNos = strNos & "," & Nvl(!NO)
                            End If
                            
                            .MoveNext
                        Loop
                        
                    End If
                Next
                If strMCNO <> "" Then strMCNO = Mid(strMCNO, 2)
            End With
        End If
    End If
    
    If ExecuteDataSave(cllPro, strMCNO) = False Then Exit Function
    
    '问题:34994
    '   进行审核操作
    If mbytFun = E申请 And chkVerfy.Visible And chkVerfy.Value = 1 Then
        If zlApplyToVerify(str审核费用ID) = False Then
            MsgBox "注意:" & vbCrLf & "    存在不能审核的申请,请通过销帐审核进行审核操作!", vbInformation + vbOKOnly, gstrSysName
        End If
    End If
    
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            '消息发送
            If Not (chkVerfy.Visible And chkVerfy.Value = 1) Then
                Call SendMsgModule(str费用IDs, strMsgDate)
            End If
            txtPatient.Text = "": txtPatient.SetFocus
            Call ClearPatientInfo
        Else
            Call cmdRefresh_Click
        End If
    Else
        If tbsType.SelectedItem.Key = "T1" Then
            Call cmdRefresh_Click
        End If
    End If
    
    stbThis.Panels(2).Text = "保存数据成功!"
    SaveData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteDataSave(cllPro As Collection, ByVal strMCNO As String) As Boolean
    '执行数据保存
    '入参：
    '   cllPro 需要保存的SQL
    '   strMCNO 医保数据上传信息，格式：NO,险类,记帐作废上传,记帐完成后上传|...
    Dim arrMCRec As Variant, i As Integer
    Dim arrMCPar As Variant
    
    On Error GoTo errHandler
    Screen.MousePointer = 11
        
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '医保，记帐作废上传，作废时上传
    arrMCRec = Split(strMCNO, "|")
    For i = 0 To UBound(arrMCRec)
        arrMCPar = Split(arrMCRec(i), ",")
        If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                gcnOracle.RollbackTrans: Exit Function
            End If
        End If
    Next
    gcnOracle.CommitTrans
    
    '医保，记帐作废上传，完成后上传
    For i = 0 To UBound(arrMCRec)
        arrMCPar = Split(arrMCRec(i), ",")
        If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
            End If
        End If
    Next
    Screen.MousePointer = 0
    ExecuteDataSave = True
    Exit Function
errHandler:
    gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
'        Resume
    End If
    Call SaveErrLog
End Function


Private Function zlCheckFeeIsValied(ByVal byt费用来源 As Byte, ByVal lng费用ID As Long, _
    ByVal lng审核部门ID As Long, Optional int申请类别 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查费用销帐是否有效
    '入参:
    '   byt费用来源 0-住院费用记录,1-门诊费用记录
    '   int申请类别-1-已执行;0-未执行
    '出参:
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-07-28 09:48:59
    '问题:24597
    '规则:1.如果当明费用未被执行，则与原来的规则不变，谁开单，谁就可以进行销帐
    '     2.如果当明费用被执行,则需要判断如下情况:
    '        a.如果审核科室与执行科室相等,则允许审核确认
    '        b.如果审核科室与执行室不相等，则需要检查执行科室是否在当前操作员人员所属病区或科室,如果是，则允许审核，否则不允许审核
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lng执行部门ID As Long
    Dim strNO As String, str部门名称 As String, rsDept As ADODB.Recordset
    Dim strSQL As String, strTable As String
    
    On Error GoTo errHandle
    '之所以要重新读取执行状态，因考虑并发操作这种情况
    strTable = IIf(byt费用来源 = 1, "门诊费用记录", "住院费用记录")
    strSQL = _
        " Select a.No, a.执行状态, a.执行部门id, a.收费细目id, a.收费类别, Nvl(b.跟踪在用, 0) As 跟踪在用" & _
        " From " & strTable & " A, 材料特性 B" & _
        " Where a.收费细目id = b.材料id(+) And a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng费用ID)
    If rsTemp.EOF Then
        MsgBox "注意:" & vbCrLf & _
               "    审核中至少一条明细费用不存在，可能被他人删除，请刷新后再试！", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    lng执行部门ID = Val(Nvl(rsTemp!执行部门ID))
    '1.如果当明费用未被执行，则与原来的规则不变，谁开单，谁就可以进行销帐
    '记录状态=1,3时：0:未执行;1:完全执行;2:部份执行；记录状态=2时：-x:第x次退费
    If Val(Nvl(rsTemp!执行状态)) = 0 Then zlCheckFeeIsValied = True: Exit Function
    
    '如果当明费用被执行,则需要判断如下情况:
    '1. 如果审核科室与执行科室相等,则允许审核确认
    If lng审核部门ID = lng执行部门ID Then zlCheckFeeIsValied = True: Exit Function
    '2  如果审核科室与执行室不相等，则需要检查执行科室是否在当前操作员人员所属病区或科室,如果是，则允许审核，否则不允许审核
    If InStr(1, "," & mstrUnitIDs & ",", "," & lng执行部门ID & ",") > 0 Then zlCheckFeeIsValied = True: Exit Function
    strNO = Nvl(rsTemp!NO)
    
    '3.如果是药品,卫材,需要检查
    If InStr(1, "5,6,7", Nvl(rsTemp!收费类别)) > 0 Or (Nvl(rsTemp!收费类别) = 4 And Nvl(rsTemp!跟踪在用) = "1") Then
        If int申请类别 = 0 Then
            zlCheckFeeIsValied = True: Exit Function
        End If
    End If
    
    strSQL = "Select 编码,名称 From 部门表 a Where id=[1]"
    Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng执行部门ID)
    If Not rsDept.EOF Then str部门名称 = Nvl(rsDept!编码) & "-" & Nvl(rsDept!名称)
    MsgBox "注意:" & vbCrLf & _
           "    单据号为“" & strNO & "”" & vbCrLf & _
           "    收费项目为“" & GetItemName(Val(Nvl(rsTemp!收费细目ID))) & "”" & vbCrLf & _
           "    已经被“" & str部门名称 & "” 执行，不能确认销帐！", vbInformation + vbDefaultButton1, gstrSysName

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowSumMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示本次销帐总额
    '编制:刘兴洪
    '日期:2011-02-15 16:57:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, vsGrid As VSFlexGrid, lngRow As Long
    Dim lngCol As Long
    Err = 0: On Error Resume Next
    If mbytFun = E申请 Then
        If tbsType.SelectedItem.Key = "T1" Then
            Set vsGrid = vsfMain(0): lngCol = ColApply("销帐金额")
        Else
            Set vsGrid = vsfMain(1): lngCol = ColApplied("销帐金额")
        End If
        With vsGrid
            For lngRow = .FixedRows To .Rows - 1
                dblMoney = dblMoney + Val(.TextMatrix(lngRow, lngCol))
            Next
        End With
        picHsc.Height = 435
        picHsc.Cls
        picHsc.CurrentY = 100: picHsc.CurrentX = 50
        picHsc.FontBold = True
        picHsc.Print "销帐金额合计:" & FormatEx(dblMoney, 5)
    Else
        picHsc.Height = 30
    End If
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKIND.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKIND.Cards.按缺省卡查找
End Sub
Private Sub LoadBabyCombox()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载婴儿费的相关信息给Combox部件
    '编制:刘兴洪
    '日期:2013-04-10 17:36:17
    '说明:
    '问题:55368
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, intCount As Integer
    
    On Error GoTo errHandle
    intCount = Val(zlDatabase.GetPara("销帐申请婴儿费显示规则", glngSys, Enum_Inside_Program.p记帐操作, "1", _
        Array(cboBaby), InStr(mstrPrivsOpt, ";记帐选项设置;") > 0))
    With cboBaby
        .Clear
        .AddItem "不包含婴儿费用"
        .ItemData(.NewIndex) = 0
        If intCount = 0 Then .ListIndex = .NewIndex
        .AddItem "包含婴儿费用"
        .ItemData(.NewIndex) = 1
        If intCount = 1 Then .ListIndex = .NewIndex
        For i = 1 To 5
            .AddItem "仅显示第" & i & "个婴儿费用"
            .ItemData(.NewIndex) = i + 1
            If intCount = i + 1 Then .ListIndex = .NewIndex
        Next
        If .ListIndex < 0 Then .ListIndex = 0
    End With
     Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetOperatorDept() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取操作员的所属科室(操作员只能为护士时)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-24 11:33:20
    '问题:60679
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mrsOperatorDept Is Nothing Then
       Set GetOperatorDept = mrsOperatorDept
       Exit Function
    End If
    Set mrsOperatorDept = GetDepartments("", "1,2,3", True, True)
    Set GetOperatorDept = mrsOperatorDept
 End Function



Private Function zlMsgModule_Init() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化消息模块
    '入参:lngModule -模块号
    '     strPivs-权限串
    '出参:objMsgModule-返回消息对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error GoTo ErrHand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModule_Init = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModule_Unload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拆卸消息模块
    '入参:objMsgModule-消息对象
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModule_Unload = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Sub SendMsgModule(ByVal str费用IDs As String, ByVal strDate As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消息发送处理
    '入参:
    '编制:刘兴洪
    '日期:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbytFun <> 0 Then Exit Sub
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    
    strSQL = "" & _
    "   Select A.费用id, A.申请类别, A.收费细目id,B.名称 as 销帐项目, B.计算单位," & _
    "       A.审核部门id,C.名称 as 审核部门, A.申请部门id,D.名称 as 申请部门, " & _
    "       A.数量, A.申请人, A.申请时间, A.状态 " & _
    "   From 病人费用销帐 A,收费项目目录 B,部门表 C,部门表 D ,Table(f_Num2List([1])) M" & _
    "   where A.收费细目ID=B.ID and A.审核部门ID=C.ID(+) and A.申请部门ID=D.ID(+)" & _
    "         And A.费用ID=M.Column_value And A.申请时间=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str费用IDs, CDate(strDate))
    If rsTemp.EOF Then Exit Sub
                        
    zlXML.ClearXmlText
        'ZLHIS_CHARGE_001 费用销帐申请通知
    '节点名称    属性    含义    重复    类型    缺省值  值域描述
    'patient_info        病人信息    1
    '   patient_id      病人id  1   N
    '   page_id     主页id  1   N
    '   patient_name        姓名    1   S
    '   patient_sex     性别    1   S
    '   patient_age     年龄    1   S
    '   identity_card       身份证号    0..1    S
    '   in_number       住院号  0..1    S
    '   out_number      门诊号  0..1    S
    'cancel_reqeust      销帐申请    1
    '   cancel_charge           1..*
    '       charge_id       费用id  1   N
    '       request_kind        申请类别    1   N
    '       request_time        申请时间    1   S
    '       request_person      申请人员    1   S
    '       cancel_item_id      销帐项目id  1   N
    '       cancel_item_title       销帐项目    1   S
    '       calcel_num      销帐数量    1   N
    '       charge_unit     费用单位    1   S
    '       audit_dept_id       审核部门id  1   N
    '       audit_dept_title        审核部门    1   S
    Call zlXML.AppendNode("patient_info")
        Call zlXML.appendData("patient_id", Val(Nvl(mrsInfo!病人ID)))
        Call zlXML.appendData("page_id", Val(Nvl(mrsInfo!主页ID)))
        Call zlXML.appendData("patient_name", Nvl(mrsInfo!姓名))
        Call zlXML.appendData("patient_sex", Nvl(mrsInfo!性别))
        Call zlXML.appendData("patient_age", Nvl(mrsInfo!年龄))
        Call zlXML.appendData("identity_card", Nvl(mrsInfo!身份证号))
        Call zlXML.appendData("in_number", Nvl(mrsInfo!住院号))
        Call zlXML.appendData("out_number", Nvl(mrsInfo!门诊号))
    Call zlXML.AppendNode("patient_info", True)
    
    Call zlXML.AppendNode("cancel_reqeust")
        
    With rsTemp
        .MoveFirst
        Do While Not .EOF
            Call zlXML.AppendNode("cancel_charge")
            '       charge_id       费用id  1   N
                Call zlXML.appendData("charge_id", Val(Nvl(!费用id)))
            '       request_kind        申请类别    1   N
                Call zlXML.appendData("request_kind", Val(Nvl(!申请类别)))
            '       request_time        申请时间    1   D
                Call zlXML.appendData("request_time", Format(!申请时间, "yyyy-mm-dd HH:MM:SS"))
            '       request_person      申请人员    1   S
                Call zlXML.appendData("request_person", Nvl(!申请人))
            '       cancel_item_id      销帐项目id  1   N
                Call zlXML.appendData("cancel_item_id", Val(Nvl(!收费细目ID)))
            '       cancel_item_title       销帐项目    1   S
                Call zlXML.appendData("cancel_item_title", Trim(Nvl(!销帐项目)))
            '       calcel_num      销帐数量    1   N
                Call zlXML.appendData("calcel_num", Val(Nvl(!数量)))
            '       charge_unit     费用单位    1   S
                Call zlXML.appendData("charge_unit", Trim(Nvl(!计算单位)))
            '       audit_dept_id       审核部门id  1   N
                Call zlXML.appendData("audit_dept_id", Val(Nvl(!审核部门id)))
            '       audit_dept_title        审核部门    1   S
                Call zlXML.appendData("audit_dept_title", Trim(Nvl(!审核部门)))
            Call zlXML.AppendNode("cancel_charge", True)
            .MoveNext
        Loop
    End With
    Call zlXML.AppendNode("cancel_reqeust", True)
    
    If Not mobjMsgModule Is Nothing Then
        If mobjMsgModule.IsConnect = True Then
        '发检查消息
            Call mobjMsgModule.CommitMessage("ZLHIS_CHARGE_001", zlXML.XmlText)
        End If
    End If
    
    Call zlDatabase.SendMsg("ZLHIS_CHARGE_001", zlXML.XmlText)
    zlXML.ClearXmlText
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetUserDeptIDs() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取操作员所属科室IDS(包含操作员所属病区对应的科室)
    '返回:返回操作员所属科室IDS
    '编制:刘兴洪
    '日期:2015-07-21 16:53:40
    '问题:65039
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    strSQL = _
        " With c_所属科室 As(Select Distinct 部门id From 部门人员 Where 人员id =[1])" & _
        " Select a.科室id As 部门id" & _
        " From 病区科室对应 A, c_所属科室 B" & _
        " Where a.病区id = B.部门id" & _
        " Union All" & _
        " Select 部门id From c_所属科室"
    strSQL = "Select Distinct 部门ID From (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    With rsTemp
        Do While Not .EOF
            strTemp = strTemp & "," & rsTemp!部门ID
            .MoveNext
        Loop
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    End With
    GetUserDeptIDs = strTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetItemName(ByVal lng收费细目ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费细目名称
    '返回:返回收费细目名称
    '编制:刘兴洪
    '日期:2015-07-21 17:04:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "Select 编码,名称  From 收费项目目录 Where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng收费细目ID)
    If rsTemp.EOF Then Exit Function
    GetItemName = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
