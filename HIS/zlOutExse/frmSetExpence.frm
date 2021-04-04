VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   6485
      TabIndex        =   36
      Top             =   1290
      Width           =   1230
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6630
      TabIndex        =   38
      Top             =   4765
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6615
      TabIndex        =   35
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6615
      TabIndex        =   34
      Top             =   345
      Width           =   1100
   End
   Begin TabDlg.SSTab stab 
      Height          =   5355
      Left            =   45
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   60
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   9446
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   564
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "单据控制(&1)"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk现金退款缺省方式"
      Tab(0).Control(1)=   "chk医保结算光标缺省定位"
      Tab(0).Control(2)=   "fraDoctor"
      Tab(0).Control(3)=   "fra病人"
      Tab(0).Control(4)=   "chk收费执行科室"
      Tab(0).Control(5)=   "txt收费执行科室"
      Tab(0).Control(6)=   "cmd收费执行科室"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrintSetup(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra退费缺省选择方式"
      Tab(0).Control(9)=   "chkLedDispDetail"
      Tab(0).Control(10)=   "chkLedWelcome"
      Tab(0).Control(11)=   "chkPayKey"
      Tab(0).Control(12)=   "fra类别"
      Tab(0).Control(13)=   "cbo费别"
      Tab(0).Control(14)=   "cbo结算方式"
      Tab(0).Control(15)=   "fra科室与医生"
      Tab(0).Control(16)=   "lbl费别"
      Tab(0).Control(17)=   "lbl结算方式"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "票据控制(&2)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "chkDefaultPrintDays"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdPrintSetup(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPrintSetup(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdPrintSetup(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkRegistInvoice"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdPrintSetup(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdPrintSetup(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdPrintSetup(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fraTitle"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdPrintSetup(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDefaultPrintDays"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fraLine"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "药房设置(&3)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl发料部门"
      Tab(2).Control(1)=   "vsfDrugStore"
      Tab(2).Control(2)=   "cbo卫材"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2685
         TabIndex        =   48
         Top             =   4260
         Width           =   285
      End
      Begin VB.TextBox txtDefaultPrintDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   210
         Left            =   2655
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "7"
         Top             =   4050
         Width           =   345
      End
      Begin VB.CheckBox chk现金退款缺省方式 
         Caption         =   "退费选择""现金""结算方式时缺省退款金额"
         Height          =   195
         Left            =   -74820
         TabIndex        =   18
         Top             =   3810
         Width           =   3585
      End
      Begin VB.CheckBox chk医保结算光标缺省定位 
         Caption         =   "医保结算光标缺省定位到“医保结算”按钮"
         Height          =   195
         Left            =   -74820
         TabIndex        =   17
         Top             =   3555
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "退费票据打印设置(&0)"
         Height          =   350
         Index           =   7
         Left            =   4185
         TabIndex        =   47
         Top             =   3960
         Width           =   1950
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "显示开单人"
         Height          =   630
         Left            =   -74820
         TabIndex        =   7
         Top             =   1560
         Width           =   3945
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按编码+姓名显示"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   2085
            TabIndex        =   9
            Top             =   285
            Width           =   1695
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按简码+姓名显示"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   8
            Top             =   285
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame fra病人 
         Caption         =   "病人来源"
         Height          =   1095
         Left            =   -72240
         TabIndex        =   4
         Top             =   390
         Width           =   1365
         Begin VB.OptionButton opt病人 
            Caption         =   "门诊病人"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   165
            TabIndex        =   5
            Top             =   345
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt病人 
            Caption         =   "住院病人"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   165
            TabIndex        =   6
            ToolTipText     =   "住院病人门诊记帐时,门诊标志为1(后面结帐等操作也将按照门诊记帐的规则处理)"
            Top             =   675
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk收费执行科室 
         Caption         =   "本机收费执行科室"
         Height          =   210
         Left            =   -74820
         TabIndex        =   25
         Top             =   5040
         Width           =   1770
      End
      Begin VB.TextBox txt收费执行科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   -73050
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4980
         Width           =   3975
      End
      Begin VB.CommandButton cmd收费执行科室 
         Caption         =   "…"
         Height          =   280
         Left            =   -69060
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4995
         Width           =   280
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "划价通知单打印设置(&1)"
         Height          =   350
         Index           =   3
         Left            =   -70830
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4350
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame fra退费缺省选择方式 
         Caption         =   "退费缺省选择方式"
         Height          =   840
         Left            =   -74820
         TabIndex        =   19
         Top             =   4065
         Width           =   3945
         Begin VB.OptionButton opt退费缺省选择方式 
            Caption         =   "缺省全选择退费项目"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   555
            Width           =   2010
         End
         Begin VB.OptionButton opt退费缺省选择方式 
            Caption         =   "缺省按单据号或发票号选择退费项目"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Value           =   -1  'True
            Width           =   3195
         End
      End
      Begin VB.ComboBox cbo卫材 
         Height          =   300
         Left            =   -73620
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   4815
         Width           =   2355
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用收费票据"
         Height          =   3105
         Left            =   150
         TabIndex        =   45
         Top             =   510
         Width           =   6000
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2760
            Left            =   75
            TabIndex        =   28
            Top             =   255
            Width           =   5790
            _cx             =   10213
            _cy             =   4868
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
            FormatString    =   $"frmSetExpence.frx":0060
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
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "收费票据打印设置(&1)"
         Height          =   350
         Index           =   0
         Left            =   150
         TabIndex        =   44
         Top             =   4365
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "收据证明打印设置(&2)"
         Height          =   350
         Index           =   1
         Left            =   2160
         TabIndex        =   43
         Top             =   4365
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "收费清单打印设置(&3)"
         Height          =   350
         Index           =   2
         Left            =   4185
         TabIndex        =   42
         Top             =   4365
         Width           =   1950
      End
      Begin VB.CheckBox chkRegistInvoice 
         Caption         =   "挂号时使用与收费相同的票据"
         Height          =   195
         Left            =   165
         TabIndex        =   29
         Top             =   3780
         Width           =   2640
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "医保回单打印设置(&4)"
         Height          =   350
         Index           =   4
         Left            =   150
         TabIndex        =   41
         Top             =   4770
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "退费回单打印设置(&5)"
         Height          =   350
         Index           =   5
         Left            =   2160
         TabIndex        =   40
         Top             =   4770
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "执行清单打印设置(&6)"
         Height          =   350
         Index           =   6
         Left            =   4185
         TabIndex        =   39
         Top             =   4770
         Width           =   1950
      End
      Begin VB.CheckBox chkLedDispDetail 
         Caption         =   "LED显示收费明细"
         Height          =   225
         Left            =   -74835
         TabIndex        =   15
         ToolTipText     =   "收费窗口,输入收费项目后是否显示信息"
         Top             =   3300
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   -73020
         TabIndex        =   16
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   3300
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkPayKey 
         Caption         =   "使用小键盘的加减(+-)来切换支付方式"
         Height          =   195
         Left            =   -74835
         TabIndex        =   14
         Top             =   3045
         Width           =   3375
      End
      Begin VB.Frame fra类别 
         Caption         =   "可用收费类别"
         Height          =   3885
         Left            =   -70800
         TabIndex        =   22
         Top             =   390
         Width           =   2055
         Begin VB.ListBox lst收费类别 
            ForeColor       =   &H00C00000&
            Height          =   3420
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   23
            ToolTipText     =   "请复选允许使用的收费类别"
            Top             =   345
            Width           =   1920
         End
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2295
         Width           =   2235
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   300
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2655
         Width           =   2235
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4140
         Left            =   -74835
         TabIndex        =   32
         Top             =   555
         Width           =   5970
         _cx             =   10530
         _cy             =   7302
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
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSetExpence.frx":013E
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
      Begin VB.Frame fra科室与医生 
         Caption         =   "科室与医生"
         Height          =   1095
         Left            =   -74835
         TabIndex        =   0
         Top             =   390
         Width           =   2580
         Begin VB.OptionButton optUnit 
            Caption         =   "通过输入科室来确定医生"
            Height          =   180
            Left            =   210
            TabIndex        =   1
            Top             =   300
            Value           =   -1  'True
            Width           =   2280
         End
         Begin VB.OptionButton optDoctor 
            Caption         =   "通过输入医生来确定科室"
            Height          =   180
            Left            =   210
            TabIndex        =   2
            Top             =   540
            Width           =   2280
         End
         Begin VB.OptionButton optSelf 
            Caption         =   "科室和医生互相独立输入"
            Height          =   195
            Left            =   210
            TabIndex        =   3
            Top             =   780
            Width           =   2280
         End
      End
      Begin VB.CheckBox chkDefaultPrintDays 
         Caption         =   "按病人补打票据时缺省打印     天的费用"
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   4050
         Width           =   3660
      End
      Begin VB.Label lbl发料部门 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省发料部门"
         Height          =   180
         Left            =   -74820
         TabIndex        =   46
         Top             =   4875
         Width           =   1080
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省病人费别"
         Height          =   180
         Left            =   -74835
         TabIndex        =   10
         Top             =   2355
         Width           =   1080
      End
      Begin VB.Label lbl结算方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省结算方式"
         Height          =   180
         Left            =   -74835
         TabIndex        =   12
         Top             =   2715
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytInFun As Byte '0=收费,1=划价,2=门诊记帐
Public mstrPrivs As String
Public mlngModul As Long
Public mblnSetDrugStore As Boolean
Private mblnAutoAddItem As Boolean
Private mblnNotClick As Boolean

Private Sub chkDefaultPrintDays_Click()
    txtDefaultPrintDays.Enabled = (chkDefaultPrintDays.Value = vbChecked)
    txtDefaultPrintDays.BackColor = IIf(txtDefaultPrintDays.Enabled, vbWhite, vbButtonFace)
End Sub

Private Sub cmdDeviceSetup_Click()
    Dim lngModule As Long
    Select Case mbytInFun
    Case 0
        lngModule = 1121
    Case 1
        lngModule = 1120
    Case 2
        lngModule = 1122
    End Select
    Call zlCommFun.DeviceSetup(Me, glngSys, lngModule)
End Sub

Private Sub cbo费别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cbo费别.ListIndex = -1
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
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
    zlDatabase.SetPara "共用收费票据批次", strValue, glngSys, mlngModul, blnHavePrivs
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
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Dim strValue As String, i As Long
    Dim str西药房窗口 As String, str中药房窗口 As String, str成药房窗口 As String
    Dim lng缺省西药房 As Long, lng缺省中药房 As Long, lng缺省成药房 As Long, lng缺省发料部门 As Long
    
    'a.数据检查
    '--------------------------------------------------------------
    'b.本机注册表存储的模块参数
    '------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If isValied = False Then Exit Sub
     
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    
    If Not mblnSetDrugStore Then
        For i = lst收费类别.ListCount - 1 To 0 Step -1
            If lst收费类别.Selected(i) Then strValue = strValue & "'" & Chr(lst收费类别.ItemData(i)) & "',"
        Next
        
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "收费类别", strValue, glngSys, mlngModul, blnHavePrivs
        
        If mbytInFun <> 2 Then
            zlDatabase.SetPara "缺省费别", cbo费别.Text, glngSys, mlngModul, blnHavePrivs
        End If
        
        If mbytInFun = 0 Then
            Call SaveInvoice
            zlDatabase.SetPara "缺省结算方式", cbo结算方式.Text, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "挂号共用收费票据", chkRegistInvoice.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED显示收费明细", chkLedDispDetail.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "医保结算光标缺省定位", chk医保结算光标缺省定位.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "现金退款缺省方式", chk现金退款缺省方式.Value, glngSys, mlngModul, blnHavePrivs
            
            '96357
            zlDatabase.SetPara "本机收费执行科室", txt收费执行科室.Tag, glngSys, mlngModul, blnHavePrivs
            
            If chkDefaultPrintDays.Value = vbUnchecked Then
                strValue = "0"
            Else
                strValue = Val(txtDefaultPrintDays.Text)
            End If
            zlDatabase.SetPara "缺省发票打印天数", strValue, glngSys, mlngModul, blnHavePrivs
        End If
    End If

    With vsfDrugStore
        For i = 1 To vsfDrugStore.Rows - 1
            If (mbytInFun = 0 Or mbytInFun = 1) And .TextMatrix(i, .ColIndex("窗口")) <> "自动分配" And .TextMatrix(i, .ColIndex("窗口")) <> "" Then
                Select Case .TextMatrix(i, 0)
                    Case "西药房"
                        str西药房窗口 = str西药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                    Case "中药房"
                        str中药房窗口 = str中药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                    Case "成药房"
                        str成药房窗口 = str成药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                End Select
            End If
            
            If Abs(Val(.TextMatrix(i, .ColIndex("缺省")))) = 1 Then
                Select Case .TextMatrix(i, .ColIndex("类别"))
                    Case "西药房"
                        lng缺省西药房 = .RowData(i)
                    Case "中药房"
                        lng缺省中药房 = .RowData(i)
                    Case "成药房"
                        lng缺省成药房 = .RowData(i)
                End Select
            End If
        Next
    End With
    
    If cbo卫材.ListIndex <> -1 Then
        lng缺省发料部门 = cbo卫材.ItemData(cbo卫材.ListIndex)
    End If
    
    
    If mbytInFun = 0 Or mbytInFun = 1 Then
        str西药房窗口 = Mid(str西药房窗口, 2)
        str中药房窗口 = Mid(str中药房窗口, 2)
        str成药房窗口 = Mid(str成药房窗口, 2)
        zlDatabase.SetPara "西药房窗口", str西药房窗口, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "中药房窗口", str中药房窗口, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "成药房窗口", str成药房窗口, glngSys, mlngModul, blnHavePrivs
    End If
    
    zlDatabase.SetPara "缺省西药房", lng缺省西药房, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省中药房", lng缺省中药房, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省成药房", lng缺省成药房, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省发料部门", lng缺省发料部门, glngSys, mlngModul, blnHavePrivs
    
    If Not mblnSetDrugStore Then
        zlDatabase.SetPara "科室医生", IIf(optDoctor.Value, 0, IIf(optUnit.Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "开单人显示方式", IIf(optDoctorKind(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "病人来源", IIf(opt病人(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
        If mbytInFun = 0 Or mbytInFun = 1 Then
            zlDatabase.SetPara "使用加减切换支付方式", IIf(chkPayKey.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
        End If
    End If
     '87489
    zlDatabase.SetPara "退费缺省选择方式", IIf(opt退费缺省选择方式(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    
    Call InitLocPar(Choose(mbytInFun + 1, 1121, 1120, 1122))     '主要是要重读存到本机注册表的参数,存在数据库的参数在保存时已重读
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSetup_Click(Index As Integer)
    Select Case Index
        Case 0 '门诊医疗费收费
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_1", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me)
                End If
            End If
        Case 1 '门诊诊断证明
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_2", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_2", Me)
            End If
        Case 2 '门诊收费清单
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me)
            End If
        Case 3 '划价通知单
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me)
        Case 4 '医保回单设置
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me)
        Case 5  '退费回单设置
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me)
        '62982:李南春,2015/5/19,收费执行单
        Case 6  '收费执行单设置
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me)
        Case 7  '退费执行单设置
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_7", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_7", Me)
                End If
            End If
    End Select
End Sub
 
Private Sub SetDrugStore()
    Dim lngType As Long, strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long, lngRow As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    With vsfDrugStore
        strTmp = "'西药房','中药房','成药房','发料部门'"
        
        If stab.TabVisible(1) = True Then
            lngType = IIf(opt病人(0).Value, 1, 2)
        Else
            lngType = gint病人来源
        End If
        Set rsTmp = GetDepartments(strTmp, lngType & ",3")
        .Rows = 1
        If mbytInFun = 2 Then .ColHidden(3) = True '门诊记帐不设窗口
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "工作性质<>'发料部门'"
            .Rows = rsTmp.RecordCount + 1
            .MergeCells = flexMergeFixedOnly
            .MergeCol(0) = True
            
            strTmp = "'西药房','中药房','成药房'"
            arrTmp = Split(strTmp, ",")
            lngRow = 1
            For j = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质=" & arrTmp(j)
                If rsTmp.RecordCount > 0 Then
                    For i = 1 To rsTmp.RecordCount
                        .TextMatrix(lngRow, 0) = Replace(arrTmp(j), "'", "")
                        .TextMatrix(lngRow, 1) = 0
                        .TextMatrix(lngRow, 2) = rsTmp!名称
                        If mbytInFun <> 2 Then .TextMatrix(lngRow, 3) = "自动分配"
                        .RowData(lngRow) = Val(rsTmp!ID)
                        lngRow = lngRow + 1
                        rsTmp.MoveNext
                    Next
                    
                    If lngRow < .Rows - 1 Then  '划分隔线
                        .Select lngRow, .FixedCols, lngRow, .COLS - 1
                        .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                    End If
                End If
            Next
            
            cbo卫材.AddItem "人工选择"
            rsTmp.Filter = "工作性质='发料部门'"
            For j = 1 To rsTmp.RecordCount
                cbo卫材.AddItem rsTmp!名称
                cbo卫材.ItemData(cbo卫材.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Next
            cbo卫材.ListIndex = 0
        End If
    
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
Private Sub Load药房ParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载药房相关参数值
    '编制:刘兴洪
    '日期:2011-12-07 15:05:10
    '问题:43775
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean
    Dim i As Long, k As Long, j As Long, intType As Integer
    Dim arrTmp  As Variant, arrWindow As Variant
    Dim str西药房窗口 As String, str中药房窗口 As String, str成药房窗口 As String
    Dim lng缺省西药房 As Long, lng缺省中药房 As Long, lng缺省成药房 As Long, lng缺省发料部门 As Long
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    With vsfDrugStore
        arrTmp = Split("缺省西药房,缺省中药房,缺省成药房", ",")
        .Cell(flexcpData, 0, 0, .Rows - 1, .COLS - 1) = "0" '存储是否允许编译.:0-不锁定,1-锁定
        
        For j = 0 To UBound(arrTmp)
            '刘兴洪:由于可能参数权限发生变更,因此,不能统一进行设置,需要设置某一部分:
            '问题:25132,intType-'返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
            strTmp = zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, "0", , blnParSet, intType)
            If Val(strTmp) > 0 Then
                Select Case arrTmp(j)
                    Case "缺省西药房"
                        lng缺省西药房 = Val(strTmp)
                    Case "缺省中药房"
                        lng缺省中药房 = Val(strTmp)
                    Case "缺省成药房"
                        lng缺省成药房 = Val(strTmp)
                End Select
                Call SetDrugStockEdit(Replace(arrTmp(j), "缺省", ""), intType, .ColIndex("缺省"), Val(strTmp))
            Else
                Call SetDrugStockEdit(Replace(arrTmp(j), "缺省", ""), intType, .ColIndex("缺省"), "")
            End If
        Next
        
        strTmp = zlDatabase.GetPara("缺省发料部门", glngSys, mlngModul, "0", Array(cbo卫材), blnParSet)
        zlControl.CboLocate cbo卫材, strTmp, True
        
        If mbytInFun <> 2 Then
                arrTmp = Split("西药房窗口,中药房窗口,成药房窗口", ",")
                For j = 0 To UBound(arrTmp)
                    strTmp = Trim(zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, , , blnParSet, intType))
                    If strTmp <> "" Then
                        '处理旧的数据,窗口参数中没有存储药房ID
                        If InStr(strTmp, ":") = 0 Then
                            Select Case arrTmp(j)
                                Case "西药房窗口"
                                    strTmp = lng缺省西药房 & ":" & strTmp
                                Case "中药房窗口"
                                    strTmp = lng缺省中药房 & ":" & strTmp
                                Case "成药房窗口"
                                    strTmp = lng缺省成药房 & ":" & strTmp
                            End Select
                        End If
                        arrWindow = Split(strTmp, ",")
                        strTmp = Replace(arrTmp(j), "窗口", "")
                        For k = 0 To UBound(arrWindow)
                            Call SetDrugStockEdit(Replace(arrTmp(j), "窗口", ""), intType, .ColIndex("窗口"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
                        Next
                    Else
                        Call SetDrugStockEdit(Replace(arrTmp(j), "窗口", ""), intType, .ColIndex("窗口"), "")
                    End If
                Next
            End If
        End With
End Sub
Private Sub LoadParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数值
    '编制:刘兴洪
    '日期:2011-09-12 15:03:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean, k As Long, rsTmp As ADODB.Recordset
    Dim i As Long, arrTmp As Variant, j As Long, intType As Integer, arrWindow As Variant
        
    
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0

    strTmp = zlDatabase.GetPara("收费类别", glngSys, mlngModul, , Array(lst收费类别), blnParSet)
    If strTmp = "" Then
        For i = 0 To lst收费类别.ListCount - 1
            lst收费类别.Selected(i) = True
        Next
    Else
        For i = 0 To lst收费类别.ListCount - 1
            If InStr(strTmp, Chr(lst收费类别.ItemData(i))) Then lst收费类别.Selected(i) = True
        Next
    End If
    If lst收费类别.ListCount > 0 Then lst收费类别.TopIndex = 0: lst收费类别.ListIndex = 0
    If mbytInFun <> 2 Then
        strTmp = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , Array(cbo费别), blnParSet)
        zlControl.CboLocate cbo费别, strTmp
    End If
    
    i = IIf(zlDatabase.GetPara("病人来源", glngSys, mlngModul, , Array(opt病人(0), opt病人(1)), blnParSet) = "1", 0, 1)
    opt病人(i).Value = True
    If mbytInFun <> 2 Then opt病人(1).ToolTipText = ""
    
    Call opt病人_Click(IIf(opt病人(0).Value, 0, 1)) '加载药品库房和卫材发料部门
    Call Load药房ParaValue
    Select Case mbytInFun
    Case 2 '记帐
    Case 1 '划价
    Case Else
        chkRegistInvoice.Value = IIf(zlDatabase.GetPara("挂号共用收费票据", glngSys, mlngModul, 0, Array(chkRegistInvoice), blnParSet) = "1", 1, 0)
        chkLedDispDetail.Value = IIf(zlDatabase.GetPara("LED显示收费明细", glngSys, mlngModul, 1, Array(chkLedDispDetail), blnParSet) = "1", 1, 0)
        chkLedWelcome.Value = IIf(zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, 1, Array(chkLedWelcome), blnParSet) = "1", 1, 0)
        
        Dim objCards As Cards '启用的三方账户的医疗卡
        Set rsTmp = Get结算方式("收费", "1,2,7,8")
        If Not gobjSquare Is Nothing Then
            ' zlGetCards(ByVal BytType As Byte)
                '入参:bytType-  0-所有医疗卡;
            '                        1-启用的医疗卡,
            '                        2-所有存在三方账户的三方卡
            '                        3-启用的三方账户的医疗卡
           Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
        End If
        With cbo结算方式
            .Clear
            Do While Not rsTmp.EOF
                If Not (Val(NVL(rsTmp!性质)) = 7 Or Val(NVL(rsTmp!性质)) = 8 Or Val(NVL(rsTmp!应付款)) = 1) Then
                    .AddItem NVL(rsTmp!名称)
                End If
                rsTmp.MoveNext
            Loop
            '加入医疗卡结算方式，对应结算方式未启用的不加入
            For i = 1 To objCards.Count
            rsTmp.Filter = "名称='" & objCards(i).结算方式 & "'"
                If Not rsTmp.EOF Then
                    .AddItem objCards(i).名称
                End If
            Next
        End With
        '问题:54923
        strTmp = zlDatabase.GetPara("缺省结算方式", glngSys, mlngModul, , Array(cbo结算方式), blnParSet)
        For i = 0 To cbo结算方式.ListCount - 1
            If cbo结算方式.List(i) = strTmp Then cbo结算方式.ListIndex = i: Exit For
        Next
        
        '加载发票相关
        Call InitShareInvoice
        chkPayKey.Value = IIf(Val(zlDatabase.GetPara("使用加减切换支付方式", glngSys, mlngModul, "1", Array(chkPayKey), blnParSet)) = 1, 1, 0)
        '87489
        strTmp = zlDatabase.GetPara("退费缺省选择方式", glngSys, mlngModul, "0", Array(opt退费缺省选择方式(0), opt退费缺省选择方式(1)), blnParSet)
        For i = 0 To 1
            If Val(strTmp) = i Then opt退费缺省选择方式(i).Value = True: Exit For
        Next
        chk医保结算光标缺省定位.Value = IIf(zlDatabase.GetPara("医保结算光标缺省定位", glngSys, mlngModul, "0", Array(chk医保结算光标缺省定位), blnParSet) = "1", 1, 0)
        chk现金退款缺省方式.Value = IIf(zlDatabase.GetPara("现金退款缺省方式", glngSys, mlngModul, "0", Array(chk现金退款缺省方式), blnParSet) = "1", 1, 0)
        
        '96357
        strTmp = zlDatabase.GetPara("本机收费执行科室", glngSys, mlngModul, , Array(chk收费执行科室, txt收费执行科室, cmd收费执行科室), blnParSet)
        mblnNotClick = True
        chk收费执行科室.Value = IIf(strTmp <> "", vbChecked, vbUnchecked)
        mblnNotClick = False
        cmd收费执行科室.Enabled = chk收费执行科室.Value = vbChecked
        txt收费执行科室.Text = GetDeptNameStr(strTmp)
        txt收费执行科室.Tag = strTmp
        
        strTmp = zlDatabase.GetPara("缺省发票打印天数", glngSys, mlngModul, "0", Array(chkDefaultPrintDays, txtDefaultPrintDays), blnParSet)
        If Val(strTmp) <= 0 Then
            chkDefaultPrintDays.Value = vbUnchecked
            txtDefaultPrintDays.Text = 7
        Else
            chkDefaultPrintDays.Value = vbChecked
            txtDefaultPrintDays.Text = strTmp
        End If
        txtDefaultPrintDays.Enabled = (chkDefaultPrintDays.Value = vbChecked)
        txtDefaultPrintDays.BackColor = IIf(txtDefaultPrintDays.Enabled, vbWhite, vbButtonFace)
    End Select
End Sub
 
 

Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
    strShareInvoice = zlDatabase.GetPara("共用收费票据批次", glngSys, mlngModul, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,使用类别1|领用IDn,使用类别n|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(1)
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
    
 
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSQL As String, objItem As ListItem, blnParSet As Boolean
    Dim strTmp As String, i As Integer, j As Long, k As Long, arrTmp As Variant, arrWindow As Variant, intType As Integer, blnSeted As Boolean '被设置了缺省值

    Dim str西药房窗口 As String, str中药房窗口 As String, str成药房窗口 As String
    Dim lng缺省西药房 As Long, lng缺省中药房 As Long, lng缺省成药房 As Long, lng缺省发料部门 As Long
    
    gblnOK = False
    On Error GoTo errH
    If mbytInFun = 0 Then
         mblnAutoAddItem = InStr(zlDatabase.GetPara("自动加收挂号费", glngSys, mlngModul), ";") > 0
    End If
    
    blnParSet = InStr(1, mstrPrivs, "参数设置") > 0
    
    'a.初始数据
    '----------------------------------------------------------------------------------------
    '收费类别(挂号除外):按序号排序
    strSQL = "Select 编码,名称 as 类别 from 收费项目类别 Where 编码<>'1' Order by 序号"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst收费类别.AddItem rsTmp!类别
        lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
    
        rsTmp.MoveNext
    Loop
    If mbytInFun <> 2 Then
        strSQL = _
            " Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别" & _
            " Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(1,3)" & _
            "       And Sysdate Between Nvl(有效开始,To_Date('1900-01-01','yyyy-mm-dd')) And Nvl(有效结束,To_Date('3000-01-01','yyyy-mm-dd'))+1-1/24/60/60" & _
            " Order by 编码"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            cbo费别.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo费别.ListIndex = cbo费别.NewIndex
            rsTmp.MoveNext
        Next
    End If
     
    'b.本机注册表存储的模块参数
    '----------------------------------------------------------------------------------------
    Call LoadParaValue
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    '--------------------------
    strTmp = zlDatabase.GetPara("科室医生", glngSys, mlngModul, , Array(optUnit, optDoctor, optSelf), blnParSet)
    If strTmp = "1" Then
        optUnit.Value = True
    ElseIf strTmp = "0" Then
        optDoctor.Value = True
    Else
        optSelf.Value = True
    End If
    
    i = IIf(zlDatabase.GetPara("开单人显示方式", glngSys, mlngModul, "1", Array(optDoctorKind(0), optDoctorKind(1)), blnParSet) = "1", 0, 1)
    optDoctorKind(i).Value = True
    
    
    'd.权限控制
    '----------------------------------------------------------------------------------------
    chkLedDispDetail.Visible = mbytInFun = 0
    chkLedWelcome.Visible = mbytInFun = 0
    chkPayKey.Visible = mbytInFun = 0
    '87489
    fra退费缺省选择方式.Visible = mbytInFun = 0
    chk医保结算光标缺省定位.Visible = mbytInFun = 0
    chk现金退款缺省方式.Visible = mbytInFun = 0
    
    txt收费执行科室.Visible = mbytInFun = 0
    chk收费执行科室.Visible = mbytInFun = 0
    cmd收费执行科室.Visible = mbytInFun = 0

    lbl结算方式.Visible = mbytInFun = 0
    cbo结算方式.Visible = mbytInFun = 0
    
    cmdPrintSetup(3).Visible = mbytInFun = 1
    lbl费别.Visible = mbytInFun <> 2
    cbo费别.Visible = mbytInFun <> 2
    
    stab.TabVisible(1) = mbytInFun = 0
 
    If mblnSetDrugStore Then
        '56963
        stab.TabCaption(2) = "药房设置"
        stab.TabVisible(0) = False
        stab.TabVisible(1) = False
    Else
        If mbytInFun = 1 Then
            stab.TabCaption(2) = "药房设置(&2)"
        ElseIf mbytInFun = 2 Then
            stab.TabCaption(2) = "药房设置(&2)"
        End If
    End If
    If stab.TabVisible(0) Then stab.Tab = 0

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDeptNameStr(ByVal strIDs As String) As String
    '将部门ID字符串装换成名称字符串
    '入参：
    '   strIDs 部门ID，格式：ID1,ID2,ID3,...
    '返回：
    '   部门名称s，格式：部门名称1;部门名称2;部门名称3;...
    Dim strSQL As String, rsTemp As Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    If strIDs = "" Then Exit Function
    strSQL = "Select /*+cardinality(b,10) */a.编码, a.名称" & vbNewLine & _
            " From 部门表 A, Table(f_Str2list([1], ',')) B" & vbNewLine & _
            " Where a.Id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据科室ID获取科室名称", strIDs)
    If rsTemp Is Nothing Then Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & ";" & NVL(rsTemp!名称)
        rsTemp.MoveNext
    Loop
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    GetDeptNameStr = strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mblnSetDrugStore = False
    If mbytInFun = 0 Then
        zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
    End If
    
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub

Private Sub opt病人_Click(Index As Integer)
    
    Call SetDrugStore

End Sub
Private Sub stab_Click(PreviousTab As Integer)
    Select Case stab.Tab
        Case 0
            If optUnit.Enabled And optUnit.Visible And optUnit.Value Then optUnit.SetFocus
            If optSelf.Enabled And optSelf.Visible And optSelf.Value Then optSelf.SetFocus
            If optDoctor.Enabled And optDoctor.Visible And optDoctor.Value Then optDoctor.SetFocus
        Case 1
            If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
        Case 2
            If vsfDrugStore.Visible And vsfDrugStore.Enabled Then vsfDrugStore.SetFocus
    End Select
End Sub
      

Private Sub txtDefaultPrintDays_GotFocus()
    zlControl.TxtSelAll txtDefaultPrintDays
End Sub

Private Sub txtDefaultPrintDays_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtDefaultPrintDays, KeyAscii, m数字式
End Sub

Private Sub txtDefaultPrintDays_Validate(Cancel As Boolean)
    If Val(txtDefaultPrintDays.Text) < 1 Then txtDefaultPrintDays.Text = 1
End Sub

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

 
Private Sub vsfDrugStore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("缺省")
           Call SetDrugStockDeFault(Row)
        Case Else
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("缺省"), .ColIndex("窗口")
            Cancel = Val(.Cell(flexcpData, Row, Col)) = 1
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    Dim strTmp As String, i As Long
    With vsfDrugStore
        If Not (.Row > 0 And .Col = 1) Then Exit Sub
        If .Cell(flexcpData, .Row, .ColIndex("缺省")) = 1 Then Exit Sub
        
        .TextMatrix(.Row, .Col) = IIf(Val(.TextMatrix(.Row, .Col)) = 0, 1, 0)
        Call SetDrugStockDeFault(.Row)
    End With
End Sub
Private Sub SetDrugStockDeFault(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置药房的缺省值
    '入参:lngRow-指定行
    '编制:刘兴洪
    '日期:2009-09-02 14:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng缺省 As Long, strType As String
    With vsfDrugStore
        lng缺省 = Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省"))))
        If lng缺省 = 1 Then
            strType = .TextMatrix(lngRow, .ColIndex("类别"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = strType And i <> lngRow Then
                    .TextMatrix(i, .ColIndex("缺省")) = 0
                End If
            Next
        End If
    End With
End Sub
Private Sub SetDrugStockEdit(ByVal strType As String, ByVal intType As Integer, ByVal lngEditCol As Long, Optional strMachValue As String = "", Optional strDefaultValue As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置药房的编辑属性
    '入参:strType-类别
    '     intType-返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    '     lngEditCol-控制的编辑列
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-02 14:53:10
    '问题:25132
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSetDefault As Boolean '设置了缺省值了,随后不能再设置缺省值
    Dim lngEditForColor As Long, blnAllowEdit As Boolean, bytLockEdit As Integer '1-锁定,0-不锁定
    
    '刘兴洪:由于可能参数权限发生变更,因此,不能统一进行设置,需要设置某一部分:
    With vsfDrugStore
        blnSetDefault = False: blnAllowEdit = InStr(1, mstrPrivs, ";参数设置;") > 0
        bytLockEdit = 0
        If InStr(1, ",1,3,15,", "," & intType & ",") > 0 Then
            lngEditForColor = IIf(blnAllowEdit, vbBlue, &H8000000C)  '授权限控制
            bytLockEdit = IIf(blnAllowEdit, 0, 1)
        ElseIf intType = 5 Then
            lngEditForColor = vbBlue    '公共模块,但不授权限控制
        Else
            lngEditForColor = &H80000008    '正常编辑
        End If
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("类别")) = strType Then
                If lngEditCol = .ColIndex("缺省") Then
                    '设置药房
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, .ColIndex("缺省")) = IIf(Val(strMachValue) > 0, 1, 0)
                        blnSetDefault = True
                    End If
                     .Cell(flexcpForeColor, i, .ColIndex("缺省")) = lngEditForColor
                     .Cell(flexcpForeColor, i, .ColIndex("药房")) = lngEditForColor:
                Else
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, lngEditCol) = strDefaultValue
                    End If
                    '设置窗口
                     .Cell(flexcpForeColor, i, .ColIndex("窗口")) = lngEditForColor
                End If
                .Cell(flexcpData, i, lngEditCol) = bytLockEdit
            End If
        Next
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("窗口") Then
                Set rsTmp = Read发药窗口(.RowData(.Row))
                strList = "自动分配|" & .BuildComboList(rsTmp, "名称")
                .ColComboList(.Col) = strList
            Else
                .ColComboList(.Col) = ""
              '  .Editable = flexEDNone
            End If
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub chk收费执行科室_Click()
    If mblnNotClick Then Exit Sub
    If chk收费执行科室.Value = vbChecked Then
        cmd收费执行科室.Enabled = True
        Call cmd收费执行科室_Click
    Else
        txt收费执行科室.Text = ""
        txt收费执行科室.Tag = ""
        cmd收费执行科室.Enabled = False
    End If
End Sub

Private Sub cmd收费执行科室_Click()
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    '96357
    strSQL = "Select Distinct A.ID, A.编码, A.名称, A.简码" & vbNewLine & _
            " From 部门表 A, 部门性质说明 B" & vbNewLine & _
            " Where B.部门ID=A.ID And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & vbNewLine & _
            "       And B.工作性质 In('中药房', '西药房', '成药房', '发料部门')" & vbNewLine & _
            "       And B.服务对象 In (1, 2, 3)" & vbNewLine & _
            "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
    vRect = zlControl.GetControlRect(txt收费执行科室.hWnd)
    Set rsDept = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "本机收费执行科室", True, "", "", False, False, False, _
        vRect.Left, vRect.Top, txt收费执行科室.Height, blnCancel, False, True, "MultiCheckReturn=1")
    If blnCancel Then Exit Sub
    If rsDept Is Nothing Then Exit Sub
    
    txt收费执行科室.Text = ""
    txt收费执行科室.Tag = ""
    Do While Not rsDept.EOF
        txt收费执行科室.Text = txt收费执行科室.Text & ";" & NVL(rsDept!名称)
        strTemp = strTemp & "," & NVL(rsDept!ID)
        rsDept.MoveNext
    Loop
    If txt收费执行科室.Text <> "" Then txt收费执行科室.Text = Mid(txt收费执行科室.Text, 2)
    If strTemp <> "" Then txt收费执行科室.Tag = Mid(strTemp, 2)
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

