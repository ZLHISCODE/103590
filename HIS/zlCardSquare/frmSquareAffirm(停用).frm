VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSquareAffirm 
   Caption         =   "病人消费结算"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareAffirm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8114.99
   ScaleMode       =   0  'User
   ScaleWidth      =   11445
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPatientInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   60
      ScaleHeight     =   1455
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   60
      Width           =   8925
      Begin VB.CommandButton cmdYB 
         Caption         =   "医保"
         Height          =   375
         Left            =   3435
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "热键：F6"
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   "预交余额:99999999.99"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   6
         Top             =   585
         Width           =   4110
      End
      Begin VB.Label lbl 
         Caption         =   "未结费用:99999999.99"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4380
         TabIndex        =   7
         Top             =   585
         Width           =   4410
      End
      Begin VB.Label lbl 
         Caption         =   "哈斯啦.王珂"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   2
         Top             =   135
         Width           =   2370
      End
      Begin VB.Label lbl 
         Caption         =   "门诊号:1810080001"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6000
         TabIndex        =   5
         Top             =   135
         Width           =   2760
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "病  人:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   135
         Width           =   1260
      End
      Begin VB.Label lbl 
         Caption         =   "性别:不明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4380
         TabIndex        =   4
         Top             =   135
         Width           =   1425
      End
      Begin VB.Label lbl 
         Caption         =   "剩余款额:99999999.99"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   8
         Top             =   1035
         Width           =   4110
      End
      Begin VB.Label lbl 
         Caption         =   "家属余额:99999999.99"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4380
         TabIndex        =   9
         Top             =   1035
         Width           =   4410
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   1470
         X2              =   4140
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   2
         X1              =   5160
         X2              =   5850
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   3
         X1              =   7140
         X2              =   8790
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   4
         X1              =   1470
         X2              =   4140
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   5
         X1              =   1470
         X2              =   4140
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   6
         X1              =   5805
         X2              =   8785
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   7
         X1              =   5805
         X2              =   8785
         Y1              =   1365
         Y2              =   1365
      End
   End
   Begin VB.CommandButton cmdYBBalance 
      Caption         =   "医保结算(&Y)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9210
      TabIndex        =   28
      ToolTipText     =   "热键：F2"
      Top             =   345
      Width           =   2055
   End
   Begin VB.CommandButton cmdInsureSet 
      Caption         =   "险类设置(&I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   31
      Top             =   3270
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&P)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   32
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   29
      ToolTipText     =   "热键：F2"
      Top             =   375
      Width           =   2055
   End
   Begin VB.PictureBox pic误差 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   9210
      ScaleHeight     =   810
      ScaleWidth      =   2040
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   2040
      Begin VB.Label lbl 
         Caption         =   "本次误差"
         Height          =   315
         Index           =   13
         Left            =   105
         TabIndex        =   26
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.0111"
         Height          =   315
         Index           =   14
         Left            =   885
         TabIndex        =   27
         Top             =   420
         Width           =   1080
      End
   End
   Begin VB.PictureBox pic剩余自付 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   60
      ScaleHeight     =   1365
      ScaleWidth      =   3360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1590
      Width           =   3390
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Index           =   8
         Left            =   2235
         TabIndex        =   12
         Top             =   585
         Width           =   1005
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTitle 
         Height          =   450
         Left            =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         Width           =   3345
         _Version        =   589884
         _ExtentX        =   5900
         _ExtentY        =   794
         _StockProps     =   6
         Caption         =   "当前未付"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox pic自付合计 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   60
      ScaleHeight     =   1320
      ScaleWidth      =   3360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3030
      Width           =   3390
      Begin XtremeSuiteControls.ShortcutCaption stcTitleTotal 
         Height          =   420
         Left            =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   3345
         _Version        =   589884
         _ExtentX        =   5900
         _ExtentY        =   741
         _StockProps     =   6
         Caption         =   "自付合计"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   9
         Left            =   2220
         TabIndex        =   15
         Top             =   600
         Width           =   1005
      End
   End
   Begin MSCommLib.MSComm mscCom 
      Left            =   11160
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picPayInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3510
      ScaleHeight     =   2745
      ScaleWidth      =   5445
      TabIndex        =   16
      Top             =   1590
      Width           =   5475
      Begin VB.TextBox txt摘要 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1290
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1320
         Width           =   3960
      End
      Begin VB.TextBox txt金额 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   764
         Width           =   2100
      End
      Begin VB.ComboBox cbo支付方式 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   772
         Width           =   1395
      End
      Begin VB.TextBox txt冲预交 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   18
         Top             =   210
         Width           =   3960
      End
      Begin zlIDKind.ucQRCodePayButton btQRCodePay 
         Height          =   450
         Left            =   4800
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "扫码付允许使用快键【F3】进行快速支付"
         Top             =   744
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         Appearance      =   2
         ToolTipString   =   "扫码付允许使用快键【F3】进行快速支付"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "摘  要"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   285
         TabIndex        =   23
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预存款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   285
         TabIndex        =   17
         Top             =   285
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴  款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   285
         TabIndex        =   19
         Top             =   832
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   30
      Top             =   930
      Width           =   2055
   End
   Begin VB.PictureBox picBlanceAndFee 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   60
      ScaleHeight     =   2460
      ScaleWidth      =   11205
      TabIndex        =   33
      Top             =   4440
      Width           =   11235
      Begin VB.PictureBox picFee 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   300
         ScaleHeight     =   1470
         ScaleWidth      =   10275
         TabIndex        =   40
         Top             =   780
         Width           =   10275
         Begin VSFlex8Ctl.VSFlexGrid vsFee 
            Height          =   1125
            Left            =   300
            TabIndex        =   41
            Top             =   270
            Width           =   10065
            _cx             =   17754
            _cy             =   1984
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
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
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSquareAffirm.frx":0442
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
            AllowUserFreezing=   1
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picBlance 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   30
         ScaleHeight     =   2415
         ScaleWidth      =   10995
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   150
         Width           =   10995
         Begin VSFlex8Ctl.VSFlexGrid vsBalance 
            Height          =   2295
            Left            =   0
            TabIndex        =   39
            Top             =   420
            Width           =   10125
            _cx             =   17859
            _cy             =   4048
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
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
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSquareAffirm.frx":05A5
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
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "医保支付:99999999.99"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   17
            Left            =   6690
            TabIndex        =   38
            Top             =   90
            Width           =   2640
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "费用合计:99999999.99"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   15
            Left            =   60
            TabIndex        =   36
            Top             =   90
            Width           =   2640
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "已付合计:99999999.99"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   16
            Left            =   3240
            TabIndex        =   37
            Top             =   90
            Width           =   2640
         End
      End
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1125
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   2355
         _Version        =   589884
         _ExtentX        =   4154
         _ExtentY        =   1984
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   42
      Top             =   7080
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmSquareAffirm.frx":06BC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11721
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Key             =   "个人帐户显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmSquareAffirm.frx":0F50
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1693
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
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSquareAffirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'入参变量
Private mfrMain As Object
Private mbytBillType As Byte '0-不区分收费或记帐单,1-收费记录;2-记帐记录
Private mlngModule As Long, mlngPatiID As Long
Private mstrInNos As String, mstr医嘱IDs As String, mstrPrivs As String
Private mlngCardTypeID As Long, mbln消费卡 As Boolean
Private mblnCliniqueRoomPay As Boolean  '诊间支付
Private mbln使用预交 As Boolean '是否允许使用预交款,104381
'---------------------------------------------------------------------
'模块变量
Private mlng结帐ID As Long, mblnOk As Boolean
Private mobjPayCards As Cards
Private mrsInfo As ADODB.Recordset
Private mblnFirst As Boolean
Private mstrTittle As String '窗体标题

'---------------------------------------------------------------------
'模块参数
Private mintFeePrecision  As Integer
Private mbytFeeMoneyPrecision  As Byte
Private Type Ty_Para
    int审核票据格式 As Integer
    int收费票据格式 As Integer
    int审核打印方式 As Integer
    int收费打印方式 As Integer
    int药品单位 As Integer
End Type
Private mbytCurType As Byte '1-门诊收费;2-门诊记帐
Private mPara As Ty_Para
Private mbln只对医保结算成功单据收费 As Boolean

Public mbln门诊自动发料 As Boolean '记帐划价单审核后自动发料

'常量值
Private Enum Pg_Index
    Blance = 0
    FeeDetail
End Enum

Private Enum Lbl_Index
    姓名 = 1
    性别 = 2
    门诊号 = 3
    预交余额 = 4
    未结费用 = 5
    剩余款额 = 6
    家属余额 = 7
    当前未付 = 8
    自付合计 = 9
    预存款 = 10
    缴款 = 11
    摘要 = 12
    误差 = 14
    费用合计 = 15
    已付合计 = 16
    医保支付 = 17
End Enum

Private Enum Pan
    C2提示信息 = 2
    C3个人帐户 = 3
End Enum
'----------------------------------------------------------------------------
'结算数据
Private mrs结算方式 As ADODB.Recordset
Private Type TY_ChargeMoney
    dbl费用合计 As Double
    dbl本次冲预交  As Double
    dbl医保支付 As Double
    dbl已付合计 As Double
    dbl当前未付 As Double
    dbl本次误差费 As Double
    
    dbl预交余额 As Double
    dbl费用余额 As Double
    dbl可用预交 As Double
    
    lng结帐ID As Long
    lng结算序号 As Long
End Type
Private mCurCharge As TY_ChargeMoney
'------------------------------------------------------------------------------------------
'卡支付相关
Private Type TY_PayMoney
    lng卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    strQRCode As String
    str交易流水号 As String
    str交易说明 As String
    bln读卡 As Boolean
    bln卡号密文  As Boolean
    int医疗卡长度 As Integer
    bln支票 As Boolean
    bln自制卡 As Boolean
    blnOneCard As Boolean '是否一卡通结算
    int性质 As Integer '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算;<0 表示第三方支付
    strNo As String
    lngID As Long '预交ID
    objCard As Card
    str支付结算 As String '字符串，格式：结算方式|支付金额||...
End Type
Private mCurCardPay As TY_PayMoney '本次卡支付
Private mcllSquareBalance As Collection '消费卡结算
Private mobjThreeSwap As clsThreeSwap

Private mstr家属IDs As String '病人家属ID,79868
Private mbyt预存款消费验卡 As Byte '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Private mdblBrushCardMoney As Double
'----------------------------------------------------------------------------
Private mstrCurNos As String
Private mrsFeeData As ADODB.Recordset   '记录本次刷卡消费的数据
Private mobjBalanceBills As BalanceBills '注意：单据顺序必须与 mstrCurNos 的顺序一致
Private mblnCommitData As Boolean
Private mblnSaveBill As Boolean
Private mblnCommitBill As Boolean
Private mblnYbBalanced As Boolean
'----------------------------------------------------------------------------
'医保相关
Private Type TY_Insure
    intInsure As Integer
    strYBPati As String 'New:空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    dbl个帐余额 As Double '当前病人个人帐户余额
    dbl个帐透支 As Double '个人帐户允许透支金额
    colBalance As Collection '记录各张单据保险结算原始值及修改值,元素:BalanceMoneys
    
    strAllNos As String '原提取出的单据，可能部分结算成功
End Type
Private mInsure As TY_Insure '本次卡支付
Private mstr个人帐户 As String '是否将个人帐户设置到收费可用
Private mInsurePara As Ty_InsurePara

Private mclsExpenceSvr As Object 'zlPublicExpense.clsExpenceSvr
Private mobjDrugStuff As clsDrugStuff

Public Function zlSquareAffirm(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    Optional ByVal lngPatiID As Long = 0, _
    Optional ByVal lngCardTypeID As Long = 0, _
    Optional ByVal bln消费卡 As Boolean = False, _
    Optional ByVal blnCliniqueRoomPay As Boolean = False, _
    Optional ByVal bytBillType As Byte, _
    Optional ByVal strNOs As String = "", _
    Optional ByVal str医嘱IDs As String = "", _
    Optional ByRef strExpand As String = "", _
    Optional ByRef lng结帐ID As Long = 0, _
    Optional ByVal bln使用预交 As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 消费确认接口 , 主要是应用于病人在各消费环境进行消费确认
    '入参:frmMain-传入调用对象
    '       lngModule:调用的模块号
    '       strPrivs:权限串
    '       lngPatiID :病人ID,可以不传,在本接口窗体中刷卡!
    '       lngCardTypeID   Long    In  卡类别ID(消费卡为消费接口序号):0为不区分;在确认窗口中处理 目前 , 只有在预交款缴款中使用,传入后,支付方式缺省为该方式.
    '       bln消费卡   Boolean In  缺省为Fase,表示是否消费卡结算
    '       bytBillType:单据类别: 0-不区分收费或记帐单,1-收费记录;2-记帐记录
    '       strNOs:格式为( 单据1,单据2),配合BytBillType单据类型使用.一次只能使用一种性质
    '                   如:  A0001,A002,A003…;
    '       str医嘱IDs:格式为:ID1,ID2,...
    '       strCardNO-主界面中刷的卡号
    '       blnCliniqueRoomPay-诊间支付(诊间支付不弹出刷卡界面),诊间支付时，只针对收费性质
    '       bln使用预交-是否允许使用预交：Ture，允许使用预交款，且存在预交款时缺省使用预交款；False，不允许使用预交款，必须要有启用的三方帐户
    '出参:
    '返回:Boolean 返回    成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-06-15 09:53:37
    '说明:
    '      如果strNos和str医嘱IDs都没传,只是对指定病人的门诊收费划价单收费和门诊记帐划价进行审核.
    '      如果病人ID不传入,则需要在窗体中先进行刷卡找到病人后,再进行消费确认.
    '调用者:
    '    1.  检查;检验;药房等.
    '    2.  其他所有需要进行消费确认的地方都应该调用该接口.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln门诊留观预交使用预交款 As Boolean
    On Error GoTo errHandle
    Set mfrMain = frmMain
    mlngModule = lngModule: mlngPatiID = lngPatiID: mstrPrivs = strPrivs
    mstrInNos = strNOs: mstr医嘱IDs = str医嘱IDs
    mbytBillType = bytBillType: mlngCardTypeID = lngCardTypeID
    mblnCliniqueRoomPay = blnCliniqueRoomPay
    
    strExpand = "": mlng结帐ID = 0
    mblnOk = False: mstr家属IDs = ""
    
    
    bln门诊留观预交使用预交款 = Val(zlDatabase.GetPara(323, glngSys)) <> 1
    If zlCheckCurPatiIsMzLg(lngPatiID) Then     '门诊留观病人使用预交款
       bln使用预交 = bln门诊留观预交使用预交款
    End If
    mbln使用预交 = bln使用预交
    
    Call InitVariableData
    
    Set mrsFeeData = GetFeeData(lngPatiID)
    If mrsFeeData Is Nothing Then Exit Function
    If mrsFeeData.State <> 1 Then Exit Function
    If mrsFeeData.RecordCount = 0 Then zlSquareAffirm = True: Exit Function
    
    If CreateOneCardComLib(frmMain, lngModule) = False Then Exit Function
    If CreateExpenceSvr(mclsExpenceSvr, lngModule) = False Then Exit Function
    
    Set mobjDrugStuff = New clsDrugStuff
    If mobjDrugStuff.InitCommon(mlngModule, mstrPrivs, mblnCliniqueRoomPay) = False Then Exit Function
    
    Call zlInitPriceGrade '初始化价格等级
    
    Call InitPara
    If GetPatient(mlngPatiID) = False Then Exit Function
    If InitThreeSwap(frmMain) = False Then Exit Function
    
    If mblnCliniqueRoomPay Then
        If CliniqueRoomPayValied = False Then Exit Function
        If ExecuteCliniqueRoomPay(frmMain) = False Then Exit Function
        lng结帐ID = mlng结帐ID
        zlSquareAffirm = True
        Exit Function
    End If
    
    On Error Resume Next
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng结帐ID = mlng结帐ID
    zlSquareAffirm = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitVariableData()
    '初始化模块变量
    Dim tyInsureTmp As TY_Insure
    Dim tyChargeTmp As TY_ChargeMoney
    
    mblnYbBalanced = False
    mblnCommitData = False
    mblnSaveBill = False
    mblnCommitBill = False
    
    mInsure = tyInsureTmp
    mCurCharge = tyChargeTmp
End Sub

Private Function CreateLocalTypeObject(ByVal lngCardTypeID As Long) As Boolean
    '功能:创建指定卡类别对象
    '入参:
    '   lngCardTypeID-卡类别ID
    '返回:创建成功返回true,否则返回False
    Dim objCard As Card, blnReturn As Boolean
    Dim tyTemp As TY_PayMoney
    
    On Error GoTo ErrHandler
    blnReturn = gobjOneCardComLib.zlGetCard(lngCardTypeID, False, objCard)
    If blnReturn = False Or objCard Is Nothing Then
        ShowMsgbox "未找到指定的三方帐户所支持的卡类别，可能该类别未启用，请检查［医疗卡类别］。"
        Exit Function
    End If
    If objCard.启用 = False Then
        ShowMsgbox objCard.名称 & "未启用，请检查。"
        Exit Function
    End If
    If objCard.是否存在帐户 = False Then
        ShowMsgbox objCard.名称 & "未设置三方帐户，请检查［医疗卡类别］。"
        Exit Function
    End If
    If objCard.结算方式 = "" Then
        ShowMsgbox objCard.名称 & "未设置结算方式，请检查［医疗卡类别］。"
        Exit Function
    End If
    If objCard.接口程序名 = "" Then
        ShowMsgbox objCard.名称 & "未设置三方接口所支持的部件，请检查［医疗卡类别］。"
        Exit Function
    End If
    
    mCurCardPay = tyTemp
    With mCurCardPay
       .lng卡类别ID = objCard.接口序号
       .bln消费卡 = objCard.消费卡
       .str结算方式 = objCard.结算方式
       .str名称 = objCard.名称
    End With
    CreateLocalTypeObject = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function CliniqueRoomPayValied() As Boolean
    '功能:诊间支付检查
    '返回:合法返回true,否则返回False
    
    On Error GoTo ErrHandler
    If mbytBillType <> 1 Then   '只针对收费单
        ShowMsgbox "诊间支付时，不允许针对记帐单据进行支付。"
        Exit Function
    End If
    If mlngCardTypeID = 0 Then
        ShowMsgbox "诊间支付时要求指定一个有效的三方帐户支付类别。"
        Exit Function
    End If
 
    '对象创建失败的,不允许支付
    If Not CreateLocalTypeObject(mlngCardTypeID) Then Exit Function
    CliniqueRoomPayValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteCliniqueRoomPay(frmMain As Object) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:诊间支付
    '返回:诊间支付成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-01-14 17:28:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curMoney As Currency, tyTmp As TY_ChargeMoney
    Dim strPrintNo As String '格式：'A001','A002',...
    
    On Error GoTo errHandle
    mCurCharge = tyTmp
    mbytCurType = 1
    mstrCurNos = GetCurFeeNos(mbytCurType)
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl费用合计 = curMoney
    Call Cacl误差金额
    
    If isValied() = False Then Exit Function
    '保存数据
    If SaveCharge(strPrintNo) = False Then Exit Function
    
    Call PrintBill(strPrintNo)
    '银医一卡通写卡，85950
    Call WriteInforToCard(frmMain, mlngModule, mstrPrivs, 0, strPrintNo)
    
    ExecuteCliniqueRoomPay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeData(ByVal lng病人ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次收取的费用数据
    '返回:获取费用数据
    '编制:刘兴洪
    '日期:2011-09-14 20:09:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTableNos As String, strTableIDs As String
    Dim varPara() As Variant, strSubTable As String
    Dim rsTemp As ADODB.Recordset
    Dim strSfTable As String, strJzTable As String
    
    On Error GoTo ErrHandler
    If lng病人ID = 0 Then Exit Function
    ReDim Preserve varPara(0 To 1) As Variant
    
    varPara(0) = lng病人ID: varPara(1) = mbytBillType
    
    If mstr医嘱IDs <> "" Then
        If zlGetSubTable(0, mstr医嘱IDs, strTableIDs, varPara(), 2) = False Then Exit Function
    End If
    If mstrInNos <> "" Then
        If zlGetSubTable(1, mstrInNos, strTableNos, varPara(), UBound(varPara) + 1) = False Then Exit Function
    End If
 
    If mstr医嘱IDs <> "" And mstrInNos <> "" Then
        strSubTable = " With  医嘱  As (" & strTableIDs & "),单据 as (" & strTableNos & ")"
    ElseIf mstr医嘱IDs <> "" Then
        strSubTable = " With  医嘱  As (" & strTableIDs & ") "
    ElseIf strTableNos <> "" Then
        strSubTable = " With   单据 as (" & strTableNos & ")"
    End If
    '110421:李南春,2017/6/23,费用执行时应使用价格父号而不是从属父号
    strSfTable = "": strJzTable = ""
    If mbytBillType <= 1 Then
        strSfTable = "" & _
        "Select '收费' As 类别, a.记录性质, a.执行部门ID, a.发药窗口, a.病人ID, " & vbNewLine & _
        "       a.NO, nvl(A.价格父号,A.序号) As 序号," & _
        "       b.编码||'-'||Decode(Decode(J1.诊疗类别||':'||J1.医嘱内容,'7:***',1,0), 1, '***', B.名称) As 项目," & vbNewLine & _
        "       b.规格, nvl(a.付数,1)*a.数次 As 数次, b.计算单位, a.收费细目ID, a.标准单价, a.应收金额, a.实收金额," & vbNewLine & _
        "       a.收费类别, a.登记时间, a.门诊标志,a.付款方式, a.病人科室ID, a.开单部门ID, a.是否急诊, a.保险项目否, a.统筹金额" & vbNewLine & _
        "From 门诊费用记录 A,收费项目目录 B ,病人医嘱记录 J1" & IIf(mstrInNos <> "", " ,单据 C", "") & vbNewLine & _
        "Where a.收费细目ID=b.ID And a.记录性质=1 And a.病人ID=[1] And (a.记录状态=0 Or a.记录状态=1 And a.结帐ID Is Null) " & _
        "      And a.医嘱序号=J1.id(+) "
        If mstr医嘱IDs <> "" And mstrInNos <> "" Then
            '问题:49593
            strSfTable = strSfTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=1  ))"
        ElseIf mstr医嘱IDs <> "" Then
            strSfTable = strSfTable & " And  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=1)"
        ElseIf strTableNos <> "" Then
            strSfTable = strSfTable & " And A.NO= C.Column_Value  "
        End If
    End If
    If mbytBillType = 2 Or mbytBillType = 0 Then
        strJzTable = "" & _
        "Select '记帐' As 类别,A.记录性质,A.执行部门ID,A.发药窗口,A.病人ID, " & vbNewLine & _
        "       a.NO, nvl(A.价格父号,A.序号) As 序号," & _
        "       b.编码||'-'||Decode(Decode(J1.诊疗类别||':'||J1.医嘱内容,'7:***',1,0), 1, '***', B.名称) As 项目," & vbNewLine & _
        "       b.规格, nvl(a.付数,1)*a.数次 As 数次, b.计算单位, a.收费细目ID, a.标准单价, a.应收金额, a.实收金额," & vbNewLine & _
        "       a.收费类别, a.登记时间, a.门诊标志, a.付款方式, a.病人科室ID, a.开单部门ID, a.是否急诊, a.保险项目否, a.统筹金额" & vbNewLine & _
        "From 门诊费用记录 A,收费项目目录 B,病人医嘱记录 J1" & IIf(mstrInNos <> "", " ,单据 C", "") & vbNewLine & _
        "Where a.收费细目ID=B.ID And a.记录性质=2 And a.病人ID=[1] And a.记录状态=0 " & _
        "      And a.医嘱序号=J1.id(+) "
        If mstr医嘱IDs <> "" And mstrInNos <> "" Then
            '问题:49593
            strJzTable = strJzTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=2  ))"
        ElseIf mstr医嘱IDs <> "" Then
            strJzTable = strJzTable & " And   A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=2  ) "
        ElseIf strTableNos <> "" Then
            strJzTable = strJzTable & " And A.NO= C.Column_Value "
        End If
        If strSfTable <> "" Then strJzTable = vbCrLf & " Union all   " & vbCrLf & strJzTable
    End If
    strSQL = strSubTable & vbCrLf & strSfTable & vbCrLf & strJzTable
    strSQL = "  Select * From (" & strSQL & ") Order by 记录性质,NO,序号"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "获取病人费用信息", varPara)
    Set GetFeeData = rsTemp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function LoadFeeData(ByVal bytType As Byte, Optional ByVal strNOs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费用数据
    ' 参数:
    '   bytType-1-门诊收费;2-记帐
    '   strNos - 格式：A001,A002,...
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-15 14:33:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo ErrHandler
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "记录性质=" & bytType
    If mrsFeeData.RecordCount = 0 Then Exit Function
    
    mrsFeeData.Sort = "NO,序号"
    With vsFee
        .Redraw = flexRDNone
        .Clear 1: .Rows = 1
        i = 1
        Do While Not mrsFeeData.EOF
            If strNOs = "" Or InStr("," & strNOs & ",", "," & nvl(mrsFeeData!NO) & ",") > 0 Then
                If i > .Rows - 1 Then .Rows = .Rows + 1
                .RowData(i) = Val(nvl(mrsFeeData!序号))
                .TextMatrix(i, .ColIndex("类别")) = nvl(mrsFeeData!类别)
                .Cell(flexcpData, i, .ColIndex("类别")) = Val(nvl(mrsFeeData!记录性质))
                .TextMatrix(i, .ColIndex("单据号")) = nvl(mrsFeeData!NO)
                .Cell(flexcpData, i, .ColIndex("单据号")) = Trim(nvl(mrsFeeData!收费类别))
                .TextMatrix(i, .ColIndex("项目")) = nvl(mrsFeeData!项目)
                .TextMatrix(i, .ColIndex("规格")) = nvl(mrsFeeData!规格)
                .TextMatrix(i, .ColIndex("数次")) = FormatEx(Val(nvl(mrsFeeData!数次)), 5)
                .TextMatrix(i, .ColIndex("单位")) = nvl(mrsFeeData!计算单位)
                .TextMatrix(i, .ColIndex("单价")) = FormatEx(Val(nvl(mrsFeeData!标准单价)), mintFeePrecision, , True)
                .TextMatrix(i, .ColIndex("应收金额")) = FormatEx(Val(nvl(mrsFeeData!应收金额)), mbytFeeMoneyPrecision, , True)
                .TextMatrix(i, .ColIndex("实收金额")) = FormatEx(Val(nvl(mrsFeeData!实收金额)), mbytFeeMoneyPrecision, , True)
                .Cell(flexcpData, i, .ColIndex("实收金额")) = Val(nvl(mrsFeeData!实收金额))
                .TextMatrix(i, .ColIndex("门诊标志")) = Val(nvl(mrsFeeData!门诊标志))
                
                i = i + 1
            End If
            mrsFeeData.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    
    LoadFeeData = True
    Exit Function
ErrHandler:
    vsFee.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetButtonVisible()
    '设置按钮的显示状态
    
    cmdYB.Visible = mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "保险收费") _
        And (mInsure.intInsure = 0 Or mInsure.intInsure <> 0 And Not mblnYbBalanced)
    '医保且医保未进行结算时,才显示
    cmdYBBalance.Visible = mInsure.intInsure <> 0 And Not mblnYbBalanced
    '医保进行结算了的,或非医保的,显示完成收费
    cmdOK.Visible = (mInsure.intInsure = 0 Or mInsure.intInsure <> 0 And mblnYbBalanced)
    cmdInsureSet.Visible = mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "保险收费") And mInsure.intInsure = 0
End Sub

Private Sub SetControlProperty()
    '设置控件属性
    On Error GoTo ErrHandler
    Call SetButtonVisible
    Call Cacl误差金额
    
    lbl(Lbl_Index.自付合计).Caption = FormatEx(mCurCharge.dbl费用合计 - mCurCharge.dbl医保支付, 6, , , 2)
    lbl(Lbl_Index.当前未付).Caption = Format(mCurCharge.dbl当前未付, "0.00")
    
    lbl(Lbl_Index.费用合计).Caption = "费用合计:" & FormatEx(mCurCharge.dbl费用合计, 6, , , 2)
    lbl(Lbl_Index.已付合计).Caption = "已付合计:" & Format(mCurCharge.dbl已付合计, "0.00")
    lbl(Lbl_Index.医保支付).Caption = "医保支付:" & Format(mCurCharge.dbl医保支付, "0.00")
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cacl误差金额()
    '显示误差金额
    Dim dblMoney As Double
    
    On Error GoTo ErrHandler
    mCurCharge.dbl本次误差费 = 0
    mCurCharge.dbl当前未付 = RoundEx(mCurCharge.dbl费用合计 - mCurCharge.dbl已付合计, 6)
    
    dblMoney = RoundEx(mCurCharge.dbl当前未付, 2)
    mCurCharge.dbl本次误差费 = RoundEx(mCurCharge.dbl当前未付 - dblMoney, 6)
    mCurCharge.dbl当前未付 = RoundEx(mCurCharge.dbl当前未付 - mCurCharge.dbl本次误差费, 6)
    
    If mblnCliniqueRoomPay Then Exit Sub
    
    pic误差.Visible = RoundEx(mCurCharge.dbl本次误差费, 6) <> 0
    lbl(Lbl_Index.误差).Caption = FormatEx(mCurCharge.dbl本次误差费, 6, , , 2)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ClearData()
    '功能:清除界面数据
    lbl(Lbl_Index.姓名).Caption = ""
    lbl(Lbl_Index.性别).Caption = "性别:"
    lbl(Lbl_Index.门诊号).Caption = "门诊号:"
    
    lbl(Lbl_Index.预交余额).Caption = "预交余额:0.00"
    lbl(Lbl_Index.未结费用).Caption = "未结费用:0.00"
    lbl(Lbl_Index.剩余款额).Caption = "剩余款额:0.00"
    lbl(Lbl_Index.家属余额).Caption = "家属余额:0.00"
    
    lbl(Lbl_Index.家属余额).Visible = False
    lineUnder(Lbl_Index.家属余额).Visible = False
    
    lbl(Lbl_Index.当前未付).Caption = "0.00"
    lbl(Lbl_Index.自付合计).Caption = "0.00"
    lbl(Lbl_Index.误差).Caption = "0.00"
    
    lbl(Lbl_Index.费用合计).Caption = "费用合计:0.00"
    lbl(Lbl_Index.已付合计).Caption = "已付合计:0.00"
    lbl(Lbl_Index.医保支付).Caption = "医保支付:0.00"
    
    txt冲预交.Text = "0.00"
    txt金额.Text = "0.00"
    txt摘要.Text = ""
    
    vsFee.Clear 1: vsFee.Rows = 2
    vsBalance.Clear 1: vsBalance.Rows = 2
    vsBalance.RowData(1) = ""
    
    staThis.Panels(Pan.C2提示信息).Text = ""
    staThis.Panels(Pan.C3个人帐户).Text = ""
    staThis.Panels(Pan.C3个人帐户).Visible = False
End Sub

Private Sub SetControlMove()
    '功能:设置控件属性
    Dim sngTop As Single, sngSplitHeight As Single, bln预交 As Boolean
    Dim sngHeght As Single
    
    sngTop = 200: sngSplitHeight = 80
    bln预交 = mCurCharge.dbl可用预交 <> 0 Or cbo支付方式.ListCount = 0
    If mbytCurType = 1 And cbo支付方式.ListCount > 0 Then
        lbl(Lbl_Index.预存款).Visible = bln预交: txt冲预交.Visible = bln预交
        If bln预交 Then
            txt冲预交.Top = sngTop: sngTop = txt冲预交.Top + txt冲预交.Height + sngSplitHeight
        End If
        cbo支付方式.Top = sngTop: sngTop = cbo支付方式.Top + cbo支付方式.Height + sngSplitHeight
        txt金额.Top = cbo支付方式.Top: btQRCodePay.Top = txt金额.Top - 20
        
        txt金额.Width = txt冲预交.Left + txt冲预交.Width - txt金额.Left - IIf(mbytCurType = 1 And btQRCodePay.Tag <> "", btQRCodePay.Width + 10, 0)
    
        txt摘要.Top = sngTop: txt摘要.Height = picPayInfo.ScaleHeight - txt摘要.Top - sngSplitHeight
        
        lbl(Lbl_Index.预存款).Top = txt冲预交.Top + (txt冲预交.Height - lbl(Lbl_Index.预存款).Height) \ 2
        lbl(Lbl_Index.缴款).Top = cbo支付方式.Top + (cbo支付方式.Height - lbl(Lbl_Index.缴款).Height) \ 2
        lbl(Lbl_Index.摘要).Top = txt摘要.Top + 45
        Exit Sub
    End If
    
    sngHeght = picPayInfo.ScaleHeight
    sngHeght = sngHeght - txt冲预交.Height
    sngTop = sngHeght / 2
    txt冲预交.Top = sngTop
    lbl(Lbl_Index.预存款).Top = txt冲预交.Top + (txt冲预交.Height - lbl(Lbl_Index.预存款).Height) \ 2
    
    lbl(Lbl_Index.预存款).Visible = True: txt冲预交.Visible = True
    lbl(Lbl_Index.缴款).Visible = False: cbo支付方式.Visible = False: txt金额.Visible = False
    lbl(Lbl_Index.摘要).Visible = False: txt摘要.Visible = False
End Sub

Private Sub cbo支付方式_Click()
    Dim tyTemp As TY_PayMoney
    Dim objCard As Card
    
    On Error GoTo ErrHandler
    mCurCardPay = tyTemp
    '记帐不处理
    If mbytCurType = 2 Then Exit Sub
    
    Call GetCurCard(objCard)
    If objCard Is Nothing Then Exit Sub
    
    With mCurCardPay
        .lng卡类别ID = objCard.接口序号
        .bln消费卡 = objCard.消费卡
        .str结算方式 = objCard.结算方式
        .str名称 = objCard.名称
        .bln自制卡 = objCard.自制卡
     End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdInsureSet_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub cmdYB_Click()
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.RecordCount = 0 Then Exit Sub
    
    Call MCPatientProcess(mrsInfo!病人ID)
End Sub

Private Function YBIdentifyCancel() As Boolean
    '取消医保病人身份验证
    Dim lng病人ID As Long
    
    YBIdentifyCancel = True
    If mInsure.intInsure = 0 Then Exit Function
    If mInsure.strYBPati = "" Then Exit Function
    If mblnYbBalanced Then Exit Function
    
    If UBound(Split(mInsure.strYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mInsure.strYBPati, ";")(8)) And Val(Split(mInsure.strYBPati, ";")(8)) <> 0 Then
            lng病人ID = Val(CLng(Split(mInsure.strYBPati, ";")(8)))
        End If
    End If
    If lng病人ID = 0 Then Exit Function
    
    YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng病人ID, mInsure.intInsure)
End Function

Private Sub MCPatientProcess(ByVal lngCur病人ID As Long)
    Dim i As Long, blnTran As Boolean, strSQL As String
    Dim lng病人ID As Long, lng病人IDOut As Long, intInsure As Integer
    Dim cur透支额 As Currency, str医保号 As String
    Dim varNos As Variant, curMoney As Currency
    
    On Error GoTo ErrHandler
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State = 0 Then Exit Sub
    
    If gblnLED Then zl9LedVoice.Speak "#50"
    mInsure.dbl个帐余额 = 0: mInsure.dbl个帐透支 = 0
    lng病人IDOut = lngCur病人ID '避免Identify接口中修改该变量后返回新值
    
    '返回：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,24就诊类型(1=急诊门诊),25开单科室名称
    mInsure.strYBPati = gclsInsure.Identify(id门诊收费, lng病人IDOut, mInsure.intInsure)
    If mInsure.strYBPati = "" Then
        mInsure.intInsure = 0: Exit Sub
    End If
    
    '获取病人信息
    If UBound(Split(mInsure.strYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mInsure.strYBPati, ";")(8)) And Val(Split(mInsure.strYBPati, ";")(8)) <> 0 Then
            lng病人ID = Val(CLng(Split(mInsure.strYBPati, ";")(8)))
            If lng病人ID <> lngCur病人ID Then
                ShowMsgbox "医保验证的病人与当前病人不是同一个病人！"
                staThis.Panels(Pan.C2提示信息) = "医保验证的病人与当前病人不是同一个病人！"
                GoTo IdentifyCancel:
            End If
        End If
    End If

    mInsurePara = initInsurePara(mInsure.intInsure, lng病人ID)  '初始化医保参数
    
    '重新加载病人信息，可能医保接口中有变动
    Call GetPatient(mlngPatiID)
    Call LoadPatient
    Call ShowLedInfor
    
    lbl(Lbl_Index.姓名).ForeColor = vbRed
    If nvl(mrsInfo!病人类型) <> "" Then
        Call SetPatiColor(lbl(Lbl_Index.姓名), nvl(mrsInfo!病人类型), vbRed)
    End If
        
    '个人帐户
    str医保号 = CStr(Split(mInsure.strYBPati, ";")(1))
    mInsure.dbl个帐余额 = gclsInsure.SelfBalance(lng病人ID, str医保号, 10, cur透支额, mInsure.intInsure)
    staThis.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mInsure.dbl个帐余额, "0.00")
    staThis.Panels(Pan.C3个人帐户).Visible = True
    mInsure.dbl个帐透支 = cur透支额
        
    '计算已提取的划价单的相关保险数据
    varNos = Split(mstrCurNos, ",")
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(varNos)
        strSQL = "zl_门诊划价记录_Update(" & mInsure.intInsure & "," & lng病人ID & ",'" & varNos(i) & "',0)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    '加入预结算结果对象
    Set mInsure.colBalance = New Collection
    For i = 0 To UBound(varNos)
        mInsure.colBalance.Add New BalanceMoneys
    Next
    
    Set mrsFeeData = GetFeeData(lng病人ID) '重新读取费用信息
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl费用合计 = curMoney
    staThis.Panels(Pan.C2提示信息) = ""
    '直接进行预结算
    If 门诊预结算() = False Then GoTo IdentifyCancel:
    
    If mInsurePara.门诊预结算 Then
        Call InsureLedSpeak
    End If
    
    tbPage.Item(Pg_Index.Blance).Selected = True
    cmdYBBalance.Enabled = True
    Call SetControlProperty
    Call SetDefaultPrepayMoney
    Call SetCtlEnable(False)
    
    zlControl.ControlSetFocus vsBalance
    
    Exit Sub
IdentifyCancel:
    '取消医保验证
    Call YBIdentifyCancel
    mInsure.intInsure = 0: mInsure.strYBPati = ""
    
    Call SetPatiColor(lbl(Lbl_Index.姓名), nvl(mrsInfo!病人类型), &HFF0000)
    staThis.Panels(Pan.C3个人帐户).Text = ""
    staThis.Panels(Pan.C3个人帐户).Visible = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Function GetBalanceBills(ByVal bytType As Byte, ByVal strNOs As String, _
    Optional ByRef cur实收合计 As Currency) As BalanceBills
    '获取费用单据信息
    '入参：
    '   bytType 1-门诊收费;2-门诊记帐
    Dim objBalanceBill As BalanceBill, objBalanceBills As BalanceBills
    Dim varNos As Variant, strNo As String
    Dim cur实收金额 As Currency
    Dim p As Integer, i As Integer
    
    On Error GoTo ErrHandler
    Set objBalanceBills = New BalanceBills
    cur实收合计 = 0
    varNos = Split(strNOs, ",")
    For p = 1 To UBound(varNos) + 1
        strNo = varNos(p - 1)
        Set objBalanceBill = New BalanceBill
        objBalanceBill.NO = strNo
        
        mrsFeeData.Filter = "记录性质=" & bytType & " And No='" & strNo & "'"
        For i = 1 To mrsFeeData.RecordCount
            cur实收金额 = Val(nvl(mrsFeeData!实收金额))
            objBalanceBill.实收合计 = objBalanceBill.实收合计 + cur实收金额
            
            '统计保险金额
            If nvl(mrsFeeData!统筹金额, 0) = 0 Or Val(nvl(mrsFeeData!保险项目否)) = 0 Then
                '以原始金额为准,不管分币处理
                objBalanceBill.全自付 = objBalanceBill.全自付 + cur实收金额
            Else
                objBalanceBill.进入统筹 = objBalanceBill.进入统筹 + Val(nvl(mrsFeeData!统筹金额))
                '以原始金额为准,不管分币处理
                objBalanceBill.先自付 = objBalanceBill.先自付 + cur实收金额 - Val(nvl(mrsFeeData!统筹金额))
            End If
            
            cur实收合计 = cur实收合计 + cur实收金额
            mrsFeeData.MoveNext
        Next
        
        objBalanceBills.AddItem objBalanceBill, "K" & strNo
    Next
    Set GetBalanceBills = objBalanceBills
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 门诊预结算() As Boolean
    '功能：门诊预结算
    Dim bytMode As Byte
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim dbl合计 As Double, cur个帐预付 As Currency, cur可用个帐 As Currency, cur个帐支付 As Currency
    Dim objItem As BalanceMoney, strNone As String
    Dim strErrMsg As String
    Dim cur实收合计 As Currency
    
    On Error GoTo ErrHandler
    '初始化结算结果表格
    vsBalance.Clear 1: vsBalance.Rows = 2
    vsBalance.RowData(1) = ""
    If mInsure.intInsure = 0 Then Exit Function
    
    If mInsurePara.门诊预结算 = False Then
        If mstr个人帐户 = "" Then 门诊预结算 = True: Exit Function
        
        '计算当前单据个人帐户支付金额:不支持预结算时
        If mInsurePara.多单据分单据结算 Then
            For p = 1 To mobjBalanceBills.Count
                With mobjBalanceBills(p)
                    cur个帐预付 = .进入统筹 + IIf(mInsurePara.先自付, .先自付, 0) + IIf(mInsurePara.全自付, .全自付, 0)
                End With
                '统计除开之前单据个帐支付后的个帐余额
                cur可用个帐 = mInsure.dbl个帐余额
                For i = 1 To p - 1
                    cur可用个帐 = cur可用个帐 - GetMedicareSum(mInsure.colBalance, mstr个人帐户, i)
                Next
                
                cur个帐支付 = Get个帐报销金额(mobjBalanceBills(p).实收合计, cur个帐预付, cur可用个帐, mInsure.dbl个帐透支)
                Call SetBalanceVal(mInsure.colBalance, p, mstr个人帐户 & "|" & cur个帐支付)
            Next
        Else
            cur个帐预付 = 0: cur实收合计 = 0
            For i = 1 To mobjBalanceBills.Count
                With mobjBalanceBills(i)
                    cur个帐预付 = cur个帐预付 + .进入统筹 + IIf(mInsurePara.先自付, .先自付, 0) + IIf(mInsurePara.全自付, .全自付, 0)
                    cur实收合计 = cur实收合计 + mobjBalanceBills(i).实收合计
                End With
            Next
            cur可用个帐 = mInsure.dbl个帐余额
            
            cur个帐支付 = Get个帐报销金额(cur实收合计, cur个帐预付, cur可用个帐, mInsure.dbl个帐透支)
            Call SetBalanceVal(mInsure.colBalance, 1, mstr个人帐户 & "|" & cur个帐支付)
        End If
    Else
    
        If mInsurePara.实时监控 Then
            '本来对于划价单才传2进行明细和汇总的检查，但是，由于以下原因，数量和实收金额在输入检查通过后可能改变，所以须再次检查明细
            '1.导入单据，2.修改单据，3.输入中药配方，4.修改中药付数后，其它行的付数同时变化，5.输入主项，自动产生从项，以及从项汇总计算折扣
            '6.修改单价，7.调整执行科室，药品价格重算，8.调整费别，实收金额重算,9.先输费用再验证医保身份,其它等等
            If gclsInsure.CheckItem(mInsure.intInsure, 0, 2, MakeDetailRecord(mobjBalanceBills)) = False Then
                staThis.Panels(Pan.C2提示信息).Text = "费用项目检查失败！"
                Exit Function
            End If
        End If
    
        If mInsurePara.多单据分单据结算 Then
            bytMode = 2
        ElseIf mInsurePara.一次结算分单据退费 Then
            bytMode = 1
        Else
            bytMode = 0
        End If
        
        If ZlExecuteInsurePreSwap(bytMode, mobjBalanceBills, mInsure.intInsure, mInsure.colBalance, strErrMsg) = False Then
            staThis.Panels(Pan.C2提示信息).Text = strErrMsg
            Exit Function
        End If
        strNone = CheckInsureBalanceValid(mrs结算方式, mInsure.colBalance)
        If strNone <> "" Then
            ShowMsgbox "当前保险结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                "在门诊未设置，请先到结算方式管理中设置这些结算方式！"
            Exit Function
        End If
    End If
    
    '全部预结完后的处理
    '-----------------------------------------------------------
    '显示预结的表格结果
    For p = 1 To mInsure.colBalance.Count
        For Each objItem In mInsure.colBalance(p)
            With vsBalance
                '定位到匹配行或空行
                k = -1
                For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("支付方式")) = objItem.结算方式 Then
                        k = j: Exit For '记录已填写的匹配行
                    ElseIf .TextMatrix(j, .ColIndex("支付方式")) = "" Then
                        If k = -1 Then k = j '记录第一可用空行
                    End If
                Next
                If j > .Rows - 1 And k = -1 Then
                    .Rows = .Rows + 1
                    k = .Rows - 1
                End If
                
                '汇总该种结算方式的金额
                .TextMatrix(k, .ColIndex("支付方式")) = objItem.结算方式
                .TextMatrix(k, .ColIndex("支付金额")) = Format(Val(.TextMatrix(k, .ColIndex("支付金额"))) + objItem.原始金额, "0.00")
                dbl合计 = dbl合计 + Val(Format(objItem.原始金额, "0.00"))
                If .RowData(k) = 0 Then
                    '多张单据中,只要有一张允许修改,则汇总的允许修改
                    .RowData(k) = IIf(objItem.允许修改, 1, 0)
                End If
            End With
        Next
    Next
    mCurCharge.dbl医保支付 = dbl合计
    mCurCharge.dbl已付合计 = dbl合计
    
    门诊预结算 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InsureLedSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结Led报价
    '编制:刘兴洪
    '日期:2011-12-15 13:40:46
    '问题:44425
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double
    
    If Not gblnLED Then Exit Sub
    dbl个帐合计 = GetMedicareSum(mInsure.colBalance, mstr个人帐户)
    zl9LedVoice.DisplayBank "医保结算:", "帐户余额" & Format(mInsure.dbl个帐余额, "0.00"), _
        "帐户支付" & Format(dbl个帐合计, "0.00"), "统筹支付" & Format(GetMedicareSum(mInsure.colBalance) - dbl个帐合计, "0.00")
    zl9LedVoice.Speak "#21 " & Format(mCurCharge.dbl费用合计, "0.00")
End Sub

Private Sub LedDisplayBank(Optional ByVal blnSpeak As Boolean = True)
    '功能:显示报价信息
    '问题:52117
    Dim i As Long
    Dim str医保 As String, str三方交易 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String
    Dim strTemp As String
    
    If Not gblnLED Then Exit Sub
    With vsBalance
        For i = 1 To .Rows - 1
            '医保交易
            If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                strTemp = .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("支付金额"))), "0.00")
                Select Case .RowData(i)
                Case Enum_BalanceType.医保
                    str医保 = str医保 & "||" & strTemp
                Case Enum_BalanceType.一卡通
                    str三方交易 = str三方交易 & "||" & strTemp
                Case Enum_BalanceType.老一卡通
                    str老一卡通 = str老一卡通 & "||" & strTemp
                Case Else
                    str普通结算 = str普通结算 & "||" & strTemp
                End Select
            End If
        Next
    End With
     
    str结算方式 = ""
    If str医保 <> "" Then str结算方式 = str结算方式 & "||医保结算:||帐户余额:" & Format(mInsure.dbl个帐余额, "0.00") & str医保
    If str三方交易 <> "" Then str结算方式 = str结算方式 & "||一卡通结算:" & str三方交易
    If str老一卡通 <> "" Then str结算方式 = str结算方式 & "||一卡通结算(老):" & str老一卡通
    If str普通结算 <> "" Then str结算方式 = str结算方式 & "" & str普通结算
    If str结算方式 = "" Then Exit Sub
    
    str结算方式 = Mid(str结算方式, 3)
    varPara = Split(str结算方式, "||")
    '目前最多只能显示10个参数值
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str结算方式 = ""
        For i = 10 To UBound(varPara)
            str结算方式 = str结算方式 & ";" & varPara(i)
        Next
        If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str结算方式
    End Select

    If blnSpeak Then zl9LedVoice.Speak "#21 " & Format(mCurCharge.dbl当前未付, "0.00")
End Sub

Private Sub cmdYBBalance_Click()
    Dim blnSpeak As Boolean, dblOld未支付 As Currency
    
    On Error GoTo ErrHandler
    dblOld未支付 = mCurCharge.dbl当前未付
    '费用数据保存
    If SaveFeeBill() = False Then Exit Sub
    '处理医保数据
    If ExecuteInsureSwap() = False Then
        Call SetButtonVisible
        Exit Sub
    End If
    
    Call LoadBalancedData(mCurCharge.lng结帐ID)
    Call SetControlProperty
    Call SetButtonVisible
    Call SetCtlEnable
     
    Call SetDefaultPrepayMoney
    Call SetBeginFocus '光标定位
    
    blnSpeak = RoundEx(dblOld未支付, 6) <> RoundEx(mCurCharge.dbl当前未付, 6)
    Call LedDisplayBank(blnSpeak)
    If RoundEx(mCurCharge.dbl当前未付, 6) = 0 Then
        '医保全部结算,直接确定完成
        If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
    End If
    
    If RoundEx(mCurCharge.dbl当前未付, 6) < 0 Then
        '医保报销金额大于了费用总金额时，需要退款给病人
        MsgBox "    本次医保报销金额大于了费用总金额，无法完成结算。" & vbCrLf & _
            "请到收费窗口进行处理！", vbExclamation + vbOKOnly, gstrSysName
        Unload Me: Exit Sub
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveFeeBill() As Boolean
    '保存费用单据数据
    '说明:
    '   调用此过程时,不需要开始事务。异常时,数据回退；保存成功时,未提交数据
    Dim objBalanceBill As BalanceBill
    Dim blnTrans As Boolean, strSQL As String
    Dim str发生时间 As String, str发药窗口 As String
    Dim cllDept As Collection, int病人来源 As Integer
    Dim varNos As Variant, strNo As String, p As Integer
    Dim strErrMsg As String, i As Integer
    
    On Error GoTo ErrHandler
    If (mblnSaveBill And mblnCommitBill) Or mblnCommitData Then
        gcnOracle.BeginTrans
        SaveFeeBill = True: Exit Function
    End If
    
    '划价单划价收费检查
    varNos = Split(mstrCurNos, ",")
    For i = 0 To UBound(varNos)
        If mclsExpenceSvr.zlPriceChargeCheck(varNos(i), mlngPatiID, strErrMsg) = False Then
            MsgBox IIf(strErrMsg = "", "执行划价收费检查错误！", strErrMsg), vbInformation, gstrSysName
            Exit Function
        End If
    Next
    
    mCurCharge.lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    mCurCharge.lng结算序号 = -1 * mCurCharge.lng结帐ID
    str发生时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    int病人来源 = IIf(Val(nvl(mrsInfo!在院)) = 1, 2, 1)
    
    For p = 1 To UBound(varNos) + 1
        strNo = varNos(p - 1)
        mrsFeeData.Filter = "记录性质=" & mbytCurType & " And No='" & strNo & "'"
        If mrsFeeData.RecordCount <> 0 Then
            '发药窗口
            Set cllDept = New Collection
            Do While Not mrsFeeData.EOF
                If InStr(",5,6,7,", nvl(mrsFeeData!收费类别)) > 0 Then
                    cllDept.Add Array(nvl(mrsFeeData!收费类别), Val(nvl(mrsFeeData!执行部门ID)), nvl(mrsFeeData!发药窗口))
                End If
                mrsFeeData.MoveNext
            Loop
            str发药窗口 = GetPayDrugWindow(mlngPatiID, CDate(str发生时间), cllDept)
            
            mrsFeeData.MoveFirst
            'Zl_病人划价收费_Insert
            strSQL = "Zl_病人划价收费_Insert("
            '  No_In         门诊费用记录.NO%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  病人id_In     门诊费用记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(mlngPatiID) & ","
            '  病人来源_In   Number,
            strSQL = strSQL & "" & int病人来源 & ","
            '  付款方式_In   门诊费用记录.付款方式%Type,
            If nvl(mrsInfo!付款方式编码) <> "" Then
               strSQL = strSQL & "'" & nvl(mrsInfo!付款方式编码) & "',"
            Else
               strSQL = strSQL & "'" & nvl(mrsFeeData!付款方式) & "',"
            End If
            '  姓名_In       门诊费用记录.姓名%Type,
            strSQL = strSQL & "'" & nvl(mrsInfo!姓名) & "',"
            '  性别_In       门诊费用记录.性别%Type,
            strSQL = strSQL & "'" & nvl(mrsInfo!性别) & "',"
            '  年龄_In       门诊费用记录.年龄%Type,
            strSQL = strSQL & "'" & nvl(mrsInfo!年龄) & "',"
            '  病人科室id_In 门诊费用记录.病人科室id%Type,
            strSQL = strSQL & "" & ZVal(nvl(mrsFeeData!病人科室ID)) & ","
            '  开单部门id_In 门诊费用记录.开单部门id%Type,
            strSQL = strSQL & "" & ZVal(nvl(mrsFeeData!开单部门ID)) & ","
            '  开单人_In     门诊费用记录.开单人%Type,
            strSQL = strSQL & "NULL,"    ' 过程内部处理,保持原来的不变
            '  结帐id_In     门诊费用记录.结帐id%Type,
            strSQL = strSQL & "" & mCurCharge.lng结帐ID & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            strSQL = strSQL & "To_Date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
            strSQL = strSQL & "'" & str发药窗口 & "',"
            '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
            strSQL = strSQL & "" & Val(nvl(mrsFeeData!是否急诊)) & ","
            '  登记时间_In   门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "To_Date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'))"
            
            mobjBalanceBills("K" & strNo).划价收费SQL = strSQL
            If mInsure.intInsure <> 0 Then
                Set mobjBalanceBills("K" & strNo).预结算 = mInsure.colBalance(p)
            End If
        End If
    Next
    
    gcnOracle.BeginTrans: blnTrans = True
    For Each objBalanceBill In mobjBalanceBills
        zlDatabase.ExecuteProcedure objBalanceBill.划价收费SQL, Me.Caption
    Next
    
    mblnSaveBill = True
    SaveFeeBill = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteInsureSwap() As Boolean
    '医保结算
    Dim bytMode As Byte, blnCommit As Boolean
    Dim strErrMsg As String
    Dim strSavedNos As String, lngSavedBillCount As Long, blnYbBalanced As Boolean
    
    On Error GoTo ErrHandler
    If mInsure.intInsure = 0 Then ExecuteInsureSwap = True: Exit Function
    
    If mInsurePara.多单据分单据结算 Then
        bytMode = 2
    ElseIf mInsurePara.一次结算分单据退费 Then
        bytMode = 1
    Else
        bytMode = 0
    End If
    
    mInsure.strAllNos = ""
    If zlExecuteInsureSwap(bytMode, mlngPatiID, mInsure.intInsure, mstr个人帐户, _
        mbln只对医保结算成功单据收费, mCurCharge.lng结帐ID, mCurCharge.lng结算序号, _
        mobjBalanceBills, blnCommit, strSavedNos, lngSavedBillCount, blnYbBalanced, strErrMsg) = False Then
        If blnCommit = False Then
            If strErrMsg <> "" Then ShowMsgbox strErrMsg
            Exit Function
        End If
        
        mblnCommitBill = True
        '重新加载数据
        If blnYbBalanced Then
            mInsure.strAllNos = mstrCurNos
            mstrCurNos = strSavedNos
            Call LoadFeeData(mbytCurType, strSavedNos)
            
            mblnYbBalanced = True '医保已经结算
            ExecuteInsureSwap = True
        End If
    Else
        mblnCommitBill = True
        mblnYbBalanced = True '医保已经结算
        ExecuteInsureSwap = True
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearBanalce()
    '清除结算数据
    With mCurCharge
        .dbl费用合计 = 0
        .dbl医保支付 = 0
        .dbl已付合计 = 0
        .dbl当前未付 = 0
        .dbl本次冲预交 = 0
        .dbl本次误差费 = 0
    End With
End Sub

Private Sub AddNewRow()
    '结算表格新增一行
    With vsBalance
        If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
            .Rows = .Rows + 1
            .RowPosition(.Rows - 1) = 1
        End If
    End With
End Sub

Private Function LoadBalancedData(ByVal lng结帐ID As Long) As Boolean

    '加载已结算成功的结算数据
    Dim strSQL As String, rsBalance As ADODB.Recordset
    Dim bln消费卡 As Boolean, lng卡类别ID As Long
    Dim rsTypes As ADODB.Recordset
    Dim bln密文 As String, str卡类别名 As String
    On Error GoTo ErrHandler
    Call ClearBanalce
    vsBalance.Clear 1: vsBalance.Rows = 2
    vsBalance.RowData(1) = ""
    
    If gobjOneCardComLib Is Nothing Then Call CreateOneCardComLib(Me, mlngModule, gcnOracle)
    If Not gobjOneCardComLib Is Nothing Then
       Call gobjOneCardComLib.zlGetOneCardTypes(rsTypes)
    End If
    
    
    
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    Set rsBalance = GetChargeBalance(lng结帐ID)
    Do While Not rsBalance.EOF
        Select Case nvl(rsBalance!类型)
        Case Enum_BalanceType.预存款
            mCurCharge.dbl本次冲预交 = RoundEx(mCurCharge.dbl本次冲预交 + Val(nvl(rsBalance!冲预交)), 6)
            mCurCharge.dbl已付合计 = RoundEx(mCurCharge.dbl已付合计 + Val(nvl(rsBalance!冲预交)), 6)
        Case Else
            If nvl(rsBalance!类型) = Enum_BalanceType.医保 Then
                mCurCharge.dbl医保支付 = RoundEx(mCurCharge.dbl医保支付 + nvl(rsBalance!冲预交, 0), 6)
            End If
            
            If Val(nvl(rsBalance!校对标志)) = 2 Then
                bln消费卡 = nvl(rsBalance!类型) = Enum_BalanceType.消费卡
                If bln消费卡 Then
                    lng卡类别ID = Val(nvl(rsBalance!结算卡序号))
                Else
                    lng卡类别ID = Val(nvl(rsBalance!卡类别ID))
                End If
                
                With vsBalance
                    Call AddNewRow
                    .RowData(1) = nvl(rsBalance!类型)
                    .TextMatrix(1, .ColIndex("支付方式")) = nvl(rsBalance!结算方式)
                    str卡类别名 = nvl(rsBalance!卡类别名称, nvl(rsBalance!结算方式))
                    bln密文 = Val(nvl(rsBalance!是否密文)) = 1
                    If Not bln消费卡 And Not rsTypes Is Nothing Then
                        rsTypes.Filter = "ID=" & lng卡类别ID
                        If Not rsTypes.EOF Then
                            bln密文 = Val(nvl(rsTypes!是否密文)) = 1
                            str卡类别名 = nvl(rsTypes!名称)
                        End If
                    End If
                    '医疗卡类别ID|消费卡(1,0)|接口名称
                    .Cell(flexcpData, 1, .ColIndex("支付方式")) = lng卡类别ID & "|" & IIf(bln消费卡, 1, 0) & "|" & str卡类别名
                    .TextMatrix(1, .ColIndex("支付金额")) = Format(Val(nvl(rsBalance!冲预交)), "0.00")
                    .TextMatrix(1, .ColIndex("备注")) = nvl(rsBalance!摘要)
                    .TextMatrix(1, .ColIndex("交易流水号")) = nvl(rsBalance!交易流水号)
                    .TextMatrix(1, .ColIndex("交易说明")) = nvl(rsBalance!交易说明)
                    
                    .TextMatrix(1, .ColIndex("卡号")) = IIf(bln密文, String(Len(nvl(rsBalance!卡号)), "*"), nvl(rsBalance!卡号))
                    .Cell(flexcpData, 1, .ColIndex("卡号")) = nvl(rsBalance!卡号)
                    .TextMatrix(1, .ColIndex("结算状态")) = 1
                    .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = Me.BackColor
                End With
                mCurCharge.dbl已付合计 = RoundEx(mCurCharge.dbl已付合计 + Val(nvl(rsBalance!冲预交)), 6)
            End If
        End Select
        
        rsBalance.MoveNext
    Loop
                   
    strSQL = "Select Sum(b.实收金额) As 实收合计 From 门诊费用记录 B Where b.结帐id = [1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    mCurCharge.dbl费用合计 = Val(rsBalance!实收合计)
        
    If mCurCharge.dbl本次冲预交 <> 0 Then
        txt冲预交.Text = Format(mCurCharge.dbl本次冲预交, "0.00")
        txt冲预交.Tag = "1"
        txt冲预交.BackColor = Me.BackColor
        txt冲预交.Enabled = False
    End If
    LoadBalancedData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If mblnCliniqueRoomPay Then Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "保险收费") Then '136681
        mInsure.intInsure = GetCustomPatiInsure(mrsInfo!病人ID)
        If mInsure.intInsure <> 0 Then
            Call MCPatientProcess(mrsInfo!病人ID)
        End If
    End If
    
    Call SetBeginFocus '光标定位
    '78773:李南春,2014-10-29,LED显示一卡通支付信息
    Call ShowLedInfor
End Sub

Private Sub SetBeginFocus()
    '设置开始时的焦点位置
    If Val(txt冲预交.Text) <> 0 And mbln使用预交 Or cbo支付方式.ListCount = 0 Or mbytCurType = 2 Then
        zlControl.ControlSetFocus txt冲预交: zlControl.TxtSelAll txt冲预交
    Else
        zlControl.ControlSetFocus txt金额: zlControl.TxtSelAll txt金额
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF6
        If cmdYB.Visible And cmdYB.Enabled Then Call cmdYB_Click
    Case vbKeyF2
        '强制完成
        If mInsure.intInsure <> 0 And mblnYbBalanced = False Then
            If cmdYBBalance.Visible And cmdYBBalance.Enabled Then Call cmdYBBalance_Click
        Else
            If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
        End If
    Case vbKeyF4
        If Me.ActiveControl Is txt金额 And txt金额.Enabled Then
            If cbo支付方式.Visible = False Or cbo支付方式.Enabled = False Then Exit Sub
            If Shift = vbShiftMask Then
                If cbo支付方式.ListIndex - 1 < 0 Then
                    cbo支付方式.ListIndex = cbo支付方式.ListCount - 1
                Else
                    cbo支付方式.ListIndex = cbo支付方式.ListIndex - 1
                End If
            Else
                If cbo支付方式.ListIndex + 1 > cbo支付方式.ListCount - 1 Then
                    cbo支付方式.ListIndex = 0
                Else
                    cbo支付方式.ListIndex = cbo支付方式.ListIndex + 1
                End If
            End If
        End If
    Case vbKeyF3    '扫码付快键
        If btQRCodePay.Visible And btQRCodePay.Enabled Then Call btQRCodePay.zlReReadQRCode
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If YBIdentifyCancel() = False Then '取消医保病人身份验证,返回假时不退出
        Cancel = 1: Exit Sub
    End If
    
    '78773:李南春,2014-10-29,LED显示一卡通支付信息
    If gblnLED Then zl9LedVoice.DisplayPatient ""
    If Not mobjThreeSwap Is Nothing Then Set mobjThreeSwap = Nothing
    If Not mobjDrugStuff Is Nothing Then Set mobjDrugStuff = Nothing
    
    Set mobjPayCards = Nothing
    Set mrsInfo = Nothing
    Set mrs结算方式 = Nothing
    Set mrsFeeData = Nothing
    SaveWinState Me, App.ProductName, mstrTittle
End Sub

Private Sub lbl_Change(Index As Integer)
    Select Case Index
    Case Lbl_Index.预存款
        lbl(Index).Tag = ""
    End Select
End Sub

Private Sub picBlance_Resize()
    On Error Resume Next
    With picBlance
        vsBalance.Left = .ScaleLeft
        vsBalance.Height = .ScaleHeight - vsBalance.Top
        vsBalance.Width = .ScaleWidth - vsBalance.Left
    End With
End Sub

Private Sub picBlanceAndFee_Resize()
    On Error Resume Next
    With picBlanceAndFee
        tbPage.Left = .ScaleLeft + 30
        tbPage.Top = .ScaleTop + 10
        tbPage.Height = .ScaleHeight - tbPage.Top - 40
        tbPage.Width = .ScaleWidth - tbPage.Left - 40
    End With
    zlControl.PicShowFlat picBlanceAndFee, -1, , 1
End Sub

Private Sub picFee_Resize()
    On Error Resume Next
    With picFee
        vsFee.Left = .ScaleLeft
        vsFee.Top = .ScaleTop
        vsFee.Height = .ScaleHeight - vsFee.Top
        vsFee.Width = .ScaleWidth - vsFee.Left
    End With
End Sub

Private Function Load预交余额(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载预交余额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-21 10:47:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim dbl病人余额 As Double, dbl费用余额 As Double, dbl家属余额 As Double
    On Error GoTo errHandle
    
    '79868,将病人家属余额加入病人剩余款
    '获得记录集最多只有两条，一条是病人本人的，一条是病人家属的
    Set rsTemp = GetMoneyInfo(lng病人ID, , , 1, , , True)
    With mCurCharge
        .dbl预交余额 = 0
        .dbl费用余额 = 0
        Do While Not rsTemp.EOF
            .dbl预交余额 = .dbl预交余额 + Val(nvl(rsTemp!预交余额))
            .dbl费用余额 = .dbl费用余额 + Val(nvl(rsTemp!费用余额))
            If nvl(rsTemp!家属, 0) = 0 Then
                dbl病人余额 = Val(nvl(rsTemp!预交余额))
                dbl费用余额 = Val(nvl(rsTemp!费用余额))
            Else
                dbl家属余额 = Val(nvl(rsTemp!预交余额)) - Val(nvl(rsTemp!费用余额))
            End If
            rsTemp.MoveNext
        Loop
        .dbl可用预交 = .dbl预交余额 - .dbl费用余额
        If .dbl可用预交 < 0 Then .dbl可用预交 = 0
    End With
    If mbln使用预交 = False And mbytCurType = 1 Then
        mCurCharge.dbl可用预交 = 0: dbl家属余额 = 0
    End If
    
    lbl(Lbl_Index.预交余额).Caption = "预交余额:" & Format(dbl病人余额, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.未结费用).Caption = "未结费用:" & Format(dbl费用余额, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.剩余款额).Caption = "剩余款额:" & Format(dbl病人余额 - dbl费用余额, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.家属余额).Caption = "家属余额:" & Format(dbl家属余额, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.家属余额).Visible = RoundEx(dbl家属余额, 6) <> 0
    lineUnder((Lbl_Index.家属余额)).Visible = RoundEx(dbl家属余额, 6) <> 0
    Load预交余额 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体大小
    '编制:刘兴洪
    '日期:2011-09-15 11:26:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '最小窗体尺寸
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hwnd, GWL_WNDPROC)
    Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub

Private Sub Form_Load()
    Dim curMoney As Currency
    
    mblnFirst = True
    If mblnCliniqueRoomPay Then Exit Sub
    
    mstrTittle = "病人消费结算"
    
    If mbytBillType = 0 Then
        mrsFeeData.Filter = "记录性质=1"
        mbytCurType = IIf(mrsFeeData.RecordCount = 0, 2, 1)
    Else
        mbytCurType = mbytBillType
    End If
    Call InitFace
    If LoadPatient() = False Then Unload Me: Exit Sub
    
    mstrCurNos = GetCurFeeNos(mbytCurType)
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl费用合计 = curMoney
    
    mstr个人帐户 = ""
    Set mrs结算方式 = Get结算方式()
    If Not mrs结算方式.EOF Then
        mrs结算方式.Filter = "性质=3"
        If Not mrs结算方式.EOF Then mstr个人帐户 = nvl(mrs结算方式!名称)
    End If
    If Load预交余额(mrsInfo!病人ID) = False Then Unload Me: Exit Sub
    If Load支付方式() = False Then Unload Me: Exit Sub
    If LoadFeeData(mbytCurType) = False Then Unload Me: Exit Sub
    
    Call SetCtlEnable
    Call SetControlMove
    Call SetControlProperty
    Call SetDefaultPrepayMoney
End Sub

Public Function Get结算方式() As ADODB.Recordset
    '获取所有结算方式数据，不分结算场合，也不区分性质
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    strSQL = _
        "Select b.编码, b.名称, b.缺省标志 As 缺省, Nvl(b.性质, 1) As 性质, Nvl(b.应付款, 0) As 应付款" & vbNewLine & _
        "From 结算方式 B" & vbNewLine & _
        "Where b.性质 <> 9"
    Set Get结算方式 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitFace()
    '初始化界面
    If mblnFirst Then
        RestoreWinState Me, App.ProductName, mstrTittle
        If Not IsDesinMode Then Call SetWindowsSize
        
        zlControl.CboSetWidth cbo支付方式.hwnd, cbo支付方式.Width * 2
        zlControl.PicShowFlat picPatientInfo, -1, , 1
        zlControl.PicShowFlat pic剩余自付, -1, , 1
        zlControl.PicShowFlat pic自付合计, -1, , 1
        zlControl.PicShowFlat picPayInfo, -1, , 1
    End If
    
    Call InitPage
    picBlance.Visible = (mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "保险收费"))
    
    Call ClearData
End Sub

Private Sub InitPage()
    '功能:初始化页面控件
    Dim objItem As TabControlItem
    
    On Error GoTo ErrHandler
    tbPage.RemoveAll
    If mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "保险收费") Then
        Set objItem = tbPage.InsertItem(Pg_Index.Blance, "结算信息", picBlance.hwnd, 0)
        objItem.Tag = Pg_Index.Blance
    End If
    Set objItem = tbPage.InsertItem(Pg_Index.FeeDetail, "费用明细", picFee.hwnd, 0)
    objItem.Tag = Pg_Index.FeeDetail
    
    With tbPage
        .Item(.ItemCount - 1).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitFactPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始发票相关的参数
    '编制:刘兴洪
    '日期:2011-08-11 00:24:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mPara
        .int收费票据格式 = Val(zlDatabase.GetPara("收费收据格式", glngSys, 1151))
        .int收费打印方式 = Val(zlDatabase.GetPara("收费打印方式", glngSys, 1151))
        .int审核票据格式 = Val(zlDatabase.GetPara("审核收据格式", glngSys, 1151))
        .int审核打印方式 = Val(zlDatabase.GetPara("审核打印方式", glngSys, 1151))
        .int药品单位 = Val(zlDatabase.SetPara("药品单位", glngSys, 1151))
    End With
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数值
    '编制:刘兴洪
    '日期:2011-06-20 16:48:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    Call InitFactPara
    '门诊病人消费时需要刷卡验证
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    mdblBrushCardMoney = Val(Split(strValue, "|")(0))
    If mdblBrushCardMoney < 0 Then
        mbyt预存款消费验卡 = 3
        mdblBrushCardMoney = -1 * mdblBrushCardMoney
    Else
        mbyt预存款消费验卡 = mdblBrushCardMoney
    End If
    
    '费用单价保留位数
    mintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    '费用金额小数点位数
    mbytFeeMoneyPrecision = Val(zlDatabase.GetPara(9, glngSys, , 2))
    
    '自动发料
    mbln门诊自动发料 = zlDatabase.GetPara(92, glngSys) = "1"

    mbln只对医保结算成功单据收费 = Val(zlDatabase.GetPara(317, glngSys, , "0")) = 1
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    picBlanceAndFee.Width = Me.ScaleWidth - picBlanceAndFee.Left * 2
    picBlanceAndFee.Height = Me.ScaleHeight - staThis.Height - picBlanceAndFee.Top
End Sub

Private Function GetPatient(ByVal lngPatiID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=表示是否就诊卡刷卡
    '出参:
    '返回:病人读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-20 16:04:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH
    '读取病人信息
    strSQL = "Select Decode(Sign(a.就诊时间 - a.登记时间), 0, 1, 0) As 初诊, a.病人id, a.病人类型, a.Ic卡号, a.就诊卡号," & vbNewLine & _
            "        a.门诊号, a.住院号, a.姓名, a.卡验证码, a.性别, a.年龄, a.出生日期, a.费别," & vbNewLine & _
            "        a.医疗付款方式, m.编码 As 付款方式编码, a.在院, Decode(B1.病人性质, Null, 0, 1, 1, 0) As 留观," & vbNewLine & _
            "        B1.入院日期, a.险类, c.名称 As 险类名称" & vbNewLine & _
            " From 病人信息 A, 病案主页 B1, 保险类别 C, 医疗付款方式 M" & vbNewLine & _
            " Where a.病人id = B1.病人id(+) And a.主页id = B1.主页id(+) And a.险类 = c.序号(+)" & vbNewLine & _
            "       And a.医疗付款方式 = m.名称(+) And a.停用时间 Is Null And a.病人id = [1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, "病人消费结算-获取病人信息", lngPatiID)
    If mrsInfo.EOF Then
        ShowMsgbox "病人信息未找到，请检查！"
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If
    GetPatient = True
    Exit Function
errH:
    Set mrsInfo = New ADODB.Recordset
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPatient() As Boolean
    '加载病人信息
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    If mrsInfo.RecordCount = 0 Then Exit Function
    
    lbl(Lbl_Index.姓名).Caption = nvl(mrsInfo!姓名)
    lbl(Lbl_Index.性别).Caption = "性别:" & nvl(mrsInfo!性别)
    lbl(Lbl_Index.门诊号).Caption = "门诊号:" & nvl(mrsInfo!门诊号)
    '74309:李南春，2014-7-7，病人姓名显示颜色处理
    Call SetPatiColor(lbl(Lbl_Index.姓名), nvl(mrsInfo!病人类型), &HFF0000)
    LoadPatient = True
End Function

Private Function Load支付方式() As Boolean
    '加载有效的支付方式，仅启用的三方卡
    Dim i As Long, objCards As Cards, lngKey As Long
    Dim strRQCardTypeIDs As String
    
    Set mobjPayCards = New Cards
     
    ' zlGetCards(ByVal BytType As Byte)
    'bytType-  0-所有医疗卡;1-启用的医疗卡, 2-所有存在三方账户的三方卡3-启用的三方账户的医疗卡
    Set objCards = gobjOneCardComLib.zlGetCards(3)

    With cbo支付方式
        .Clear
        For i = 1 To objCards.Count
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
            
            .AddItem objCards(i).名称
            .ItemData(.NewIndex) = i
        Next
    End With
    If cbo支付方式.ListCount > 0 Then cbo支付方式.ListIndex = 0
    
    If mbytCurType = 1 Then
        strRQCardTypeIDs = mobjThreeSwap.GetRQCardTypeIDsFromCards(mobjPayCards)
        If Not btQRCodePay.zlInit(Me, strRQCardTypeIDs, glngSys, mlngModule, gcnOracle, gstrDBUser) Then strRQCardTypeIDs = ""
        btQRCodePay.Tag = strRQCardTypeIDs  '存储有效的卡类别IDs
        btQRCodePay.Visible = btQRCodePay.Tag <> ""
    Else
        btQRCodePay.Visible = False
    End If
    
    Load支付方式 = True
End Function

Private Sub SetCtlEnable(Optional ByVal blnEdit As Boolean = True)
    '设置控件的可用状态
    
    If blnEdit Then blnEdit = (mInsure.intInsure = 0 Or mInsure.intInsure <> 0 And mblnYbBalanced)
    picPayInfo.Enabled = blnEdit
    txt冲预交.Enabled = blnEdit And UsedPrepayMoney() = False And mCurCharge.dbl可用预交 > 0
    btQRCodePay.Enabled = blnEdit And btQRCodePay.Tag <> "" 'Tag:存储的是有效支持的扫码付的卡类别Ids
    txt金额.Enabled = blnEdit
    txt摘要.Enabled = blnEdit
    vsBalance.Editable = IIf(mInsure.intInsure <> 0 And mblnYbBalanced = False, flexEDKbdMouse, flexEDNone)
    
    '控制显示颜色
    txt冲预交.BackColor = IIf(txt冲预交.Enabled, &H80000005, Me.BackColor)
    cbo支付方式.BackColor = IIf(txt冲预交.Enabled, &H80000005, Me.BackColor)
    txt金额.BackColor = IIf(txt金额.Enabled, &H80000005, Me.BackColor)
    txt摘要.BackColor = IIf(txt摘要.Enabled, &H80000005, Me.BackColor)
End Sub

Private Function UsedPrepayMoney() As Boolean
    '判断是否已使用预交款
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = 1 To vsBalance.Rows - 1
        If vsBalance.RowData(i) = 1 Then
            UsedPrepayMoney = True: Exit Function
        End If
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Calc" Then Call ShowWindowsCalculator
End Sub

Private Sub txt冲预交_Change()
    txt冲预交.Tag = ""
    txt金额.Text = "0.00"
End Sub

Private Sub txt冲预交_GotFocus()
    zlControl.TxtSelAll txt冲预交
End Sub

Private Sub txt冲预交_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txt冲预交.Text) = 0 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If BrushcardStrikePrepay(False) = False Then
        zlControl.ControlSetFocus txt冲预交: zlControl.TxtSelAll txt冲预交
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt冲预交_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt冲预交, KeyAscii, m金额式)
End Sub

Private Function CheckPrepayValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查预存款数据是否有效
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-14 22:30:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If txt冲预交.Text = "" Then
        txt冲预交.Text = "0.00"
    ElseIf Not IsNumeric(txt冲预交.Text) And txt冲预交.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf Val(txt冲预交.Text) < 0 Then
        MsgBox "预存款冲款金额不能为负！", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf Val(txt冲预交.Text) > 0 And RoundEx(mCurCharge.dbl当前未付, 6) < 0 Then
        MsgBox "当前未付金额为负时不能使用预存款！", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf RoundEx(Val(txt冲预交.Text), 6) > RoundEx(mCurCharge.dbl可用预交, 6) Then
        MsgBox "预存款冲款金额不能超过病人的预存余额:" & FormatEx(mCurCharge.dbl可用预交, 6, , , 2) & " ！", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf RoundEx(Val(txt冲预交.Text), 6) > RoundEx(mCurCharge.dbl当前未付, 6) Then
        MsgBox "预存款冲款金额不能大于未付金额:" & Format(mCurCharge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    Else
        txt冲预交.Text = Format(Val(txt冲预交.Text), "0.00")
    End If
    CheckPrepayValied = True
    Exit Function
InvalidDataHandler:
    Call SetDefaultPrepayMoney
    zlControl.ControlSetFocus txt冲预交: zlControl.TxtSelAll txt冲预交
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt冲预交_LostFocus()
    If txt冲预交.Tag = "1" Then Exit Sub
    Call SetControlProperty
End Sub

Private Sub txt冲预交_Validate(Cancel As Boolean)
    If txt冲预交.Tag = "1" Then Exit Sub
    If CheckPrepayValied = False Then Cancel = True: Exit Sub
End Sub

Private Sub txt金额_GotFocus()
    txt金额.Text = Format(mCurCharge.dbl当前未付 - Val(txt冲预交.Text), "0.00")
    zlControl.TxtSelAll txt金额
End Sub

Private Sub txt金额_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Val(txt金额.Text) = 0 Then txt金额.Text = "0.00"
    txt金额.Text = Format(Val(txt金额.Text), "0.00")
    If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
End Sub

Private Sub txt金额_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt金额, KeyAscii, m金额式)
End Sub

Private Sub txt金额_Validate(Cancel As Boolean)
    txt金额.Text = Format(Val(txt金额.Text), "0.00")
End Sub

Private Sub SetDefaultPrepayMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省预交金额
    '编制:刘兴洪
    '日期:2011-08-13 17:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt金额.Text = "0.00"
    txt冲预交.Text = "0.00"
    With mCurCharge
        If mbytCurType = 2 Then
            txt冲预交.Text = Format(.dbl当前未付, "###0.00;###0.00;0.00;0.00")
            Exit Sub
        End If
        If RoundEx(.dbl可用预交, 6) <> 0 Then
            If RoundEx(.dbl可用预交, 6) > RoundEx(.dbl当前未付, 6) Then
                txt冲预交.Text = Format(.dbl当前未付, "###0.00;###0.00;0.00;0.00")
            Else
                txt冲预交.Text = Format(.dbl可用预交, "###0.00;###0.00;0.00;0.00")
            End If
        End If
    End With
End Sub

Private Function CheckThreeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方交易金额输入是否合法
    '返回:合法成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-15 00:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If txt金额.Visible = False Or txt金额.Enabled = False Then CheckThreeValied = True: Exit Function
    
    If Val(txt金额) = 0 Then
        ShowMsgbox "未输入交易金额，请检查！"
        zlControl.ControlSetFocus txt金额: zlControl.TxtSelAll txt金额
        Exit Function
    End If
    If Not IsNumeric(txt金额.Text) And txt金额.Text <> "" Then
        ShowMsgbox "无效数值！"
        zlControl.ControlSetFocus txt金额: zlControl.TxtSelAll txt金额
        Exit Function
    ElseIf Val(txt金额.Text) < 0 Then
        ShowMsgbox "交易金额不能为负！"
    ElseIf RoundEx(Val(txt金额.Text) + Val(txt冲预交.Text), 2) > RoundEx(mCurCharge.dbl当前未付, 2) And Val(txt金额.Text) <> 0 Then
        ShowMsgbox "交易金额不能大于本次未付金额:" & Format(mCurCharge.dbl当前未付 - Val(txt冲预交.Text), "0.00") & " ！"
        zlControl.ControlSetFocus txt金额: zlControl.TxtSelAll txt金额
        Exit Function
    End If
    CheckThreeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetClassMoney(ByVal bytType As Byte, ByVal strNOs As String, _
    ByRef rsClass As ADODB.Recordset) As Boolean
    '获取分类汇总金额
    '入参：
    '   bytType 1-门诊收费;2-门诊记帐
    Dim i As Integer
    Dim varNos As Variant
    
    On Error GoTo ErrHandler
    Set rsClass = New ADODB.Recordset
    rsClass.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsClass.Fields.Append "金额", adDouble, , adFldIsNullable
    rsClass.CursorLocation = adUseClient
    rsClass.LockType = adLockOptimistic
    rsClass.CursorType = adOpenStatic
    rsClass.Open
    
    varNos = Split(strNOs, ",")
    For i = 0 To UBound(varNos)
        mrsFeeData.Filter = "记录性质=" & bytType & " And No='" & varNos(i) & "'"
        Do While Not mrsFeeData.EOF
            rsClass.Find "收费类别='" & nvl(mrsFeeData!收费类别) & "'", , adSearchForward, 1
            If rsClass.EOF Then rsClass.AddNew
            rsClass!收费类别 = nvl(mrsFeeData!收费类别)
            rsClass!金额 = RoundEx(Val(nvl(rsClass!金额)) + Val(nvl(mrsFeeData!实收金额)), 6)
            rsClass.Update
            
            mrsFeeData.MoveNext
        Loop
    Next
    GetClassMoney = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function BrushCardThreeSwapCheck(ByVal strNOs As String, _
    ByVal dblMoney As Double, ByVal str费用来源 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '入参:strNos -本次支付的单据号
    '       dblMoney-支付的总金额
    '返回:返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsClassMoney As ADODB.Recordset, cllSquareBalance As Collection
    
    On Error GoTo errHandle
    mCurCardPay.str支付结算 = ""
    If mbytCurType = 2 Then BrushCardThreeSwapCheck = True: Exit Function
    If mCurCardPay.lng卡类别ID = 0 Then BrushCardThreeSwapCheck = True: Exit Function
    
    If mblnCliniqueRoomPay = False Then
        If CheckThreeValied() = False Then Exit Function
    End If
    
    If mCurCardPay.bln消费卡 Then
        If GetClassMoney(mbytCurType, strNOs, rsClassMoney) = False Then Exit Function
        '构建消费卡的刷卡信息
        Set cllSquareBalance = mcllSquareBalance
    End If
    
    If mobjThreeSwap.CheckPayValid(mCurCardPay.lng卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str结算方式, _
        dblMoney, strNOs, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, , mCurCardPay.str支付结算, _
        rsClassMoney, str费用来源, cllSquareBalance, mCurCardPay.strQRCode) = False Then Exit Function
    
    If mCurCardPay.bln消费卡 Then Set mcllSquareBalance = cllSquareBalance
    
    BrushCardThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCurFeeNos(ByVal bytType As Byte) As String
    '获取单据号
    '入参：
    '   bytType 1-门诊收费;2-门诊记帐
    '返回:单据号,单据之间用逗号分离,如:A0001,A0002....
    Dim strNOs As String
    
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "记录性质=" & bytType
    If mrsFeeData.RecordCount = 0 Then Exit Function
    
    mrsFeeData.Sort = "NO"
    With mrsFeeData
        Do While Not .EOF
            If InStr(strNOs & ",", "," & nvl(!NO) & ",") = 0 Then
                strNOs = strNOs & "," & nvl(!NO)
            End If
            .MoveNext
        Loop
    End With
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    GetCurFeeNos = strNOs
End Function

Private Function Get费用来源(ByVal bytType As Byte) As String
    '获取单据号
    '入参：
    '   bytType 1-门诊收费;2-门诊记帐
    '返回:
    Dim str费用来源 As String
    
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "记录性质=" & bytType
    If mrsFeeData.RecordCount = 0 Then Exit Function
    
    With mrsFeeData
        Do While Not .EOF
            If InStr(str费用来源, Decode(Val(!门诊标志), 4, 3, 2, 2, 1)) = 0 Then
                str费用来源 = str费用来源 & "," & Decode(Val(!门诊标志), 4, 3, 2, 2, 1)
            End If
            .MoveNext
        Loop
    End With
    If str费用来源 <> "" Then str费用来源 = Mid(str费用来源, 2)
    Get费用来源 = str费用来源
End Function

Private Function GetSelectNOsAndSerialNum(ByRef strNOs As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择的单据号和单据中的序号
    '返回:单据号,单据之间用逗号分离,如:A0001:1,2|A0002:1,2,3|....
    '编制:刘兴洪
    '日期:2011-06-23 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strNo As String
    Dim str序号 As String, strData As String
    Dim j As Long
    
    With vsFee
        strData = "": strNOs = ""
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, .ColIndex("单据号")))
            If InStr(1, strNOs & ",", "," & strNo & ",") = 0 Then
                str序号 = ""
                For j = 1 To .Rows - 1
                    If strNo = Trim(.TextMatrix(j, .ColIndex("单据号"))) Then
                        str序号 = str序号 & "," & .RowData(j)
                    End If
                Next
                If str序号 <> "" Then str序号 = Mid(str序号, 2)
                strNOs = strNOs & "," & strNo
                strData = strData & "|" & strNo & ":" & str序号
            End If
        Next
    End With
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    If strData <> "" Then strData = Mid(strData, 2)
    GetSelectNOsAndSerialNum = strData
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据合法性检查
    '返回:数据合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-22 15:28:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl冲预交  As Double, dblThreeMoney  As Double
    Dim str费用来源 As String
    
    If mrsInfo Is Nothing Then
        ShowMsgbox "病人信息不能确定，请检查！"
        zlControl.ControlSetFocus cmdCancel: Exit Function
    End If
    
    If mrsInfo.State <> 1 Then
        ShowMsgbox "病人信息不能确定，请检查！"
        zlControl.ControlSetFocus cmdCancel: Exit Function
    End If
    
    If mbytCurType = 1 Then
        dbl冲预交 = 0: dblThreeMoney = 0
        If mblnCliniqueRoomPay = False Then '非诊间支付时，需要检查相关的数据合法性
            If Not CheckTextLength("摘要", txt摘要) Then Exit Function
            
            If txt冲预交.Visible And txt冲预交.Enabled Then dbl冲预交 = Val(txt冲预交.Text)
            If txt金额.Visible And txt金额.Enabled Then dblThreeMoney = Val(txt金额.Text)
        
            If cbo支付方式.ListIndex >= 0 Then
                If mCurCardPay.str结算方式 = "" Then
                    ShowMsgbox mCurCardPay.str名称 & " 未设置结算方式，请与系统管理员联系！"
                    Exit Function
                End If
            ElseIf RoundEx(dblThreeMoney, 6) <> 0 Then
                ShowMsgbox "未选择支付方式！"
                Exit Function
            End If
            
            If RoundEx(dbl冲预交 + dblThreeMoney, 6) <> RoundEx(mCurCharge.dbl当前未付, 6) Then
                If Val(txt金额.Text) = 0 And txt冲预交.Visible Then
                    ShowMsgbox "病人的预存款余额不足，请充值！"
                    zlControl.ControlSetFocus txt冲预交: zlControl.TxtSelAll txt冲预交
                ElseIf txt冲预交.Visible = False Then
                    ShowMsgbox "本次" & cbo支付方式.Text & "支付金额(" & _
                        Format(dblThreeMoney, "0.00") & ")与本次未付金额(" & _
                        Format(mCurCharge.dbl当前未付, "0.00") & ")不等，请检查！"
                    zlControl.ControlSetFocus txt金额: zlControl.TxtSelAll txt金额
                Else
                    ShowMsgbox "本次支付金额合计(预存款+" & cbo支付方式.Text & ":" & _
                        Format(dbl冲预交 + dblThreeMoney, "0.00") & ")与本次未付金额(" & _
                        Format(mCurCharge.dbl当前未付, "0.00") & ")不等，请检查！"
                    zlControl.ControlSetFocus txt金额: zlControl.TxtSelAll txt金额
                End If
                Exit Function
            End If
            
            If RoundEx(dbl冲预交, 6) > 0 And Val(txt冲预交.Tag) = 0 Then
                '证明没有验证卡，需要输入密码验证
                If BrushcardStrikePrepay(True) = False Then Exit Function
            End If
        Else
            dblThreeMoney = mCurCharge.dbl当前未付
        End If
        
        str费用来源 = Get费用来源(mbytCurType)
        If RoundEx(dblThreeMoney, 6) <> 0 Then
            If BrushCardThreeSwapCheck(mstrCurNos, dblThreeMoney, str费用来源) = False Then Exit Function
        End If
    Else
        If Val(txt冲预交.Tag) = 0 Then
            '证明没有验证卡，需要输入密码验证
            If BrushcardStrikePrepay(True) = False Then Exit Function
        End If
    End If
    isValied = True
End Function

Private Function BrushcardStrikePrepay(ByVal blnOKClick As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证刷卡冲预交
    '返回:冲销成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    
    On Error GoTo ErrHandler
    If Val(txt冲预交.Tag) = 1 Then BrushcardStrikePrepay = True: Exit Function
    If Val(txt冲预交.Text) = 0 And mbytCurType = 1 Then BrushcardStrikePrepay = True: Exit Function
    
    If mbytCurType <> 2 Then
        If CheckPrepayValied() = False Then Exit Function
    End If
    dblMoney = Val(txt冲预交.Text)
    
    gblnNotCloseWindows = True
    '刷卡确认
    If zlPatiIdentify(mlngModule, Me, mlngPatiID, dblMoney, False, 1, mlngCardTypeID, True, , mstr家属IDs) Then
        gblnNotCloseWindows = False
                    
        txt冲预交.Tag = "1" '标记已验证
        
        '修正病人预交记录关联交易ID数据
        If mobjThreeSwap.UpgradeHistoryData(mlngModule, mlngPatiID, mstr家属IDs) = False Then Exit Function
        '59412
        If blnOKClick Then BrushcardStrikePrepay = True: Exit Function
        
        If RoundEx(dblMoney, 6) = RoundEx(mCurCharge.dbl当前未付, 6) Or mbytCurType = 2 Then
            '相等时,保存数据
            Call cmdOK_Click
            If mblnOk Then BrushcardStrikePrepay = True: Exit Function
        ElseIf mbytCurType = 1 And cbo支付方式.ListCount = 0 Then
            ShowMsgbox "病人的预存款余额不足，请充值！"
            Exit Function
        End If
        
        Call SetControlProperty
        BrushcardStrikePrepay = True
    Else
        gblnNotCloseWindows = False
        BrushcardStrikePrepay = False
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Dim ty_Tmp As TY_Insure
    
    On Error GoTo ErrHandler
    If mbytCurType = 1 And mInsure.intInsure <> 0 And mInsure.strYBPati <> "" Then
        If mblnCommitBill Then
            If MsgBox("    当前正在对医保病人收费，退出后本次结算将保存为异常状态，" & vbCrLf & _
                "需要到收费窗口进行处理，确实要退出吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            '取消医保验证
            If YBIdentifyCancel() = False Then Exit Sub
            mInsure = ty_Tmp
            mCurCharge.dbl医保支付 = 0
            mCurCharge.dbl已付合计 = 0
            
            vsBalance.Clear 1: vsBalance.Rows = 2
            vsBalance.RowData(1) = ""
            tbPage.Item(Pg_Index.FeeDetail).Selected = True
            cmdYBBalance.Enabled = False
            
            Call SetPatiColor(lbl(Lbl_Index.姓名), nvl(mrsInfo!病人类型), &HFF0000)
            staThis.Panels(Pan.C3个人帐户).Text = ""
            staThis.Panels(Pan.C3个人帐户).Visible = False
            
            Call SetControlProperty
            Call SetCtlEnable
            Call SetDefaultPrepayMoney
            Call SetBeginFocus '光标定位
            Exit Sub
        End If
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintBill(ByVal strPrintNo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印票据
    '入参：
    '   strPrintNO 格式：'A001','A002',...
    '编制:刘兴洪
    '日期:2014-01-20 11:01:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, strFormat As String
    Dim frmMain As Object
    
    If mblnCliniqueRoomPay Then
        Set frmMain = mfrMain
    Else
        Set frmMain = Me
    End If
    Select Case mbytCurType
    Case 1
        blnPrint = mPara.int收费打印方式 = 1
        If mPara.int收费打印方式 = 2 Then
            If MsgBox("你是否真的要打印清单吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int收费票据格式 = 0, "", "ReportFormat=" & mPara.int收费票据格式)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & strPrintNo, "药品单位=" & mPara.int药品单位, "PrintEmpty=0", strFormat, 2)
        End If
    Case 2
        blnPrint = mPara.int审核打印方式 = 1
        If mPara.int审核打印方式 = 2 Then
            If MsgBox("你是否真的要打印清单吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int审核票据格式 = 0, "", "ReportFormat=" & mPara.int审核票据格式)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & strPrintNo, "药品单位=" & mPara.int药品单位, "PrintEmpty=0", strFormat, 2)
        End If
    End Select
End Sub

Private Sub cmdOK_Click()
    Dim blnPartialSaved As Boolean '部分保存成功
    Dim curMoney As Currency, bln继续收费 As Boolean
    Dim strPrintNo As String '格式：'A001','A002',...
    
    On Error GoTo errHandle
    '数据校对
    If isValied = False Then Exit Sub
    If SaveData(strPrintNo, blnPartialSaved) = False Then Exit Sub
    If blnPartialSaved Then Unload Me: Exit Sub
    
    '打印票据
    Call PrintBill(strPrintNo)
    
    '银医一卡通写卡，85950
    If mbytCurType = 1 Then
        Call WriteInforToCard(Me, mlngModule, mstrPrivs, 0, strPrintNo)
    End If
    
    bln继续收费 = False
    If mbytCurType = 1 And mInsure.strAllNos <> "" Then
        If MsgBox("当前只成功收取了" & UBound(Split(mstrCurNos, ",")) + 1 & "张单据的费用，" & _
                  "是否对未收取成功的单据进行重新收费？", _
            vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            
            mstrCurNos = GetRemainNos(mInsure.strAllNos, mstrCurNos)
            Set mrsFeeData = GetFeeData(mlngPatiID)
            mrsFeeData.Filter = "记录性质=1"
            If mrsFeeData.RecordCount > 0 Then
                bln继续收费 = True
            End If
        End If
    End If
    
    If bln继续收费 = False Then
        '0-不区分收费或记帐单时，对记账单进行审核
        If mbytBillType = 0 And mbytCurType = 1 Then
            mbytCurType = 2
            bln继续收费 = True
        End If
    End If
        
    If bln继续收费 = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    '继续处理剩余费用
    Call InitVariableData
    Call InitFace
    If LoadFeeData(mbytCurType) = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    Call LoadPatient
    If Load预交余额(mrsInfo!病人ID) = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    If Load支付方式() = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    mstrCurNos = GetCurFeeNos(mbytCurType)
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl费用合计 = curMoney
    
    Call SetCtlEnable
    Call SetControlMove
    Call SetControlProperty
    Call SetDefaultPrepayMoney
    Call SetBeginFocus '光标定位
    
    Call ShowLedInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetRemainNos(ByVal strAllNos As String, ByVal strSavedNos As String) As String
    '获取剩余单据号
    '入参：
    '   strAllNos 所有单据号，A001,A002,...
    '   strSavedNos 以保存单据号，A001,A002,...
    '返回：剩余单据号，A001,A002,...
    Dim varAllNos As Variant, strNOs As String
    Dim i As Integer
    
    varAllNos = Split(strAllNos, ",")
    For i = 0 To UBound(varAllNos)
        If InStr("," & strSavedNos & ",", "," & varAllNos(i) & ",") = 0 Then
            strNOs = strNOs & "," & varAllNos(i)
        End If
    Next
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    GetRemainNos = strNOs
End Function

Private Function VerifyFee(ByRef strPrintNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:审核费用
    '入参:
    '   strPrintNO 打印单据号，格式：'A001','A002',...
    '返回:审核成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2011-06-23 09:59:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNOs As String
    Dim strNosData As String '格式:A0001:1,2|A0002:1,2,3|....
    Dim str审核时间 As String
    
    strPrintNo = ""
    strNosData = GetSelectNOsAndSerialNum(strNOs)
     '记帐的话,要费用报警
    If Not zlAuditingWarn(mstrPrivs, strNOs, Val(nvl(mrsInfo!病人ID))) Then Exit Function
    
    '记帐审核
    str审核时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    'strNos-单据信息, 格式：NO1:序号1,序号2,...|NO1:序号1,序号2,...|...
    If mclsExpenceSvr.zlVerfyBillingPriceBill(Val("1-门诊"), strNosData, str审核时间) = False Then Exit Function
    
    '药品已收费状态确认
    Call mclsExpenceSvr.zlDrugRecipeAffirm(strNOs, 1, 2)
    '卫材已收费状态确认
    Call mclsExpenceSvr.zlStuffBillAffirm(strNOs, 1, 2, mbln门诊自动发料)
    
    strPrintNo = "'" & Replace(strNOs, ",", "','") & "'"
    
    VerifyFee = True
    
    '调用包药机
    Call mobjDrugStuff.DrugMachine_Charge(2, strNOs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCharge(ByRef strPrintNo As String, Optional ByRef blnPartialSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:划价收费
    '出参：
    '   strPrintNO 打印单据号，格式：'A001','A002',...
    '   blnPartialSaved - 是否部分保存成功
    '返回:收费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-23 11:38:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblThreeMoney As Double, dbl冲预交 As Double, dbl误差费 As Double
    Dim strSQL As String, cllUpdate As Collection, cllOthers As Collection
    Dim str交易流水号 As String, str交易说明 As String
    Dim str结算方式 As String, j As Integer
    Dim blnTrans As Boolean, strExpend As String, dblOutMoney As Double
    Dim lng关联交易ID As Long, blnHaveMoney As Boolean
    Dim cll结算方式 As Collection, i As Integer
    Dim bln正在交易 As Boolean, strErrMsg As String, blnCommit As Boolean
    Dim lng病人ID  As Long, lng结帐ID As Long
 
    Err = 0: On Error GoTo ErrHandler
    strPrintNo = "": blnPartialSaved = False
    
    If mblnCliniqueRoomPay Then
        dblThreeMoney = mCurCharge.dbl当前未付
    Else
        If txt冲预交.Visible And txt冲预交.Enabled Then dbl冲预交 = Val(txt冲预交.Text)
        If txt金额.Visible And txt金额.Enabled Then dblThreeMoney = Val(txt金额.Text)
    End If
    dbl误差费 = mCurCharge.dbl本次误差费
    
    blnTrans = True
    If SaveFeeBill() = False Then Exit Function
    
    lng病人ID = mlngPatiID
    lng结帐ID = mCurCharge.lng结帐ID
    
    If RoundEx(dblThreeMoney, 6) = 0 Then
        '全部使用预交款支付
        strSQL = SetCurBalanceSQL(0, lng病人ID, lng结帐ID, "", dbl冲预交, mstr家属IDs, dbl误差费, True)
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Else
        'bytType-1-三方接口支付;2-消费卡支付,0-其他
        If mCurCardPay.bln消费卡 Then
            If mcllSquareBalance Is Nothing Then Exit Function
            If mcllSquareBalance.Count = 0 Then Exit Function
            '卡类别ID|卡号|消费卡ID|消费金额||...
            '消费卡ID可以不传,传为0时,以卡号自动查找
            str结算方式 = ""
            For j = 1 To mcllSquareBalance.Count
                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                str结算方式 = str结算方式 & "||" & Val(mcllSquareBalance(j)(0))
                str结算方式 = str结算方式 & "|" & mcllSquareBalance(j)(3)
                str结算方式 = str结算方式 & "|" & Val(mcllSquareBalance(j)(1))
                str结算方式 = str结算方式 & "|" & Val(mcllSquareBalance(j)(2))
            Next
            If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 3)
            strSQL = SetCurBalanceSQL(2, lng病人ID, lng结帐ID, str结算方式, dbl冲预交, _
                mstr家属IDs, dbl误差费, True, mCurCardPay.lng卡类别ID, mCurCardPay.str刷卡卡号)
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Else '三方卡支付,预交款和误差费放在完成支付时
            '结算方式|结算金额|结算号码|结算摘要
            If mCurCardPay.str支付结算 = "" Then
                str结算方式 = mCurCardPay.str结算方式
                str结算方式 = str结算方式 & "|" & dblThreeMoney
            Else
                str结算方式 = mCurCardPay.str支付结算
            End If
            str结算方式 = str结算方式 & "| |" & IIf(Trim(txt摘要.Text) = "", " ", Trim(txt摘要.Text))
            lng关联交易ID = zlDatabase.GetNextId("病人预交记录")
            strSQL = SetCurBalanceSQL(3, lng病人ID, lng结帐ID, str结算方式, 0, _
                "", 0, False, mCurCardPay.lng卡类别ID, mCurCardPay.str刷卡卡号, lng关联交易ID)
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    End If
    
    If mblnCliniqueRoomPay = False Then
        If RoundEx(dblThreeMoney, 6) = 0 Or mCurCardPay.bln消费卡 Then
            '这个肯定是冲预交或者为消费卡在医院的卡帐户
            gcnOracle.CommitTrans
            
            GoTo SuccessHandler:
            Exit Function
        End If
    End If
    
    If gblnAsyncCharge Then '费用结算异步控制，先提交数据
        gcnOracle.CommitTrans: blnTrans = False
        blnCommit = True
    End If
    
    If mobjThreeSwap.ExecutePay(mCurCardPay.lng卡类别ID, mCurCardPay.bln消费卡, _
        mCurCardPay.str刷卡卡号, lng结帐ID, dblThreeMoney, str交易流水号, str交易说明, _
        strExpend, dblOutMoney, cll结算方式, bln正在交易, strErrMsg, mCurCardPay.strQRCode) = False Then
        If blnTrans Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then ShowMsgbox strErrMsg
            Exit Function
        End If
        
        If bln正在交易 Then
            MsgBox IIf(strErrMsg = "", "", strErrMsg & vbCrLf) & _
                "    " & mCurCardPay.str结算方式 & " 支付交易出现异常，不确定交易是否成功，无法完成结算。" & vbCrLf & _
                "请到收费窗口进行处理！", vbExclamation + vbOKOnly, gstrSysName
            blnPartialSaved = True: SaveCharge = True
            Exit Function
        Else
            If strErrMsg <> "" Then ShowMsgbox strErrMsg
            
            gcnOracle.BeginTrans: blnTrans = True
            '1.删除病人预交记录
            'Zl_病人结算记录_Delete(
            strSQL = "Zl_病人结算记录_Delete("
            '  结帐id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & lng结帐ID & ","
            '  关联交易id_In 病人预交记录.关联交易id%Type
            strSQL = strSQL & "" & lng关联交易ID & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            
            '2.删除费用结算对照,恢复为划价单
            'Zl_门诊收费结算_Cancel(
            strSQL = "Zl_门诊收费结算_Cancel("
            '  结帐id_In   门诊费用记录.结帐id%Type,
            strSQL = strSQL & "" & lng结帐ID & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            gcnOracle.CommitTrans: blnTrans = False
        End If
        
        Exit Function
    End If
    blnHaveMoney = RoundEx(dblThreeMoney, 6) <> RoundEx(dblOutMoney, 6)
    
    Set cllUpdate = New Collection
    If cll结算方式 Is Nothing Then
        Call zlAddUpdateSwapSQL(False, lng结帐ID, mCurCardPay.lng卡类别ID, mCurCardPay.bln消费卡, _
            mCurCardPay.str刷卡卡号, str交易流水号, str交易说明, cllUpdate, 2)
    Else
        'Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算|卡号",交易流水号,交易说明)
        For i = 1 To cll结算方式.Count
            If Trim(Split(cll结算方式(i)(0), "|")(6)) <> "" Then mCurCardPay.str刷卡卡号 = Split(cll结算方式(i)(0), "|")(6)
            strSQL = SetCurBalanceSQL(3, lng病人ID, lng结帐ID, cll结算方式(i)(0), 0, _
                "", 0, False, mCurCardPay.lng卡类别ID, mCurCardPay.str刷卡卡号, _
                lng关联交易ID, (i = 1), cll结算方式(i)(1), cll结算方式(i)(2), 2)
            zlAddArray cllUpdate, strSQL
        Next
    End If
    
    If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, blnNoBeginTrans:=True
    blnTrans = False
    blnCommit = True
    
    Err = 0: On Error GoTo ErrOthers
    Set cllOthers = New Collection
    Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, mCurCardPay.lng卡类别ID, mCurCardPay.bln消费卡, _
        mCurCardPay.str刷卡卡号, strExpend, cllOthers)
    zlExecuteProcedureArrAy cllOthers, Me.Caption
    
ChargeOver:
    Err = 0: On Error GoTo ErrHandler
    If blnHaveMoney Then
        MsgBox "    " & mCurCardPay.str结算方式 & " 实际支付金额(" & Format(dblOutMoney, "0.00") & ")不等于应付金额(" & Format(dblThreeMoney, "0.00") & ")，无法完成结算。" & vbCrLf & _
            "请到收费窗口进行处理！", vbExclamation + vbOKOnly, gstrSysName
        blnPartialSaved = True: SaveCharge = True
        Exit Function
    End If
    
    '完成结算
    strSQL = SetCurBalanceSQL(0, lng病人ID, lng结帐ID, "", dbl冲预交, mstr家属IDs, dbl误差费, True)
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '收费成功后的处理
SuccessHandler:
    '执行药品卫材处理：费用状态更新，自动发药、发料
    Call mclsExpenceSvr.zlDrugRecipeAffirm(Replace(mstrCurNos, "'", ""), 1, 1)
    Call mclsExpenceSvr.zlStuffBillAffirm(Replace(mstrCurNos, "'", ""), 1, 1)
    
    mlng结帐ID = lng结帐ID
    strPrintNo = "'" & Replace(mstrCurNos, ",", "','") & "'"
    SaveCharge = True
    
    '调用包药机
    Call mobjDrugStuff.DrugMachine_Charge(1, Replace(mstrCurNos, "'", ""))
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnCommit Then
        MsgBox IIf(Err.Description = "", "", Err.Description & vbCrLf) & _
            "    " & mCurCardPay.str结算方式 & " 支付交易出现异常，无法完成结算。" & vbCrLf & _
            "请到收费窗口进行处理！", vbExclamation + vbOKOnly, gstrSysName
        blnPartialSaved = True: SaveCharge = True
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrOthers:
    gcnOracle.CommitTrans   '能保存多少保存多少
    Call ErrCenter
    GoTo ChargeOver:
End Function

Public Function GetBrushCardXMLExpend(ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef strBalance As String) As Boolean
    '功能：解析三方支付刷卡验证数据
    '入参：
    '   strXMLExpend:XML串
    '    <OUTPUT>
    '        <JS> //结算信息(目前只支持返回一种方式)
    '            <JYFS>交易方式</JYFS> //交易方式:即结算方式.名称
    '            <JYJE>交易金额</JYJE>
    '        </JS>
    '        ...
    '    </OUTPUT>
    '出参：
    '   dblOutMoney - 实际支付金额
    '   strBalance - 结算数据，格式：结算方式|结算金额||...
    Dim lngCount As Long, strValue As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    strBalance = ""
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    '结算信息
    Call zlXML_GetRows("JS", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("JYFS", i, strValue)
        strBalance = strBalance & "||" & strValue '结算方式
        Call zlXML_GetNodeValue("JYJE", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '结算金额
        dblOutMoney = dblOutMoney + Val(strValue)
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    GetBrushCardXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData(ByRef strPrintNo As String, Optional ByRef blnPartialSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '   strPrintNO 打印单据号，格式：'A001','A002',...
    '编制:刘兴洪
    '日期:2011-06-22 16:01:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1-收费记录;2-记帐记录
    
    Select Case mbytCurType
    Case 1  '收费划价处理
        If SaveCharge(strPrintNo, blnPartialSaved) = False Then Exit Function
        '打印相关的票据
    Case 2 '划价记帐审核
        If VerifyFee(strPrintNo) = False Then Exit Function
        SaveData = True
    Case Else
        Exit Function
    End Select
    
    SaveData = True
End Function

Private Sub cmdPrintSet_Click()
    If frmSquareAffirmParaSet.SetPara(Me) = False Then Exit Sub
    Call InitFactPara
End Sub

Private Function SetCurBalanceSQL(ByVal bytType As Byte, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal str结算方式 As String, ByVal dbl冲预交 As Double, ByVal str家属IDs As String, _
    ByVal dbl本次误差费 As Double, ByVal bln完成结算 As Boolean, _
    Optional ByVal lngCardTypeID As Long, Optional ByVal str卡号 As String, _
    Optional ByVal lng关联交易ID As Long, Optional ByVal bln删除原结算 As Boolean, _
    Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
    Optional ByVal byt校对标志 As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置当前结算的SQL给cllpro过程
    '入参:  bytType-1-三方接口支付;2-消费卡支付;3-三方接口多种结算方式支付;0-其他
    '       dbl冲预交-预交款支付
    '       dbl本次误差费-本次产生的误差费
    '出参:
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String
    
    On Error GoTo errHandle
    ' Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --功能:收费结算时,修改结算的相关信息
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退支票额_In:传入零
    '  --   3-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
    '  --     ②退支票额_In:传入零
    '  --   4-三方卡结算，多种结算方式:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  -- 冲预交_In: 存在冲预交时,传入
    '  -- 误差金额_In:存在误差费时,传入
    '  -- 完成结算_In:1-完成收费;0-未完成收费
    '  ------------------------------------------------------------------------------------------------------------------------------
    Select Case bytType
    Case 1  '1-三方接口支付
        strSQL = strSQL & "1" & ","
    Case 2 ' 2-消费卡支付
        strSQL = strSQL & "3" & ","
    Case 3 ' 3-三方接口多种结算方式支付
        strSQL = strSQL & "4" & ","
    Case Else
        strSQL = strSQL & "0" & ","
    End Select
    '    病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & lng病人ID & ","
    '    结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & lng结帐ID & ","
    '    结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '    冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & ZVal(dbl冲预交) & ","
    '    退支票额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & ZVal(lngCardTypeID) & ","
    '    卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(lngCardTypeID = 0, "NULL", "'" & str卡号 & "'") & ","
    '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '    交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '    缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '    找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '    误差金额_In   门诊费用记录.实收金额%Type := Null,
    '    -- 误差金额_In:存在误差费时,传入
    strSQL = strSQL & "" & dbl本次误差费 & ","
    '    完成结算_In Number:=0
    '    -- 完成结算_In:1-完成收费;0-未完成收费
    strSQL = strSQL & IIf(bln完成结算, "1", "0") & ","
    '  缺省结算方式_In  结算方式.名称%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '79868,冉俊明,2015-06-10,使用病人家属预交
    '  冲预交病人ids_In Varchar2:=Null,
    strSQL = strSQL & "'" & lng病人ID & "," & str家属IDs & "',"
    '  更新交款余额_In  Number := 1,
    strSQL = strSQL & "" & 1 & ","
    '  关联交易id_In    病人预交记录.关联交易id%Type := Null
    strSQL = strSQL & "" & ZVal(lng关联交易ID) & ","
    '  删除原结算_In    Number := 0
    strSQL = strSQL & "" & IIf(bln删除原结算, "1", "0") & ","
    '  校对标志_In      病人预交记录.校对标志%Type := 0
    strSQL = strSQL & "" & byt校对标志 & ")"
    SetCurBalanceSQL = strSQL
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示病人信息以及消费情况
    '编制:李南春
    '日期:2014-10-29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String, lngPatient As Long
    
    On Error GoTo Errhand
    If Not gblnLED Then Exit Sub
    
    zl9LedVoice.Reset mscCom
    strInfo = nvl(mrsInfo!姓名) & " " & nvl(mrsInfo!性别) & " " & nvl(mrsInfo!年龄)
    lngPatient = Val("" & mrsInfo!病人ID)
    zl9LedVoice.DisplayPatient strInfo, lngPatient
    
    '消费总额:本次需要支付的金额，预交余额:病人当前的预交余额
    Call zl9LedVoice.DisplayBank("消费总额:" & mCurCharge.dbl费用合计 & "元" & _
        IIf(mCurCharge.dbl预交余额 = 0, "", ",预交余额:" & mCurCharge.dbl预交余额 & "元"))
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function InitThreeSwap(frmMain As Object) As Boolean
    '初始化卡支付对象
    On Error GoTo ErrHandler
    If Not mobjThreeSwap Is Nothing Then InitThreeSwap = True: Exit Function
    
    Set mobjThreeSwap = New clsThreeSwap
    mobjThreeSwap.Init frmMain, mlngModule, nvl(mrsInfo!病人ID), nvl(mrsInfo!姓名), nvl(mrsInfo!性别), nvl(mrsInfo!年龄)
    InitThreeSwap = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPayDrugWindow(ByVal lng病人ID As Long, ByVal dt收费时间 As Date, _
    ByVal cllDept As Collection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：分配发药窗口
    '入参:lng病人ID-病人ID
    '     dt收费时间-收费时间
    '     cllDept-具体执行部门:array(收费类别,执行部门ID,发药窗口)
    '返回：发药窗口名称
    '编制：李南春
    '入参:strNO
    '时间：2014-6-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发药窗口 As String, strPayDrugWins As String
    Dim i As Long, varData As Variant
    Dim blnFirst As Boolean
    
    On Error GoTo ErrHandler
    blnFirst = True
    strPayDrugWins = ""
    For i = 1 To cllDept.Count
        varData = cllDept(i)
        str发药窗口 = varData(2)
        If str发药窗口 = "" Then
            str发药窗口 = mobjDrugStuff.Get发药窗口(lng病人ID, Trim(varData(0)), Val(varData(1)), blnFirst)
            If blnFirst Then blnFirst = False
        End If
        If InStr(1, strPayDrugWins & ";", ";" & Val(varData(1)) & "|") = 0 Then
            strPayDrugWins = strPayDrugWins & ";" & Val(varData(1)) & "|" & str发药窗口
        End If
    Next
    GetPayDrugWindow = strPayDrugWins
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
      ByVal lngCardTypeID As Long, ByVal strNOs As String) As Boolean
    '功能:将门诊信息写入卡中
    '入参：
    '    frmMain - 调用窗体
    '    lngModul - 模块号
    '    strPrivs - 权限串
    '    objSquareCard - 医疗卡对象
    '    strNOs - 单据号，格式：'A0001','A0002','A0003',...或A0001,A0002,A0003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng病人ID As Long, lng结算序号 As Long
    
    Err = 0: On Error GoTo errH:
    '问题:56615
    'If InStr(strPrivs, ";门诊信息写卡;") = 0 Then Exit Function
    
    strSQL = "Select /*+Cardinality(j,10)*/ Distinct A.病人ID,B.结算序号" & _
            " From 门诊费用记录 A,病人预交记录 B,Table(f_Str2list([1])) J" & _
            " Where A.结帐ID=B.结帐ID And A.NO=J.Column_Value And A.记录性质 = 1 And A.记录状态 in(1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取单据结算序号", Replace(strNOs, "'", ""))
    If rsTemp.EOF Then Exit Function
    
    Do While Not rsTemp.EOF
        lng病人ID = Val(nvl(rsTemp!病人ID))
        lng结算序号 = Val(nvl(rsTemp!结算序号))
        '调用健康卡写卡接口
        If lng病人ID <> 0 And lng结算序号 <> 0 Then
            Call gobjOneCardComLib.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng病人ID, lng结算序号)
        End If
        rsTemp.MoveNext
    Loop
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt摘要_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt摘要
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlControl.ControlSetFocus cmdOK
End Sub

Private Sub txt摘要_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrHandler
    With vsBalance
        If mInsure.intInsure = 0 Or mInsurePara.多单据分单据结算 Then Cancel = True: Exit Sub
        If mblnYbBalanced Then Cancel = True: Exit Sub
        
        If Row < .FixedRows Or Col < .FixedCols Then Cancel = True: Exit Sub
        If .TextMatrix(Row, .ColIndex("支付方式")) = "" Then Cancel = True: Exit Sub
        If Col <> .ColIndex("支付金额") Then Cancel = True: Exit Sub
        
        '不允许修改的医保项目
        If Val(.RowData(Row)) = 0 Then Cancel = True: Exit Sub
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsBalance_EnterCell()
    If vsBalance.Editable = flexEDNone Then Exit Sub
    vsBalance.EditCell
End Sub

Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call Grid.CheckKeyPress(vsBalance, Row, Col, KeyAscii, m负金额式)
End Sub

Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strMoney As String, str支付方式 As String
    Dim curOrig As Currency, curTotal As Currency
    Dim p As Integer, objItem As BalanceMoney
    
    On Error GoTo ErrHandler
    With vsBalance
        If Row < 0 Then Exit Sub
        If Col <> 1 Or Col < 0 Then Exit Sub
        
        If zlCommFun.DblIsValid(.EditText, 10, False, False) = False Then Cancel = True: Exit Sub
        .EditText = Format(Val(.EditText), "0.00")
            
        strMoney = Trim(.EditText)
        If Not IsNumeric(strMoney) Then
            ShowMsgbox "输入了非法的支付金额：""" & strMoney & """！"
            .EditCell
            .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True: Exit Sub
        End If
        
        str支付方式 = Trim(.TextMatrix(.Row, .ColIndex("支付方式")))
        If str支付方式 = "" Then Exit Sub
        
        If str支付方式 = mstr个人帐户 Then '个人帐户检查
            '不允许超过允许透支金额
            If Val(strMoney) > mInsure.dbl个帐余额 + mInsure.dbl个帐透支 Then
                ShowMsgbox "帐户余额:" & Format(mInsure.dbl个帐余额, "0.00") & _
                    IIf(mInsure.dbl个帐透支 = 0, "", "(" & "允许透支:" & Format(mInsure.dbl个帐透支, "0.00") & ")") & _
                    "不足要支付的金额。"
                .EditCell
                .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True: Exit Sub
            End If
        Else
            '结算金额不允许超过返回的原始金额(个人帐户允许透支时再判断)
            curOrig = GetMedicareSum(mInsure.colBalance, str支付方式, , True)   '该结算方式所有原始返回金额和
            If Val(strMoney) > curOrig Then
                ShowMsgbox "输入的""" & .TextMatrix(Row, 0) & """支付金额不能超过 " & Format(curOrig, "0.00") & " ！"
                .EditCell
                .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True: Exit Sub
            End If
        End If
        
        '不允许超出单据剩余可结算金额
        curTotal = mCurCharge.dbl当前未付
        For p = 1 To mInsure.colBalance.Count
            For Each objItem In mInsure.colBalance(p)
                If objItem.结算方式 <> str支付方式 Then
                    curTotal = curTotal - objItem.有效金额
                End If
            Next
        Next
        If Val(strMoney) > curTotal Then
            ShowMsgbox "支付金额过大，超过单据允许支付金额:" & Format(curTotal, "0.00") & "。"
            .EditCell
            .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True: Exit Sub
        End If
        
        Call SetBalanceVal(mInsure.colBalance, 1, str支付方式 & "|" & CCur(Val(strMoney)))
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '获取当前卡
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    GetCurCard = True
    Exit Function
errHandle:
    Set objCard = New Card
End Function

Private Sub RestorePrePayTypeFromTag()
    '恢复到上次选择的支付项
    '说明:cbo支付方式.Tag存储的是上次选择的支付项内容,格式:Index:缴款金额
    Dim varTemp As Variant, intIndex As Integer
    
    On Error GoTo ErrHandler
    mCurCardPay.strQRCode = ""
    If cbo支付方式.Tag = "" Then Exit Sub
    
    '有上次选择的卡类别ID,恢复
    varTemp = Split(cbo支付方式.Tag & ":", ":")
    cbo支付方式.Tag = ""
    
    intIndex = Val(varTemp(0))
    cbo支付方式.ListIndex = intIndex
    txt金额.Text = varTemp(1)
    zlControl.ControlSetFocus txt金额
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetPayTypeFromCardTypeID(ByVal lngCardTypeID As Long, Optional ByVal bln消费卡 As Boolean) As Boolean
    '根据卡类别ID,定位到指定的支付类别上
    '入参:
    '   lngCardTypeID-卡类别ID
    '   bln消费卡-是否消费卡
    '   blnOnlyChangePayType '是否仅改变支付类别
    '返回:定位成功返回true,否则返回False
    Dim objCard As Card, blnFind As Boolean, i As Integer
    Dim intIndex As Integer
    
    On Error GoTo ErrHandler
    If lngCardTypeID <= 0 Then Exit Function
    For i = 1 To mobjPayCards.Count
        Set objCard = mobjPayCards(i)
        If objCard.接口序号 = lngCardTypeID And objCard.消费卡 = bln消费卡 Then intIndex = i: Exit For
    Next
    If intIndex = 0 Then Exit Function
    
    '卡类别ID必须在有效的支持扫码付的卡类别中
    If InStrEx(btQRCodePay.Tag, lngCardTypeID) = False Then Exit Function
    
    With cbo支付方式
        For i = 0 To .ListCount - 1
            If .ItemData(i) = intIndex Then
                .ListIndex = i
                blnFind = True: Exit For
            End If
        Next
    End With
    SetPayTypeFromCardTypeID = blnFind
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub btQRCodePay_zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
    On Error GoTo ErrHandler
    Call RestorePrePayTypeFromTag '恢复上次选择项
    If strErrMsg = "" Then Exit Sub
    ShowMsgbox strErrMsg
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub btQRCodePay_zlGetPayMoney(dblMoney As Double, strExpend As String, blnCancel As Boolean)
    Dim varTemp As Variant
    
    On Error GoTo ErrHandler
    cbo支付方式.Tag = cbo支付方式.ListIndex & ":" & txt金额.Text '先记录原支付信息
    zlControl.ControlSetFocus txt金额
    Call txt金额_GotFocus
    
    '定位到指定卡类别
    dblMoney = Val(txt金额.Text)
    varTemp = Split(btQRCodePay.Tag & ",", ",") '存储了有效的卡类别IDs
    If SetPayTypeFromCardTypeID(Val(varTemp(0))) = False Then
        ShowMsgbox "不存在指定的扫码付类别，请检查！"
        blnCancel = True
        Call RestorePrePayTypeFromTag '恢复上次选择项
        Exit Sub
    End If
    
    '获取本次支付金额
    txt金额.Text = Format(dblMoney, "0.00")
    
    If RoundEx(dblMoney, 6) <= 0 Then
        If RoundEx(dblMoney, 6) = 0 Then
            ShowMsgbox "病人未付金额为零，不需要进行扫码付款！"
        Else
            ShowMsgbox "当前为退款，扫码付不支持退款操作！"
        End If
        blnCancel = True
        Call RestorePrePayTypeFromTag '恢复上次选择项
        Exit Sub
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    blnCancel = True
    Call RestorePrePayTypeFromTag '恢复上次选择项
End Sub

Private Sub btQRCodePay_zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, blnCancel As Boolean)
    '进行扫码付款
    '入参:
    '   lngCardTypeID-卡类别ID
    '   strPayMentQRCode-二维码付款内码
    '   strExpendXML-暂无
    '出参:strExpendXML-暂无
    '     blnCancel-true表示取消本次扫码付,False-表示本次扫码付成功
    
    On Error GoTo ErrHandler
    If lngCardTypeID = 0 Or blnCancel Then
        blnCancel = True
        Call RestorePrePayTypeFromTag '恢复上次选择项
        Exit Sub
    End If
    
    blnCancel = False
    If SetPayTypeFromCardTypeID(lngCardTypeID, False) = False Then    '定位到扫码付的指定类别上
        ShowMsgbox "不能有效识别当前扫码付的类别，可能本机不支持该类别的扫码付，请与管理员联系！"
        blnCancel = True
        Call RestorePrePayTypeFromTag '恢复上次选择项
        Exit Sub
    End If
    
    mCurCardPay.strQRCode = strPayMentQRCode
    Call cmdOK_Click
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    blnCancel = True
    Call RestorePrePayTypeFromTag '恢复上次选择项
End Sub

