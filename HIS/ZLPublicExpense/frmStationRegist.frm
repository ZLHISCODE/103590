VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmStationRegist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医生站挂号"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   Icon            =   "frmStationRegist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrice 
      Caption         =   "保存划价单(&J)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3780
      TabIndex        =   49
      Top             =   5760
      Width           =   1725
   End
   Begin VB.CheckBox chkAll 
      Height          =   360
      Left            =   8055
      Picture         =   "frmStationRegist.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "显示不当班别"
      Top             =   60
      Width           =   345
   End
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   2940
      Picture         =   "frmStationRegist.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "新增病人(F4)"
      Top             =   600
      Width           =   350
   End
   Begin VB.PictureBox picPayMoney 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   6645
      ScaleHeight     =   360
      ScaleWidth      =   1695
      TabIndex        =   37
      Top             =   4942
      Width           =   1755
      Begin VB.Label lblPayMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   930
         TabIndex        =   38
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.PictureBox picInfo 
      Height          =   2925
      Left            =   15
      ScaleHeight     =   2865
      ScaleWidth      =   8310
      TabIndex        =   31
      Top             =   1950
      Width           =   8370
      Begin VB.CommandButton cmdOther 
         Height          =   345
         Left            =   4440
         Picture         =   "frmStationRegist.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "其他医生号别"
         Top             =   45
         Width           =   345
      End
      Begin VB.CheckBox chk复诊 
         Caption         =   "复诊"
         Height          =   255
         Left            =   7500
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   930
      End
      Begin VB.TextBox txtReg 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         TabIndex        =   4
         Top             =   45
         Width           =   3360
      End
      Begin VB.CommandButton cmdReg 
         Height          =   345
         Left            =   4020
         Picture         =   "frmStationRegist.frx":1788
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "当前医生号别"
         Top             =   45
         Width           =   345
      End
      Begin VB.CheckBox chkBook 
         Caption         =   "购买病历"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6990
         TabIndex        =   9
         Top             =   2543
         Width           =   1485
      End
      Begin VB.ComboBox cboDoctor 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   525
         Width           =   3390
      End
      Begin VB.ComboBox cboAppointStyle 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   525
         Width           =   1875
      End
      Begin VB.ComboBox cboRemark 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         TabIndex        =   8
         Top             =   2490
         Width           =   6120
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   1440
         Left            =   75
         TabIndex        =   32
         Top             =   975
         Width           =   8205
         _cx             =   1985886345
         _cy             =   1985874412
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStationRegist.frx":218A
         ScrollTrack     =   0   'False
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
      Begin VB.Label lblSn 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "序号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7635
         TabIndex        =   50
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
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
         Left            =   120
         TabIndex        =   40
         Top             =   585
         Width           =   480
      End
      Begin VB.Label lblAppointStyle 
         AutoSize        =   -1  'True
         Caption         =   "预约方式"
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
         Left            =   4365
         TabIndex        =   36
         Top             =   585
         Width           =   960
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "号别"
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
         Left            =   120
         TabIndex        =   35
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lblLimit 
         AutoSize        =   -1  'True
         Caption         =   "限号:"
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
         Left            =   4980
         TabIndex        =   34
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "备注"
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
         Left            =   105
         TabIndex        =   33
         Top             =   2550
         Width           =   480
      End
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   795
      ScaleHeight     =   360
      ScaleWidth      =   1575
      TabIndex        =   29
      Top             =   4950
      Width           =   1635
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   825
         TabIndex        =   30
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   3
      Left            =   -45
      TabIndex        =   25
      Top             =   5490
      Width           =   11000
   End
   Begin VB.ComboBox cboPayMode 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4935
      Width           =   1950
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   2
      Left            =   -30
      TabIndex        =   20
      Top             =   1440
      Width           =   11000
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   1
      Left            =   -30
      TabIndex        =   19
      Top             =   480
      Width           =   11000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6975
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5595
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   90
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   0
      Left            =   -60
      TabIndex        =   17
      Top             =   960
      Width           =   11000
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   360
      Left            =   705
      TabIndex        =   16
      Top             =   600
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   635
      Appearance      =   2
      IDKindStr       =   "姓|姓名或就诊卡|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|"
      BorderStyle     =   1
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
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   1650
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6510
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4140
      TabIndex        =   28
      Top             =   1568
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   360
      Left            =   3045
      TabIndex        =   3
      Top             =   1560
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   93323266
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picRoom 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5700
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   44
      Top             =   1560
      Width           =   2655
      Begin VB.Label lblRoomName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   45
         Top             =   15
         Width           =   120
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   675
      TabIndex        =   2
      Top             =   1560
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   93323267
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picDept 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   675
      ScaleHeight     =   300
      ScaleWidth      =   3330
      TabIndex        =   42
      Top             =   1560
      Width           =   3390
      Begin VB.Label lblDeptName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   43
         Top             =   15
         Width           =   120
      End
   End
   Begin VB.Label lbl时段 
      AutoSize        =   -1  'True
      Caption         =   "上午"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4500
      TabIndex        =   51
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lbl急 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "急"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   135
      TabIndex        =   46
      Top             =   45
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblPayMode 
      AutoSize        =   -1  'True
      Caption         =   "支付方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3315
      TabIndex        =   24
      Top             =   4995
      Width           =   1320
   End
   Begin VB.Label lblSum 
      AutoSize        =   -1  'True
      Caption         =   "合计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   23
      Top             =   4995
      Width           =   660
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "门诊预交余额:0.00     "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3615
      TabIndex        =   18
      Top             =   645
      Width           =   2880
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "病人"
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
      Left            =   135
      TabIndex        =   15
      Top             =   645
      Width           =   480
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "单据号"
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
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "性别:     年龄:       门诊号:              费别: "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   135
      TabIndex        =   39
      Top             =   1110
      Width           =   5880
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "日期"
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
      Left            =   135
      TabIndex        =   26
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "时间"
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
      Left            =   2505
      TabIndex        =   27
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "诊室"
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
      Left            =   5130
      TabIndex        =   21
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "科室"
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
      Left            =   135
      TabIndex        =   22
      Top             =   1620
      Width           =   480
   End
End
Attribute VB_Name = "frmStationRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset, mblnStartFactUseType As Boolean
Private mblnCard As Boolean, mintSysAppLimit As Integer
Private mfrmPatiInfo As frmPatiInfo
Private mstrYBPati As String, mlng挂号ID As Long, mlng领用ID As Long
Private mblnOlnyBJYB As Boolean, mblnSharedInvoice As Boolean
Private mstr险类 As String, mblnAppointment As Boolean, mblnChangeFeeType As Boolean
Private mstrAge As String, mstr费别 As String, mstr性别 As String, mstr门诊号 As String
Private mstrPassWord As String, mblnUnload As Boolean, mstrInsure As String
Private mlngDept As Long
Private Const SNCOLS = 10
Private Const SnArgCols = 7
Private mrsPlan As ADODB.Recordset, mblnInit As Boolean
Private mrsSNState As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsItems As ADODB.Recordset
Private mrs时间段 As ADODB.Recordset
Private mrsInComes As ADODB.Recordset
Private mcolCardPayMode As Collection
Private mcolArrangeNo As Collection
Private mblnIntact As Boolean, mstrUseType As String
Private mlng病人ID As Long, mintIDKind As Integer
Private mcur个帐余额 As Currency
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mlngSN As Long
Private mintInsure As Integer, mblnUpdateAge As Boolean
Private mdatLast As Date
Private mblnChangeByCode As Boolean
Private mstrCardNO As String
Private mcur个帐透支 As Currency
Private Enum EM_REGISTFEE_MODE  '挂号费用收取方式
        EM_RG_现收 = 0
        EM_RG_划价 = 1
        EM_RG_记帐 = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '病人收费模式
    EM_先结算后诊疗 = 0
    EM_先诊疗后结算 = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '挂号费用收取方式
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '病人收费模式

Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    使用个人帐户   As Boolean  'support挂号使用个人帐户
    连续挂号  As Boolean    'support连续挂号
    不收病历费 As Boolean   'support挂号不收取病历费
    挂号检查项目 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum ViewMode
     V_普通号
     v_专家号
     v_专家号分时段
     V_普通号分时段
End Enum
Private mViewMode As ViewMode

Private Type ty_ModulePara
    bln姓名模糊查找 As Boolean
    lng姓名查找天数 As Long
    bln默认购买病历 As Boolean
    bln默认输入摘要 As Boolean
    byt挂号模式 As Byte
    bln挂号必须刷卡 As Boolean
    bln优先使用预交 As Boolean
    bln住院病人挂号 As Boolean
    bln挂号包含科室安排 As Boolean
    bln预约包含科室安排 As Boolean
    int挂号发票打印 As Integer
    int挂号凭条打印 As Integer
    int预约挂号打印 As Integer
    bln随机序号选择 As Boolean
    lng预约有效时间 As Long
    bln共用收费票据 As Boolean
    bln退号重用 As Boolean
    bln预约时收款 As Boolean
    dbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
    bln输入医生 As Boolean
    int同科限约数           As Integer  '同科室限约
    int同科限挂数           As Integer
    bln同科限挂急诊         As Boolean
    int病人预约科室数       As Integer
    int病人挂号科室数       As Integer
    int专家号挂号限制       As Integer
    int专家号预约限制       As Integer
    strStationRegOrder As String    '医生站挂号排序字符串
    blnShowAllPlan      As Boolean   ' 是否显示不当班号别
End Type
Private mty_Para As ty_ModulePara
Private mstr家属IDs As String
Private mstrPriceGrade As String, mintPriceGradeStartType As Integer
Private mobjRegister As clsRegist
Private mstrDef费别 As String  '缺省费别

Public Sub zlShowMe(ByVal frmMain As Object, ByVal objRegister As clsRegist, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByVal blnAppointment As Boolean, ByVal lng病人ID As Long, ByRef strOutNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医生站挂号入口
    '入参:strDeptIDs-挂号科室,支持多个,用逗号分隔
    '     blnAppointment-是否预约调用
    '     objRegister-clsRegist对象
    '出参:strOutNO-挂号成功后,传出挂号单据号
    '编制:刘尔旋
    '日期:2016-7-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    mblnAppointment = blnAppointment
    Set mobjRegister = objRegister
    mlngModul = lngModul
    mlng病人ID = lng病人ID
    
    If frmMain Is Nothing Then
        Me.Show
    Else
         Me.Show 1, frmMain
    End If
    If mblnOK = True Then
        strOutNO = mstrNO
        Unload Me
    End If
End Sub

Private Sub InitPara()
    Dim strValue As String
    On Error GoTo errH
    With mty_Para
        .bln姓名模糊查找 = Val(gobjDatabase.GetPara("姓名模糊查找", glngSys, 9000, "0")) = 1
        .lng姓名查找天数 = Val(gobjDatabase.GetPara("姓名查找天数", glngSys, 9000, 0))
        .bln默认购买病历 = Val(gobjDatabase.GetPara("默认购买病历", glngSys, 9000, "0")) = 1
        .bln默认输入摘要 = Val(gobjDatabase.GetPara("默认输入摘要", glngSys, 9000, "1")) = 1
        .byt挂号模式 = Val(gobjDatabase.GetPara("挂号模式", glngSys, 9000, "0"))
        .bln优先使用预交 = Val(gobjDatabase.GetPara("优先使用预交", glngSys, 9000, "0")) = 1
        .bln住院病人挂号 = Val(gobjDatabase.GetPara("允许住院病人挂号", glngSys, 9000, "0")) = 1
        .int挂号发票打印 = Val(gobjDatabase.GetPara("挂号发票打印方式", glngSys, 9000, "0"))
        .int挂号凭条打印 = Val(gobjDatabase.GetPara("挂号凭条打印方式", glngSys, 9000, "0"))
        .int预约挂号打印 = Val(gobjDatabase.GetPara("预约挂号单打印方式", glngSys, 9000, "0"))
        .bln随机序号选择 = Val(gobjDatabase.GetPara("随机序号选择", glngSys, 9000, "0")) = 1
        .bln共用收费票据 = Val(gobjDatabase.GetPara("挂号共用收费票据", glngSys, 1121)) = 1
        .bln退号重用 = Val(gobjDatabase.GetPara("已退序号允许挂号", glngSys, 1111)) = 1
        .bln预约时收款 = Val(gobjDatabase.GetPara("预约时收款", glngSys, 9000, "0")) = 1
        .bln挂号必须刷卡 = Val(gobjDatabase.GetPara("挂号必须刷卡", glngSys, 9000)) = 1
        strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
        If InStr(strValue, "|") = 0 Then strValue = "1|0"
        .dbl预存款消费验卡 = Val(Split(strValue, "|")(0))
        .bln输入医生 = Val(gobjDatabase.GetPara("输入医生", glngSys, 9000)) = 1
        .int同科限约数 = Val(gobjDatabase.GetPara("病人同科限约N个号", glngSys, 1111, 0))
        .int同科限挂数 = Val(Split(gobjDatabase.GetPara("病人同科限挂N个号", glngSys, 1111, 0) & "|", "|")(0))
        .bln同科限挂急诊 = Split(gobjDatabase.GetPara("病人同科限挂N个号", glngSys, 1111, 0) & "|", "|")(1) = "1"
        .int病人挂号科室数 = Val(gobjDatabase.GetPara("病人挂号科室限制", glngSys, 1111, 0))
        .int病人预约科室数 = Val(gobjDatabase.GetPara("病人预约科室数", glngSys, 1111, 0))
        .int专家号挂号限制 = Val(gobjDatabase.GetPara("专家号挂号限制", glngSys, , 0))
        .int专家号预约限制 = Val(gobjDatabase.GetPara("专家号预约限制", glngSys, , 0))
        strValue = gobjDatabase.GetPara("包含科室安排", glngSys, 9000, "0|0") & "|"
        .bln挂号包含科室安排 = Val(Split(strValue, "|")(0)) = 1
        .bln预约包含科室安排 = Val(Split(strValue, "|")(1)) = 1
        .strStationRegOrder = gobjDatabase.GetPara("医生站挂号排序控制", glngSys, 9000, "医生,1|执行时间,1|科室,1|号别,1|项目,1")
        If .bln默认输入摘要 Then
            cboRemark.TabStop = True
        Else
            cboRemark.TabStop = False
        End If
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_现收
        Else
            If (.byt挂号模式 = 0 Or .byt挂号模式 = 2) And gSysPara.bln免挂号模式 = False Then
                mRegistFeeMode = EM_RG_现收
            Else
                mRegistFeeMode = EM_RG_划价
            End If
        End If
        '是否显示不当班号别
        .blnShowAllPlan = Val(gobjDatabase.GetPara("显示不当班号别", glngSys, 9000, "0")) = 1
    End With
    
    '刷卡要求输入密码
    mstrCardPass = gobjDatabase.GetPara(46, glngSys, , "0000000000")
    Call gobjControl.PicShowFlat(picInfo)
    '收费和挂号共用票据
    mblnSharedInvoice = gobjDatabase.GetPara("挂号共用收费票据", glngSys, 1121) = "1"
    '本地共用挂号批次ID
    If mblnSharedInvoice Then
        mlng挂号ID = Val(gobjDatabase.GetPara("共用收费票据批次", glngSys, 1121, ""))
    Else
        mlng挂号ID = Val(gobjDatabase.GetPara("共用挂号票据批次", glngSys, mlngModul, ""))
    End If
    mlngDept = Val(gobjDatabase.GetPara("接诊科室", glngSys, 1260, ""))
    If mlng挂号ID > 0 Then
        If Not ExistBill(mlng挂号ID, IIf(mblnSharedInvoice, 1, 4)) Then
            If mblnSharedInvoice Then
                gobjDatabase.SetPara "共用收费票据批次", "0", glngSys, 1121
            Else
                gobjDatabase.SetPara "共用挂号票据批次", "0", glngSys, mlngModul
            End If
            mlng挂号ID = 0
        End If
    End If
    '票号严格控制
    strValue = gobjDatabase.GetPara(24, glngSys, , "00000")
    gblnBill挂号 = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    mintSysAppLimit = Val(gobjDatabase.GetPara("挂号允许预约天数", glngSys))
    If mblnSharedInvoice Then
        '挂号用门诊票据:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    
    '价格等级
    mintPriceGradeStartType = GetPriceGradeStartType()
    If mintPriceGradeStartType > 0 Then
        Call GetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function zlStartFactUseType(ByVal int票种 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strSql = "Select  1 as 存在 From 票据领用记录 where 票种=[1] and nvl(使用类别,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查票据是否启用了使用类别的", int票种)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zl_GetInvoiceUserType(ByVal lng病人ID As Long, ByVal lng主页Id As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的使用类别
    '返回:发票的使用类别
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "Select  Zl_Billclass([1],[2],[3]) as 使用类别 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取票据使用类别", lng病人ID, lng主页Id, intInsure)
    zl_GetInvoiceUserType = Nvl(rsTemp!使用类别)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'功能：判断是否存在指定的票据领用
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    
    strSql = "Select ID From 票据领用记录 Where ID=[1] And 票种=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "检查领用ID", lngID, bytKind)
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function RefreshFact(Optional ByRef strFact As String) As Boolean
'参数：blnNew=是否新单保存时调用,这时对于非严格控制的票据是保存当前号
    If mblnStartFactUseType Then
        mstrUseType = zl_GetInvoiceUserType(Val(mrsInfo!病人ID), 0, mintInsure)
    End If
    If gblnBill挂号 Then
        mlng领用ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(mlng领用ID > 0, mlng领用ID, mlng挂号ID), , mstrUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            strFact = "": Exit Function
        Else
            '严格：取下一个号码
            strFact = GetNextBill(mlng领用ID)
        End If
    Else
        If mblnSharedInvoice Then
            strFact = gobjDatabase.GetPara("当前收费票据号", glngSys, 1121)
        Else
            strFact = gobjDatabase.GetPara("当前挂号票据号", glngSys, 1111)
        End If
        strFact = IncStr(strFact)
        If mblnSharedInvoice Then
            gobjDatabase.SetPara "当前收费票据号", strFact, glngSys, 1121
        Else
            gobjDatabase.SetPara "当前挂号票据号", strFact, glngSys, 1111
        End If
    End If
    RefreshFact = True
End Function

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long, strTemp As String
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0;医|医保号|0;身|身份证号|0;门|门诊号|0;住|住院号|0;手|手机号|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If

    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function

Private Sub cboAppointStyle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cmdOther_Click()
    Call LoadRegPlans(3, , True)
End Sub

Private Sub cmdPrice_Click()
    If SaveData_Price = False Then Exit Sub
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
End Sub
Private Sub cmdReg_Click()
    Call LoadRegPlans(3)
End Sub

Private Sub SetDefultRegTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的挂号时间
    '日期:2018-02-02 15:05:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtCurDate As Date, dtSysDate As Date, strNO As String, str星期 As String, strSql As String
    Dim rsTmp As ADODB.Recordset, rsTime As ADODB.Recordset, str发生时间 As String
    Dim lng安排ID As Long, lng计划ID As Long
    Dim lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean
    On Error GoTo errH
    
    
    lblSn.Caption = ""
    str星期 = zlGet当前星期几(dtpDate.Value)
    
    If Nvl(mrsPlan.Fields(str星期)) = "" Then
        dtpTime.Value = Format(GetWorkTimeDefualtTime("白天", Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
        Exit Sub
    End If
   
    Select Case mViewMode
    Case V_普通号分时段, v_专家号分时段
        
        strSql = "Select Distinct a.序号 As ID, To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
                "From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
                "Where a.安排id = b.Id And b.号码 = [1] And" & vbNewLine & _
                " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                "      Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六'," & vbNewLine & _
                "             Null) = a.星期(+) And Not Exists" & vbNewLine & _
                " (Select Count(1)" & vbNewLine & _
                "       From 挂号序号状态" & vbNewLine & _
                "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
                "        Count(1) - a.限制数量 >= 0) And Not Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From 挂号安排计划 E" & vbNewLine & _
                "       Where e.安排id = b.Id And e.审核时间 Is Not Null And" & vbNewLine & _
                "             [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                "             e.失效时间)"
        strSql = strSql & " Union " & _
                "Select Distinct a.序号 As ID, To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
                "From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C," & vbNewLine & _
                "     (Select Max(a.生效时间) 生效" & vbNewLine & _
                "       From 挂号安排计划 A, 挂号安排 B" & vbNewLine & _
                "       Where a.安排id = b.Id And b.号码 = [1] And a.审核时间 Is Not Null And" & vbNewLine & _
                "             [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                "             a.失效时间) D" & vbNewLine & _
                "Where a.计划id = b.Id And b.安排id = c.Id And c.号码 = [1] And b.生效时间 = d.生效 And b.审核时间 Is Not Null And" & vbNewLine & _
                " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                "      [2] Between Nvl(b.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                "      b.失效时间 And Not Exists" & vbNewLine & _
                " (Select Count(1)" & vbNewLine & _
                "       From 挂号序号状态" & vbNewLine & _
                "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
                "        Count(1) - a.限制数量 >= 0) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
                "                                           '周四', '6', '周五', '7', '周六', Null) = a.星期(+)" & vbNewLine & _
                "Order By 开始时间"
    
        dtCurDate = Format(dtpDate, "yyyy-mm-dd")
        strNO = txtReg.Tag
        
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, dtCurDate)
        If Not rsTmp.EOF Then
            '时段当班有时段,取最小时段
            dtpTime.Value = Format(Nvl(rsTmp!开始时间), "hh:mm:ss")
            lblSn.Caption = "序号:" & Val(Nvl(rsTmp!ID))
            Exit Sub
        End If
        
        If GetRegData(lngSN, str发生时间, blnAdd, blnNotWork) Then
            If lngSN <> 0 Then lblSn.Caption = "序号:" & lngSN
            If str发生时间 <> "" Then
                If IsDate(str发生时间) Then dtpTime.Value = Format(CDate(str发生时间), "hh:mm:ss"): Exit Sub
            End If
        End If
       dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str星期)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
        Exit Sub
    Case v_专家号
        lng计划ID = Val(Nvl(mrsPlan!计划ID))
        lng安排ID = Val(Nvl(mrsPlan!ID))
        
        dtCurDate = Format(dtpDate, "yyyy-mm-dd")
        If mobjRegister.zlGetRegisterNextSn__Tradition(lng安排ID, lng计划ID, dtCurDate, InStr(gstrPrivs, ";加号;"), mblnAppointment, False, lngSN, str发生时间) = False Then
            dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str星期)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss"): Exit Sub
        End If
        If lngSN <> 0 Then lblSn.Caption = "序号:" & lngSN
        If mblnAppointment Then
            If Format(dtpDate.Value, "yyyy-mm-dd") > Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                str发生时间 = GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str星期)), Format(dtpDate.Value, "yyyy-mm-dd"))
            End If
        End If
        If str发生时间 <> "" Then
            If IsDate(str发生时间) Then dtpTime.Value = Format(CDate(str发生时间), "hh:mm:ss"): Exit Sub
        End If
        dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str星期)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
    Case Else
        dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str星期)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
    End Select
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetWorkTimeDefualtTime(ByVal strWorkName As String, ByVal strRegDate As String, Optional ByVal strCurSysDate As String = "") As Date
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取工作时段的缺省时间
    '入参:strWorkName-工作时段
    '    strRegDate-挂号日期（yyyy-mm-dd)
    '    strCurSysDate-当前的缺省时间
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-02 15:26:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysDate As Date, strDate As String
    Dim rsTime As ADODB.Recordset
    Dim dtRegDate As Date
    
    On Error GoTo errHandle
    If strCurSysDate = "" Then
        dtSysDate = gobjDatabase.Currentdate
    Else
        dtSysDate = CDate(strCurSysDate)
    End If
    dtRegDate = CDate(strRegDate)
    
    If Format(dtRegDate, "yyyy-mm-dd") = Format(dtSysDate, "yyyy-mm-dd") Then
        '当天
       GetWorkTimeDefualtTime = dtSysDate
    End If
    
    If mobjRegister.zlGetRegisterWorkTime_Record(rsTime) = False Then
        '当天不当班,取当前时间
        GetWorkTimeDefualtTime = CDate(Format(dtRegDate, "yyyy-mm-dd" & " " & Format(dtSysDate, "hh:mm:ss")))
        Exit Function
    End If
    rsTime.Filter = "时间段='" & strWorkName & "' and 号类=NULL and 站点=NULL"
    If rsTime.EOF Then
        rsTime.Filter = 0
        GetWorkTimeDefualtTime = dtSysDate: Exit Function
    End If
    
    If IsNull(rsTime!缺省时间) Then
        strDate = Format(dtRegDate, "yyyy-mm-dd") & " " & Format(rsTime!开始时间, "hh:mm:ss")
    Else
        strDate = Format(dtRegDate, "yyyy-mm-dd") & " " & Format(rsTime!缺省时间, "hh:mm:ss")
    End If
    rsTime.Filter = 0
    GetWorkTimeDefualtTime = CDate(strDate)
    Exit Function
errHandle:
    GetWorkTimeDefualtTime = gobjDatabase.Currentdate
End Function






Private Sub GetAll医生()
    Dim strSql As String
    On Error GoTo errH
    
    strSql = "Select a.Id, a.姓名, Upper(a.简码) As 简码,b.部门id,a.编号" & _
            " From 人员表 a, 部门人员 b, 人员性质说明 c" & _
            " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order By a.简码 Desc"
    Set mrsDoctor = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, "医生")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadDoctor()
    With cboDoctor
        .Clear
        If Nvl(mrsPlan!医生) = "" Then
            If mty_Para.bln输入医生 Then
                mrsDoctor.Filter = "部门id=" & Val(Nvl(mrsPlan!科室ID))
                
                Do While Not mrsDoctor.EOF
                    .AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    If Nvl(mrsDoctor!姓名) = UserInfo.姓名 Then .ListIndex = .NewIndex
                    mrsDoctor.MoveNext
                Loop
                If .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                .Enabled = True
                lblDoctor.Enabled = True
            Else
                mrsDoctor.Filter = "姓名='" & UserInfo.姓名 & "'"
                If mrsDoctor.RecordCount <> 0 Then
                    .AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    .ListIndex = 0
                End If
                .Enabled = False
                lblDoctor.Enabled = False
            End If
        Else
            mrsDoctor.Filter = "姓名='" & Nvl(mrsPlan!医生) & "'"
            If mrsDoctor.RecordCount <> 0 Then
                .AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
                .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                .ListIndex = 0
            End If
            .Enabled = False
            lblDoctor.Enabled = False
        End If
    End With
End Sub

Private Sub cboPayMode_Click()
    If MCPAR.不收病历费 And cboPayMode.Text = mstrInsure Then
        chkBook.Enabled = False
        chkBook.Value = 0
    Else
        chkBook.Enabled = True
    End If
End Sub

Private Sub cboPayMode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboRemark_Change()
    cboRemark.Tag = ""
End Sub

Private Sub cboRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRemark.Tag <> "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(cboRemark.Text) = "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If SelectMemo(Trim(cboRemark.Text)) = False Then
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub

Private Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String

    If Val(gobjDatabase.GetPara("输入匹配")) = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择常用摘要
    '入参:strInput-输入串;为空时,表示全部
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSql As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If gobjCommFun.IsCharChinese(cboRemark.Text) Then
             strWhere = " And  名称 like [1] "
        ElseIf gobjCommFun.IsNumOrChar(cboRemark.Text) Then
             strWhere = " And (简码 like upper([1]) or 编码 like upper([1]))"
        End If
    End If
    
    strSql = "" & _
     "   Select RowNum AS ID,编码,名称,简码  " & _
     "   From 常用挂号摘要 " & _
     "   Where 1=1 " & strWhere & _
     "   Order by 缺省标志"
     vRect = GetControlRect(cboRemark.hWnd)
     On Error GoTo Hd
     Set rsInfo = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "常用挂号摘要", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cboRemark.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "没有设置常用挂号摘要,请在字典管理中设置", vbOKOnly + vbInformation, gstrSysName
        End If
        gobjCommFun.PressKey vbKeyTab: Exit Function
     End If
     gobjControl.CboSetText Me.cboRemark, Nvl(rsInfo!名称)
     cboRemark.Tag = Nvl(rsInfo!名称)
     gobjCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then Resume
    gobjComlib.SaveErrLog
End Function

Private Sub chkBook_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub
Private Sub cmdNewPati_Click()
    Call zlExcuteMorePatiInfor
End Sub
Private Sub ResetDefault复诊()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置缺省复诊信息
    '编制:刘兴洪
    '日期:2017-10-27 15:16:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, bln复诊 As Boolean
    Dim lng科室id As Long
    
    On Error GoTo errHandle
    lng科室id = 0
    If mrsInfo Is Nothing Then chk复诊.Value = 0: Exit Sub
    If mrsInfo.State <> 1 Then chk复诊.Value = 0: Exit Sub
    If mrsPlan Is Nothing Then GoTo ReSet:
    If mrsPlan.RecordCount = 0 Then GoTo ReSet:
    lng科室id = Val(Nvl(mrsPlan!科室ID))
    
ReSet:
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    bln复诊 = zlPatiIsReturnVisit(lng病人ID, lng科室id)
    chk复诊.Value = IIf(bln复诊, 1, 0)
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlExcuteMorePatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行更多的病人信息修改
    '编制:刘兴洪
    '日期:2017-10-27 14:35:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytFun As Byte, lng病人ID As Long
    Dim lngOut病人ID As Long
    Dim lng科室id As Long
    
    On Error GoTo errHandle
    If mfrmPatiInfo Is Nothing Then Set mfrmPatiInfo = New frmPatiInfo
    bytFun = 2: lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        bytFun = 0
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        If lng病人ID = 0 Then Exit Sub
    End If
    If mrsPlan.RecordCount <> 0 Then lng科室id = Val(Nvl(mrsPlan!科室ID))
    If mfrmPatiInfo.ShowMe(Me, bytFun, lng病人ID, lngOut病人ID, lng科室id) = False Then Exit Sub
    If Not mfrmPatiInfo Is Nothing Then Unload mfrmPatiInfo
    Set mfrmPatiInfo = Nothing
    
    txtPatient.Text = "-" & lngOut病人ID
    GetPatient IDKind.GetCurCard, txtPatient.Text, False
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdTime_Click()
    If SelectTimeSn = False Then Exit Sub
End Sub

Private Sub dtpDate_Change()
    Call LoadRegPlans(1)
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub dtpTime_GotFocus()
    Call cmdTime_Click
End Sub

Private Sub dtpTime_Validate(Cancel As Boolean)
    If Format(dtpDate.Value, "YYYY-MM-DD") = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") Then
        If Format(dtpTime.Value, "hh:mm:ss") < Format(gobjDatabase.Currentdate, "hh:mm:ss") Then
            MsgBox "预约时间不能小于当前时间!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnload Then mblnUnload = False: Unload Me: Exit Sub
    If mblnInit And Not mrsInfo Is Nothing Then
        If txtReg.Enabled And txtReg.Visible Then txtReg.SetFocus
    End If
    mblnInit = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    If Not mobjIDCard Is Nothing Then
         Call mobjIDCard.SetEnabled(False)
         Set mobjIDCard = Nothing
     End If
     mstr费别 = ""
     mstrDef费别 = ""
     If Not mobjICCard Is Nothing Then
         Call mobjICCard.SetEnabled(False)
         Set mobjICCard = Nothing
     End If
     mintIDKind = IDKind.IDKind
     Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
     gobjDatabase.SetPara "显示不当班号别", IIf(mty_Para.blnShowAllPlan, 1, 0), glngSys, 9000
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strExpand As String
    Dim strOutCardNO As String, strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        '系统IC卡
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
            End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    
    If lng卡类别ID <= 0 Then Exit Sub
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
'    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
'    txtPatient.Text = strOutCardNO
'
'    If txtPatient.Text <> "" Then
'        Call GetPatient(objCard, txtPatient.Text, True)
'    End If
End Sub

Private Sub txtRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub chkBook_Click()
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    If mrsPlan.RecordCount = 0 Then Exit Sub
    Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, mstrPriceGrade)
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" Then
        If MsgBox("是否清空当前病人信息？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearPatient
        End If
        Exit Sub
    End If
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp "zl9RegEvent", Me.hWnd, "frmRegistEdit"
    Exit Sub
End Sub

Private Function CheckBrushCard(ByVal dblMoney As Double, ByVal lng医疗卡类别ID As Long, ByVal bln消费卡 As Boolean, _
                                ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMoney As ADODB.Recordset, str年龄 As String, lng病人ID As Long
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode <> EM_RG_现收 Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        CheckBrushCard = True: Exit Function
    End If
    If Not (cboPayMode.Visible And cboPayMode.Enabled) Then
        CheckBrushCard = True: Exit Function
    End If
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then
        CheckBrushCard = True: Exit Function
    End If
    If lng医疗卡类别ID = 0 Then
        MsgBox cboPayMode.Text & "异常,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "使用" & cboPayMode.Text & "支付必须先初始化接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call zlGetClassMoney(rsMoney, rsItems, rsIncomes)
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    str年龄 = Trim(mstrAge)
    If Not mrsInfo Is Nothing Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
   If gobjSquare.objSquareCard.zlBrushCard(Me, glngModul, rsMoney, lng医疗卡类别ID, bln消费卡, _
    txtPatient.Text, NeedName(mstr性别), str年龄, dblMoney, mstrCardNO, mstrPassWord, _
    False, True, False, True, Nothing, False, True, "", "1", lng病人ID) = False Then Exit Function
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, glngModul, lng医疗卡类别ID, _
        bln消费卡, mstrCardNO, dblMoney, "", "") = False Then Exit Function

    CheckBrushCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "项目ID=" & rsItems!项目ID
            rsMoney.Filter = "收费类别='" & Nvl(rsItems!类别, "无") & "'"
            If rsMoney.EOF Then
                .AddNew
            Else
                rsMoney.Filter = 0
            End If
            !收费类别 = Nvl(rsItems!类别, "无")
            Do While Not rsIncomes.EOF
                !金额 = Val(Nvl(!金额)) + Val(Nvl(rsIncomes!实收))
                rsIncomes.MoveNext
            Loop
            .Update
            rsItems.MoveNext
        Loop
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function CheckIsPatiBlacklist() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查黑名单
    '返回:合法返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:21:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln专家号 As Boolean, bytMode As Byte
    Dim strSql As String, dat预约时间 As Date
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mrsInfo Is Nothing Then Exit Function

    If mblnAppointment Then
        bytMode = 1
        dat预约时间 = CDate(Format(dtpDate.Value, "yyyy-mm-dd"))
    Else
        bytMode = 0
        dat预约时间 = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
    End If
    
    bln专家号 = Nvl(mrsPlan!医生) <> ""
    
    strSql = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],Null,[4],[5]) As 检查结果 From Dual"
    Set rsCheck = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, bytMode, Val(Nvl(mrsInfo!病人ID)), Trim(txtReg.Tag), dat预约时间, IIf(bln专家号, 1, 0))
    If rsCheck.EOF Then
        MsgBox "有效性检查失败,无法继续！", vbInformation, gstrSysName
        Exit Function
    End If

   strSql = Nvl(rsCheck!检查结果)
   If Val(Mid(strSql, 1, 1)) <> 0 Then
       MsgBox Mid(strSql, 3), vbInformation, gstrSysName
       Exit Function
   End If
    CheckIsPatiBlacklist = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDataIsValied(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '入参:bytMode-0-现收;1-记帐;2-划价
    '出参:
    '返回:数据合法返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:17:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '1.检查输入数据的合法性
    If CheckValied(bytMode) = False Then Exit Function
    
    '2.检查黑名单
    If CheckIsPatiBlacklist = False Then Exit Function
    
    '3.记帐相关检查
    If bytMode = 1 Then
       If mRegistFeeMode = EM_RG_记帐 And mty_Para.bln预约时收款 And mblnAppointment Then
           MsgBox "不支持先诊疗后结算病人的预约收款挂号！", vbInformation, gstrSysName
           Exit Function
       End If
    End If
    
    '4.费用模式检查
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.bln预约时收款) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!结算模式))) = False Then Exit Function
    End If
     
    CheckDataIsValied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPrintProofIsPrint(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取打印凭条是否需要打印
    '入参:bytMode-0-现收;1-记帐;2-划价
    '出参:
    '返回:需要打印返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo errHandle
    
    '预约时不收款，则不打印挂号凭条
    If mblnAppointment And mty_Para.bln预约时收款 = False Then Exit Function
    
    If mty_Para.int挂号凭条打印 = 0 Then Exit Function '不打印
    
    If InStr(gstrPrivs, ";病人挂号凭条;") = 0 Then
        '检查是否存在权限
        MsgBox "你没有" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mty_Para.int挂号凭条打印 = 1 Then '自动打印
        GetPrintProofIsPrint = True: Exit Function
    End If
    
    '提示打印
    If MsgBox("要打印" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "凭条吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    GetPrintProofIsPrint = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetInvoiceIsPrint(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取打印发票是否需要打印
    '入参:bytMode-0-现收;1-记帐;2-划价
    '出参:
    '返回:需要打印返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo errHandle
    
    If bytMode = 1 Or bytMode = 2 Or mblnAppointment And mty_Para.bln预约时收款 = False Then Exit Function '划价及记帐或预约时不收款不打印发票
    If mintInsure <> 0 And MCPAR.医保接口打印票据 Then Exit Function
    
    
    
    If mty_Para.int挂号发票打印 = 0 Then Exit Function '不打印
    
    If InStr(gstrPrivs, ";挂号发票打印;") = 0 Then
        '检查是否存在权限
        MsgBox "你没有" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "发票打印的权限，请联系管理员！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mty_Para.int挂号发票打印 = 1 Then '自动打印
        GetInvoiceIsPrint = True: Exit Function
    End If
    
    '提示打印
    If MsgBox("要打印" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "发票吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    GetInvoiceIsPrint = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetDepositBillIsPrint(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取预约单是否需要打印
    '入参:bytMode-0-现收;1-记帐;2-划价
    '出参:
    '返回:需要打印返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo errHandle
    
    If Not (mblnAppointment And mty_Para.bln预约时收款 = False) Then Exit Function  '只有预约单(未收款)才打印
     
    
    If mty_Para.int预约挂号打印 = 0 Then Exit Function '不打印
    
    If InStr(gstrPrivs, ";预约挂号单;") = 0 Then
        '检查是否存在权限
        MsgBox "你没有预约挂号单打印的权限，请联系管理员！！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mty_Para.int预约挂号打印 = 1 Then '自动打印
        GetDepositBillIsPrint = True: Exit Function
    End If
    
    '提示打印
    If MsgBox("要打印预约挂号单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    
    GetDepositBillIsPrint = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRegData(ByRef lngSn_Out As Long, ByRef str发生时间_Out As String, ByRef blnAdd_Out As Boolean, _
                            ByRef blnNotWork_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号相关的数据
    '入参:
    '出参:lngSn_Out-返回序号
    '     str发生时间_Out-返回发生时间
    '     blnAdd_Out-是否返回的加号
    '     blnNotWork_Out-挂号时间是否不在排班时间(False-有效挂号时间,True-不当班)
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 14:36:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str星期 As String, blnAdd As Boolean, strWorkTimeName As String, blnValied As Boolean
    Dim str发生时间 As String, lngSN As Long, strWeekName As String
    Dim lng计划ID As Long, lng安排ID As Long
    Dim dtRegDate  As Date
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    
    On Error GoTo errHandle

    str星期 = zlGet当前星期几(IIf(mblnAppointment, dtpDate.Value, ""))
  
    
    lng计划ID = Val(Nvl(mrsPlan!计划ID))
    lng安排ID = Val(Nvl(mrsPlan!ID))
  
    '获取发生时间
    blnAdd = False: lngSN = 0
    
    If mblnAppointment Then '预约处理
        
        dtRegDate = CDate(Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss"))
        
        str发生时间 = Format(dtpDate, "yyyy-mm-dd")
        strWeekName = mobjRegister.zlGetWeekNameFromDate(dtRegDate)
        
        
        If mViewMode = v_专家号分时段 Then
            If lng计划ID <> 0 Then
                    strSql = "" & _
                    " Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
                    "       Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
                    " From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码," & vbNewLine & _
                    "              To_Date('" & str发生时间 & "' || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
                    "              To_Date('" & str发生时间 & "' || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
                    "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
                    "       From 挂号安排计划 Jh, 挂号计划时段 Sd" & vbNewLine & _
                    "       Where Jh.Id = Sd.计划id And Jh.Id = [1] And" & vbNewLine & _
                    "             Sd.星期 =[3]) Jh," & vbNewLine & _
                    "     挂号序号状态 Zt" & vbNewLine & _
                    " Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Jh.开始时间 = [2] And Zt.序号(+) = Jh.序号 And Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
                    " Order By 序号"
                        
            Else
                    strSql = "" & _
                    " Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
                    "       Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
                    " From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码," & vbNewLine & _
                    "              To_Date('" & str发生时间 & "' ||' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
                    "              To_Date('" & str发生时间 & "' ||' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
                    "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
                    "       From 挂号安排 Ap, 挂号安排时段 Sd" & vbNewLine & _
                    "       Where Ap.Id = Sd.安排id And Ap.Id = [1] And" & vbNewLine & _
                    "             Sd.星期 =[3] ) Ap, 挂号序号状态 Zt" & vbNewLine & _
                    " Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Ap.开始时间 =[2] And Zt.序号(+) = Ap.序号 And Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
                    " Order By 序号"
            End If
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng计划ID <> 0, lng计划ID, lng安排ID), dtRegDate, strWeekName)
            
            '124298：李南春，2018/4/13，不当班加号时不判断权限
            If Not rsTmp.EOF Then
                rsTmp.Filter = "剩余数 <> 0"
            Else
                blnValied = True
            End If
            If rsTmp.RecordCount <> 0 Then
                lngSN = Val(Nvl(rsTmp!序号))
            Else
                strSql = "Select Max(序号) As 序号 From 挂号序号状态 Where 号码 = [1] And Trunc(日期) = Trunc(" & str发生时间 & ")"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Nvl(mrsPlan!号别))
                If rsTmp.RecordCount <> 0 Then lngSN = Val(Nvl(rsTmp!序号))
                
                
                If lng计划ID <> 0 Then
                    strSql = "" & _
                    "   Select Max(序号) As 序号 From 挂号计划时段  " & _
                    "   Where 计划ID = [1] And 星期 = [3]"
                Else
                    strSql = "" & _
                    "   Select Max(序号) As 序号 From 挂号安排时段 " & _
                    "   Where 安排ID = [1] And 星期 =[3]"
                End If
                
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng计划ID <> 0, lng计划ID, lng安排ID), dtRegDate, strWeekName)
                If lngSN = 0 Then
                    If rsTmp.RecordCount <> 0 Then lngSN = Val(Nvl(rsTmp!序号))
                Else
                    If Val(Nvl(rsTmp!序号)) > lngSN Then lngSN = Val(Nvl(rsTmp!序号))
                End If
                
                lngSN = lngSN + 1
                blnAdd = True
            End If
        End If
        If IsNull(mrsPlan.Fields(str星期).Value) Then blnValied = True
        
        str发生时间 = Format(dtRegDate, "yyyy-mm-dd HH:MM:SS")
        blnAdd_Out = blnAdd
        blnNotWork_Out = blnValied
        str发生时间_Out = str发生时间
        lngSn_Out = lngSN
        GetRegData = True
        Exit Function
    End If
    
    '挂号处理
    If mViewMode <> v_专家号分时段 Or (mViewMode = v_专家号分时段 And IsNull(mrsPlan.Fields(str星期).Value)) Then
        
        dtRegDate = gobjDatabase.Currentdate
        str发生时间 = Format(dtRegDate, "yyyy-mm-dd HH:MM:SS")
        
        If IsNull(mrsPlan.Fields(str星期).Value) Then blnValied = True
        
        blnNotWork_Out = blnValied
        str发生时间_Out = str发生时间
        lngSn_Out = lngSN
        GetRegData = True
        Exit Function
    End If
    
    dtRegDate = gobjDatabase.Currentdate
    str发生时间 = Format(dtRegDate, "yyyy-mm-dd HH:MM:SS")
    
    '检查发时间是否有效
    strWorkTimeName = Nvl(mrsPlan.Fields(str星期).Value)
    If mobjRegister.zlCheckIsValiedWorkTimeFromWorkTimeName(str发生时间, strWorkTimeName, "", "", blnValied) = False Then Exit Function
    
    If blnValied Then   '不当班
        blnNotWork_Out = blnValied
        str发生时间_Out = str发生时间
        lngSn_Out = lngSN
        GetRegData = True
        Exit Function
    End If
    
    strWeekName = mobjRegister.zlGetWeekNameFromDate(dtRegDate)
    
    
    If lng计划ID <> 0 Then
            strSql = "Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
            "       Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
            "From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
            "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
            "       From 挂号安排计划 Jh, 挂号计划时段 Sd" & vbNewLine & _
            "       Where Jh.Id = Sd.计划id And Jh.Id = [1] And" & vbNewLine & _
            "             Sd.星期 =[2]) Jh," & vbNewLine & _
            "     挂号序号状态 Zt" & vbNewLine & _
            "Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
            "Order By 序号"
    Else
        strSql = "Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
            "       Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
            "From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
            "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
            "       From 挂号安排 Ap, 挂号安排时段 Sd" & vbNewLine & _
            "       Where Ap.Id = Sd.安排id And Ap.Id = [1] And" & vbNewLine & _
            "             Sd.星期 =  [2]) Ap, 挂号序号状态 Zt" & vbNewLine & _
            "Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
            "Order By 序号"
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng计划ID <> 0, lng计划ID, lng安排ID), strWeekName)
    If Not rsTmp.EOF Then rsTmp.Filter = "剩余数 <> 0"
    
    '取最小可用时间段
    If rsTmp.RecordCount <> 0 Then
        lngSN = Val(Nvl(rsTmp!序号))
        str发生时间 = Format(dtRegDate, "yyyy-mm-dd") & " " & Format(Nvl(rsTmp!开始时间), "hh:mm:ss")
    End If
     
    blnAdd_Out = blnAdd
    str发生时间_Out = str发生时间
    lngSn_Out = lngSN
    GetRegData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Get付款方式编码(ByVal str付款方式 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取付款方式编码
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 18:38:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rs付款方式 As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "Select 编码 From 医疗付款方式 Where 名称 = [1]"
    Set rs付款方式 = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, str付款方式)
    If rs付款方式.RecordCount <> 0 Then
        Get付款方式编码 = Nvl(rs付款方式!编码)
    Else
        strSql = "Select 编码 From 医疗付款方式 Where 缺省标志 = 1"
        Set rs付款方式 = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
        If rs付款方式.RecordCount <> 0 Then
            Get付款方式编码 = Nvl(rs付款方式!编码)
        End If
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub PrintInvoic(ByVal strNO As String, ByVal strFactNO As String, ByVal dat登记时间 As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印发票
    '入参:strNo-单据号
    '     strFactNo-发票号
    '编制:刘兴洪
    '日期:2018-02-01 11:12:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    'If mintInsure <> 0 And MCPAR.医保接口打印票据 Then Exit Sub:不应该判断(刘兴洪)RePrint:
RePrint:
    Load frmPrint
    Call frmPrint.ReportPrint(1, strNO, "", mlng领用ID, mlng挂号ID, strFactNO, dat登记时间, , , , mintInsure <> 0 And MCPAR.医保接口打印票据, False, mstrUseType)
    If Not gblnBill挂号 Then Exit Sub
    
    If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
        If MsgBox(IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "单号为[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
         Exit Sub
    End If
    
End Sub

Private Function SaveData_Cash() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存现收数据
    '入参:
    '出参:
    '返回:保存成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:13:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim i As Long, strFactNO As String, blnBalance As Boolean
    Dim blnProofPrint As Boolean  '凭条打印
    Dim blnInvoiceIsPrint  As Boolean '发票打印
    Dim blnDepositBillIsPrint As Boolean '预约挂号单打印
    Dim cur预交 As Currency, cur个帐 As Currency, cur现金 As Currency
    Dim dat登记时间 As Date, str登记时间 As String, dt发生时间 As Date
    Dim str发生时间 As String, lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean, lngValue As Long
    Dim lng医疗卡类别ID As Long, bln消费卡 As Boolean, blnNoDoc As Boolean, str结算方式 As String
    Dim blnTrans As Boolean, lng结帐ID As Long
    Dim cllProAfter As Collection, cllPro As Collection
    Dim str交易流水号 As String, str交易说明   As String
    Dim strNO As String, blnOneCard As Boolean
    Dim strSql As String, blnNotCommit As Boolean, strAdvance As String
    Dim cllTheeSwap As Collection, cllTheeSwapOther As Collection
    
    On Error GoTo errHandle
    'bytMode-0-现收;1-记帐;2-划价
    If CheckDataIsValied(0) = False Then Exit Function
    
    blnProofPrint = GetPrintProofIsPrint(0) '凭条打印
    blnInvoiceIsPrint = GetInvoiceIsPrint(0)    '发票打印
    blnDepositBillIsPrint = GetDepositBillIsPrint(0)   '预约挂号单打印
    
    '确定发票号
    If blnInvoiceIsPrint Or (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
        If RefreshFact(strFactNO) = False Then Exit Function
    End If
    
    blnBalance = False
    
    If mblnAppointment = False Or mblnAppointment And mty_Para.bln预约时收款 Then
        If cboPayMode.Text = "预交金" Then
            cur预交 = Val(lblTotal.Caption)
        Else
            If cboPayMode.Text = mstrInsure Then
                cur个帐 = Val(lblTotal.Caption)
            Else
                blnBalance = True
                cur现金 = Val(lblTotal.Caption)
            End If
        End If
    End If
    If Val(cur预交) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!病人ID), Val(cur预交), mlngModul, 1, , _
                             IIf(-1 * mty_Para.dbl预存款消费验卡 >= Val(cur预交), False, True), True, mstr家属IDs, (mty_Para.dbl预存款消费验卡 <> 0), (mty_Para.dbl预存款消费验卡 = 2)) Then Exit Function
    End If
    
    
    ReadRegistPrice Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, False, mstr费别, rsItems, rsIncomes, _
        Nvl(mrsInfo!病人ID), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
        
    '其他检查
    str结算方式 = ""
    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lng医疗卡类别ID = mcolCardPayMode.Item(i)(3)
                bln消费卡 = Val(mcolCardPayMode.Item(i)(5)) = 1
                str结算方式 = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur现金), lng医疗卡类别ID, bln消费卡, rsItems, rsIncomes) = False Then Exit Function
        If str结算方式 = "" Then str结算方式 = cboPayMode.Text
    End If

    dat登记时间 = gobjDatabase.Currentdate
    str登记时间 = Format(dat登记时间, "yyyy-mm-dd HH:MM:SS")
    
    '124431:李南春,2018/5/17，序号排重检查
    lngValue = Val(Split(lblSn.Caption & ":", ":")(1))
    If lngValue <> 0 Then
        If mblnAppointment Then
            dt发生时间 = CDate(Format(dtpDate, "yyyy-mm-dd"))
        Else
            dt发生时间 = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
        End If
        '医生站不判断锁号和预留
        strSql = "Select 1 From 挂号序号状态 " & _
                "  Where 号码 = [1] And Trunc(日期) = [2] And 序号 = [3]" & vbNewLine & _
                IIf(mty_Para.bln退号重用, " And 状态 <> 4", "")
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "序号检查", Nvl(mrsPlan!号别), dt发生时间, lngValue)
        If rsTemp.RecordCount > 1 Then
            If MsgBox("号序 " & lngValue & " 已被他人使用，是否自动获取可用时段进行" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    '124298：李南春，2018/4/13，不当班加号时不判断权限
    If GetRegData(lngSN, str发生时间, blnAdd, blnNotWork) = False Then Exit Function
    
    '137272:李南春,2019/2/20,挂号锁号
    If mViewMode = v_专家号 Or mViewMode = v_专家号分时段 Then
        If ReserveRegNo(Nvl(mrsPlan!号别), True, mViewMode = v_专家号分时段, str发生时间, lngSN, "医生站锁号") = False Then Exit Function
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!已挂)) >= Val(Nvl(mrsPlan!限号)) And Val(Nvl(mrsPlan!限号)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!已约)) >= Val(Nvl(mrsPlan!限约)) And Val(Nvl(mrsPlan!限约)) <> 0 And mblnAppointment Then
            blnAdd = True
        End If
    End If
    
    If blnAdd And InStr(gstrPrivs, ";加号;") = 0 Then
        MsgBox "你没有加号权限，无法对当前号别进行" & IIf(gSysPara.bln免挂号模式 And mblnAppointment = False, "就诊", "挂号") & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";加号;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
    
    blnOneCard = zlOldOneCardIsStart(cboPayMode.Text)
    If GetSaveRegDataSQL(0, rsItems, rsIncomes, False, str结算方式, cur预交, cur个帐, cur现金, lngSN, blnAdd, str发生时间, str登记时间, cllPro, cllProAfter, lng结帐ID, strNO, "", lng医疗卡类别ID, bln消费卡, mstrCardNO, str交易流水号, str交易说明) = False Then Exit Function
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

    blnTrans = True
    If blnOneCard And lng医疗卡类别ID <> 0 And cur现金 <> 0 Then
        If Not mobjICCard.PaymentSwap(Val(cur现金), Val(cur现金), Val(lng医疗卡类别ID), 0, mstrCardNO, "", lng结帐ID, Nvl(mrsInfo!病人ID)) Then
            gcnOracle.RollbackTrans
            MsgBox "一卡通结算挂号费失败", vbInformation, gstrSysName
            Exit Function
        End If
        strSql = "zl_一卡通结算_Update(" & lng结帐ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lng医疗卡类别ID & "','" & "" & "'," & cur现金 & ")"
        Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If

    
    
    blnNotCommit = False
    If mintInsure <> 0 And mstrYBPati <> "" And cur个帐 <> 0 Then
        '68991:strAdvance:结算模式(0或1)|挂号费收取方式(0或1) |挂号单号
        strAdvance = ""
        If mPatiChargeMode = EM_先诊疗后结算 Then
            strAdvance = IIf(mPatiChargeMode = EM_先诊疗后结算, "1", "0")
            strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_记帐, "1", "0")
            strAdvance = strAdvance & "|" & strNO
        End If
        If Not gclsInsure.RegistSwap(lng结帐ID, cur个帐, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans:   Exit Function
        End If
        blnNotCommit = True
    End If
        
        
    '问题:31187 调用医保成功后,最后作一些数据更新:内部过程中已有提交语句,所以不用再写
    zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
    
    Set cllTheeSwap = New Collection: Set cllTheeSwapOther = New Collection
    If Not blnOneCard And Not mPatiChargeMode = EM_先诊疗后结算 And cur现金 <> 0 Then
        If zlInterfacePrayMoney(Me, mlngModul, lng结帐ID, cllTheeSwap, cllTheeSwapOther, Val(cur现金), mstrCardNO, lng医疗卡类别ID, bln消费卡) = False Then gcnOracle.RollbackTrans: Exit Function
        '修正三方交易
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
    End If
    
    
    Err = 0: On Error GoTo OthersCommit:
    zlExecuteProcedureArrAy cllTheeSwapOther, Me.Caption, False, False
        
OthersCommit:
    gcnOracle.CommitTrans: blnTrans = False
    
    If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
    
    '145198:李南春,2019/12/26,挂号成功后调用外挂接口，目前用于预约后产生支付二维码
    Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.bln预约时收款))
    If blnInvoiceIsPrint Then
        Call PrintInvoic(strNO, strFactNO, dat登记时间)
    End If
    If blnProofPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '记录打印的凭条
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    
    If blnDepositBillIsPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    mstrNO = strNO
    
    SaveData_Cash = True
    Exit Function
errHandle:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function




Private Function SaveData_Accounting() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存现收数据
    '入参:
    '出参:
    '返回:保存成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:13:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    Dim i As Long
    Dim blnProofPrint As Boolean  '凭条打印
    Dim blnDepositBillIsPrint  As Boolean   '预约挂号单打印
    Dim dat登记时间 As Date, str登记时间 As String
    Dim str发生时间 As String, lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean
    Dim blnTrans As Boolean, lng结帐ID As Long
    Dim cllProAfter As Collection, cllPro As Collection
    Dim cur预交 As Currency, cur个帐 As Currency, cur现金 As Currency
    Dim str结算方式 As String, blnBalance As Boolean
    Dim lng医疗卡类别ID As Long, bln消费卡 As Boolean, blnNoDoc As Boolean
    Dim strNO As String, str交易流水号 As String, str交易说明 As String
    Dim blnNotCommit As Boolean, strAdvance As String
    
    
    On Error GoTo errHandle
    
    'bytMode-0-现收;1-记帐;2-划价
    If CheckDataIsValied(1) = False Then Exit Function
    blnProofPrint = GetPrintProofIsPrint(1) '凭条打印
    blnDepositBillIsPrint = GetDepositBillIsPrint(1)   '预约挂号单打印
    

    ReadRegistPrice Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, False, mstr费别, rsItems, rsIncomes, _
        Nvl(mrsInfo!病人ID), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
       
       
    dat登记时间 = gobjDatabase.Currentdate
    str登记时间 = Format(dat登记时间, "yyyy-mm-dd HH:MM:SS")
        
    '124298：李南春，2018/4/13，不当班加号时不判断权限
    If GetRegData(lngSN, str发生时间, blnAdd, blnNotWork) = False Then Exit Function
    
    '137272:李南春,2019/2/20,挂号锁号
    If mViewMode = v_专家号 Or mViewMode = v_专家号分时段 Then
        If ReserveRegNo(Nvl(mrsPlan!号别), True, mViewMode = v_专家号分时段, str发生时间, lngSN, "医生站锁号") = False Then Exit Function
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!已挂)) >= Val(Nvl(mrsPlan!限号)) And Val(Nvl(mrsPlan!限号)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!已约)) >= Val(Nvl(mrsPlan!限约)) And Val(Nvl(mrsPlan!限约)) <> 0 And mblnAppointment Then
            blnAdd = True
        End If
    End If
    
    If blnAdd And InStr(gstrPrivs, ";加号;") = 0 Then
        MsgBox "你没有加号权限，无法对当前号别进行" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";加号;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
  
    If cboPayMode.Text = "预交金" Then
        cur预交 = Val(lblTotal.Caption)
    Else
        If cboPayMode.Text = "个人帐户" Then
            cur个帐 = Val(lblTotal.Caption)
        Else
            blnBalance = True
            cur现金 = Val(lblTotal.Caption)
        End If
    End If
    If Val(cur预交) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!病人ID), Val(cur预交), mlngModul, 1, , _
                           IIf(-1 * mty_Para.dbl预存款消费验卡 >= Val(cur预交), False, True), True, mstr家属IDs, (mty_Para.dbl预存款消费验卡 <> 0), (mty_Para.dbl预存款消费验卡 = 2)) Then Exit Function
    End If
    
    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lng医疗卡类别ID = mcolCardPayMode.Item(i)(3)
                bln消费卡 = Val(mcolCardPayMode.Item(i)(5)) = 1
                str结算方式 = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur现金), lng医疗卡类别ID, bln消费卡, rsItems, rsIncomes) = False Then Exit Function
        If str结算方式 = "" Then str结算方式 = str结算方式 = cboPayMode.Text
    End If
     
    If GetSaveRegDataSQL(1, rsItems, rsIncomes, False, str结算方式, cur预交, cur个帐, cur现金, lngSN, blnAdd, str发生时间, str登记时间, cllPro, cllProAfter, lng结帐ID, strNO, "", lng医疗卡类别ID, bln消费卡, mstrCardNO, str交易流水号, str交易说明) = False Then Exit Function

    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, False
    
    
    blnNotCommit = False
    If mintInsure <> 0 And mstrYBPati <> "" And cur个帐 <> 0 Then
        '68991:strAdvance:结算模式(0或1)|挂号费收取方式(0或1) |挂号单号
        strAdvance = IIf(mPatiChargeMode = EM_先诊疗后结算, "1", "0")
        strAdvance = strAdvance & "|" & "1"
        strAdvance = strAdvance & "|" & strNO
        If Not gclsInsure.RegistSwap(lng结帐ID, cur个帐, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans:   Exit Function
        End If
        blnNotCommit = True
    End If
        
    zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
    gcnOracle.CommitTrans: blnTrans = False
    If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
    
    '145198:李南春,2019/12/26,挂号成功后调用外挂接口，目前用于预约后产生支付二维码
    Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.bln预约时收款))
    If blnProofPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '记录打印的凭条
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    If blnDepositBillIsPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    mstrNO = strNO
    SaveData_Accounting = True
    Exit Function
errHandle:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function


Private Function SaveData_Price() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存为划价单
    '入参:
    '出参:
    '返回:保存成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 11:41:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    Dim blnProofPrint As Boolean  '凭条打印
    Dim blnDepositBillIsPrint As Boolean    '预约挂号单打印
    Dim dat登记时间 As Date, str登记时间 As String
    Dim str发生时间 As String, lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean
    Dim str划价NO As String, blnTrans As Boolean, lng结帐ID As Long, strNO As String
    Dim cllProAfter As Collection, cllPro As Collection
    
    On Error GoTo errHandle
    
    'bytMode-0-现收;1-记帐;2-划价
    If CheckDataIsValied(2) = False Then Exit Function
    blnProofPrint = GetPrintProofIsPrint(2) '凭条打印
    blnDepositBillIsPrint = GetDepositBillIsPrint(2)   '预约挂号单打印
    

    ReadRegistPrice Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, False, mstr费别, rsItems, rsIncomes, _
     Nvl(mrsInfo!病人ID), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
     
     
    dat登记时间 = gobjDatabase.Currentdate
    str登记时间 = Format(dat登记时间, "yyyy-mm-dd HH:MM:SS")
    
    '124298：李南春，2018/4/13，不当班加号时不判断权限
    If GetRegData(lngSN, str发生时间, blnAdd, blnNotWork) = False Then Exit Function
    
    '137272:李南春,2019/2/20,挂号锁号
    If mViewMode = v_专家号 Or mViewMode = v_专家号分时段 Then
        If ReserveRegNo(Nvl(mrsPlan!号别), True, mViewMode = v_专家号分时段, str发生时间, lngSN, "医生站锁号") = False Then Exit Function
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!已挂)) >= Val(Nvl(mrsPlan!限号)) And Val(Nvl(mrsPlan!限号)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!已约)) >= Val(Nvl(mrsPlan!限约)) And Val(Nvl(mrsPlan!限约)) <> 0 And mblnAppointment Then
            blnAdd = True
        End If
    End If
    
    If blnAdd And InStr(gstrPrivs, ";加号;") = 0 Then
        MsgBox "你没有加号权限，无法对当前号别进行" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";加号;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
    
    mlngSN = lngSN
    
    If GetSaveRegDataSQL(2, rsItems, rsIncomes, False, "", 0, 0, 0, lngSN, blnAdd, str发生时间, str登记时间, cllPro, cllProAfter, lng结帐ID, strNO, str划价NO) = False Then Exit Function
    
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, False
    zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
    gcnOracle.CommitTrans: blnTrans = False
    
    '145198:李南春,2019/12/26,挂号成功后调用外挂接口，目前用于预约后产生支付二维码
    Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.bln预约时收款))
    If blnProofPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '记录打印的凭条
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    If blnDepositBillIsPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    mstrNO = strNO
    SaveData_Price = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
 


Private Function GetSaveRegDataSQL(ByVal bytMode As Byte, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, ByVal blnInvoicePrint As Boolean, _
    ByVal str结算方式 As String, ByVal dbl冲预交 As Double, ByVal dbl个人帐户 As Double, ByVal dbl现金 As Double, ByVal lngSN As Long, ByVal blnAddNum As Boolean, _
    ByVal str发生时间 As String, ByVal str登记时间 As String, ByRef cllPro_out As Collection, ByRef cllProAffter_out As Collection, _
    Optional lng结帐ID_Out As Long, Optional strNO_Out As String, Optional strPriceNo_Out As String, _
    Optional ByVal lng支付类别ID As Long = 0, Optional ByVal bln消费卡 As Boolean = False, Optional ByVal str卡号 As String = "", Optional ByVal str交易流水号 As String = "", Optional ByVal str交易说明 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号保存数据
    '入参:bytMode-0-现收;1-记帐;2-划价
    '     rsItem-项目集
    '     rsInComes-收入项目集
    '     blnInvoicePrint-是否发票打印
    '     strBalances-结算方式
    '     blnAddNum:是否加号
    '出参:cllPro_out-返回数据保存集
    '     cllProAffter_out-返回后执行的SQL集
    '     lng结帐ID_Out_Out-结帐ID
    '     strNO_out-单据号
    '     strPriceNo_Out-划价单号
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 17:46:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, k As Long, i As Long, j As Long, int价格父号 As Integer
    Dim lng挂号科室ID As Long, byt复诊 As Byte
    Dim lng医生ID As Long, str医生姓名 As String
    Dim str付款方式编码 As String
    Dim dblTotal As Double, blnNoDoc As Boolean
    
    
    On Error GoTo errHandle
    strPriceNo_Out = ""
    
    rsItems.Filter = ""
    str医生姓名 = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lng医生ID = 0
    Else
        lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    
    str付款方式编码 = Get付款方式编码(Nvl(mrsInfo!医疗付款方式))
    
    lng挂号科室ID = Val(Nvl(mrsPlan!科室ID))
    byt复诊 = IIf(chk复诊.Value = 1, 1, 0)
    
    If bytMode = 2 Then '存为划价单
        dblTotal = GetRegistMoney(True, False)
        '挂号费存为零且保存为划价单，才产生划价NO
        If dblTotal <> 0 Then strPriceNo_Out = gobjDatabase.GetNextNo(13)
    End If
    
    If cllPro_out Is Nothing Then Set cllPro_out = New Collection
    If cllProAffter_out Is Nothing Then Set cllProAffter_out = New Collection
    
    lng结帐ID_Out = 0
    If bytMode <> 1 And (Not mblnAppointment Or (mblnAppointment And mty_Para.bln预约时收款)) Then      '预约收款
        lng结帐ID_Out = gobjDatabase.GetNextId("病人结帐记录")
    End If
    strNO_Out = gobjDatabase.GetNextNo(12)
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int价格父号 = k
        rsIncomes.Filter = "项目ID=" & rsItems!项目ID
        For j = 1 To rsIncomes.RecordCount
        
            strSql = "zl_病人挂号记录_INSERT("
            '  病人id_In        门诊费用记录.病人id%Type,
            strSql = strSql & "" & ZVal(Nvl(mrsInfo!病人ID)) & ","
            '  门诊号_In        门诊费用记录.标识号%Type,
            strSql = strSql & "" & IIf(mstr门诊号 = "", "NULL", mstr门诊号) & ","
            '  姓名_In          门诊费用记录.姓名%Type,
            strSql = strSql & "'" & txtPatient.Text & "',"
            '  性别_In          门诊费用记录.性别%Type,
            strSql = strSql & "'" & mstr性别 & "',"
            '  年龄_In          门诊费用记录.年龄%Type,
            strSql = strSql & "'" & mstrAge & "',"
            '  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
            strSql = strSql & "'" & str付款方式编码 & "',"
            '  费别_In          门诊费用记录.费别%Type,
            strSql = strSql & "'" & mstr费别 & "',"
            '  单据号_In        门诊费用记录.No%Type,
            strSql = strSql & "'" & strNO_Out & "',"
            '  票据号_In        门诊费用记录.实际票号%Type,
            strSql = strSql & "'" & IIf(blnInvoicePrint = False, "", "") & "',"
            '  序号_In          门诊费用记录.序号%Type,
            strSql = strSql & "" & k & ","
            '  价格父号_In      门诊费用记录.价格父号%Type,
            strSql = strSql & "" & IIf(int价格父号 = k, "NULL", int价格父号) & ","
            '  从属父号_In      门诊费用记录.从属父号%Type,
            strSql = strSql & "" & IIf(rsItems!性质 = 2, 1, "NULL") & ","
            '  收费类别_In      门诊费用记录.收费类别%Type,
            strSql = strSql & "'" & rsItems!类别 & "',"
            '  收费细目id_In    门诊费用记录.收费细目id%Type,
            strSql = strSql & "" & rsItems!项目ID & ","
            '  数次_In          门诊费用记录.数次%Type,
            strSql = strSql & "" & rsItems!数次 & ","
            '  标准单价_In      门诊费用记录.标准单价%Type,
            strSql = strSql & "" & rsIncomes!单价 & ","
            '  收入项目id_In    门诊费用记录.收入项目id%Type,
            strSql = strSql & "" & rsIncomes!收入项目ID & ","
            '  收据费目_In      门诊费用记录.收据费目%Type,
            strSql = strSql & "'" & rsIncomes!收据费目 & "',"
            '  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
            strSql = strSql & "'" & str结算方式 & "',"
            '  应收金额_In      门诊费用记录.应收金额%Type,
            strSql = strSql & "" & IIf(bytMode = 2 And dblTotal <> 0, 0, Val(Nvl(rsIncomes!应收))) & ","
            '  实收金额_In      门诊费用记录.实收金额%Type,
            strSql = strSql & "" & IIf(bytMode = 2, 0, Val(Nvl(rsIncomes!实收))) & ","
            '  病人科室id_In    门诊费用记录.病人科室id%Type,
            strSql = strSql & "" & lng挂号科室ID & ","
            '  开单部门id_In    门诊费用记录.开单部门id%Type,
            strSql = strSql & "" & lng挂号科室ID & ","
            '  执行部门id_In    门诊费用记录.执行部门id%Type,
            strSql = strSql & "" & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & ","
            '  操作员编号_In    门诊费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In    门诊费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  发生时间_In      门诊费用记录.发生时间%Type,
            strSql = strSql & "to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  登记时间_In      门诊费用记录.登记时间%Type,
            strSql = strSql & "to_date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  医生姓名_In      挂号安排.医生姓名%Type,
            strSql = strSql & "'" & str医生姓名 & "',"
            '  医生id_In        挂号安排.医生id%Type,
            strSql = strSql & "" & ZVal(lng医生ID) & ","
            '  病历费_In Number, --该条记录是否病历工本费
            strSql = strSql & "" & IIf(rsItems!性质 = 3, 1, IIf(rsItems!性质 = 4, 2, 0)) & ","
            '  急诊_In          Number,
            strSql = strSql & "" & IIf(lbl急.Visible, 1, 0) & ","
            '  号别_In          挂号安排.号码%Type,
            strSql = strSql & "'" & txtReg.Tag & "',"
            '  诊室_In          门诊费用记录.发药窗口%Type,
            strSql = strSql & "'" & IIf(str医生姓名 = UserInfo.姓名, lblRoomName.Caption, "") & "',"
            '  结帐id_In        门诊费用记录.结帐id%Type,
            strSql = strSql & "" & ZVal(lng结帐ID_Out) & ","
            '  领用id_In        票据使用明细.领用id%Type,
            strSql = strSql & "" & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng领用ID)) & ","
            '  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
            strSql = strSql & "" & ZVal(IIf(k = 1, dbl冲预交, 0)) & ","
            '  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
            strSql = strSql & "" & ZVal(IIf(k = 1, dbl现金, 0)) & ","
            '  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
            strSql = strSql & "" & ZVal(IIf(k = 1, dbl个人帐户, 0)) & ","
            '  保险大类id_In    门诊费用记录.保险大类id%Type,
            strSql = strSql & "" & ZVal(Nvl(rsItems!保险大类ID, 0)) & ","
            '  保险项目否_In    门诊费用记录.保险项目否%Type,
            strSql = strSql & "" & ZVal(Nvl(rsItems!保险项目否, 0)) & ","
            '  统筹金额_In      门诊费用记录.统筹金额%Type,
            strSql = strSql & "" & ZVal(Nvl(rsIncomes!统筹金额, 0)) & ","
            '  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
            strSql = strSql & "'" & Trim(cboRemark.Text) & "',"
            '  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
            strSql = strSql & "" & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 0, 1), 0) & ","
            '  收费票据_In      Number := 0, --挂号是否使用收费票据
            strSql = strSql & "" & IIf(mty_Para.bln共用收费票据, 1, 0) & ","
            '  保险编码_In      门诊费用记录.保险编码%Type,
            strSql = strSql & "'" & rsItems!保险编码 & "',"
            '  复诊_In          病人挂号记录.复诊%Type := 0,
            strSql = strSql & "" & byt复诊 & ","
            '  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
            strSql = strSql & "" & ZVal(lngSN) & ","
            '  社区_In          病人挂号记录.社区%Type := Null,
            strSql = strSql & "" & "NULL" & ","
            '  预约接收_In      Number := 0,
            strSql = strSql & "" & IIf(mblnAppointment, 1, 0) & ","
            '  预约方式_In      预约方式.名称%Type := Null,
            strSql = strSql & "'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "',"
            '  生成队列_In      Number := 0,
            strSql = strSql & "" & 0 & ","
            '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & IIf(lng支付类别ID <> 0 And bln消费卡 = False, lng支付类别ID, "NULL") & ","
            '  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
            strSql = strSql & "" & IIf(lng支付类别ID <> 0 And bln消费卡, lng支付类别ID, "NULL") & ","
            '  卡号_In          病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & mstrCardNO & "',"
            '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
            strSql = strSql & "'" & str交易流水号 & "',"
            '  交易说明_In      病人预交记录.交易说明%Type := Null,
            strSql = strSql & "'" & str交易说明 & "',"
            '  合作单位_In      病人预交记录.合作单位%Type := Null,
            strSql = strSql & " NULL,"
            '  操作类型_In      Number := 0,
            strSql = strSql & IIf(blnAddNum, 1, 0) & ","
            '  险类_In          病人挂号记录.险类%Type := Null,
            strSql = strSql & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  结算模式_In      Number := 0,
            strSql = strSql & IIf(mPatiChargeMode = EM_先诊疗后结算, 1, 0) & ","
            '  记帐费用_In      Number := 0,
            strSql = strSql & IIf(bytMode = 1, 1, 0) & ","
            '  退号重用_In      Number := 1,
            strSql = strSql & IIf(mty_Para.bln退号重用, 1, 0) & ","
            '  冲预交病人ids_In Varchar2 := Null,
            strSql = strSql & "'" & Nvl(mrsInfo!病人ID) & "," & mstr家属IDs & "',"
            '  修正病人费别_In  Number := 0,
            strSql = strSql & "" & IIf(mblnChangeFeeType, 1, 0) & ","
            '  修正病人年龄_In  Number := 0,
            strSql = strSql & "" & IIf(mblnUpdateAge, 1, 0) & ","
            '  收费单_In        病人挂号记录.收费单%Type := Null,
            strSql = strSql & "'" & strPriceNo_Out & "')"
            '  更新交款余额_In Number:=1
            
            Call zlAddArray(cllPro_out, strSql)
            
            '问题:31187:将挂号汇总单独出来
            If txtReg.Tag <> "" And k = 1 Then
                If Nvl(mrsPlan!医生) = "" Then blnNoDoc = True
                strSql = "zl_病人挂号汇总_Update("
                '  医生姓名_In   挂号安排.医生姓名%Type,
                strSql = strSql & IIf(blnNoDoc, "Null,", "'" & str医生姓名 & "',")
                '  医生id_In     挂号安排.医生id%Type,
                strSql = strSql & "" & IIf(blnNoDoc, "0,", ZVal(lng医生ID) & ",")
                '  收费细目id_In 门诊费用记录.收费细目id%Type,
                strSql = strSql & "" & Val(Nvl(rsItems!项目ID)) & ","
                '  执行部门id_In 门诊费用记录.执行部门id%Type,
                strSql = strSql & "" & IIf(Val(Nvl(rsItems!执行科室ID)) = 0, lng挂号科室ID, Val(Nvl(rsItems!执行科室ID))) & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                strSql = strSql & "to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'),"
       
                '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
                strSql = strSql & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 3, 1), 0) & ","
                '  号码_In       挂号安排.号码%Type := Null
                strSql = strSql & "'" & txtReg.Tag & "')"
                Call zlAddArray(cllProAffter_out, strSql)
            End If
            
            If bytMode = 2 And dblTotal <> 0 Then
                strSql = _
                "zl_门诊划价记录_Insert('" & strPriceNo_Out & "'," & k & "," & ZVal(Nvl(mrsInfo!病人ID)) & ",NULL," & _
                         IIf(mstr门诊号 = "", "NULL", mstr门诊号) & ",'" & str付款方式编码 & "'," & _
                         "'" & txtPatient.Text & "','" & mstr性别 & "','" & mstrAge & "'," & _
                         "'" & mstr费别 & "',NULL," & lng挂号科室ID & "," & _
                         IIf(lng挂号科室ID <> 0, lng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                         rsItems!项目ID & ",'" & rsItems!类别 & "','" & rsItems!计算单位 & "'," & _
                         "NULL,1," & rsItems!数次 & ",NULL," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & _
                         rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "'," & rsIncomes!单价 & "," & _
                         rsIncomes!应收 & "," & rsIncomes!实收 & ",to_Date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'),NULL,'" & UserInfo.姓名 & "','挂号:" & strNO_Out & "')"
                Call zlAddArray(cllPro_out, strSql)
            End If
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i

    If Not mblnAppointment Then
        If str医生姓名 = UserInfo.姓名 Then
            strSql = "ZL_病人挂号记录_更新诊室('" & strNO_Out & "'," & Nvl(mrsInfo!病人ID) & ",'" & lblRoomName.Caption & "','" & UserInfo.姓名 & "','','','" & zl_Get预约方式ByNo(strNO_Out) & "')"    '问题号:48350
            Call zlAddArray(cllPro_out, strSql)
            strSql = "zl_病人接诊(" & Nvl(mrsInfo!病人ID) & ",'" & strNO_Out & "',NULL,'" & UserInfo.姓名 & "','" & lblRoomName.Caption & "')"
            Call zlAddArray(cllPro_out, strSql)
        End If
    End If
    
    GetSaveRegDataSQL = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Select Case mRegistFeeMode
    Case EM_RG_划价
        If SaveData_Price = False Then GoTo ToFail
    Case EM_RG_记帐
        If SaveData_Accounting = False Then GoTo ToFail
    Case EM_RG_现收
        If SaveData_Cash = False Then GoTo ToFail
    End Select
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
    Exit Sub
ToFail:
    If mViewMode = v_专家号 Or mViewMode = v_专家号分时段 Then Call CancelRegNo
End Sub

Private Sub ReloadPage()
    On Error GoTo errHandle
    Call LoadRegPlans(1)
    Call ClearPatient
    Call ClearRegInfo
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function zlIsAllowPatiChargeFeeMode(ByVal lng病人ID As Long, ByVal int原结算模式 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许改变病人收费模式
    '入参:lng病人ID-病人ID
    '       int原结算模式-0表示先结算后诊疗;1表示先诊疗后结算
    '返回:允许调整收费模式,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-12-25 10:06:49
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
'    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function '预约不处理
    '模式未调整，直接返回true
    If int原结算模式 = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If int原结算模式 = 1 Then
        '原为先诊疗后结算且存在未结费用的,则必须采用记帐模式
        strSql = "" & _
        "   Select 1 " & _
        "   From 病人未结费用 " & _
        "   Where 病人id = [1] And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID)
        If rsTemp.EOF = False Then
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算且存在未结费用，" & _
                                          vbCrLf & "不允许调整该病人的就诊模式,你可以先对未结费用结帐后" & _
                                          vbCrLf & "再挂号或不调整病人的就诊模式", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = -1 * Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        dtDate = DateAdd("d", intDay, gobjDatabase.Currentdate)
        ' 上次为"先诊疗后结算",本次为"先结算后诊疗"的,同时满足未发生医嘱业务数据的 ,
        '   则不允许更改就诊模式
        strSql = "Select 1 " & _
        " From 病人挂号记录 A, 病人医嘱记录 B " & _
        " Where a.病人id + 0 = b.病人id And a.No || '' = b.挂号单  " & _
        "               And a.记录状态 = 1 And a.记录性质 = 1 And a.登记时间 - 0 >= [2] " & _
        "               And  a.病人id = [1] And rownum<2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, dtDate)
        If rsTemp.EOF Then
            '未发生医嘱数据
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算," & vbCrLf & "  不允许调整该病人的就诊模式!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    zlIsAllowPatiChargeFeeMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub ClearRegInfo()
    txtReg.Text = ""
    txtReg.Tag = ""
    lblDeptName.Caption = ""
    lblRoomName.Caption = ""
    cboRemark.Text = ""
    chkBook.Value = IIf(mty_Para.bln默认购买病历, 1, 0)
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = "0.00"
    lblPayMoney.Caption = "0.00"
    txtPatient.SetFocus
    lbl急.Visible = False
End Sub

Private Function zlIsNotSucceedPrintBill(ByVal bytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否已经正常打印
    '入参:bytType-1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '       strNos-本次打印票据的单据,用逗号分离
    '出参:strOutValidNos-打印失败的单据号
    '返回:存在不存功票据的打印,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-16 18:06:01
    '问题:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSql As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    strSql = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From 票据使用明细 A,票据打印内容 B,Table( f_Str2list([2])) J" & _
        " Where A.打印ID =b.ID And B.数据性质=[1] And B.No=J.Column_value "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查票据是否打印", bytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckValied(ByVal bytMode As Byte) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检检输入值 的合法性
    '入参:bytMode-0-现收;1-记帐;2-划价
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-02 16:18:48
    '---------------------------------------------------------------------------------------------------------------------------------------------


    Dim i As Integer
    '保存前检查
    If mrsInfo Is Nothing Then
        MsgBox "无法确定病人信息,请先选择一个病人！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "无法确定病人信息,请先选择一个病人！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan Is Nothing Then
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.State = 0 Then
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.RecordCount = 0 Then
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mstr费别 = "" Then
        MsgBox "病人费别不能为空,请先选择一个费别！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If bytMode = 0 Then  'bytMode-0-现收;1-记帐;2-划价
        If cboPayMode.Text = "" And cboPayMode.Visible And Val(lblTotal.Caption) <> 0 Then
            MsgBox "没有确定可用的结算方式,不能完成挂号!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        If IsNull(mrsPlan!排班) Then
            MsgBox "预约不收款模式下,不能挂不当班的号别!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Nvl(mrsInfo!姓名) <> txtPatient.Text Then
        If MsgBox("当前病人姓名已经发生变化,是否重新读取病人信息?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Call GetPatient(IDKind.GetCurCard, txtPatient.Text, False)
            Exit Function
        Else
            txtPatient.Text = Nvl(mrsInfo!姓名)
        End If
    End If
    
    If InStr(gstrPrivs, ";挂号费别打折;") = 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "你没有权限给病人使用打折费别,不能完成" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    '服务对象检查
    If Not mrsItems Is Nothing Then
        mrsItems.MoveFirst
        Do While Not mrsItems.EOF
            If Val(Nvl(mrsItems!项目ID)) <> 0 Then
                If CheckServeRange(0, Val(Nvl(mrsItems!项目ID))) = False Then Exit Function
            End If
            mrsItems.MoveNext
        Loop
        mrsItems.MoveFirst
    End If
    
    CheckValied = True
End Function

Private Function CheckServeRange(intType As Integer, lng收费细目ID As Long, Optional intRow As Integer = 0) As Boolean
'功能:检查收费项目的服务对象,intType:0-门诊调用;1-住院调用
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = "Select 名称,Nvl(服务对象,0) As 服务对象 From 收费项目目录 Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "CheckServeRange", lng收费细目ID)
    If rsTmp.EOF Then
        MsgBox "不能确定" & IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目的服务对象,请检查项目是否正确录入!", vbInformation, gstrSysName
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!服务对象) = 2 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于门诊,请检查!", vbInformation, gstrSysName
                Exit Function
            End If
        Case 1
            If Val(rsTmp!服务对象) = 1 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于住院,请检查!", vbInformation, gstrSysName
                Exit Function
            End If
        Case Else
            If Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于病人,请检查!", vbInformation, gstrSysName
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function


Private Sub chkAll_Click()
    mty_Para.blnShowAllPlan = chkAll.Value <> 0
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub SetControl()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTemp As String, i As Integer
    If mblnAppointment Then
        lblRoom.Visible = False
        picRoom.Visible = False
        lblDept.Left = lblRoom.Left
        picDept.Left = picRoom.Left
        picDept.Width = picRoom.Width
        chkBook.Value = 0
        chkBook.Visible = False
        cboRemark.Width = 7170
        If mty_Para.bln预约时收款 Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
            cmdPrice.Visible = True
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
            cmdPrice.Visible = False
        End If
        cboAppointStyle.Clear
        strSql = "Select 名称,缺省标志 From 预约方式"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
        Do While Not rsTmp.EOF
            cboAppointStyle.AddItem Nvl(rsTmp!名称)
            If Val(Nvl(rsTmp!缺省标志)) = 1 Then cboAppointStyle.ListIndex = cboAppointStyle.NewIndex
            rsTmp.MoveNext
        Loop
        strTemp = gobjDatabase.GetPara("缺省预约方式", glngSys, 9000, "")
        For i = 0 To cboAppointStyle.ListCount - 1
            If cboAppointStyle.List(i) = strTemp Then
                cboAppointStyle.ListIndex = i
            End If
        Next i
    Else
        lblDate.Visible = False
        lblTime.Visible = False
        dtpDate.Visible = False
        dtpTime.Visible = False
        cmdTime.Visible = False
        
        If (mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2) And gSysPara.bln免挂号模式 = False Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
            cmdPrice.Visible = mty_Para.byt挂号模式 = 2
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
            cmdPrice.Visible = False
            
        End If
    End If
End Sub

Private Sub Form_Load()
    Err = 0
    mblnInit = True
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    
    lblSn.Caption = ""
    If mblnAppointment Then
        Me.Caption = "医生站预约"
        lblAppointStyle.Visible = True
        cboAppointStyle.Visible = True
    Else
        Me.Caption = "医生站" & IIf(gSysPara.bln免挂号模式, "直接就诊", "挂号")
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
    End If
    
    gobjDatabase.ExecuteProcedure "Zl_挂号安排_Autoupdate", Me.Caption
    Call Init费别
    Call InitPara
    chkBook.Value = IIf(mty_Para.bln默认购买病历, 1, 0)
    Call InitIDKind
    Call InitAppointmentTime
    Call GetAll医生
    chkAll.Value = IIf(mty_Para.blnShowAllPlan, 1, 0)
    If LoadRegPlans(1) = False Then
        mblnUnload = True
    End If
    Call LoadPayMode
    Call SetControl
    If mblnAppointment And mlng病人ID <> 0 Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False)
    End If
    cmdNewPati.Enabled = InStr(gstrPrivs, ";挂号病人建档;") > 0
    cmdOther.Enabled = InStr(gstrPrivs, ";允许挂其他医生的号源;") > 0
    '137272:李南春,2019/2/20,防止锁号后系统意外崩溃的情况
    Call CancelRegNo
End Sub

Private Sub InitAppointmentTime()
    '初始化预约时间
    Dim int预约天数 As Integer
    Dim dtNow As Date
    
    On Error GoTo ErrHandler
    int预约天数 = mintSysAppLimit
    If mblnAppointment Then
        Call mobjRegister.zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition( _
            gstrDeptIDs, UserInfo.姓名, mty_Para.bln预约包含科室安排, int预约天数)
    End If

    dtNow = gobjDatabase.Currentdate
    dtpDate.minDate = Format(dtNow, "yyyy-mm-dd")
    dtpDate.MaxDate = Format(dtNow + int预约天数, "yyyy-mm-dd")
    dtpDate.Value = Format(dtNow + 1, "yyyy-mm-dd")
    dtpTime.Value = Format(dtNow, "hh:mm:ss")
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.卡号
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

 

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：医保身份验卡
    '编制：刘兴洪
    '日期：2010-07-14 11:32:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim str病人类型 As String
    Dim rsTmp As ADODB.Recordset
    Dim cur余额 As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng病人ID = 0
        str病人类型 = ""
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        str病人类型 = Nvl(mrsInfo!病人类型)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '结算模式(0-先结算后诊疗或1-先诊疗后结算)|挂号费收取方式(0-现收或1-记帐)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng病人ID, mintInsure, strAdvance)
    
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_现收
    Else
        If (mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2) And gSysPara.bln免挂号模式 = False Then
            mRegistFeeMode = EM_RG_现收
            picPayMoney.Visible = True
            cboPayMode.Visible = True
            lblPayMode.Visible = True
            cmdPrice.Visible = mty_Para.byt挂号模式 = 2
        Else
            mRegistFeeMode = EM_RG_划价
            picPayMoney.Visible = False
            cboPayMode.Visible = False
            lblPayMode.Visible = False
            cmdPrice.Visible = False
        End If
    End If
    
    mPatiChargeMode = EM_先结算后诊疗
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng病人ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If

    If zlPatiCardCheck(1, lng病人ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    MCPAR.挂号检查项目 = gclsInsure.GetCapability(support挂号检查项目, lng病人ID, mintInsure)
    txtPatient.Text = "-" & lng病人ID
    Call txtPatient_Validate(False)    '其中的Setfocus调用使本事件(txtPatient_KeyPress)执行完后,不会再次自动执行txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str病人类型, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True

    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_先诊疗后结算, EM_先结算后诊疗)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_记帐, EM_RG_现收)
    End If
    MCPAR.不收病历费 = gclsInsure.GetCapability(support挂号不收取病历费, lng病人ID, mintInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    mlng领用ID = 0
    curMoney = GetRegistMoney
    Set rsTmp = GetMoneyInfoRegist(lng病人ID, , , 1, , , True)
    Dim dbl家属余额 As Double
    cur余额 = 0
    Do While Not rsTmp.EOF
        cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
        cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
        If Val(Nvl(rsTmp!家属)) = 1 Then
            dbl家属余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
        End If
        rsTmp.MoveNext
    Loop
    If cur余额 > 0 Then
        lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00") & _
                        IIf(FormatEx(dbl家属余额, 6) <> 0, "(含家属:" & Format(dbl家属余额, "0.00") & ")", "")
        If cur余额 >= curMoney Then
            blnDeposit = True
        Else
            blnDeposit = False
        End If
    End If
    
    mcur个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
    lblMoney.Caption = lblMoney.Caption & "/个人帐户余额:" & Format(mcur个帐余额, "0.00")
    If gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure) = False Then
        blnInsure = False
    Else
        If mcur个帐余额 + mcur个帐透支 >= curMoney Then
            blnInsure = True
        Else
            blnInsure = False
        End If
    End If
    Call LoadPayMode(blnDeposit, blnInsure)
    If mRegistFeeMode = EM_RG_记帐 Then
        lblSum.Caption = "记帐"
        picPayMoney.Visible = False
        cboPayMode.Visible = False
        lblPayMode.Visible = False
        cmdPrice.Visible = False
        
    Else
        lblSum.Caption = "合计"
    End If
    If mRegistFeeMode = EM_RG_现收 Then
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_现收
        Else
            If (mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2) And gSysPara.bln免挂号模式 = False Then
                mRegistFeeMode = EM_RG_现收
                picPayMoney.Visible = True
                cboPayMode.Visible = True
                lblPayMode.Visible = True
                cmdPrice.Visible = mty_Para.byt挂号模式 = 2
            Else
                mRegistFeeMode = EM_RG_划价
                picPayMoney.Visible = False
                cboPayMode.Visible = False
                lblPayMode.Visible = False
                cmdPrice.Visible = False
                
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-门诊号,1-姓名,2-挂号单,3-就诊卡号,4-医保号
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    '医保验证
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    If KeyAscii <> 0 And KeyAscii > 32 And mty_Para.bln挂号必须刷卡 Then
        sngNow = Timer
        If txtPatient.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(txtPatient.Text) + 1), "0.000") >= 0.04 Then    '>0.007>=0.01
            txtPatient.Text = Chr(KeyAscii)
            txtPatient.SelStart = 1
            KeyAscii = 0
            sngBegin = sngNow
        End If
    End If
    
    strKind = IDKind.GetCurCard.名称
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.卡号密文规则 <> "", "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.bln缺省卡号密文)
        intLen = gobjSquare.int缺省卡号长度
    Case "门诊号"
        If InStr("0123456789-" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "身份证"
    Case "医保号"
    Case Else
            If IDKind.GetCurCard.接口序号 <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.卡号密文规则 <> "")
                intLen = IDKind.GetCurCard.卡号长度
            End If
    End Select
    
    '刷卡完毕或输入号码后回车
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0: mblnCard = True
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        mblnCard = False
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    CheckNoValied = True
End Function

Private Function zl_Get预约方式ByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据挂号单据号获取病人预约方式
    '入参:strNo-挂号单据号
    '返回:预约方式
    '编制:王吉
    '日期:2012-07-03
    '问题号:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim str预约方式 As String
    Dim rsTemp As Recordset
    strSql = "" & _
        "Select 预约方式 From 病人挂号记录 Where 记录状态=1 And No=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取预约方式", strNO)
    If rsTemp Is Nothing Then zl_Get预约方式ByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_Get预约方式ByNo = "": Exit Function
    While rsTemp.EOF = False
        str预约方式 = Nvl(rsTemp!预约方式)
        rsTemp.MoveNext
    Wend
    zl_Get预约方式ByNo = str预约方式
End Function

Public Function Get失约号(ByVal str号别 As String, ByVal datThis As Date) As Long
   '获取安排在某一天.预约失约数
    Dim strSql  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strDat  As String
'    If mty_Para.bln失约用于挂号 = False Or mty_Para.lng预约有效时间 <= 0 Then Exit Function
    strSql = "                " & " SELECT count(1) AS 失约号 "
    strSql = strSql & vbNewLine & " FROM 挂号序号状态 "
    strSql = strSql & vbNewLine & " WHERE 号码=[1] AND 状态=2 AND 日期-[3]/24/60 <SYSDATE AND To_Char(日期,'yyyy-MM-dd')=[2]"
    strDat = Format(datThis, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str号别, strDat, mty_Para.lng预约有效时间)
    If rsTmp.EOF Then
        Get失约号 = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    Get失约号 = Val(Nvl(rsTmp!失约号, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIndex As Long
    Dim strSql As String
    Dim lng病人ID As Long
    Dim strAge As String, strBirth As String
    Dim str费别 As String, str付款方式 As String
    
    On Error GoTo errH
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strID:
        If txtPatient.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True, , True)
        IDKind.IDKind = lngPreIndex
        '未找到病人,自动建档
        If txtPatient.Text = "" And strName <> "" Then
            If InStr(gstrPrivs, ";挂号病人建档;") > 0 Then
                If MsgBox("未找到身份证号为[" & strID & "]的病人,是否自动建档?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    lng病人ID = gobjDatabase.GetNextNo(1)
                    If IsDate(datBirthDay) Then
                        strAge = ReCalcOld(datBirthDay, , , False)
                        strBirth = "To_Date('" & Format(datBirthDay, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    Else
                        strBirth = "Null"
                    End If
                    str费别 = GetDefault("费别")
                    str付款方式 = GetDefault("医疗付款方式")
                    strSql = "ZL_挂号病人病案_INSERT(1," & lng病人ID & ",Null,Null,Null,'" & strName & "','" & strSex & "','" & strAge & "','" & _
                            str费别 & "','" & str付款方式 & "','" & strNation & "',Null,Null,Null,'" & strID & "',Null,Null,Null,Null,'" & strAddress & "',Null," & _
                            "Null,Sysdate,Null," & strBirth & ",Null,Null,Null,Null,'" & strAddress & "')"
                    gobjDatabase.ExecuteProcedure strSql, Me.Caption
                    txtPatient.Text = "-" & lng病人ID
                    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False, , , True)
                End If
            Else
                MsgBox "未找到身份证号为[" & strID & "]的病人,不能继续!", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Function GetDefault(strItem As String) As String
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = "Select 名称 From " & strItem & " Where 缺省标志 = 1"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then GetDefault = Nvl(rsTmp!名称)
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF3
            If txtPatient.Visible = True And txtPatient.Enabled Then
                Call txtPatient.SetFocus
            End If
        Case vbKeyF4
            If cmdNewPati.Enabled And cmdNewPati.Visible Then Call cmdNewPati_Click
        Case Else
            IDKind.ActiveFastKey
    End Select
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strNO
        If txtPatient.Text = "" Then
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, _
                    Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean, _
                    Optional blnNoPrompt As Boolean = False)
    '功能：获取病人信息
    '参数：blnCard=是否就诊卡刷卡
    '      blnInputIDCard-是否身份证刷卡
    '出参:Cancel-为true表示返回的放弃读取病人信息
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim strInputInfo As String '保存传入的输入文本 避免在使用身份证号 对病人进行查找后 被替换成"-" 病人ID的情况
    Dim i As Integer, strPati As String, dat出生日期
    Dim vRect As RECT, str非在院 As String
    Dim bln医保号 As Boolean, rsFeeType As ADODB.Recordset
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '非法卡类别

    strInputInfo = strInput
    
    On Error GoTo errH
    bln医保号 = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard
    

      strSql = "Select  A.病人ID,A.门诊号,A.住院号,A.就诊卡号,A.费别,A.医疗付款方式,A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号,A.其他证件,A.身份,A.职业,A.民族,A.病人类型, " & _
               "A.国籍,A.籍贯,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.户口地址, " & _
               "A.户口地址邮编,A.Email,A.QQ,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人,A.担保额,A.担保性质,A.就诊时间,A.就诊状态, " & _
               "A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间,A.在院,A.IC卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号, " & _
               "B.名称 险类名称,A.查询密码 As 卡验证码,A.结算模式,a.主页ID From 病人信息 A,保险类别 B Where A.险类 = B.序号(+) And A.停用时间 is NULL "

    If mty_Para.bln住院病人挂号 = False Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID   And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If
   
    If blnCard And objCard.名称 Like "姓名*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
        Else
            lng卡类别ID = -1
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        If IDKind.IsMobileNo(strInput) And lng病人ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        End If
        If lng病人ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSql = strSql & " And A.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSql = strSql & " And A.病人ID=[2]" & _
        IIf(mstrYBPati <> "", "", str非在院)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then
        '住院号
        strSql = strSql & " And A.病人ID=(Select Max(病人ID) As 病人ID From 病案主页 Where 住院号 = [2])" & str非在院
    ElseIf blnInputIDCard Then  '单独的身份证识别
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        strInput = "-" & lng病人ID
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    ElseIf objCard.名称 Like "姓名*" And IDKind.IsMobileNo(strInput) = True Then
        If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Sub
        strInput = "-" & lng病人ID
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If Not mty_Para.bln姓名模糊查找 Or mty_Para.bln姓名模糊查找 And Len(txtPatient.Text) < 2 Then
                    Set mrsInfo = Nothing: Exit Sub
                End If
                strPati = _
                    " Select distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                    " From 病人信息 A " & _
                    " Where Rownum <101 And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院 & _
                    IIf(mty_Para.lng姓名查找天数 = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                    
'                strPati = strPati & " Union ALL " & _
'                        "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by 排序ID,姓名"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mty_Para.lng姓名查找天数)
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '当作新病人
                        txtPatient.Text = ""
                        If Not blnNoPrompt Then MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '以病人ID读取
                        strInput = rsTmp!病人ID
                        strSql = strSql & " And A.病人ID=[1]"
                    End If
                Else '取消选择
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "医保号"
                strInput = UCase(strInput)
                bln医保号 = True
                If mblnOlnyBJYB And gobjCommFun.ActualLen(strInput) >= 9 Then
                    strSql = strSql & " And A.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSql = strSql & " And A.医保号=[1]" & str非在院
                End If
            Case "手机号"
                If IDKind.IsMobileNo(strInput) = False Then Exit Sub
                If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Sub
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
                 
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.门诊号=[1]" & str非在院
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.病人ID=(Select Max(病人ID) As 病人ID From 病案主页 Where 住院号 = [1])" & str非在院
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                strSql = strSql & " And A.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If Mid(mstrCardPass, 1, 1) = "1" And strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "病人身份验证失败！", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        txtPatient.Text = Nvl(mrsInfo!姓名) '会调用Change事件
        txtPatient.BackColor = &H80000005
        lblSum.Caption = "合计"
        Call SetControl
        '在调用txtPatient_Change事件后在门诊号和病人姓名都为空的情况下 无法识别该病人信息 出现错误
        '对这类数据库数据错误不再进行后续的处理
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(Trim(mstr险类) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        mstr性别 = Nvl(mrsInfo!性别)
        txtPatient.PasswordChar = ""
        
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        If Load费别(Nvl(mrsInfo!费别)) = False Then mstr费别 = ""
        
        mstrAge = Nvl(mrsInfo!年龄)
        
        mblnUpdateAge = False
        If Not IsNull(mrsInfo!出生日期) Then
            strSql = "Select Zl_Age_Calc([1],[2],Null) As Old From Dual"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, CDate(mrsInfo!出生日期))
            If mstrAge <> Nvl(rsTmp!old) And Nvl(rsTmp!old) <> "" Then
                mblnUpdateAge = True
                mstrAge = Nvl(rsTmp!old)
            End If
        End If
        
        mstr门诊号 = Nvl(mrsInfo!门诊号)
        If mstr门诊号 = "" Then
            mstr门诊号 = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        lblInfo.Caption = "性别:" & mstr性别 & "   年龄:" & mstrAge & "   门诊号:" & mstr门诊号 & "   费别:" & mstr费别
        
        '病人预交款信息
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!病人ID, , , 1, , , True)
        Dim dbl家属余额 As Double
        cur余额 = 0
        Do While Not rsTmp.EOF
            cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
            cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
            If Val(Nvl(rsTmp!家属)) = 1 Then
                dbl家属余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
            End If
            rsTmp.MoveNext
        Loop
        If cur余额 > 0 Then
            lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00") & _
                            IIf(FormatEx(dbl家属余额, 6) <> 0, "(含家属:" & Format(dbl家属余额, "0.00") & ")", "")
            curMoney = GetRegistMoney
            If cur余额 >= curMoney Then
                Call LoadPayMode(True)
            Else
                Call LoadPayMode
            End If
        Else
            lblMoney.Caption = "门诊预交余额:0.00"
            Call LoadPayMode
        End If
        
        Call ResetDefault复诊 '缺省读取复诊标志
        
        '根据病人重新读取项目费用
        If mintPriceGradeStartType >= 2 Then
           Call GetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Nvl(mrsInfo!医疗付款方式), , , mstrPriceGrade)
        End If
                
        If Not mrsPlan Is Nothing Then
            If Not mrsPlan.EOF Then Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, mstrPriceGrade)
        End If
        
        cmdNewPati.ToolTipText = "详细信息(F4)"
        cmdNewPati.Enabled = True
        If txtReg.Enabled And txtReg.Visible Then txtReg.SetFocus
    Else
NewPati:
        If Not blnNoPrompt Then MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    mstr性别 = ""
    mstrAge = ""
    cmdNewPati.ToolTipText = "新增病人(F4)"
    cmdNewPati.Enabled = InStr(gstrPrivs, ";挂号病人建档;") > 0
    mstr门诊号 = ""
    mstr费别 = ""
    lblInfo.Caption = "性别:     年龄:       门诊号:              费别:  "
    lblMoney.Caption = "门诊预交余额:0.00  "
    lblSum.Caption = "合计"
    mintInsure = 0
    mlng领用ID = 0
    chkBook.Enabled = True
    LoadPayMode False, False
    Set mrsInfo = Nothing
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_现收
    Else
        If (mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2) And gSysPara.bln免挂号模式 = False Then
            mRegistFeeMode = EM_RG_现收
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
            cmdPrice.Visible = mty_Para.byt挂号模式 = 2
        Else
            mRegistFeeMode = EM_RG_划价
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
            cmdPrice.Visible = False
        End If
    End If
End Sub


Private Sub LoadPayMode(Optional ByVal blnPrepay As Boolean = False, Optional ByVal blnInsure As Boolean = False)
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSql As String, str性质 As String
    
    strSql = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 And Instr([2] ,','||B.性质||',')>0" & _
        " Order by B.编码"
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, "挂号", ",3,7,8,")
    
    Set mcolCardPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cboPayMode
        .Clear: j = 0
'        Do While Not rsTemp.EOF
'            blnFind = False
'            For i = 0 To UBound(varData)
'                varTemp = Split(varData(i) & "|||||", "|")
'                If varTemp(6) = Nvl(rsTemp!名称) Then
'                    blnFind = True
'                    Exit For
'                End If
'            Next
'
'            If Not blnFind Then
'                .AddItem Nvl(rsTemp!名称)
'                mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
'                If Val(Nvl(rsTemp!缺省)) = 1 Then
'                    If .ListIndex = -1 Then
'                         .ItemData(.NewIndex) = 1: .ListIndex = .NewIndex
'                    End If
'                End If
'                j = j + 1
'            End If
'            rsTemp.MoveNext
'        Loop
     
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                rsTemp.Filter = "名称='" & varTemp(6) & "'"
                If Not rsTemp.EOF Then
                    mcolCardPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    
    If blnPrepay Then
        cboPayMode.AddItem "预交金"
        If mty_Para.bln优先使用预交 Then
            cboPayMode.ListIndex = cboPayMode.NewIndex
        End If
    End If
    
    If blnInsure Then
        rsTemp.Filter = "性质 = 3"
        If rsTemp.EOF Then
            mstrInsure = ""
            MsgBox "不能加载医保结算方式,请检查!", vbInformation, gstrSysName
        Else
            cboPayMode.AddItem Nvl(rsTemp!名称)
            mstrInsure = Nvl(rsTemp!名称)
            If Not mty_Para.bln优先使用预交 Or blnPrepay = False Then
                cboPayMode.ListIndex = cboPayMode.NewIndex
            End If
            If (mintInsure <> 0 And MCPAR.不收病历费) And cboPayMode.Text = mstrInsure And cboPayMode.Visible Then
                chkBook.Enabled = False
                chkBook.Value = 0
            Else
                chkBook.Enabled = True
            End If
        End If
    End If
    
    If cboPayMode.ListCount > 0 And cboPayMode.ListIndex = -1 Then
        cboPayMode.ListIndex = 0
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function LoadRegPlans(ByVal intSelMode As Integer, Optional ByVal strFilter As String, Optional ByVal blnOtherDoctor As Boolean) As Boolean
'功能:读取挂号安排
'入参:intSelMode:读取模式=1-默认读取;2-过滤读取;3-所有读取
'       blnOtherDoctor:是否读取其他医生号别
    Dim strTime As String, strState As String, strWhere As String
    Dim strSql As String, strIF As String, rsPlan As ADODB.Recordset
    Dim i As Integer, k As Integer
    Dim DateThis As Date, strZero As String
    Dim str挂号安排 As String, strViewSQL As String
    Dim str挂号安排计划 As String, str挂号汇总计划 As String
    Dim str排序         As String, str挂号汇总安排 As String
    Dim vRect          As RECT
    Dim varTemp As Variant, varData As Variant
    On Error GoTo errH
    
    If chkAll.Value = 0 Then
        varTemp = Split(mty_Para.strStationRegOrder, "|")
        For i = 0 To UBound(varTemp)
            varData = Split(varTemp(i), ",")
            Select Case varData(0)
                Case "医生"
                    str排序 = str排序 & ",Decode(医生,Null,Decode(科室ID," & mlngDept & ",3,4),Decode(科室ID," & mlngDept & ",1,2)),医生 " & IIf(varData(1) = 1, "", "desc")
                Case "科室"
                    str排序 = str排序 & ",科室 " & IIf(varData(1) = 1, "", "desc")
                Case "执行时间"
                    str排序 = str排序 & ",开始时间 " & IIf(varData(1) = 1, "", "desc")
                Case "号别"
                    str排序 = str排序 & ",号别 " & IIf(varData(1) = 1, "", "desc")
                Case "项目"
                    str排序 = str排序 & ",项目 " & IIf(varData(1) = 1, "", "desc")
            End Select
        Next
        str排序 = Mid(str排序, 2)
    Else
        str排序 = "Decode(医生,'" & UserInfo.姓名 & "',1,2),Decode(科室ID," & mlngDept & ",1,2),号别,项目,已挂"
    End If
    
    If gstrDeptIDs <> "" And Not blnOtherDoctor Then strIF = " And Instr(','||[4]||',',','||P.科室ID||',')>0"
    If mblnAppointment Then
        If mty_Para.bln预约包含科室安排 Then
            strIF = strIF & IIf(blnOtherDoctor, " And (p.医生姓名 <> [1] or p.医生姓名 Is Null )", " And (p.医生姓名 = [1] or p.医生姓名 Is Null )")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (p.医生姓名 <> [1] )", " And (p.医生姓名 = [1])")
        End If
    Else
        If mty_Para.bln挂号包含科室安排 Then
            strIF = strIF & IIf(blnOtherDoctor, " And (p.医生姓名 <> [1] or p.医生姓名 Is Null)", " And (p.医生姓名 = [1] or p.医生姓名 Is Null)")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (p.医生姓名 <> [1] )", " And (p.医生姓名 = [1])")
        End If
    End If
    
    If intSelMode = 2 Then
        strIF = strIF & " And (p.号码 Like [8] Or Upper(b.名称) Like Upper([8]) Or Upper(zlSpellCode(b.名称)) Like Upper([8]) Or Upper(p.医生姓名) Like Upper([8]) Or Upper(zlSpellCode(p.医生姓名)) Like Upper([8]))"
    End If
     
    str挂号安排 = "" & _
        "            Select A.ID, A.号码, A.号类, A.科室id, A.项目id, A.医生id, A.医生姓名, A.病案必须, A. 周日, A.周一, A.周二, A.周三, " & _
        "                   A.周四 , A.周五, A.周六, A.分诊方式,A.序号控制, B.限号数, B.限约数,a.停用日期 " & IIf(chkAll.Value <> 1, ",c.开始时间 ", "") & vbNewLine & _
        "            From 挂号安排 A, 挂号安排限制 B " & IIf(chkAll.Value <> 1, ", 时间段 C ", "") & vbNewLine & _
        "            Where a.停用日期 Is Null And [5] Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
        "                 Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
        "                  And a.ID = B.安排id(+) " & IIf(mblnAppointment, " And Trunc(Sysdate)+Nvl(A.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]", "") & _
        "                  And Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = B.限制项目(+)" & vbNewLine & _
        IIf(chkAll.Value <> 1, " And Decode(To_Char([5], 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',a.周六, Null)  = c.时间段 And c.站点 Is Null And c.号类 Is Null ", "")
    
    '挂号安排 限号数限约数 挂号安排限制中获取
    str挂号汇总安排 = str挂号安排 & " And Not Exists (Select 1 From 挂号安排计划 Where 安排id = a.Id) "
    '挂号安排计划 限号数限约数 挂号计划限制中获取
    str挂号汇总计划 = " Union All " & _
        "            Select C.ID, A.号码, C.号类, C.科室id, A.项目id, A.医生id, A.医生姓名, C.病案必须, A. 周日, A.周一, A.周二, A.周三, " & _
        "                   A.周四 , A.周五, A.周六, A.分诊方式,A.序号控制, B.限号数, B.限约数,C.停用日期 " & IIf(chkAll.Value <> 1, ",NULL as 开始时间 ", "") & vbNewLine & _
        "            From 挂号安排计划 A, 挂号计划限制 B,挂号安排 C " & vbNewLine & _
        "            Where c.停用日期 Is Null And [5] Between Nvl(a.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
        "                 a.失效时间 And a.审核时间 Is Not Null And " & _
        "           a.生效时间 = (Select Max(生效时间)" & vbNewLine & _
        "                           From 挂号安排计划" & vbNewLine & _
        "                           Where 安排id = a.安排id And [5] Between" & vbNewLine & _
        "                           Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And 失效时间 And" & vbNewLine & _
        "                           审核时间 Is Not Null)" & _
        "                  And a.ID = B.计划id(+) And a.安排id = c.Id " & IIf(mblnAppointment, "   And Trunc(Sysdate)+Nvl(C.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]", "") & _
        "                  And Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = B.限制项目(+)" & vbNewLine & _
        IIf(chkAll.Value <> 1, " And Decode(To_Char([5], 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',a.周六, Null) Is Not Null", "")
    
    
    If mblnAppointment Then
        DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
    Else
        DateThis = gobjDatabase.Currentdate
    End If
    '取对应日期安排的时间段
    strSql = "Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)"
    
    '该部分语句取现在所对应的时间段
    strTime = _
        "Select 时间段 From 时间段 Where 号类 Is Null And 站点 Is Null And " & _
        "    ('3000-01-10 '||To_Char([5],'HH24:MI:SS') >" & _
        "               Decode(Sign(开始时间-终止时间),1,'3000-01-09 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS'),'3000-01-10 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS')))" & _
        " Or" & _
        " ('3000-01-10 '||To_Char([5],'HH24:MI:SS')  >" & _
        "   '3000-01-10 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS')) "
    
    If Not (mblnAppointment And Format(DateThis, "yyyy-mm-dd") > Format(gobjDatabase.Currentdate, "yyyy-mm-dd")) Then
        strWhere = IIf(chkAll.Value = 0, " And " & strSql & " IN(" & strTime & ")", "")
    End If

    '该部分语句当时读取各种安排的挂号情况
    strState = _
    "   Select A.ID as 安排ID,B.已挂数,B.已约数" & _
    "   From (" & str挂号汇总安排 & str挂号汇总计划 & ") A,病人挂号汇总 B" & _
    "   Where A.科室ID = B.科室ID And A.项目ID = B.项目ID" & _
    "               And Nvl(A.医生ID,0)=Nvl(B.医生ID,0) " & _
    "               And Nvl(A.医生姓名,'医生')=Nvl(B.医生姓名,'医生') " & _
    "               And (A.号码=B.号码 or B.号码 is Null )  And B.日期=[6]"
    
    If mblnAppointment Then
        str挂号安排计划 = " " & _
            "             Select A.ID,A.ID as 计划ID, A.安排id, A.号码, A.项目id, A.安排人, A.安排时间, A. 周日, A.周一, A.周二, A.周三, A.周四, A.周五," & _
            "                    A.周六 , A.分诊方式, A.序号控制, B.限号数, B.限约数, A.生效时间, A.失效时间 ,A.医生姓名,A.医生ID" & IIf(chkAll.Value <> 1, ",D.开始时间 ", "") & _
            "             From 挂号安排计划 A, 挂号计划限制 B," & vbNewLine & _
            "                  (" & vbNewLine & _
            "                      Select Max(生效时间) As 生效时间, 安排id" & _
            "                      From 挂号安排计划 " & vbNewLine & _
            "                      Where 审核时间 Is Not Null And  [5] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
            "                          失效时间  " & vbNewLine & _
            "                       Group By 安排id" & vbNewLine & _
            "                   ) C" & IIf(chkAll.Value <> 1, ",时间段 D", "") & _
            "             Where A.审核时间 Is Not Null And ([5] Between  A.生效时间 + 0 And A.失效时间)" & _
            "                   And A.ID = B.计划id(+) And " & vbNewLine & _
            "                   Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6'," & _
            "                  '周五', '7', '周六', Null) = B.限制项目(+) And A.生效时间 = C.生效时间 And A.安排id = C.安排id " & _
             IIf(chkAll.Value <> 1, "And Decode(To_Char([5], 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',a.周六, Null) = d.时间段 And d.站点 Is Null And d.号类 Is Null ", "")
    
        strSql = _
        " Select P.ID,0 as 计划ID,P.号码 ,P.号类,P.科室ID,P.项目ID," & _
        "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
        "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
        "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)  as 排班" & IIf(chkAll.Value <> 1, ",p.开始时间 ", "") & _
        " From (" & str挂号安排 & ") P" & _
        " Where    Not Exists(Select 1 From 挂号安排计划 where 安排ID=P.id And ([5] BETWEEN 生效时间 + 0 and 失效时间)  And 审核时间 is not NULL  ) " & _
        "          And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=P.ID and [5] between 开始停止时间 and 结束停止时间 )" & _
        " Union ALL " & _
        " Select   C.ID,P.计划ID,C.号码,C.号类,C.科室ID,P.项目ID," & _
        "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(C.病案必须,0) as 病案必须," & _
        "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
        "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)  as 排班" & IIf(chkAll.Value <> 1, ",p.开始时间 ", "") & _
        " From (" & str挂号安排计划 & ") P, 挂号安排 C" & _
        " Where P.安排ID=C.ID  And C.停用日期 Is  NULL  And Trunc(Sysdate)+Nvl(C.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]  " & _
        "           And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=C.ID and [5] between 开始停止时间 and 结束停止时间 )"
        strSql = "(" & strSql & ") P"
    Else
        strSql = _
                    " (Select P.ID,0 as 计划ID,P.号码 ,P.号类,P.科室ID,P.项目ID," & _
                    "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
                    "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
                    "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL) as 排班" & IIf(chkAll.Value <> 1, ",p.开始时间 ", "") & _
                    " From (" & str挂号安排 & ") P "
        strSql = strSql & vbNewLine & "  ) P"
    End If
    
    strViewSQL = _
                "Select Distinct " & _
                "       P.ID,P.号码 as 号别,P.号类,P.科室ID,B.名称 As 科室,C.名称 As 项目," & _
                "       P.医生姓名 as 医生,Nvl(A.已挂数,0) as 已挂,Nvl(A.已约数,0) as 已约," & _
                "       P.限号数 as 限号,P.限约数 as 限约,Decode(Nvl(P.病案必须,0),1,'√','') as 病案,Decode(Nvl(C.项目特性,0),1,'√','') as 急诊," & _
                "       Decode(P.分诊方式,1,'指定',2,'动态',3,'平均',NULL) as 分诊,Decode(Nvl(P.序号控制,0),1,'√','') As 序号控制,P.排班" & IIf(chkAll.Value <> 1, ",p.开始时间", "") & _
                " From " & strSql & "," & vbCrLf & _
                "           (" & strState & ") A,部门表 B,收费项目目录 C" & _
                " Where P.ID=A.安排ID(+) And Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.科室ID=B.ID And P.项目ID=C.ID" & strIF & strZero & _
                "           And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & strWhere & _
                "           And (Nvl(P.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where P.医生ID=Q.ID And (Q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.撤档时间 Is Null)))" & _
                " Order by " & str排序
    If chkAll.Value <> 1 Then
        strViewSQL = _
                    "Select  ID,号别,号类,科室ID,科室,项目,医生,已挂, 已约,限号, 限约,病案, 急诊," & _
                    "           分诊, 序号控制,排班 " & _
                    "From (" & strViewSQL & ")"
    End If
    
    strSql = _
                "Select Distinct " & _
                "       P.ID,p.计划ID,P.号码 as 号别,P.号类,P.科室ID,B.名称 As 科室,P.项目ID,C.名称 As 项目," & _
                "       P.医生ID,P.医生姓名 as 医生,Nvl(A.已挂数,0) as 已挂,Nvl(A.已约数,0) as 已约," & _
                "       P.限号数 as 限号,P.限约数 as 限约,Nvl(P.病案必须,0) as 病案,Nvl(C.项目特性,0) as 急诊," & _
                "       P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & _
                "       Decode(P.分诊方式,1,'指定',2,'动态',3,'平均',NULL) as 分诊,P.序号控制,P.排班" & IIf(chkAll.Value <> 1, ",p.开始时间", "") & _
                " From " & strSql & "," & vbCrLf & _
                "           (" & strState & ") A,部门表 B,收费项目目录 C" & _
                " Where P.ID=A.安排ID(+) And Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.科室ID=B.ID And P.项目ID=C.ID" & strIF & strZero & _
                "           And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & strWhere & _
                "           And (Nvl(P.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where P.医生ID=Q.ID And (Q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.撤档时间 Is Null)))" & _
                " Order by " & str排序
                
    Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, _
            UserInfo.姓名, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, "%" & strFilter & "%")
 
    
    If mrsPlan.RecordCount <> 0 Then
        If intSelMode = 1 Or mrsPlan.RecordCount = 1 Then
            '默认读取
            Call ReadLimit(Nvl(mrsPlan!号别))
        Else
            vRect = GetControlRect(txtReg.hWnd)
            Set rsPlan = gobjDatabase.ShowSQLSelect(Me, strViewSQL, 0, "号码选择", False, "", "号码选择", _
                                                False, False, True, vRect.Left, vRect.Top - 250, 600, False, True, False, _
                                                UserInfo.姓名, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, "%" & strFilter & "%")
            If rsPlan Is Nothing Then
                Call ReadLimit(Nvl(mrsPlan!号别))
            Else
                If Not rsPlan.EOF Then
                    Call ReadLimit(Nvl(rsPlan!号别))
                Else
                    Call ReadLimit(Nvl(mrsPlan!号别))
                End If
            End If
        End If
        Call LoadDoctor
        Call ResetDefault复诊
        Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, mstrPriceGrade)
        Call GetActiveView
        If mblnAppointment Then
            Select Case mViewMode
                Case V_普通号分时段, v_专家号分时段
                    cmdTime.Visible = True
                Case Else
                    cmdTime.Visible = False
            End Select
            Call SetDefultRegTime
        Else
            cmdTime.Visible = False
        End If
        lblDeptName.Caption = Nvl(mrsPlan!科室)
        If txtReg.Visible And txtReg.Enabled Then txtReg.SetFocus
    End If
    LoadRegPlans = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub txtReg_Change()
    If mblnChangeByCode = True Then Exit Sub
    mblnIntact = False
End Sub

Private Sub txtReg_GotFocus()
    Call gobjControl.TxtSelAll(txtReg)
End Sub

Private Sub txtReg_KeyPress(KeyAscii As Integer)
    If mblnIntact Then
        If KeyAscii = 13 Then gobjCommFun.PressKeyEx vbKeyTab
    Else
        If KeyAscii = 13 Then Call LoadRegPlans(2, txtReg.Text)
    End If
End Sub

Private Sub ReadLimit(ByVal strRegNo As String)
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    
    mrsPlan.Filter = "号别='" & strRegNo & "'"
    
    If mrsPlan.RecordCount = 0 Then Exit Sub
    mblnIntact = True
    mblnChangeByCode = True
    If Nvl(mrsPlan!医生) = "" Then
        txtReg.Text = "[" & Nvl(mrsPlan!号别) & "]" & Nvl(mrsPlan!项目)
    Else
        txtReg.Text = "[" & Nvl(mrsPlan!号别) & "]" & Nvl(mrsPlan!项目) & "(" & Nvl(mrsPlan!医生) & ")"
    End If
    mblnChangeByCode = False
    txtReg.Tag = Nvl(mrsPlan!号别)
    
    If mblnAppointment Then
        If Nvl(mrsPlan!限约) = "" Then
            lblLimit.Caption = "已约:" & Nvl(mrsPlan!已约, 0)
        Else
            lblLimit.Caption = "限约:" & Nvl(mrsPlan!限约) & "  已约:" & Nvl(mrsPlan!已约, 0)
        End If
    Else
        If Nvl(mrsPlan!限号) = "" Then
            lblLimit.Caption = "已挂:" & Nvl(mrsPlan!已挂, 0)
        Else
            lblLimit.Caption = "限号:" & Nvl(mrsPlan!限号) & "  已挂:" & Nvl(mrsPlan!已挂, 0)
        End If
    End If
    If Val(Nvl(mrsPlan!急诊)) = 0 Then
        lbl急.Visible = False
    Else
        lbl急.Visible = True
    End If
    
    lbl时段.Caption = Nvl(mrsPlan!排班)
    lbl时段.Visible = Nvl(mrsPlan!排班) <> ""
    If Not mrsInfo Is Nothing Then Call Load费别(Nvl(mrsInfo!费别))
End Sub

Private Function GetActiveView()
    '得到当前挂号业务  采取那种类型的流程
    Dim strSql          As String
    Dim rsTmp           As ADODB.Recordset
    Dim str号码         As String
    Dim dat            As Date
    
    On Error GoTo errH
    str号码 = txtReg.Tag
    If mblnAppointment Then
        dat = dtpDate.Value
    Else
        dat = gobjDatabase.Currentdate
    End If
    
    strSql = _
    "       Select   Havedata, 安排id" & vbNewLine & _
    "       From (" & vbNewLine & _
    "               Select 1 As Havedata, b.Id As 安排id " & vbNewLine & _
    "               From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
    "               Where B.号码=[1] And A.安排id = b.ID " & _
    "                And   Decode(To_Char([2], 'D'), '1', '周日', '2'," & _
    "                   '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =a.星期 " & vbNewLine & _
    "                       And Not Exists" & vbNewLine & _
    "                     (Select 1 From 挂号安排计划 C " & vbNewLine & _
    "                         Where c.安排id = b.Id And c.审核时间 Is Not Null And [2] Between " & _
    "                               Nvl(c.生效时间, [2]) And" & _
    "                          c.失效时间)" & vbNewLine & _
    "               Union All " & vbNewLine & _
    "               Select 1 As Havedata, c.Id As 安排id" & vbNewLine & _
    "               From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C,(" & vbNewLine & _
    "                   SELECT MAX(a.生效时间 ) 生效 FROM 挂号安排计划 a,挂号安排 B  WHERE a.安排Id=b.ID AND b.号码=[1] AND a.审核时间 IS NOT NULL" & vbNewLine & _
    "             And [2] Between nvl(a.生效时间,to_date('1900-01-01','yyyy-mm-dd')) And a.失效时间" & vbNewLine & _
    "           ) D  " & vbNewLine & _
    "               Where  C.号码=[1] And c.Id = b.安排id And b.Id = a.计划id And b.生效时间=d.生效 And b.审核时间 Is Not Null" & _
    "                    And   Decode(To_Char([2], 'D'), '1', '周日', '2'," & _
    "                   '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =a.星期 " & vbNewLine & _
    "                       And [2] Between Nvl(b.生效时间,[2]) And b.失效时间) B"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str号码, dat)
    If rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!序号控制)) = 1 Then
       '*********************
       '专家号分时段
       '*********************
       mViewMode = v_专家号分时段

    ElseIf rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!序号控制)) = 0 Then
       '*********************
       '普通号分时段
       '*********************
       mViewMode = V_普通号分时段

    ElseIf Val(Nvl(mrsPlan!序号控制)) = 1 And Nvl(mrsPlan!限号) <> "" Then
       '*********************
       '专家号不分时段
       '*********************
       mViewMode = v_专家号

     Else
       '*********************
       '普通号
       '*********************
       mViewMode = V_普通号

    End If
    Set rsTmp = Nothing
Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
         Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function SelectTimeSn() As Boolean
    '**************************************
    '加载时段
    '返回时段是否加载成功或是否有分时段
    '**************************************
     Dim strSql         As String
     Dim dateCur        As Date
     Dim strNO          As String
     Dim vRect          As RECT
    If Not mblnAppointment Then Exit Function
    
    strSql = "" & _
    " Select Distinct a.序号 As ID, A.序号,To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
    " From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
    " Where a.安排id = b.Id And b.号码 = [1] And" & vbNewLine & _
    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
    "      Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',Null) = a.星期(+)  " & _
    "      And Not Exists (Select Count(1) From 挂号序号状态 Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having Count(1) - a.限制数量 >= 0) " & _
    "      And Not Exists (Select 1 From 挂号安排计划 E Where e.安排id = b.Id And e.审核时间 Is Not Null And [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And e.失效时间)"
    
    strSql = strSql & " Union " & _
    "Select Distinct a.序号 As ID,A.序号,To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
    "From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C," & vbNewLine & _
    "     (Select Max(a.生效时间) 生效" & vbNewLine & _
    "       From 挂号安排计划 A, 挂号安排 B" & vbNewLine & _
    "       Where a.安排id = b.Id And b.号码 = [1] And a.审核时间 Is Not Null And" & vbNewLine & _
    "             [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
    "             a.失效时间) D" & vbNewLine & _
    "Where a.计划id = b.Id And b.安排id = c.Id And c.号码 = [1] And b.生效时间 = d.生效 And b.审核时间 Is Not Null And" & vbNewLine & _
    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
    "      [2] Between Nvl(b.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
    "      b.失效时间 And Not Exists" & vbNewLine & _
    " (Select Count(1)" & vbNewLine & _
    "       From 挂号序号状态" & vbNewLine & _
    "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
    "        Count(1) - a.限制数量 >= 0) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
    "                                           '周四', '6', '周五', '7', '周六', Null) = a.星期(+)" & vbNewLine & _
            "Order By 开始时间"


    dateCur = Format(dtpDate, "yyyy-mm-dd")
    If strSql = "" Then Exit Function
    strNO = txtReg.Tag
    vRect = GetControlRect(dtpTime.hWnd)
    
    On Error GoTo errH
    
    Set mrs时间段 = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "预约时间选择", False, "", "预约时间选择", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, strNO, dateCur)
    
    If mrs时间段 Is Nothing Then Exit Function
    If mrs时间段.EOF Then Exit Function
    
    lblSn.Caption = ""
    dtpTime.Value = Format(mrs时间段!开始时间, "hh:mm:ss")
    lblSn.Caption = "序号:" & Val(Nvl(mrs时间段!序号))
    SelectTimeSn = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadFeeItem(ByVal lngItemID As Long, ByVal blnBook As Boolean, ByVal strPriceGrade As String)
    Dim strSql As String, i As Integer, dblTotal As Double
    Dim rsIncomes As ADODB.Recordset, cur应收 As Currency, cur实收 As Currency
    Dim j As Integer, rsItems As ADODB.Recordset, lng病人ID As Long
    If lngItemID = 0 Then Exit Sub
    '性质:1-主挂号费用 2-从项费用 3-病历费
    ReadRegistPrice lngItemID, blnBook, False, mstr费别, rsItems, rsIncomes, , , , 1, _
        Val(Nvl(mrsPlan!科室ID)), strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    If mintInsure <> 0 Then
        If MCPAR.挂号检查项目 = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "医保病人收费项目检查失败，不能继续 " & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    ReadRegistPrice lngItemID, blnBook, False, mstr费别, rsItems, rsIncomes, lng病人ID, mintInsure, _
        txtReg.Tag, IIf(mblnAppointment, 1, 0), , strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = Format(0, "0.00")
    lblPayMoney.Caption = Format(0, "0.00")
    dblTotal = 0
    If rsItems.RecordCount = 0 Then Exit Sub
    rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        With vsfMoney
            .RowData(.Rows - 1) = Nvl(rsItems!项目ID)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = Nvl(rsItems!项目名称)
            rsIncomes.Filter = "项目ID=" & rsItems!项目ID
            cur应收 = 0: cur实收 = 0
            For j = 1 To rsIncomes.RecordCount
                cur应收 = cur应收 + rsIncomes!应收
                cur实收 = cur实收 + rsIncomes!实收
                rsIncomes.MoveNext
            Next j
            .TextMatrix(.Rows - 1, .ColIndex("应收金额")) = Format(cur应收, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("实收金额")) = Format(cur实收, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("性质")) = Nvl(rsItems!性质)
            .Rows = .Rows + 1
        End With
        rsItems.MoveNext
    Next i
    If vsfMoney.Rows > 2 Then vsfMoney.Rows = vsfMoney.Rows - 1
    For i = 1 To vsfMoney.Rows - 1
        dblTotal = dblTotal + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("实收金额")))
    Next i
    vsfMoney.RowHeightMin = 350
    lblTotal.Caption = Format(dblTotal, "0.00")
    lblPayMoney.Caption = Format(dblTotal, "0.00")
    lblRoomName.Caption = gstrRooms
End Sub


Private Function GetSNState(str号别 As String, datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    Dim strSql           As String
    Dim datStart         As Date
    Dim datEnd           As Date
    On Error GoTo errH
    datStart = CDate(Format(datThis, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    strSql = "    " & vbNewLine & " Select 序号,状态,操作员姓名,Nvl(预约,0) as 预约,TO_Char(日期,'hh24:mi:ss') as 日期  "
    strSql = strSql & vbNewLine & " From 挂号序号状态 "
    strSql = strSql & vbNewLine & " Where 号码=[1]"
    strSql = strSql & vbNewLine & IIf(datThis = CDate(0), " And 日期 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 ", " And 日期 Between  [2] And [3]")
    strSql = strSql & vbNewLine & IIf(lngSN > 0, " And 序号=[4]", "")
    Set GetSNState = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str号别, datStart, datEnd, lngSN)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function zlGet当前星期几(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当日是星期几
    '编制:刘兴洪
    '日期:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, bln当前日期 As Boolean, strTemp As String
    If strDate = "" Then
        strSql = "Select Decode(To_Char(Sysdate,'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六',NULL) as 星期  From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        strSql = "Select Decode(To_Char([1],'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六','') As 星期 From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!星期)
    zlGet当前星期几 = strTemp
End Function





Private Function GetTotalFromMshMoney(Optional ByVal str项目名称 As String = "") As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取汇总金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 16:57:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    
    On Error GoTo errHandle
    With vsfMoney
        For i = 1 To .Rows - 1
            If str项目名称 = "" Or Trim(.TextMatrix(i, 0)) = str项目名称 Then
                dblMoney = dblMoney + Val(.TextMatrix(i, 2))
            End If
        Next
    End With
    GetTotalFromMshMoney = dblMoney
    Exit Function
errHandle:
    GetTotalFromMshMoney = 0
End Function



Private Function GetRegistMoney(Optional blnOnlyReg As Boolean = False, Optional blnNoBook As Boolean = False) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前挂号单的合计金额
    '入参:blnOnlyReg-是否仅仅读取挂号费用
    '     blnNoBook-读取病历费
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 16:53:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl合计 As Double, i As Integer
    Dim k As Integer
    
    If Not blnOnlyReg Then
        dbl合计 = FormatEx(GetTotalFromMshMoney, 5)
    Else
        If mrsItems Is Nothing Then
             GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        mrsItems.Filter = " 性质 <> 4"
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        With mrsItems
            Do While Not .EOF
                dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
                .MoveNext
            Loop
        End With
        mrsItems.Filter = 0
    End If
    If blnNoBook Then
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = " 性质 = 3"
            Do While Not mrsItems.EOF
                dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
                mrsItems.MoveNext
            Loop
            mrsItems.Filter = 0
        End If
    End If
    GetRegistMoney = FormatEx(dbl合计, 5)
End Function
 
Private Sub Init费别()
    '初始化缺省费别
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    
    strSql = "Select 名称 From 费别 Where 缺省标志 = 1 And Nvl(服务对象, 3) In (1, 3)"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        mstrDef费别 = Nvl(rsTmp!名称)
    Else
        MsgBox "无法读取缺省费别，请检查缺省费别是否正确设置！", vbInformation, gstrSysName
    End If
    
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function Load费别(Optional ByVal str费别 As String) As Boolean
    '功能:根据科室加载病人费别
    '参数:str费别-病人上次使用的的费别
    '返回:成功,返回true,否则返回False
    
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    If mrsInfo Is Nothing Then Exit Function
    If mrsPlan Is Nothing Then Exit Function
    If mrsPlan.EOF Then Exit Function
    If str费别 <> "" Then
        strSql = " Select 1 From 费别 A, 费别适用科室 B" & _
                 " Where a.名称 = b.费别(+) And a.属性 = 1" & _
                 "      And Trunc(Sysdate) Between Nvl(a.有效开始, To_Date('1900-01-01', 'YYYY-MM-DD'))" & _
                 "      And Nvl(a.有效结束, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                 "      And Nvl(a.服务对象, 3) In (1, 3) And (B.科室ID=[1] or B.科室ID is NULL) and A.名称=[2]"
        
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!科室ID)), str费别)
        If Not rsTmp.EOF Then
            mstr费别 = str费别
        Else
            mstr费别 = mstrDef费别
        End If
    Else
        mstr费别 = mstrDef费别
    End If
    If mstr费别 = "" Then
        MsgBox "未找的适用于【" & Nvl(mrsPlan!科室) & "】的缺省费别,请在『病人详细信息』界面中设置病人费别！", vbInformation, gstrSysName
        Load费别 = False
        Exit Function
    End If
    lblInfo.Caption = "性别:" & mstr性别 & "   年龄:" & mstrAge & "   门诊号:" & mstr门诊号 & "   费别:" & mstr费别
    Load费别 = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
