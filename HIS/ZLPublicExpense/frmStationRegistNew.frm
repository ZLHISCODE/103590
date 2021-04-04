VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmStationRegistNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医生站挂号"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   Icon            =   "frmStationRegistNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8460
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
      Left            =   3720
      TabIndex        =   49
      Top             =   5625
      Width           =   1725
   End
   Begin VB.CheckBox chkAll 
      Height          =   360
      Left            =   8070
      Picture         =   "frmStationRegistNew.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "显示不当班别"
      Top             =   45
      Width           =   345
   End
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   2940
      Picture         =   "frmStationRegistNew.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "新增病人(F4)"
      Top             =   600
      Width           =   350
   End
   Begin VB.PictureBox picPayMoney 
      BackColor       =   &H80000005&
      Height          =   435
      Left            =   6300
      ScaleHeight     =   375
      ScaleWidth      =   1995
      TabIndex        =   37
      Top             =   4942
      Width           =   2055
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
         Left            =   1200
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
         Picture         =   "frmStationRegistNew.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "其他医生号别"
         Top             =   45
         Width           =   345
      End
      Begin VB.CheckBox chk复诊 
         Caption         =   "复诊"
         Height          =   255
         Left            =   7590
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   930
      End
      Begin VB.CommandButton cmdReg 
         Height          =   345
         Left            =   4005
         Picture         =   "frmStationRegistNew.frx":1788
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "当前医生号别"
         Top             =   45
         Width           =   345
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
         Left            =   7020
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
         Width           =   2055
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
         Width           =   6180
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
         FormatString    =   $"frmStationRegistNew.frx":218A
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
         Caption         =   "序号:22 "
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
         Left            =   7425
         TabIndex        =   50
         Top             =   90
         Width           =   960
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
         Top             =   120
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
      Top             =   5430
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
      Left            =   4335
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
      Left            =   7080
      TabIndex        =   12
      Top             =   5625
      Width           =   1300
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
      Left            =   5730
      TabIndex        =   11
      Top             =   5625
      Width           =   1300
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助"
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
      Left            =   30
      TabIndex        =   13
      Top             =   5625
      Width           =   1300
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
      Left            =   4125
      TabIndex        =   28
      Top             =   1568
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   360
      Left            =   3030
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
      Left            =   5745
      ScaleHeight     =   300
      ScaleWidth      =   2580
      TabIndex        =   44
      Top             =   1560
      Width           =   2640
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
      Left            =   4440
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
      Left            =   2970
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
      Left            =   5790
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
      Left            =   2550
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
Attribute VB_Name = "frmStationRegistNew"
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
Private mstrAge As String, mstrFeeType As String, mstrGender As String, mstrClinic As String
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
Private mcolArrangeNo As Collection, mblnUpdateAge As Boolean
Private mlng病人ID As Long, mintIDKind As Integer
Private mcur个帐余额 As Currency, mblnIntact As Boolean
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mlngSN As Long
Private mintInsure As Integer, mstrUseType As String
Private mdatLast As Date
Private mblnChangeByCode As Boolean
Private mstrCardNO As String
Private mcur个帐透支 As Currency
Private mlng锁号记录ID As Long '挂号锁号时的记录ID
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
    int可预约天数 As Integer
    int同科限约数           As Integer  '同科室限约
    int同科限挂数           As Integer
    bln同科限挂急诊         As Boolean
    int病人预约科室数       As Integer
    int病人挂号科室数       As Integer
    int专家号挂号限制       As Integer
    int专家号预约限制       As Integer
    strStationRegOrder As String  '医生站挂号排序字符串
    blnShowAllPlan      As Boolean   ' 是否显示不当班号别
End Type

Private mty_Para As ty_ModulePara
Private mstrPriceGrade As String, mintPriceGradeStartType As Integer
Private mobjRegister As clsRegist
Private mstrDef费别 As String   '缺省费别

Public Sub zlShowMe(ByVal frmMain As Object, ByVal objRegister As clsRegist, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByVal blnAppointment As Boolean, ByVal lng病人ID As Long, ByRef strOutNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医生站挂号入口
    '入参:strDeptIDs-挂号科室,支持多个,用逗号分隔
    '     blnAppointment-是否预约调用
    '出参:strOutNO-挂号成功后,传出挂号单据号
    '编制:刘尔旋
    '日期:2016-7-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    mblnAppointment = blnAppointment
    mlngModul = lngModul
    mlng病人ID = lng病人ID
    Set mobjRegister = objRegister
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
        .int可预约天数 = Val(gobjDatabase.GetPara(66, glngSys, , 15))
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
        Call GetPriceGrade(gstrNodeNo, 0, 0, "", "", "", mstrPriceGrade)
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
                    MsgBox "你没有自用和共用的" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
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

Private Sub cmdPrice_Click()
    Dim bytRegistFeeMode As EM_REGISTFEE_MODE
    bytRegistFeeMode = mRegistFeeMode
    
    mRegistFeeMode = EM_RG_划价
    If SaveData = False Then mRegistFeeMode = bytRegistFeeMode: Exit Sub
    
    mRegistFeeMode = bytRegistFeeMode
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Call LoadRegPlans(3)
End Sub
Private Function GetWorkTimeDefualtTime(ByVal strWorkName As String, ByVal str号类 As String, ByVal strRegDate As String, Optional ByVal strCurSysDate As String = "") As Date
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
    Dim dtRegDate As Date, strFilter As String
    
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
    strFilter = "时间段='" & strWorkName & "'"
    strFilter = strFilter & IIf(str号类 = "", " And 号类=NULL", " And =号类='" & str号类 & "'")
    strFilter = strFilter & " and 站点='" & gstrNodeNo & "'"
    
    rsTime.Filter = strFilter
    If rsTime.EOF Then
        strFilter = "时间段='" & strWorkName & "'"
        strFilter = strFilter & " and 站点='" & gstrNodeNo & "'"
        rsTime.Filter = strFilter
        If rsTime.EOF Then
            strFilter = "时间段='" & strWorkName & "' And 站点=NULL "
            strFilter = strFilter & IIf(str号类 = "", " And 号类=NULL", " And =号类='" & str号类 & "'")
            rsTime.Filter = strFilter
            If rsTime.EOF Then
                rsTime.Filter = "时间段='" & strWorkName & "' and 号类=NULL and 站点=NULL"
                If rsTime.EOF Then
                    rsTime.Filter = 0
                    GetWorkTimeDefualtTime = dtSysDate: Exit Function
                End If
            End If
        End If
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



Private Sub SetDefultRegTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的挂号时间
    '日期:2018-02-05 10:00:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, dtSysDate As Date
    Dim rsTmp As ADODB.Recordset, rsTime As ADODB.Recordset
    Dim lng记录ID As Long, lng序号 As Long, str发生时间 As String
    
    On Error GoTo errH
    
    
    dtSysDate = gobjDatabase.Currentdate
    
    lblSn.Caption = ""
        
    lng记录ID = Val(Nvl(mrsPlan!记录ID))
    If lng记录ID = 0 Then
      dtpTime.Value = Format(GetWorkTimeDefualtTime("白天", "", Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
      Exit Sub
    End If
    
    strSql = "Select 开始时间,终止时间,缺省预约时间 As 缺省时间 From 临床出诊记录 Where ID=[1]"
    Set rsTime = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng记录ID)

    If mViewMode = v_专家号分时段 Or (mViewMode = V_普通号分时段 And mblnAppointment) Then
        
        strSql = "Select 序号,开始时间 From 临床出诊序号控制 Where 记录ID=[1] And Nvl(挂号状态,0) = 0 Order By 序号"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng记录ID)
        If Not rsTmp.EOF Then
            lblSn.Caption = "序号:" & Val(Nvl(rsTmp!序号))
            '时段当班有时段,取最小时段
            dtpTime.Value = Format(Nvl(rsTmp!开始时间), "hh:mm:ss"): Exit Sub
        End If
                
        
        If Format(dtpDate.Value, "yyyy-mm-dd") = Format(dtSysDate, "yyyy-mm-dd") Then
            dtpTime.Value = Format(dtSysDate, "hh:mm:ss"): Exit Sub
        End If

        '时段当班无时段,取开始时间
        If rsTime.EOF Then
            dtpTime.Value = Format(dtSysDate, "hh:mm:ss")
        Else
            If IsNull(rsTime!缺省时间) Then
                dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
            Else
                dtpTime.Value = Format(Nvl(rsTime!缺省时间), "hh:mm:ss")
            End If
        End If
        Exit Sub
    End If
    
    If mViewMode = v_专家号 Then
    
        If mobjRegister.zlGetRegisterNextSn__Visits(lng记录ID, Format(dtSysDate, "yyyy-mm-dd HH:MM:SS"), InStr(gstrPrivs, ";加号;"), mblnAppointment, False, lng序号, str发生时间) Then
            If lng序号 <> 0 Then lblSn.Caption = "序号:" & lng序号
            If mblnAppointment Then
                If Format(dtpDate.Value, "yyyy-mm-dd") > Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                    str发生时间 = GetWorkTimeDefualtTime("白天", "", Format(dtpDate.Value, "yyyy-mm-dd"))
                End If
            End If
            If IsDate(str发生时间) Then dtpTime.Value = Format(CDate(str发生时间), "hh:mm:ss"): Exit Sub
        End If
    End If
    
    If Not mblnAppointment Then
        dtpTime.Value = Format(dtSysDate, "hh:mm:ss"): Exit Sub
    End If
         
    If rsTime.EOF Then
        dtpTime.Value = Format(dtSysDate, "hh:mm:ss"): Exit Sub
    End If
    
    If IsNull(rsTime!缺省时间) Then
        dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
    Else
        If CDate(Format(dtpDate.Value, "yyyy-mm-dd") & " " & Format(Nvl(rsTime!缺省时间), "hh:mm:ss")) > CDate(Format(rsTime!终止时间, "yyyy-mm-dd hh:mm:ss")) Then
            dtpTime.Value = Format(Nvl(rsTime!终止时间), "hh:mm:ss")
        ElseIf CDate(Format(dtpDate.Value, "yyyy-mm-dd") & " " & Format(Nvl(rsTime!缺省时间), "hh:mm:ss")) < CDate(Format(rsTime!开始时间, "yyyy-mm-dd hh:mm:ss")) Then
            dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
        Else
            dtpTime.Value = Format(Nvl(rsTime!缺省时间), "hh:mm:ss")
        End If
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

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
        If Nvl(mrsPlan!医生姓名) = "" Then
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
            mrsDoctor.Filter = "姓名='" & Nvl(mrsPlan!医生姓名) & "'"
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
            MsgBox "没有设置常用" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "摘要,请在字典管理中设置", vbOKOnly + vbInformation, gstrSysName
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

Private Sub chkAll_Click()
    mty_Para.blnShowAllPlan = chkAll.Value <> 0
End Sub

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

Private Sub cmdOther_Click()
    Call LoadRegPlans(3, , True)
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
     mstrDef费别 = ""
     mstrFeeType = ""
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
    txtPatient.Text, NeedName(mstrGender), str年龄, dblMoney, mstrCardNO, mstrPassWord, _
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

Private Sub cmdOK_Click()
    If SaveData = False Then
        If (mViewMode = v_专家号 Or mViewMode = v_专家号分时段) And mlng锁号记录ID <> 0 Then Call CancelRegNo(mlng锁号记录ID)
        Exit Sub
    End If
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存挂号数据
    '返回:保存成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-01 15:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSlipPrint As Boolean, blnInvoicePrint As Boolean, int价格父号 As Integer, blnBalance As Boolean
    Dim k As Integer, i As Integer, j As Integer, strNO As String, strFactNO As String
    Dim cllPro As New Collection, strSql As String, str登记时间 As String, str发生时间 As String
    Dim cur预交 As Currency, cur个帐 As Currency, cur现金 As Currency, str划价NO As String
    Dim lngSN As Long, lng挂号科室ID As Long, lng结帐ID As Long, byt复诊 As Byte
    Dim lng医疗卡类别ID As Long, bln消费卡 As Boolean, blnNoDoc As Boolean, strBalanceStyle As String
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, cllProAfter As New Collection
    Dim blnTrans As Boolean, blnNotCommit As Boolean, strAdvance As String, rsTemp As ADODB.Recordset
    Dim lng医生ID As Long, blnOneCard As Boolean, rsTmp As ADODB.Recordset, dat发生时间 As Date
    Dim cllCardPro As Collection, cllTheeSwap As Collection, strNotValiedNos As String
    Dim strDay As String, blnAppointPrint As Boolean, str付款方式 As String, mstr家属IDs As String
    Dim rs付款方式 As ADODB.Recordset, str医生 As String, blnAdd As Boolean, blnNotWork As Boolean
    Dim dat登记时间 As Date, rs时间段 As ADODB.Recordset, str时间段 As String
    Dim bytMode As Byte, rsCheck As ADODB.Recordset, dat预约时间 As Date
    Dim strResult As String, bln专家号 As Boolean
    Dim dblTotal  As Double
    
    If CheckValied = False Then Exit Function
    If Not mrsInfo Is Nothing Then
        strSql = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],[4],[5],[6]) As 检查结果 From Dual"
        If mblnAppointment Then
            bytMode = 1
            dat预约时间 = CDate(Format(dtpDate.Value, "yyyy-mm-dd"))
        Else
            bytMode = 0
            dat预约时间 = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
        End If
        
        bln专家号 = Nvl(mrsPlan!医生姓名) <> ""
        Set rsCheck = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, bytMode, Val(Nvl(mrsInfo!病人ID)), Trim(txtReg.Tag), Val(Nvl(mrsPlan!记录ID)), dat预约时间, IIf(bln专家号, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!检查结果)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "有效性检查失败,无法继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSql = "Select 编号,名称,医院编码,结算方式 From 一卡通目录 Where 启用 = 1 And 结算方式 = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, cboPayMode.Text)
    blnOneCard = rsTmp.RecordCount <> 0
    
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        blnSlipPrint = False
    Else
        Select Case Val(mty_Para.int挂号凭条打印)
            Case 0    '不打印
                blnSlipPrint = False
            Case 1    '自动打印
                If InStr(gstrPrivs, ";病人挂号凭条;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "你没有" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2    '选择打印
                If MsgBox("要打印" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "凭条吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(gstrPrivs, ";病人挂号凭条;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "你没有" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If mRegistFeeMode = EM_RG_划价 Or mRegistFeeMode = EM_RG_记帐 Or (mblnAppointment And mty_Para.bln预约时收款 = False) Then
        blnInvoicePrint = False
    Else
        If Not (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
            Select Case Val(mty_Para.int挂号发票打印)
                Case 0    '不打印
                    blnInvoicePrint = False
                Case 1    '自动打印
                    If InStr(gstrPrivs, ";挂号发票打印;") > 0 Then
                        blnInvoicePrint = True
                    Else
                        blnInvoicePrint = False
                        MsgBox "你没有" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "发票打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Case 2    '选择打印
                    If MsgBox("要打印" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "发票吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If InStr(gstrPrivs, ";挂号发票打印;") > 0 Then
                            blnInvoicePrint = True
                        Else
                            blnInvoicePrint = False
                            MsgBox "你没有" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "发票打印的权限，请联系管理员！", vbInformation, gstrSysName
                        End If
                    Else
                        blnInvoicePrint = False
                    End If
            End Select
        End If
    End If
    
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        Select Case Val(mty_Para.int预约挂号打印)
            Case 0
                blnAppointPrint = False
            Case 1
                If InStr(gstrPrivs, ";预约挂号单;") > 0 Then
                    blnAppointPrint = True
                Else
                    blnAppointPrint = False
                    MsgBox "你没有预约" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "单打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2
                If InStr(gstrPrivs, ";预约挂号单;") > 0 Then
                    If MsgBox("要打印预约" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnAppointPrint = True
                    Else
                        blnAppointPrint = False
                    End If
                Else
                    MsgBox "你没有预约" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "单打印的权限，请联系管理员！", vbInformation, gstrSysName
                    blnAppointPrint = False
                End If
        End Select
    Else
        blnAppointPrint = False
    End If
    
    If blnInvoicePrint Or (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
        If RefreshFact(strFactNO) = False Then Exit Function
    End If
    
    If mblnAppointment Then
        If mRegistFeeMode = EM_RG_记帐 And mty_Para.bln预约时收款 Then
            MsgBox "不支持先诊疗后结算病人的预约收款挂号！", vbInformation, gstrSysName
            Exit Function
        End If
        If mty_Para.bln预约时收款 Then
            If Not mRegistFeeMode = EM_RG_划价 Then
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
        Else
            blnBalance = False
        End If
    Else
        If Not mRegistFeeMode = EM_RG_划价 Then
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
        End If
    End If
    
    If Val(cur预交) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!病人ID), Val(cur预交), mlngModul, 1, , _
                             IIf(-1 * mty_Para.dbl预存款消费验卡 >= Val(cur预交), False, True), True, mstr家属IDs, (mty_Para.dbl预存款消费验卡 <> 0), (mty_Para.dbl预存款消费验卡 = 2)) Then Exit Function
    End If
    
    ReadRegistPrice Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, False, mstrFeeType, rsItems, rsIncomes, _
        Val(Nvl(mrsInfo!病人ID)), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.bln预约时收款) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!结算模式))) = False Then Exit Function
    End If

    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lng医疗卡类别ID = mcolCardPayMode.Item(i)(3)
                bln消费卡 = Val(mcolCardPayMode.Item(i)(5)) = 1
                strBalanceStyle = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur现金), lng医疗卡类别ID, bln消费卡, rsItems, rsIncomes) = False Then Exit Function
        If strBalanceStyle <> "" Then
            strBalanceStyle = strBalanceStyle & "," & Val(cur现金) & ",,1"
        Else
            strBalanceStyle = cboPayMode.Text & "," & Val(cur现金) & ",,0"
        End If
    End If
    
    str登记时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    dat登记时间 = gobjDatabase.Currentdate
    
    If mblnAppointment Then
        strDay = zlGet当前星期几(dtpDate.Value)
    Else
        strDay = zlGet当前星期几
    End If
    
    '获取发生时间
    blnAdd = False
    If mblnAppointment Then
        mlngSN = 0
        If IsNull(mrsPlan!开始时间) = False Then
            str发生时间 = "To_Date('" & Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
            dat发生时间 = CDate(Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss"))
        Else
            str发生时间 = "To_Date('" & Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
            dat发生时间 = CDate(Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss"))
        End If
        If mViewMode = v_专家号分时段 Then
            str时间段 = "Select Rownum As Id, 序号, To_Char(开始时间, 'hh24') || ':00' As 时间点, To_Char(开始时间, 'hh24:mi') As 开始时间," & vbNewLine & _
                    "       To_Char(终止时间, 'hh24:mi') As 结束时间, 开始时间 As 详细开始时间, 终止时间 As 详细结束时间 " & vbNewLine & _
                    "From 临床出诊序号控制" & vbNewLine & _
                    "Where 记录id = [1] And Nvl(挂号状态,0) = 0 And Nvl(是否预约,0)=1 And Trunc(开始时间) = [2]" & vbNewLine & _
                    "Order By 详细开始时间"
            Set rs时间段 = gobjDatabase.OpenSQLRecord(str时间段, Me.Caption, Val(Nvl(mrsPlan!记录ID)), CDate(Format(dtpDate.Value, "yyyy-mm-dd")))
            If rs时间段.RecordCount = 0 Then
                MsgBox "当前选择的分时段号别无可用时段，无法预约！", vbInformation, gstrSysName
                Exit Function
            End If
            strSql = "Select a.序号,a.开始时间 From 临床出诊序号控制 A Where a.记录ID=[1] And Nvl(a.挂号状态,0) = 0 And Not Exists (Select 1 From 临床出诊挂号控制记录 Where 类型=1 And 记录ID=[1] And 序号=a.序号 And 控制方式=3) Order By a.序号"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!记录ID)))
            If rsTmp.RecordCount <> 0 Then
                Dim blnFind As Boolean
                blnFind = False
                Do While blnFind = False
                    If rsTmp.EOF Then
                        Exit Do
                    Else
                        If str发生时间 = "To_Date('" & Format(Nvl(rsTmp!开始时间), "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')" Then blnFind = True: Exit Do
                        rsTmp.MoveNext
                    End If
                Loop
                If blnFind Then
                    mlngSN = Val(Nvl(rsTmp!序号))
                Else
                    If MsgBox("因为并发原因,选择的时段已不可用,是否自动获取可用时段进行" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
                        Exit Function
                    End If
                    rsTmp.MoveFirst
                    mlngSN = Val(Nvl(rsTmp!序号))
                    str发生时间 = "To_Date('" & Format(Nvl(rsTmp!开始时间), "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
                    dat发生时间 = CDate(Format(Nvl(rsTmp!开始时间), "yyyy-mm-dd hh:mm:ss"))
                End If
            Else
                blnAdd = True
                strSql = "Select Max(序号) As 序号 From 临床出诊序号控制 Where 记录ID=[1]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!记录ID)))
                mlngSN = Val(Nvl(rsTmp!序号)) + 1
            End If
        End If
        If Val(Nvl(mrsPlan!记录ID)) = 0 Then blnNotWork = True
    Else
        Select Case mViewMode
            Case V_普通号
                str发生时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                If Val(Nvl(mrsPlan!记录ID)) = 0 Then blnNotWork = True
            Case V_普通号分时段
                str发生时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                If Val(Nvl(mrsPlan!记录ID)) = 0 Then blnNotWork = True
            Case v_专家号
                str发生时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                If Val(Nvl(mrsPlan!记录ID)) = 0 Then blnNotWork = True
            Case v_专家号分时段
                If Val(Nvl(mrsPlan!记录ID)) = 0 Then
                    str发生时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    blnNotWork = True
                Else
                    '取最小可用时间段
                    strSql = "Select 序号,开始时间 From 临床出诊序号控制 Where 记录ID=[1] And Nvl(挂号状态,0) = 0 Order By 序号"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!记录ID)))
                    If rsTmp.RecordCount <> 0 Then
                        mlngSN = Val(Nvl(rsTmp!序号))
                        str发生时间 = "To_Date('" & Format(Nvl(rsTmp!开始时间), "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        blnAdd = True
                        strSql = "Select Max(序号) As 序号 From 临床出诊序号控制 Where 记录ID=[1]"
                        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!记录ID)))
                        mlngSN = Val(Nvl(rsTmp!序号)) + 1
                        str发生时间 = "To_Date('" & gobjDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss')"
                    End If
                End If
        End Select
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!已挂)) >= Val(Nvl(mrsPlan!限号)) And Val(Nvl(mrsPlan!限号)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!已约)) >= Val(Nvl(mrsPlan!限约)) And Val(Nvl(mrsPlan!限约)) <> 0 Then
            blnAdd = True
        End If
    End If
    
    mlng锁号记录ID = 0
    If (mViewMode = v_专家号 Or mViewMode = v_专家号分时段) And Val(Nvl(mrsPlan!记录ID)) <> 0 Then
        If ReserveRegNo(Nvl(mrsPlan!号别), True, mViewMode = v_专家号分时段, str发生时间, mlngSN, "医生站锁号", Val(Nvl(mrsPlan!记录ID))) = False Then Exit Function
        mlng锁号记录ID = Val(Nvl(mrsPlan!记录ID))
    End If
    
    If blnAdd And InStr(gstrPrivs, ";加号;") = 0 Then
        MsgBox "你没有加号权限，无法对当前号别进行" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";加号;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
    
    If cboAppointStyle.Visible And mblnAppointment And blnAdd = False Then
        strSql = "Select Zl_Fun_Get临床出诊预约状态([1],[2],[3],[4]) As 预约检查 From Dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!记录ID)), dat发生时间, mlngSN, NeedName(cboAppointStyle.Text))
        If rsTemp.EOF Then
            MsgBox "当前选择的预约方式无法预约,请选择其他预约方式!", vbInformation, gstrSysName
            If cboAppointStyle.Enabled And cboAppointStyle.Visible Then cboAppointStyle.SetFocus
            Exit Function
        Else
            If Val(Mid(Nvl(rsTemp!预约检查), 1, 1)) <> 0 Then
                MsgBox "当前选择的预约方式无法预约,请选择其他预约方式!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!预约检查), InStr(Nvl(rsTemp!预约检查), "|") + 1), vbInformation, gstrSysName
                If cboAppointStyle.Enabled And cboAppointStyle.Visible Then cboAppointStyle.SetFocus
                Exit Function
            End If
        End If
    End If
    
    strSql = "Select Zl_临床出诊限制_Check([1],[2],[3]) As 适用性检查 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!记录ID)), mstrGender, mstrAge)
    If rsTemp.EOF Then
        MsgBox "当前选择的病人不适用该号别!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!适用性检查), 1, 1)) <> 0 Then
            MsgBox "当前选择的病人不适用该号别!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!适用性检查), InStr(Nvl(rsTemp!适用性检查), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng挂号科室ID = Val(Nvl(mrsPlan!科室ID))
    lng结帐ID = gobjDatabase.GetNextId("病人结帐记录")
    byt复诊 = IIf(chk复诊.Value = 1, 1, 0)
    

    lngSN = mlngSN
    strNO = gobjDatabase.GetNextNo(12)
    
    rsItems.Filter = ""
    str医生 = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lng医生ID = 0
    Else
        lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    strSql = "Select 编码 From 医疗付款方式 Where 名称 = [1]"
    Set rs付款方式 = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, Nvl(mrsInfo!医疗付款方式))
    If rs付款方式.RecordCount <> 0 Then
        str付款方式 = Nvl(rs付款方式!编码)
    Else
        strSql = "Select 编码 From 医疗付款方式 Where 缺省标志 = 1"
        Set rs付款方式 = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
        If rs付款方式.RecordCount <> 0 Then
            str付款方式 = Nvl(rs付款方式!编码)
        End If
    End If
    
    dblTotal = 0
    If mRegistFeeMode = EM_RG_划价 Then
        dblTotal = GetRegistMoney(True, False)
        '挂号费存为零且保存为划价单，才产生划价NO
       If dblTotal <> 0 Then str划价NO = gobjDatabase.GetNextNo(13)
    End If
    
    
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int价格父号 = k
        rsIncomes.Filter = "项目ID=" & rsItems!项目ID
        For j = 1 To rsIncomes.RecordCount
            strSql = _
            "zl_病人挂号记录_出诊_INSERT(" & ZVal(Nvl(mrsPlan!记录ID)) & "," & ZVal(Nvl(mrsInfo!病人ID)) & "," & IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & txtPatient.Text & "','" & mstrGender & "'," & _
                     "'" & mstrAge & "','" & str付款方式 & "','" & mstrFeeType & "','" & strNO & "'," & _
                     "'" & IIf(blnInvoicePrint = False, "", "") & "'," & k & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                     "'" & rsItems!类别 & "'," & rsItems!项目ID & "," & rsItems!数次 & "," & rsIncomes!单价 & "," & _
                     rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "','" & IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), "") & "'," & _
                     IIf(mRegistFeeMode = EM_RG_划价, 0, rsIncomes!应收) & "," & IIf(mRegistFeeMode = EM_RG_划价, 0, rsIncomes!实收) & "," & _
                     lng挂号科室ID & "," & lng挂号科室ID & "," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                     str发生时间 & "," & str登记时间 & "," & _
                     "'" & str医生 & "'," & ZVal(lng医生ID) & "," & IIf(rsItems!性质 = 3, 1, IIf(rsItems!性质 = 4, 2, 0)) & "," & IIf(lbl急.Visible, 1, 0) & "," & _
                     "'" & mrsPlan!号别 & "','" & IIf(str医生 = UserInfo.姓名, lblRoomName.Caption, "") & "'," & ZVal(lng结帐ID) & "," & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng领用ID)) & "," & _
                     ZVal(IIf(k = 1, cur预交, 0)) & "," & ZVal(IIf(k = 1, cur现金, 0)) & "," & _
                     ZVal(IIf(k = 1, cur个帐, 0)) & "," & ZVal(Nvl(rsItems!保险大类ID, 0)) & "," & _
                     ZVal(Nvl(rsItems!保险项目否, 0)) & "," & ZVal(Nvl(rsIncomes!统筹金额, 0)) & "," & _
                     "'" & Trim(cboRemark.Text) & "'," & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 0, 1), 0) & "," & IIf(mty_Para.bln共用收费票据, 1, 0) & ",'" & rsItems!保险编码 & "'," & byt复诊 & "," & ZVal(lngSN) & ",Null," & _
                     IIf(mblnAppointment, 1, 0) & ",'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "'," & _
                     0 & ","
            '卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & IIf(lng医疗卡类别ID <> 0 And bln消费卡 = False, lng医疗卡类别ID, "NULL") & ","
            '结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
            strSql = strSql & "" & IIf(lng医疗卡类别ID <> 0 And bln消费卡, lng医疗卡类别ID, "NULL") & ","
            '卡号_In       病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & mstrCardNO & "',"
            '交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSql = strSql & " NULL,"
            '交易说明_In   病人预交记录.交易说明%Type := Null,
            strSql = strSql & " NULL,"
            '合作单位_In   病人预交记录.合作单位%Type := Null
            strSql = strSql & " NULL,"
            '  操作类型_In   Number:=0
            strSql = strSql & IIf(blnAdd, 1, 0) & ","
            '  险类_IN       病人挂号记录.险类%type:=null,
            strSql = strSql & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  结算模式_IN   NUMBER :=0,
            strSql = strSql & IIf(mPatiChargeMode = EM_先诊疗后结算, 1, 0) & ","
            '  记帐费用_IN Number:=0
            strSql = strSql & IIf(mRegistFeeMode = EM_RG_记帐, 1, 0) & ","
            '  退号重用_IN Number:=1
            strSql = strSql & IIf(mty_Para.bln退号重用, 1, 0) & ","
            '  冲预交病人ids_In Varchar2 := Null
            strSql = strSql & "'" & Nvl(mrsInfo!病人ID) & "," & mstr家属IDs & "',"
            '  修正病人费别_In Number := 0
            strSql = strSql & "" & IIf(mblnChangeFeeType, 1, 0) & ",Null,"
            '  修正病人年龄_In Number := 0
            strSql = strSql & "" & IIf(mblnUpdateAge, 1, 0) & ","
            '  收费单_In        病人挂号记录.收费单%Type := Null
            strSql = strSql & "'" & str划价NO & "')"
            
            Call zlAddArray(cllPro, strSql)
            '问题:31187:将挂号汇总单独出来
            If Nvl(mrsPlan!号别) <> "" And k = 1 Then
                If Nvl(mrsPlan!医生姓名) = "" Then blnNoDoc = True
                strSql = "zl_病人挂号汇总_Update("
                '  医生姓名_In   挂号安排.医生姓名%Type,
                strSql = strSql & IIf(blnNoDoc, "Null,", "'" & str医生 & "',")
                '  医生id_In     挂号安排.医生id%Type,
                strSql = strSql & "" & IIf(blnNoDoc, "0,", ZVal(lng医生ID) & ",")
                '  收费细目id_In 门诊费用记录.收费细目id%Type,
                strSql = strSql & "" & Val(Nvl(rsItems!项目ID)) & ","
                '  执行部门id_In 门诊费用记录.执行部门id%Type,
                strSql = strSql & "" & IIf(Val(Nvl(rsItems!执行科室ID)) = 0, lng挂号科室ID, Val(Nvl(rsItems!执行科室ID))) & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                strSql = strSql & "" & str发生时间 & ","
                '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
                strSql = strSql & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 3, 1), 0) & ","
                '  号码_In       挂号安排.号码%Type := Null
                strSql = strSql & "'" & Nvl(mrsPlan!号别) & "',0,"
                strSql = strSql & "" & ZVal(Nvl(mrsPlan!记录ID)) & ")"
                Call zlAddArray(cllProAfter, strSql)
            End If
            
            If mRegistFeeMode = EM_RG_划价 And dblTotal <> 0 Then
                strSql = _
                "zl_门诊划价记录_Insert('" & str划价NO & "'," & k & "," & ZVal(Nvl(mrsInfo!病人ID)) & ",NULL," & _
                         IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & str付款方式 & "'," & _
                         "'" & txtPatient.Text & "','" & mstrGender & "','" & mstrAge & "'," & _
                         "'" & mstrFeeType & "',NULL," & lng挂号科室ID & "," & _
                         IIf(lng挂号科室ID <> 0, lng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                         rsItems!项目ID & ",'" & rsItems!类别 & "','" & rsItems!计算单位 & "'," & _
                         "NULL,1," & rsItems!数次 & ",NULL," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & _
                         rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "'," & rsIncomes!单价 & "," & _
                         rsIncomes!应收 & "," & rsIncomes!实收 & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                Call zlAddArray(cllPro, strSql)
            End If
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i
    
    If Not mblnAppointment Then
        If str医生 = UserInfo.姓名 Then
            strSql = "ZL_病人挂号记录_更新诊室('" & strNO & "'," & Nvl(mrsInfo!病人ID) & ",'" & lblRoomName.Caption & "','" & UserInfo.姓名 & "','','','" & zl_Get预约方式ByNo(strNO) & "')"    '问题号:48350
            Call zlAddArray(cllPro, strSql)
            strSql = "zl_病人接诊(" & Nvl(mrsInfo!病人ID) & ",'" & strNO & "',NULL,'" & UserInfo.姓名 & "','" & lblRoomName.Caption & "')"
            Call zlAddArray(cllPro, strSql)
        End If
    End If
    
    Err = 0: On Error GoTo ErrFirt:
    
    If cllPro.Count > 0 Then
        Err = 0: On Error GoTo ErrFirt:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

        Err = 0: On Error GoTo errH:
        blnTrans = True
        If blnOneCard And lng医疗卡类别ID <> 0 And mRegistFeeMode = EM_RG_现收 And cur现金 <> 0 Then
            If Not mobjICCard.PaymentSwap(Val(cur现金), Val(cur现金), Val(lng医疗卡类别ID), 0, mstrCardNO, "", lng结帐ID, Nvl(mrsInfo!病人ID)) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通结算挂号费失败", vbInformation, gstrSysName
                Exit Function
            Else
                strSql = "zl_一卡通结算_Update(" & lng结帐ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lng医疗卡类别ID & "','" & "" & "'," & cur现金 & ")"
                Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If

        '医保改动
        blnNotCommit = False
        If mintInsure <> 0 And mstrYBPati <> "" And cur个帐 <> 0 Then
            '68991:strAdvance:结算模式(0或1)|挂号费收取方式(0或1) |挂号单号
            strAdvance = ""
            If mRegistFeeMode = EM_RG_记帐 Or mPatiChargeMode = EM_先诊疗后结算 Then
                strAdvance = IIf(mPatiChargeMode = EM_先诊疗后结算, "1", "0")
                strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_记帐, "1", "0")
                strAdvance = strAdvance & "|" & strNO
            End If
            If Not gclsInsure.RegistSwap(lng结帐ID, cur个帐, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
            End If
            blnNotCommit = True
        End If
        '问题:31187 调用医保成功后,最后作一些数据更新:内部过程中已有提交语句,所以不用再写
        zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
        Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
        If mRegistFeeMode = EM_RG_现收 And Not blnOneCard And Not mPatiChargeMode = EM_先诊疗后结算 And cur现金 <> 0 Then
            If zlInterfacePrayMoney(lng结帐ID, cllCardPro, cllTheeSwap, Val(cur现金), lng医疗卡类别ID, bln消费卡) = False Then
                gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True
                Exit Function
            End If
            '修正三方交易
            zlExecuteProcedureArrAy cllCardPro, Me.Caption, False, False
        End If
        
        Err = 0: On Error GoTo OthersCommit:
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
OthersCommit:
        gcnOracle.CommitTrans
        blnTrans = False
        On Error GoTo 0
        
        If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
        '145198:李南春,2019/12/26,挂号成功后调用外挂接口，目前用于预约后产生支付二维码
        Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.bln预约时收款))
    End If
    '打印单据
    If blnInvoicePrint Then
RePrint:
        If Not (mintInsure <> 0 And MCPAR.医保接口打印票据) And mRegistFeeMode = EM_RG_现收 Then
            Dim blnEnterPrint As Boolean
            blnEnterPrint = True
            Load frmPrint
            Call frmPrint.ReportPrint(1, strNO, "", mlng领用ID, mlng挂号ID, strFactNO, dat登记时间, , , , mintInsure <> 0 And MCPAR.医保接口打印票据, False, mstrUseType)
            If gblnBill挂号 Then
                If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                    If MsgBox("" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "单号为[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    If blnSlipPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
    End If
    
    If blnAppointPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    End If
    
    If (blnSlipPrint Or blnInvoicePrint) And Not blnEnterPrint Then
        '记录打印的凭条
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    mstrNO = strNO
    SaveData = True
    Exit Function
ErrFirt:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Function
errH:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Function
ErrGo:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


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

Private Function CheckValied() As Boolean
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
    
    If mstrFeeType = "" Then
        MsgBox "病人费别不能为空,请先选择一个费别！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mRegistFeeMode <> EM_RG_划价 Then
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

Private Function zlInterfacePrayMoney(ByVal lng挂号结帐ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double, lng医疗卡类别ID As Long, bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lng医疗卡类别ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lng医疗卡类别ID, bln消费卡, mstrCardNO, lng挂号结帐ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If lng挂号结帐ID <> 0 Then
        '问题:58322
        'mbytMode As Integer '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
        If Not bln消费卡 Then
            '消费卡已经在插入挂号记录时,已经扣款
            Call zlAddUpdateSwapSQL(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, mstrCardNO, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, mstrCardNO, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    ' 出参:cllPro-返回SQL集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSql As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    Dim str交易信息 As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSql = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSql = strSql & "" & lng卡类别ID & ","
                    '消费卡_In   Number,
                    strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSql = strSql & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSql = strSql & "'" & strIDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSql = strSql & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSql = strSql & IIf(bln预交款, "1", "0") & ")"
                    zlAddArray cllPro, strSql
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSql = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSql = strSql & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSql = strSql & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSql = strSql & "'" & strIDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSql = strSql & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSql = strSql & IIf(bln预交款, "1", "0") & ")"
        zlAddArray cllPro, strSql
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSql = strSql & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSql = strSql & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSql = strSql & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSql = strSql & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSql = strSql & "'" & str交易说明 & "',"
    '预交款缴款_In Number := 0
    strSql = strSql & "" & IIf(bln预交, 1, 0) & ","
    '退费标志 :1-退费;0-付费
    strSql = strSql & "0,"
    '校对标志
    strSql = strSql & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ")"
    zlAddArray cllPro, strSql
End Function

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
            cmdPrice.Visible = mty_Para.byt挂号模式 = 2
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
        Me.Caption = "医生站" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & ""
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
    End If
    
    gobjDatabase.ExecuteProcedure "zl1_auto_buildingregisterplan", Me.Caption
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
    If mblnAppointment And mlng病人ID <> 0 And mblnUnload = False Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False)
    End If
    cmdNewPati.Enabled = InStr(gstrPrivs, ";挂号病人建档;") > 0
    cmdOther.Enabled = InStr(gstrPrivs, ";允许挂其他医生的号源;") > 0
    '137272:李南春,2019/2/20,防止锁号后系统意外崩溃的情况
    Call CancelRegNo
End Sub

Private Sub InitAppointmentTime()
    '初始化预约时间
    Dim rsDay As ADODB.Recordset, strSql As String
    Dim int预约天数 As Integer
    Dim dtNow As Date
  
    On Error GoTo ErrHandler
    int预约天数 = mintSysAppLimit
    If mblnAppointment Then
        Call mobjRegister.zlGetRegisterMaxDaysFromDeptAndDoctor_Visits( _
            gstrDeptIDs, UserInfo.姓名, mty_Para.bln预约包含科室安排, int预约天数)
    End If
    
    dtNow = gobjDatabase.Currentdate
    dtpDate.MaxDate = Format(dtNow + int预约天数, "yyyy-mm-dd")
    dtpDate.minDate = Format(dtNow, "yyyy-mm-dd")
    dtpTime.Value = Format(dtNow, "hh:mm:ss")
    
    strSql = _
        "Select Nvl(Min(a.出诊日期), Trunc(Sysdate) + 1) As 出诊日期" & vbNewLine & _
        "From 临床出诊记录 A" & vbNewLine & _
        "Where a.出诊日期 > Trunc(Sysdate) And a.科室id = [1] And a.医生id = [2]"
    Set rsDay = gobjDatabase.OpenSQLRecord(strSql, "", UserInfo.部门ID, UserInfo.ID)
    dtpDate.Value = Format(rsDay!出诊日期, "yyyy-mm-dd")
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
    Set rsTmp = GetMoneyInfoRegist(lng病人ID, , , 1)
    cur余额 = 0
    Do While Not rsTmp.EOF
        cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
        cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
        rsTmp.MoveNext
    Loop
    If cur余额 > 0 Then
        lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00")
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
'    If KeyCode = vbKeyReturn Then
'        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
'        gobjControl.TxtSelAll txtPatient
'    End If
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
    '
    '         blnInputIDCard-是否身份证刷卡
    '出参:Cancel-为true表示返回的放弃读取病人信息
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim strInputInfo As String '保存传入的输入文本 避免在使用身份证号 对病人进行查找后 被替换成"-" 病人ID的情况
    Dim i As Integer, strPati As String
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
             "B.名称 险类名称,A.查询密码 As 卡验证码,A.结算模式,a.主页ID From 病人信息 A,保险类别 B  Where A.险类 = B.序号(+) And A.停用时间 is NULL "

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
        mstrGender = Nvl(mrsInfo!性别)
        txtPatient.PasswordChar = ""
        
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        If Load费别(Nvl(mrsInfo!费别)) = False Then mstrFeeType = ""
        
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
        mstrClinic = Nvl(mrsInfo!门诊号)
        If mstrClinic = "" Then
            mstrClinic = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        lblInfo.Caption = "性别:" & mstrGender & "   年龄:" & mstrAge & "   门诊号:" & mstrClinic & "   费别:" & mstrFeeType
        
        '病人预交款信息
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!病人ID, , , 1)
        cur余额 = 0
        Do While Not rsTmp.EOF
            cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
            cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
            rsTmp.MoveNext
        Loop
        If cur余额 > 0 Then
            lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00")
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
        
        Call ResetDefault复诊
        
        '根据病人重新读取项目费用
        If mintPriceGradeStartType >= 2 Then
            Call GetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Nvl(mrsInfo!医疗付款方式, 0), , , mstrPriceGrade)
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
    mstrGender = ""
    mstrAge = ""
    cmdNewPati.ToolTipText = "新增病人(F4)"
    cmdNewPati.Enabled = InStr(gstrPrivs, ";挂号病人建档;") > 0
    mstrClinic = ""
    mstrFeeType = ""
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
    Dim i As Integer, k As Integer, datNow As Date
    Dim DateThis As Date, strZero As String
    Dim str挂号安排 As String, strViewSQL As String
    Dim str挂号安排计划 As String, strCondition As String
    Dim str排序         As String
    Dim vRect          As RECT
    Dim varTemp As Variant, varData As Variant
    On Error GoTo errH
    
    If chkAll.Value = 0 Then
        varTemp = Split(mty_Para.strStationRegOrder, "|")
        For i = 0 To UBound(varTemp)
            varData = Split(varTemp(i), ",")
            Select Case varData(0)
                Case "医生"
                    str排序 = str排序 & ",Decode(医生姓名,Null,Decode(科室ID," & mlngDept & ",3,4),Decode(科室ID," & mlngDept & ",1,2)),医生姓名 " & IIf(varData(1) = 1, "", "desc")
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
        str排序 = "Decode(出诊日期,Null,2,1),出诊日期 Desc" & str排序
    Else
        str排序 = "Decode(出诊日期,Null,2,1),出诊日期 Desc,Decode(医生姓名,'" & UserInfo.姓名 & "',1,2),Decode(科室ID," & mlngDept & ",1,2),已挂,号别,项目"
    End If
    
    If gstrDeptIDs <> "" And Not blnOtherDoctor Then strIF = " And Instr(','||[4]||',',','||a.科室ID||',')>0"
    If mblnAppointment Then
        If mty_Para.bln预约包含科室安排 Then
            strIF = strIF & IIf(blnOtherDoctor, " And (a.医生姓名 <> [1] or a.医生姓名 Is Null)", " And (a.医生姓名 = [1] or a.医生姓名 Is Null)")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (a.医生姓名 <> [1] )", " And (a.医生姓名 = [1])")
        End If
    Else
        If mty_Para.bln挂号包含科室安排 Then
            strIF = strIF & IIf(blnOtherDoctor, " And (a.医生姓名 <> [1] or a.医生姓名 Is Null)", " And (a.医生姓名 = [1] or a.医生姓名 Is Null)")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (a.医生姓名 <> [1] )", " And (a.医生姓名 = [1])")
        End If
    End If
    
    If intSelMode = 2 Then
        strCondition = " And (b.号码 Like [11] Or Upper(c.名称) Like Upper([11]) Or Upper(zlSpellCode(c.名称)) Like Upper([11]) Or Upper(a.医生姓名) Like Upper([11]) Or Upper(zlSpellCode(a.医生姓名)) Like Upper([11]))"
    End If
    
    strSql = "Select a.Id As 记录ID, b.号码 As 号别, b.号类, b.科室id, c.名称 As 科室, a.项目id, d.名称 As 项目, Nvl(a.替诊医生id,a.医生id) As 医生id, Nvl(a.替诊医生姓名,a.医生姓名) As 医生姓名, Nvl(a.已挂数, 0) As 已挂," & vbNewLine & _
            "       Nvl(a.已约数, 0) As 已约, a.限号数 As 限号, a.限约数 As 限约, Nvl(b.是否建病案, 0) As 病案, Nvl(d.项目特性, 0) As 急诊, a.分诊方式 As 分诊," & vbNewLine & _
            "       a.是否序号控制 As 序号控制, a.上班时段 As 排班, a.号源id, a.是否分时段 As 分时段, a.开始时间, a.终止时间, a.出诊日期  " & vbNewLine & _
            "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E" & vbNewLine & _
            "Where (a.出诊日期 = [6] Or a.出诊日期 = [8]) And Nvl(C.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.号源id = b.Id And b.科室id = c.Id And a.项目id = d.Id And Nvl(a.是否锁定, 0) = 0 " & vbNewLine & _
            "       And a.医生id = e.Id(+) And (d.撤档时间 is NULL Or d.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & _
            "       And Nvl(a.是否发布,0) = 1 "
    strSql = strSql & " And (a.开始时间 < Nvl(a.停诊开始时间,a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间,a.开始时间) Or Exists (Select 1 From 临床出诊序号控制 C,临床出诊记录 D Where D.ID=A.ID And C.记录ID=D.ID And Nvl(C.是否停诊,0) = 0 And D.是否序号控制 =1 And D.是否分时段 = 1 And C.开始时间 <> C.终止时间)) "
    If chkAll.Value <> 1 Then
        strSql = strSql & " And [5] Not Between  Nvl(a.停诊开始时间,a.终止时间) And Nvl(a.停诊终止时间,a.开始时间) "
    End If
    If mblnAppointment Then
        strSql = strSql & " And Nvl(a.预约控制,0) <> 1 "
        DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
    Else
        DateThis = gobjDatabase.Currentdate
    End If
    datNow = gobjDatabase.Currentdate
    
    If mblnAppointment Then
        If Format(DateThis, "yyyy-mm-dd") = Format(datNow, "yyyy-mm-dd") Then
            strSql = strSql & "       And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [9])"
        Else
            strSql = strSql & "       And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [6])"

        End If
    Else
        strSql = strSql & " And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [5]) "
    End If
    
    strSql = strSql & strIF & strCondition
    
    If intSelMode = 2 Then
        strCondition = " And (a.号码 Like [11] Or Upper(c.名称) Like Upper([11]) Or Upper(zlSpellCode(c.名称)) Like Upper([11]) Or Upper(a.医生姓名) Like Upper([11]) Or Upper(zlSpellCode(a.医生姓名)) Like Upper([11]))"
    End If
    
    strTime = " Union All " & _
            "Select 0 As 记录id, a.号码 As 号别, a.号类, a.科室id, c.名称 As 科室, a.项目id, d.名称 As 项目, a.医生id, a.医生姓名, 0 As 已挂, 0 As 已约, Null As 限号," & vbNewLine & _
            "       Null As 限约, Nvl(a.是否建病案, 0) As 病案, Nvl(d.项目特性, 0) As 急诊, 0 As 分诊, 0 As 序号控制, Null As 排班, a.Id As 号源id, 0 As 分时段, Null As 开始时间, Null As 终止时间, Null As 出诊日期 " & vbNewLine & _
            "From 临床出诊号源 A, 部门表 C, 收费项目目录 D" & vbNewLine & _
            "Where a.科室id = c.Id And Nvl(C.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.项目id = d.Id " & vbNewLine & _
            "      And Exists (Select 1 From 临床出诊安排 M,临床出诊表 N Where M.号源ID=A.ID And M.出诊ID=N.ID And N.发布时间 Is Not Null) " & vbNewLine & _
            "      And Sysdate < Nvl(a.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) " & strIF & _
            "      And Not Exists(Select 1" & vbNewLine & _
            "                     From 临床出诊记录" & vbNewLine & _
            "                     Where 号源id = a.Id And (出诊日期 = [6] Or 出诊日期 = [8])" & vbNewLine & _
            "                           And [5] Between 开始时间 And 终止时间" & vbNewLine & _
            "                           And (开始时间 < Nvl(停诊开始时间, 终止时间) Or 终止时间 > Nvl(停诊终止时间, 开始时间))" & vbNewLine & _
            "                           And Nvl(是否锁定, 0) = 0 And Nvl(是否发布, 0) = 1)"
        
    If mblnAppointment Then
        '预约挂号
        strSql = strSql & " And (a.限约数 > 0 Or a.限约数 Is Null)"
        strSql = strSql & " And Nvl(a.预约控制,0) <> 1 "
        strSql = strSql & " And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.预约天数," & mty_Para.int可预约天数 & "),0,15,Nvl(B.预约天数," & mty_Para.int可预约天数 & ")" & ") > [6] "
    Else
        '挂号
        If chkAll.Value = 1 Then strSql = strSql & strTime
    End If
    
    strViewSQL = "Select RowNum As Id,A.记录id,A.号别,a.号类,a.科室,a.科室id,a.项目,a.医生姓名,a.已挂,a.已约,a.限号,a.限约," & _
                 "      Decode(nvl(a.病案,0),1,'√','') As 病案,Decode(nvl(a.急诊,0),1,'√','') As 急诊,Decode(nvl(a.序号控制,0),1,'√','') As 序号控制," & _
                 "      a.排班,Decode(nvl(a.分时段,0),1,'√','') As 分时段,a.出诊日期 From (" & strSql & ") A Order By " & str排序
    strSql = "Select * From (" & strSql & ") Order By " & str排序
                
    Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, _
            UserInfo.姓名, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, CDate(Format(DateThis - 1, "yyyy-MM-dd")), gobjDatabase.Currentdate, gdatRegistTime, "%" & strFilter & "%")
                

 
    If mrsPlan.RecordCount <> 0 Then
        If intSelMode = 1 Or mrsPlan.RecordCount = 1 Then
            '默认读取
            Call ReadLimit(Val(Nvl(mrsPlan!记录ID)), Nvl(mrsPlan!号别))
        Else
            vRect = GetControlRect(txtReg.hWnd)
            Set rsPlan = gobjDatabase.ShowSQLSelect(Me, strViewSQL, 0, "号码选择", False, "", "号码选择", _
                                                False, False, True, vRect.Left, vRect.Top - 250, 600, False, True, False, _
                                                UserInfo.姓名, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), _
                                                CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, CDate(Format(DateThis - 1, "yyyy-MM-dd")), _
                                                gobjDatabase.Currentdate, gdatRegistTime, "%" & strFilter & "%")
            If rsPlan Is Nothing Then
                Call ReadLimit(Val(Nvl(mrsPlan!记录ID)), Nvl(mrsPlan!号别))
            Else
                If Not rsPlan.EOF Then
                    Call ReadLimit(Val(Nvl(rsPlan!记录ID)), Nvl(rsPlan!号别))
                Else
                    Call ReadLimit(Val(Nvl(mrsPlan!记录ID)), Nvl(mrsPlan!号别))
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
                    dtpTime.Enabled = False
                    cmdTime.Visible = True
                Case Else
                    dtpTime.Enabled = True
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

Private Sub ReadLimit(ByVal lng记录ID As Long, str号码 As String)

    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    
    mrsPlan.Filter = "记录ID=" & lng记录ID & " And 号别='" & str号码 & "'"
    
    If mrsPlan.RecordCount = 0 Then Exit Sub
    mblnIntact = True
    mblnChangeByCode = True
    If Nvl(mrsPlan!医生姓名) = "" Then
        txtReg.Text = "[" & Nvl(mrsPlan!号别) & "]" & Nvl(mrsPlan!项目)
    Else
        txtReg.Text = "[" & Nvl(mrsPlan!号别) & "]" & Nvl(mrsPlan!项目) & "(" & Nvl(mrsPlan!医生姓名) & ")"
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

Private Function GetActiveView()
    '得到当前挂号业务  采取那种类型的流程
    Dim strSql          As String
    Dim rsTmp           As ADODB.Recordset
    If mrsPlan Is Nothing Then Exit Function
    If mrsPlan.RecordCount = 0 Then Exit Function
    On Error GoTo errH
    
    strSql = "Select 1 From 临床出诊记录 Where ID=[1] And Nvl(是否分时段,0)=1 "
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsPlan!记录ID))
    
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
     Dim vRect          As RECT
     
   
    If Not mblnAppointment Then Exit Function
    
    If mrsPlan Is Nothing Then Exit Function
    If mrsPlan.State <> 1 Then Exit Function
    If mrsPlan.EOF Then Exit Function
    
    If Not (Val(Nvl(mrsPlan!序号控制)) = 1 And Val(Nvl(mrsPlan!分时段)) = 1) Then Exit Function
    
    lblSn.Caption = ""
    strSql = "" & _
    " Select Rownum As Id, 序号, To_Char(开始时间, 'hh24') || ':00' As 时间点, To_Char(开始时间, 'hh24:mi') As 开始时间," & vbNewLine & _
    "       To_Char(终止时间, 'hh24:mi') As 结束时间, 开始时间 As 详细开始时间, 终止时间 As 详细结束时间 " & vbNewLine & _
    " From 临床出诊序号控制" & vbNewLine & _
    " Where 记录id = [1] And Nvl(挂号状态,0) = 0 And Nvl(是否预约,0)=1 And Trunc(开始时间) = [2]" & vbNewLine & _
    "Order By 详细开始时间"

    If strSql = "" Then Exit Function
    vRect = GetControlRect(dtpTime.hWnd)
    
    On Error GoTo errH
    
    Set mrs时间段 = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "预约时间选择", False, "", "预约时间选择", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, mrsPlan!记录ID, CDate(Format(dtpDate.Value, "yyyy-mm-dd")))
    If mrs时间段 Is Nothing Then Exit Function
    If mrs时间段.EOF Then Exit Function
    
    lblSn.Caption = "序号:" & Val(Nvl(mrs时间段!序号))
    dtpTime.Value = Format(mrs时间段!开始时间, "hh:mm:ss")
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
    ReadRegistPrice lngItemID, blnBook, False, mstrFeeType, rsItems, rsIncomes, , , , 1, Val(Nvl(mrsPlan!科室ID)), strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    If mintInsure <> 0 Then
        If MCPAR.挂号检查项目 = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "医保病人收费项目检查失败，不能继续" & IIf(gSysPara.bln免挂号模式, "就诊", "挂号") & "！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    ReadRegistPrice lngItemID, blnBook, False, mstrFeeType, rsItems, rsIncomes, lng病人ID, mintInsure, _
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


Private Function GetSNState(lng记录ID As Long) As ADODB.Recordset
    Dim strSql           As String
    On Error GoTo errH

    strSql = "    " & vbNewLine & " Select A.序号,A.挂号状态,A.操作员姓名,Decode(A.挂号状态,2,1,0) as 预约,To_Char(B.出诊日期,'hh24:mi:ss') as 日期  "
    strSql = strSql & vbNewLine & " From 临床出诊序号控制 A, 临床出诊记录 B "
    strSql = strSql & vbNewLine & " Where B.ID=[1] And B.ID=A.记录ID"
    Set GetSNState = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng记录ID)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
            mstrFeeType = str费别
        Else
            mstrFeeType = mstrDef费别
        End If
    Else
        mstrFeeType = mstrDef费别
    End If
    If mstrFeeType = "" Then
        MsgBox "未找的适用于【" & Nvl(mrsPlan!科室) & "】的缺省费别,请在『病人详细信息』界面中设置病人费别！", vbInformation, gstrSysName
        Load费别 = False
        Exit Function
    End If
    lblInfo.Caption = "性别:" & mstrGender & "   年龄:" & mstrAge & "   门诊号:" & mstrClinic & "   费别:" & mstrFeeType
    Load费别 = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function




