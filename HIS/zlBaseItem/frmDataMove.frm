VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDataMove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据转移管理"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "frmDataMove.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5000
      Index           =   0
      Left            =   0
      TabIndex        =   30
      Tag             =   "数据转移"
      Top             =   885
      Width           =   9405
      Begin VB.TextBox txtDatePre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F4E4&
         ForeColor       =   &H00000000&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   7740
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   86
         Text            =   "2011-01-01"
         Top             =   2775
         Width           =   1020
      End
      Begin VB.CommandButton cmdDateThis 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   8805
         Picture         =   "frmDataMove.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox txtDateThis 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   280
         IMEMode         =   3  'DISABLE
         Left            =   7740
         MaxLength       =   10
         TabIndex        =   84
         Text            =   "2012-01-01"
         Top             =   2400
         Width           =   1020
      End
      Begin MSComCtl2.MonthView monSel 
         Height          =   2460
         Left            =   3000
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollRate      =   1
         StartOfWeek     =   39583745
         TitleBackColor  =   8421504
         TitleForeColor  =   16777215
         CurrentDate     =   38003
         MaxDate         =   73415
         MinDate         =   -18260
      End
      Begin VB.TextBox txtDateLast 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   280
         IMEMode         =   3  'DISABLE
         Left            =   7740
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "2015-01-01"
         Top             =   1995
         Width           =   1020
      End
      Begin VB.Frame framode 
         Caption         =   "转出参数"
         Height          =   4425
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   6135
         Begin VB.PictureBox picBakspace 
            BorderStyle     =   0  'None
            Height          =   1425
            Left            =   240
            ScaleHeight     =   1425
            ScaleWidth      =   3210
            TabIndex        =   87
            Top             =   2880
            Width           =   3210
            Begin VB.ComboBox cboBakspace 
               Height          =   300
               Index           =   0
               Left            =   1180
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   300
               Width           =   1920
            End
            Begin VB.ComboBox cboBakspace 
               Height          =   300
               Index           =   1
               Left            =   1180
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   660
               Width           =   1920
            End
            Begin VB.ComboBox cboBakspace 
               Height          =   300
               Index           =   2
               Left            =   1180
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   960
               Width           =   1920
            End
            Begin VB.Label lblBakSpace 
               Alignment       =   1  'Right Justify
               Caption         =   "标准版"
               Height          =   300
               Index           =   0
               Left            =   360
               TabIndex        =   94
               Top             =   330
               Width           =   645
            End
            Begin VB.Label lblBakSpace 
               AutoSize        =   -1  'True
               Caption         =   "历史表空间"
               Height          =   180
               Index           =   3
               Left            =   135
               TabIndex        =   93
               Top             =   0
               Width           =   900
            End
            Begin VB.Label lblBakSpace 
               Alignment       =   1  'Right Justify
               Caption         =   "体检系统"
               Height          =   300
               Index           =   1
               Left            =   240
               TabIndex        =   92
               Top             =   720
               Width           =   765
            End
            Begin VB.Label lblBakSpace 
               Alignment       =   1  'Right Justify
               Caption         =   "手麻系统"
               Height          =   300
               Index           =   2
               Left            =   240
               TabIndex        =   91
               Top             =   1080
               Width           =   765
            End
         End
         Begin VB.CheckBox chkBakTbsDisable 
            Caption         =   "禁用历史库的约束和索引"
            Height          =   180
            Left            =   1440
            TabIndex        =   80
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.TextBox txtSplit 
            Alignment       =   1  'Right Justify
            Height          =   280
            Left            =   1440
            TabIndex        =   6
            Text            =   "30"
            Top             =   2280
            Width           =   375
         End
         Begin VB.CheckBox chkjob 
            Caption         =   "禁用当前系统所有者的后台作业"
            Height          =   180
            Left            =   1440
            TabIndex        =   3
            Top             =   1080
            Width           =   2895
         End
         Begin VB.CheckBox chkTrigger 
            Caption         =   "禁用转出表上的触发器"
            Height          =   180
            Left            =   1440
            TabIndex        =   4
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton optmode 
            Caption         =   "在线模式(不用中断业务，可以在客户端正常使用的情况下进行。)"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   0
            Top             =   285
            Value           =   -1  'True
            Width           =   5775
         End
         Begin VB.OptionButton optmode 
            Caption         =   "离线模式(需要中断业务，要求在所有客户端停用的情况下进行。)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   1
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label13 
            Caption         =   "数据转出期间"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   1065
            Width           =   1455
         End
         Begin VB.Label lblSplit 
            Caption         =   "每批转出     天的数据"
            Height          =   255
            Left            =   680
            TabIndex        =   5
            Top             =   2325
            Width           =   2175
         End
      End
      Begin VB.Timer TIMStatus 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   6000
         Top             =   0
      End
      Begin VB.CommandButton cmdPrompt 
         Caption         =   "查看转出须知"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
      Begin VB.CheckBox chkAffirm 
         Caption         =   "我已详细阅读转出须知，并完成了相关的准备和调整。"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   0
         Width           =   4695
      End
      Begin VB.CommandButton cmdDateLast 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   8805
         Picture         =   "frmDataMove.frx":0680
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1995
         Width           =   280
      End
      Begin VB.CommandButton cmdMoveMark 
         Caption         =   "标记转出"
         Height          =   350
         Left            =   6555
         TabIndex        =   12
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdMoveOut 
         Caption         =   "转出(&M)"
         Height          =   350
         Left            =   7995
         TabIndex        =   14
         Top             =   4530
         Width           =   1100
      End
      Begin VB.TextBox txtPrompt 
         Height          =   2175
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   56
         Text            =   "frmDataMove.frx":0776
         Top             =   360
         Visible         =   0   'False
         Width           =   6315
      End
      Begin VB.Label lblDateLast 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最终截止日期"
         Height          =   180
         Left            =   6600
         TabIndex        =   83
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblDatePre 
         Caption         =   "上次截止日期"
         Height          =   180
         Left            =   6600
         TabIndex        =   11
         Top             =   2820
         Width           =   1080
      End
      Begin VB.Label lblDateThis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "本次截止日期"
         Height          =   180
         Left            =   6600
         TabIndex        =   9
         Top             =   2445
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "标记 2011-01-01 到 2011-02-01 之间的数据时发生中断,须继续标记转出这些数据后才能执行新的操作。"
         ForeColor       =   &H00C00000&
         Height          =   1305
         Left            =   6360
         TabIndex        =   7
         Top             =   600
         Width           =   2865
      End
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Height          =   5000
      Index           =   4
      Left            =   0
      TabIndex        =   57
      Tag             =   "转后处理"
      Top             =   840
      Width           =   9375
      Begin VB.CommandButton cmdRebIndexOther 
         Caption         =   "重建其他索引"
         Height          =   350
         Left            =   3360
         TabIndex        =   63
         ToolTipText     =   $"frmDataMove.frx":14B7
         Top             =   2055
         Width           =   1425
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "收缩数据文件"
         Height          =   350
         Left            =   1800
         TabIndex        =   81
         ToolTipText     =   "收缩以ZL为前缀的表空间的所有数据文件，一般应在重整表或重建索引后立即执行才能释放空闲空间（避免在文件尾部产生新的数据）。"
         Top             =   3840
         Width           =   1425
      End
      Begin VB.OptionButton optmode_Index 
         Caption         =   "离线重建(需停业务)"
         Height          =   180
         Index           =   1
         Left            =   4560
         TabIndex        =   78
         Top             =   1230
         Width           =   2055
      End
      Begin VB.OptionButton optmode_Index 
         Caption         =   "在线重建(非常耗时)"
         Height          =   180
         Index           =   0
         Left            =   1560
         TabIndex        =   77
         Top             =   1230
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame fraRebuild 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   1440
         TabIndex        =   73
         Top             =   1605
         Width           =   5295
         Begin VB.OptionButton optRebScope_Manual 
            Caption         =   "全部"
            Height          =   375
            Index           =   2
            Left            =   4320
            TabIndex        =   76
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optRebScope_Manual 
            Caption         =   "经济核算类、医嘱类"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   75
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optRebScope_Manual 
            Caption         =   "经济核算类"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdRebJobTrigger 
         Caption         =   "恢复被禁用的后台作业和触发器"
         Height          =   350
         Left            =   6480
         TabIndex        =   68
         ToolTipText     =   "请在离线转出操作全部完成后执行"
         Top             =   4440
         Width           =   2745
      End
      Begin VB.CommandButton cmdRebOnline 
         Caption         =   "恢复在线空间被禁用的约束及索引"
         Height          =   350
         Left            =   3360
         TabIndex        =   67
         ToolTipText     =   "如果转出操作为在线转出，则采用在线重建（比较耗时，但不影响业务运行），否则采用离线重建"
         Top             =   4440
         Width           =   2985
      End
      Begin VB.CommandButton cmdMoveTable 
         Caption         =   "重整转出表"
         Height          =   350
         Left            =   240
         TabIndex        =   66
         ToolTipText     =   "对所有历史数据转出表执行Move操作，然后恢复失效的索引"
         Top             =   3840
         Width           =   1425
      End
      Begin VB.CommandButton cmdRebBakSpace 
         Caption         =   "恢复历史空间被禁用的约束及索引"
         Height          =   350
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "请在全部转出操作完成后执行，以便历史空间的查询业务能够正常进行(固定采用离线重建模式)"
         Top             =   4440
         Width           =   2985
      End
      Begin VB.TextBox txtParallel 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   2880
         TabIndex        =   64
         Text            =   "12"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdRebIndexForTag 
         Caption         =   "重建标记转出所需的索引"
         Height          =   350
         Left            =   240
         TabIndex        =   62
         Top             =   2055
         Width           =   2985
      End
      Begin VB.Frame fraMove 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   240
         TabIndex        =   58
         Top             =   3480
         Width           =   4305
         Begin VB.OptionButton optMove 
            Caption         =   "经济核算类"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   60
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optMove 
            Caption         =   "全部"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   59
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "重整范围(标准版)"
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   30
            Width           =   1575
         End
      End
      Begin VB.Label lblPrompt 
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   3360
         TabIndex        =   82
         Top             =   3645
         Width           =   5775
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   6600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblReIndexMode 
         Caption         =   "索引重建方式"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label lblReIndexScope 
         Caption         =   "索引重建范围"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblParallel 
         Caption         =   "索引重建和表数据重整的并行度"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   285
         Width           =   6975
      End
      Begin VB.Label lblReIndex 
         Caption         =   "转出数据后，索引的碎片比较多，将会影响后续查询待转出数据的SQL性能，也会影响在线业务中的相关查询性能，建议重建索引。"
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblMove 
         Caption         =   $"frmDataMove.frx":154E
         Height          =   615
         Left            =   240
         TabIndex        =   69
         Top             =   2760
         Width           =   6855
      End
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4830
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Tag             =   "抽选返回"
      Top             =   960
      Width           =   9285
      Begin VB.CommandButton cmdMoveIn 
         Caption         =   "抽回(&I)"
         Height          =   350
         Left            =   6960
         TabIndex        =   26
         Top             =   3840
         Width           =   1100
      End
      Begin VB.TextBox txtPati 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   25
         ToolTipText     =   "通过以下方式输入：直接刷卡，-病人ID,*门诊号,.挂号单,+住院号,姓名"
         Top             =   2895
         Width           =   2535
      End
      Begin VB.ComboBox cboPatiType 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2895
         Width           =   1770
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5520
         TabIndex        =   20
         ToolTipText     =   "请输入完整的单据号"
         Top             =   1545
         Width           =   2535
      End
      Begin VB.ComboBox cboBillType 
         Height          =   300
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1545
         Width           =   1770
      End
      Begin VB.OptionButton optInType 
         Caption         =   "按某病人抽选返回(仅含病人相关诊疗数据和未结帐费用)"
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   21
         Top             =   2535
         Width           =   5100
      End
      Begin VB.OptionButton optInType 
         Caption         =   "按单据号抽选返回"
         Height          =   195
         Index           =   0
         Left            =   1605
         TabIndex        =   13
         Top             =   1185
         Value           =   -1  'True
         Width           =   5100
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   4980
         TabIndex        =   24
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
         Height          =   180
         Left            =   2160
         TabIndex        =   22
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   4800
         TabIndex        =   19
         Top             =   1605
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
         Height          =   180
         Left            =   2160
         TabIndex        =   17
         Top             =   1605
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "    通常已经转出的数据是不再操作的，只能查询。但在一些特殊的情况下，可以抽选某些特定的数据返回在线数据表，以便实施必要的操作。"
         Height          =   540
         Left            =   1680
         TabIndex        =   35
         Top             =   405
         Width           =   6195
      End
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Index           =   3
      Left            =   0
      TabIndex        =   51
      Tag             =   "转移日志"
      Top             =   960
      Width           =   9255
      Begin VSFlex8Ctl.VSFlexGrid vsflog 
         Height          =   4800
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Width           =   9015
         _cx             =   1995324957
         _cy             =   1995317523
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
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDataMove.frx":1632
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
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Tag             =   "未转查询"
      Top             =   915
      Width           =   9165
      Begin VB.CommandButton cmdData 
         Caption         =   "体检任务数据(&P)"
         Height          =   350
         Index           =   4
         Left            =   7050
         TabIndex        =   53
         Top             =   4080
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "直接收费数据(&A)"
         Height          =   350
         Index           =   0
         Left            =   7050
         TabIndex        =   41
         Top             =   1065
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "结帐费用数据(&B)"
         Height          =   345
         Index           =   1
         Left            =   7050
         TabIndex        =   40
         Top             =   1815
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "门诊就诊数据(&L)"
         Height          =   350
         Index           =   2
         Left            =   7050
         TabIndex        =   39
         Top             =   2565
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "住院就诊数据(&P)"
         Height          =   350
         Index           =   3
         Left            =   7050
         TabIndex        =   38
         Top             =   3330
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2940
         TabIndex        =   28
         Top             =   615
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   39452675
         CurrentDate     =   38471
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1245
         TabIndex        =   27
         Top             =   600
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   39452675
         CurrentDate     =   38471
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   480
         X2              =   6375
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label lblData 
         Caption         =   "查询无法转出体检任务数据及具体原因。"
         Height          =   210
         Index           =   4
         Left            =   465
         TabIndex        =   54
         Top             =   4170
         Width           =   4575
      End
      Begin VB.Label lblData 
         Caption         =   "查询无法转出的已结帐门诊、住院记帐，自动记帐的单据数据及具体原因。"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   44
         Top             =   1890
         Width           =   6015
      End
      Begin VB.Label lblData 
         Caption         =   "查询无法转出医疗数据的门诊就诊病人信息及具体原因。"
         Height          =   225
         Index           =   2
         Left            =   480
         TabIndex        =   43
         Top             =   2625
         Width           =   4575
      End
      Begin VB.Label lblData 
         Caption         =   "查询无法转出医疗数据的住院病人信息及具体原因。"
         Height          =   210
         Index           =   3
         Left            =   480
         TabIndex        =   42
         Top             =   3405
         Width           =   4575
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   480
         X2              =   6360
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   480
         X2              =   6360
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   480
         X2              =   6360
         Y1              =   2865
         Y2              =   2865
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   480
         X2              =   6360
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询时间                至"
         Height          =   180
         Left            =   495
         TabIndex        =   37
         Top             =   675
         Width           =   2340
      End
      Begin VB.Label lblData 
         Caption         =   "查询无法转出的门诊挂号，收费的单据数据及具体原因。"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   45
         Top             =   1155
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "请输入合适的查询时间范围。结束时间应在已转移的时间范围内，结束时间的设置可能影响分析无法转出的原因。"
         Height          =   405
         Left            =   135
         TabIndex        =   36
         Top             =   240
         Width           =   9240
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9390
      TabIndex        =   47
      Top             =   5925
      Width           =   9390
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   6600
         TabIndex        =   49
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "退出(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   48
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.TabStrip tabFunc 
         Height          =   345
         Left            =   120
         TabIndex        =   50
         Tag             =   "转移查询"
         Top             =   165
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   609
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         Style           =   2
         TabFixedWidth   =   2027
         TabFixedHeight  =   616
         Placement       =   1
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "数据转移(&1)"
               Key             =   "数据转移"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "抽选返回(&2)"
               Key             =   "抽选返回"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "未转查询(&3)"
               Key             =   "未转查询"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "转移日志(&4)"
               Key             =   "转移日志"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "转后处理(&5)"
               Key             =   "转后处理"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   0
         X2              =   10500
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   10500
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9390
      TabIndex        =   29
      Top             =   0
      Width           =   9390
      Begin MSComctlLib.ImageList img48 
         Left            =   -375
         Top             =   -330
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":1750
               Key             =   "数据转移"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":262A
               Key             =   "抽选返回"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":3504
               Key             =   "未转查询"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":43DE
               Key             =   "转移日志"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":E473
               Key             =   "转后处理"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据转移"
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
         Left            =   1425
         TabIndex        =   34
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    为保持系统高效运行、减少备份数据量、缩短重建索引和统计信息收集等在线空间维护的时间，建议定期将历史数据转移到历史空间中。"
         Height          =   360
         Left            =   1425
         TabIndex        =   33
         Top             =   390
         Width           =   8025
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   240
         Picture         =   "frmDataMove.frx":11905
         Top             =   60
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11040
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   11040
         Y1              =   840
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmDataMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mrsMovelog As ADODB.Recordset   '用于转出操作时按标记转出的批次进行转出处理
Private mdatBegin As Date
Private mstrPeisPrivs As String         '体检子系统的数据转移权限
Private mlngPeisSys As Long             '体检子系统编号
Private mlngOperSys As Long             '手麻子系统编号
Private mblnDBA As Boolean
Private mlngMinDays As Long, mlngMaxDays As Long
Private mblnOffLineMoved As Boolean            '是否执行了转出操作并且没有恢复在线空间的约束和索引

Private Sub cboBakspace_Click(Index As Integer)
    If Index = 0 And Me.Visible Then
        Dim strText As String
        Dim i As Long, j As Long
        
        strText = cboBakspace(Index).Text
        For i = 1 To 2
            For j = 0 To cboBakspace(i).ListCount
                If cboBakspace(i).List(j) = strText Then
                    cboBakspace(i).ListIndex = j
                    Exit For
                End If
            Next
        Next
    End If
End Sub

Private Sub cboBillType_Click()
    txtNO.Text = ""
End Sub

Private Sub cboPatiType_Click()
    txtPati.Text = ""
    txtPati.Tag = ""
    cboPatiType.Tag = ""
    
    Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
    Case 0, 1
        txtPati.ToolTipText = "通过以下方式输入：直接刷卡，-病人ID,*门诊号,.挂号单,+住院号,姓名"
        lblPatient.Caption = "病人"
    Case 2
        txtPati.ToolTipText = "通过以下方式输入：直接刷卡，-病人ID,*门诊号,+健康号,姓名"
        lblPatient.Caption = "病人"
    Case 3
        txtPati.ToolTipText = "通过以下方式输入：-团体ID,姓名"
        lblPatient.Caption = "团体"
    End Select
    
End Sub

Private Sub chkAffirm_Click()
    If txtPrompt.Visible Then txtPrompt.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdData_Click(Index As Integer)
    If dtpBegin.value > dtpEnd.value Then
        MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
        dtpBegin.SetFocus: Exit Sub
    End If
    
    If MsgBox("如果指定时间中的未转出数据较多，查询可能需要较长时间。" & vbCrLf & "继续执行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call frmDataMoveQuery.ShowMe(Index, dtpBegin, dtpEnd, Split(cmdData(Index).Caption, "(")(0), lblData(Index).Caption, Me)
End Sub

Private Sub cmdDateThis_Click()
'功能：打开日期选择器
        If IsDate(txtDateThis.Text) Then monSel.value = CDate(txtDateThis.Text)
        
        monSel.Tag = "txtDateThis"
        monSel.Left = Me.ScaleLeft + Me.ScaleWidth - monSel.Width - 120
        monSel.Top = txtDateThis.Top + txtDateThis.Height + 30
        monSel.ZOrder
        monSel.Visible = True
        monSel.SetFocus
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdMoveIn_Click()
    Dim lng病人ID As Long, str就诊 As String
    Dim blnMoved As Boolean
    
    If optInType(0).value Then
        If txtNO.Text = "" Then
            MsgBox "请输入单据号。", vbInformation, gstrSysName
            txtNO.SetFocus: Exit Sub
        End If
        
        '   "1-收费单据","2-记帐单据","3-自动记帐","4-挂号单据","5-就诊卡","6-预交单据","7-结帐单据"
        '检查单据是否已经转出
        Select Case cboBillType.ItemData(cboBillType.ListIndex)
        Case 2 '记帐单据(可能存在门诊记帐和住院记帐的情况,所以需要访问两个表)
            blnMoved = MovedByNO(txtNO.Text, "病人费用记录", "记录性质=[2] ")
        Case 3, 5           ',自动记帐,就诊卡
            blnMoved = MovedByNO(txtNO.Text, "住院费用记录", "记录性质=[2] ")
        Case 1, 4   '收费单据,,挂号单据
            blnMoved = MovedByNO(txtNO.Text, "门诊费用记录", "记录性质=[2] ")
        Case 6 '预交单据
            blnMoved = MovedByNO(txtNO.Text, "病人预交记录", "记录性质=1")
        Case 7 '结帐单据
            blnMoved = MovedByNO(txtNO.Text, "病人结帐记录")
        Case 8  '体检任务
            blnMoved = MovedByPeis(1, txtNO.Text)
        End Select
        
        If Not blnMoved Then
            MsgBox Replace(Mid(cboBillType.Text, 3), "单据", "") & "单据 " & txtNO.Text & " 没有转出。", vbInformation, gstrSysName
            txtNO.SetFocus: Exit Sub
        End If
        
        If MsgBox("现在将把" & Replace(Mid(cboBillType.Text, 3), "单据", "") & "单据 " & txtNO.Text & " 的数据抽回在线数据库，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    ElseIf optInType(1).value Then
        If txtPati.Tag = "" Then
            MsgBox "请输入病人。", vbInformation, gstrSysName
            txtPati.SetFocus: Exit Sub
        End If
        
        Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
        Case 0, 1
            lng病人ID = Val(Split(txtPati.Tag, ",")(0))
            str就诊 = CStr(Split(txtPati.Tag, ",")(1))
            
            Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
            Case 0          '门诊病人(病人id,挂号单号)
                blnMoved = MovedByNO(str就诊, "病人挂号记录")
            Case 1          '住院病人(病人id,主页id)
                blnMoved = MovedByPati(lng病人ID, Val(str就诊))
            End Select
            
        Case 2         '受检人员(病人id)
            lng病人ID = Val(txtPati.Tag)
            blnMoved = MovedByPeis(2, Val(txtPati.Tag))
        Case 3         '受检团体(团体id)
            lng病人ID = Val(txtPati.Tag)
            blnMoved = MovedByPeis(3, Val(txtPati.Tag))
            
        End Select
        
        If Not blnMoved Then
            MsgBox Mid(cboPatiType.Text, 3) & " " & txtPati.Text & "的相关数据没有转出。", vbInformation, gstrSysName
            txtPati.SetFocus: Exit Sub
        End If
        
        If MsgBox("现在将把" & Mid(cboPatiType.Text, 3) & " " & txtPati.Text & " 的相关诊疗数据和未结帐费用抽回在线数据库，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
   
    End If
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    If optInType(0).value Then
        Select Case cboBillType.ItemData(cboBillType.ListIndex)
        Case 1, 2, 3, 4, 5, 6, 7
            gstrSQL = "Zl_Retu_Exes('" & txtNO.Text & "'," & cboBillType.ItemData(cboBillType.ListIndex) & ")"
        Case 8
            gstrSQL = "zl_Return_Peis(3,'" & txtNO.Text & "')"
        End Select
        
    ElseIf optInType(1).value Then
        Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
        Case 0, 1
            gstrSQL = "Zl_Retu_Clinic(" & lng病人ID & ",'" & str就诊 & "'," & cboPatiType.ItemData(cboPatiType.ListIndex) & ")"
        Case 2
            gstrSQL = "zl_Return_Peis(1,'" & lng病人ID & "')"
        Case 3
            gstrSQL = "zl_Return_Peis(2,'" & lng病人ID & "')"
        End Select
    Else
        gstrSQL = "Zl_Retu_Clinic(0,'" & str就诊 & "',2)"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    MsgBox "数据抽选过程已执行完成。", vbInformation, gstrSysName
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RefreshMove() As Boolean
'功能：刷新转出天数等信息
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strMsg As String, strTagStartDate As String
    Dim datCurr As Date, blnFirst As Boolean, blnWaitMove As Boolean, blnWaitTag As Boolean
    Dim blnDo As Boolean
    Dim lngTmpSysNO As Long, lngDays As Long
        
    On Error GoTo errH
    
   
     '进入窗体时设置缺省并行度
    If Me.Visible = False Then
        If mblnDBA Then
            gstrSQL = "Select Value From V$parameter Where Name = 'cpu_count'"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
            If rsTmp.EOF Then
                txtParallel.Text = "0"
                txtParallel.Enabled = False
                lblParallel.Caption = "DDL操作并行度        (未找到并行度参数cpu_count)"
            Else
                txtParallel.Tag = "" & rsTmp!value
                If Val(rsTmp!value) < 3 Then
                    txtParallel.Text = "0"
                    txtParallel.Enabled = False
                    lblParallel.Caption = "DDL操作并行度        (cpu个数小于3，不必采用并行执行)"
                    
                ElseIf Val(rsTmp!value) < 13 Then
                    txtParallel.Text = Val(rsTmp!value) \ 2 '一半取整
                Else
                    txtParallel.Text = "12"  '即使cpu足够，但仍可能受限于磁盘性能，并行度并非越大越好
                End If
            End If
        Else
            txtParallel.Text = "0"
        End If
                
        
        '历史表空间的初始化
        gstrSQL = "Select 系统, 编号, 名称, 所有者, 当前 From Zlbakspaces "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If rsTmp.EOF Then
            MsgBox "系统中未发现历史表空间，本功能不能使用。", vbExclamation, gstrSysName
            Exit Function
        End If
        For i = 0 To 2
            cboBakspace(i).Clear
            lngTmpSysNO = Decode(i, 0, glngSys, 1, mlngPeisSys, 2, mlngOperSys)
            If lngTmpSysNO > 0 Then
                rsTmp.Filter = "系统=" & lngTmpSysNO
                rsTmp.Sort = "编号"
                Do While Not rsTmp.EOF
                    cboBakspace(i).AddItem NVL(rsTmp!名称)
                    cboBakspace(i).ItemData(cboBakspace(i).NewIndex) = Val(NVL(rsTmp!编号))
                    If NVL(rsTmp!当前, 0) = 0 Then cboBakspace(i).ListIndex = cboBakspace(i).NewIndex
                    rsTmp.MoveNext
                Loop
                If cboBakspace(i).ListCount > 0 And cboBakspace(i).ListIndex < 0 Then cboBakspace(i).ListIndex = 0
            End If
            cboBakspace(i).Visible = cboBakspace(i).ListCount > 1
            lblBakSpace(i).Visible = cboBakspace(i).ListCount > 1
            '根据显示调整位置
            If (Not cboBakspace(i).ListCount > 1) And i < 2 Then
                If i = 0 Then
                    cboBakspace(2).Top = cboBakspace(1).Top
                    lblBakSpace(2).Top = lblBakSpace(1).Top
                End If
                cboBakspace(i + 1).Top = cboBakspace(i).Top
                lblBakSpace(i + 1).Top = lblBakSpace(i).Top
            End If
        Next i
        picBakspace.Visible = False
        For i = 0 To 2
            If cboBakspace(i).ListCount > 1 Then
                picBakspace.Visible = True
                Exit For
            End If
        Next i
    End If
        
    gstrSQL = "Select 上次日期,本次最终日期 From zlDataMove Where 系统=[1] And 组号=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
    If rsTmp.EOF Then
        MsgBox "系统中未发现有效的数据转移定义，本功能不能使用。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If IsNull(rsTmp!本次最终日期) = False Then
        txtDateLast.Text = Format(rsTmp!本次最终日期, "yyyy-mm-dd")
        txtDateLast.Enabled = False
    Else
        txtDateLast.Enabled = True
    End If
    cmdDateLast.Enabled = txtDateLast.Enabled
    
    gstrSQL = "Select 批次,序列,截止时间,待转出,标记开始时间,标记结束时间,转出开始时间,转出结束时间,重建结束时间" & _
            " From zlDataMovelog Where 系统 = [1] Order by 批次"
    Set mrsMovelog = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
    If mrsMovelog.RecordCount > 0 Then
        '如果存在标记转出出错的，需对该时间段重新标记转出，否则转出后会导致数据不一致。
        mrsMovelog.Filter = "待转出=2"
        If mrsMovelog.RecordCount > 0 Then
            mdatBegin = mrsMovelog!截止时间
            blnWaitTag = True
        Else
            mrsMovelog.Filter = "待转出=1"
            If mrsMovelog.RecordCount > 0 Then
                mrsMovelog.MoveLast
                mdatBegin = mrsMovelog!截止时间
                blnWaitMove = True
            End If
            mrsMovelog.Filter = ""
            mrsMovelog.MoveFirst
        End If
    End If
    
    
    '标记转出后，没有更新上次日期
    If IsNull(rsTmp!上次日期) Or blnWaitTag Then
        If Not blnWaitMove And Not blnWaitTag Then blnFirst = True
        
        If blnWaitTag Then
            '取上一次标记转出或转出的时间
            gstrSQL = "Select 截止时间 as 上次日期 From zlDataMovelog Where 系统 = [1] And Nvl(待转出,0)<>2 Order by 截止时间 Desc "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
            blnDo = rsTmp.RecordCount = 0
        Else
            blnDo = True
        End If
        
        If blnDo Then
            gstrSQL = "Select Min(登记时间) 上次日期 From (Select Min(登记时间) 登记时间 From 门诊费用记录 Union All Select Min(登记时间) From 病人挂号记录)"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
            
            If IsNull(rsTmp!上次日期) Then
                MsgBox "当前系统没有发生门诊挂号或收费业务数据，本功能不能使用。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    strTagStartDate = Format(rsTmp!上次日期, "yyyy-MM-dd")
    txtDatePre.Text = strTagStartDate
    
    If blnWaitTag Then
        lblDateThis.Caption = "继续标记日期"
    ElseIf blnWaitMove Then
        lblDateThis.Caption = "继续转出日期"
    Else
        mdatBegin = rsTmp!上次日期
        lblDateThis.Caption = "本次截止日期"
    End If
   
    datCurr = zlDatabase.Currentdate
    
    
    If blnWaitTag Then
        mlngMaxDays = DateDiff("d", rsTmp!上次日期, datCurr)
    Else
        mlngMaxDays = DateDiff("d", mdatBegin, datCurr)
    End If
    
    mlngMinDays = 365   '至少保留一年的数据
    If mlngMinDays > mlngMaxDays Then mlngMinDays = mlngMinDays - 1
        
    If blnWaitTag Or blnWaitMove Then
        cmdMoveOut.Enabled = blnWaitMove
        cmdMoveMark.Enabled = blnWaitTag
        
        txtDateThis.Text = Format(mdatBegin, "yyyy-MM-dd")
        txtDateThis.Enabled = False
    Else
        cmdMoveOut.Enabled = True
        cmdMoveMark.Enabled = True
        
        txtDateThis.Enabled = True
    
        '缺省的最终截止日期为保留三年数据
        If txtDateLast.Enabled Then
            gstrSQL = "Select Trunc(add_months(Sysdate,-24*3) ,'yyyy') As year_firstday From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            lngDays = DateDiff("d", Format(rsTmp!year_firstday, "yyyy-mm-dd"), datCurr)
        
            If mlngMaxDays < lngDays Then lngDays = mlngMaxDays - 365
            If lngDays < 0 Then lngDays = mlngMaxDays - 180
            If lngDays < 0 Then lngDays = 0
            
            txtDateLast.Text = Format(datCurr - lngDays, "yyyy-mm-dd")
        End If
        
        '缺省一次转一年
        txtDateThis.Text = Format(datCurr - mlngMaxDays + 365, "yyyy-mm-dd")
        If CDate(txtDateThis.Text) > CDate(txtDateLast.Text) Then txtDateThis.Text = txtDateLast.Text
    End If
    cmdDateThis.Enabled = txtDateThis.Enabled
    
            
    If blnFirst Then
        strMsg = "从存在挂号或收费数据的 " & Format(mdatBegin, "yyyy-MM-dd") & " 开始转移数据"
        dtpEnd.MaxDate = Int(DateAdd("d", -90, datCurr) - 1)
    Else
        If blnWaitMove Then
            strMsg = "已经标记了 " & strTagStartDate & " 到 " & Format(mdatBegin, "yyyy-MM-dd") & " 之间的数据" & vbCrLf & "须先转出这些数据后才能执行新的转出操作。"
        ElseIf blnWaitTag Then
            strMsg = "标记 " & strTagStartDate & " 到 " & Format(mdatBegin, "yyyy-MM-dd") & " 之间的数据时发生中断" & vbCrLf & "须继续标记转出这些数据后才能执行新的操作。"
        Else
            strMsg = "上次已经转出了 " & Format(mdatBegin, "yyyy-MM-dd") & " 以前的数据"
        End If
        dtpEnd.MaxDate = Int(mdatBegin - 1)
    End If
    lblStatus.Caption = strMsg
    
    
    '设置未转查询
    dtpBegin.MaxDate = dtpEnd.MaxDate
    If Not Visible Then
        dtpEnd.value = dtpEnd.MaxDate
        dtpBegin.value = DateAdd("d", -30, dtpEnd.value)
    End If
    For i = 0 To cmdData.UBound
        cmdData(i).Enabled = Not blnFirst
    Next
    dtpBegin.Enabled = Not blnFirst
    dtpEnd.Enabled = Not blnFirst
        
    RefreshMove = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAffirm() As Boolean
    If chkAffirm.value = 0 Then
        MsgBox "请确认你已仔细阅读并完成了历史数据转出前须进行的准备和调整。", vbInformation, gstrSysName
        chkAffirm.SetFocus
    Else
        CheckAffirm = True
    End If
End Function

Private Sub cmdMoveMark_Click()
'功能：执行标记转出
    Dim datCurr As Date, datBegin As Date, strTime As String, lngTotaltime As Long
    Dim lngBeginDays As Long, i As Long, lngEndDays As Long, lngCurrDays As Long, bytSpeedMode As Byte
    Dim lngSplit As Long, lngAddDay As Long
    Dim rsTmp As ADODB.Recordset
    
    Dim strBakUser As String, strPeisBakUser As String, strOperBakUser As String, blnNoData As Boolean
    
    If CheckAffirm = False Then Exit Sub
    If Not CheckDate(1) Then Exit Sub
          
    If Not IsNumeric(txtSplit.Text) Then
        MsgBox "请输入有效的间隔天数。", vbInformation, gstrSysName
        txtSplit.SetFocus: Exit Sub
    End If
    lngSplit = Val(txtSplit.Text)
    
    If CheckData = False Then Exit Sub
    
    
    If MsgBox("如果标记转出的数据较多，可能需要较长时间。" & vbCrLf & "你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
    On Error GoTo errH
    
    blnNoData = cboBakspace(0).ListCount = 0
    If blnNoData = False Then
        gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys, cboBakspace(0).ItemData(cboBakspace(0).ListIndex))
        blnNoData = rsTmp.RecordCount = 0
    End If
    If blnNoData = True Then
        MsgBox "在表zlBakSpaces中未找到当前历史空间！", vbInformation, gstrSysName
        Exit Sub
    End If
    strBakUser = rsTmp!所有者
    
    
     '体检子系统的判断
    If mlngPeisSys > 0 Then
        blnNoData = cboBakspace(1).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPeisSys, cboBakspace(1).ItemData(cboBakspace(1).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "在表zlBakSpaces中未找到当前体检子系统历史空间！", vbInformation, gstrSysName
            Exit Sub
        End If
        strPeisBakUser = rsTmp!所有者
    End If
    
    '手麻子系统的判断
    If mlngOperSys > 0 Then
       blnNoData = cboBakspace(2).ListCount = 0
       If blnNoData = False Then
           gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
           Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOperSys, cboBakspace(2).ItemData(cboBakspace(2).ListIndex))
           blnNoData = rsTmp.RecordCount = 0
       End If
        If blnNoData = True Then
           MsgBox "在表zlBakSpaces中未找到当前手麻子系统历史空间！", vbInformation, gstrSysName
           Exit Sub
       End If
       strOperBakUser = rsTmp!所有者
    End If
    
    
    datBegin = zlDatabase.Currentdate
    
    lngEndDays = DateDiff("d", CDate(txtDateThis.Text), datBegin)
    lngBeginDays = mlngMaxDays
    bytSpeedMode = IIF(optmode(0).value, 0, 1)
    
    
    If (lngBeginDays - lngEndDays) Mod lngSplit = 0 Then
        lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit
    Else
        lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit + 1
    End If
    
    Screen.MousePointer = 11
    TIMStatus.Enabled = True    '启用定时刷新以显示进度
    TIMStatus.Tag = "标记转出"
    SetCommandEnable False
    
    If txtDateLast.Enabled Then
        gcnOracle.Execute "Update zlDataMove Set 本次最终日期 = to_date('" & txtDateLast.Text & "','yyyy-mm-dd')"
    End If
    
    For i = 1 To lngTotaltime
        datCurr = zlDatabase.Currentdate
        lngAddDay = DateDiff("d", datBegin, datCurr)    '转出期间可能跨天
        
        lngBeginDays = lngBeginDays - lngSplit
        lngCurrDays = IIF(lngBeginDays > lngEndDays, lngBeginDays, lngEndDays) + lngAddDay
        
        lblStatus.Caption = "正在标记" & Format(DateAdd("d", -lngCurrDays, datCurr), "yyyy-MM-dd") & "前的数据(" & i & "/" & lngTotaltime & ")，请耐心等待 … …"
        lblStatus.Refresh
        gstrSQL = "zl1_DataMoveOut1(" & lngCurrDays & ",1," & i & "," & lngTotaltime & "," & bytSpeedMode & "," & chkTrigger.value & "," & chkjob.value & "," & _
                     "0,'" & strBakUser & "','" & strPeisBakUser & "','" & strOperBakUser & "')"
        DoEvents
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    
    strTime = GetTimeString(datBegin, zlDatabase.Currentdate)
    Screen.MousePointer = 0
    MsgBox "标记转出执行完成，本次共耗时：" & strTime & "。", vbInformation, gstrSysName
    
    Call RefreshMove
    
    Exit Sub
    
errH:
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
'        Screen.MousePointer = 11
'        Resume
    End If
    Call SaveErrLog
    Call RefreshMove
End Sub

Private Function CheckPrivilegeOfTrigger(ByRef cnThis As ADODB.Connection) As Boolean
'功能：检查当前连接对象的用户，是否有创建触发器的权限
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select 1 From User_Sys_Privs Where Privilege in ('CREATE TRIGGER','CREATE ANY TRIGGER')"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
    
    CheckPrivilegeOfTrigger = rsTmp.RecordCount > 0
End Function


Private Function GrantPrivilegeOfTrigger() As Boolean
'功能：对应用系统的所有者，检查并执行创建触发器的授权
    Dim rsTmp As ADODB.Recordset
    Dim cnDBA As ADODB.Connection
    Dim strOwner As String
    
        
    If CheckPrivilegeOfTrigger(gcnOracle) = False Then
        Call zlDatabase.UserIdentify(Me, "使用DBA用户向应用系统的所有者授予创建触发器的权限。", glngSys, 0, "system", cnDBA, True)
        If cnDBA Is Nothing Then
            MsgBox "用户登录失败，放弃当前操作！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '取应用系统的所有者(手麻和体检系统是标准版的子系统，是相同的所有者)
        gstrSQL = "Select Trunc(编号 / 100) as 编号, 所有者 From zlSystems Where Trunc(编号 / 100) = 1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rsTmp.RecordCount = 0 Then
            MsgBox "从ZLSYSTEMS获取系统信息失败！", vbInformation, gstrSysName
            Exit Function
        End If
        strOwner = rsTmp!所有者
        
        cnDBA.Execute "grant create trigger to " & strOwner
        cnDBA.Close
    End If
    
    GrantPrivilegeOfTrigger = True
End Function


Private Sub cmdMoveOut_Click()
'功能：执行转出
    Dim datCurr As Date, datBegin As Date, strTime As String, lngTotaltime As Long
    Dim strMsg As String
    Dim i As Long, lngAddDay As Long
    Dim strBakUser As String, strPeisBakUser As String, strOperBakUser As String
    Dim cnBakDB As ADODB.Connection, cnPeisBakDB As New ADODB.Connection, cnOperBakDB As New ADODB.Connection
        
    Dim rsTmp As ADODB.Recordset
    Dim blnRollBack As Boolean, lngTag As Long
    Dim lngBeginDays As Long, lngEndDays As Long, lngCurrDays As Long, bytSpeedMode As Byte
    Dim lngSplit As Long
    Dim blnNoData As Boolean
    
    On Error GoTo errH
    If CheckAffirm = False Then Exit Sub
    
    
    If Not IsNumeric(txtSplit.Text) Then
        MsgBox "请输入有效的间隔天数。", vbInformation, gstrSysName
        txtSplit.SetFocus: Exit Sub
    End If
    lngSplit = Val(txtSplit.Text)
    
    bytSpeedMode = IIF(optmode(0).value, 0, 1)
    mrsMovelog.Filter = "待转出=1"
    lngTag = mrsMovelog.RecordCount
    
    '对已标记转出的数据执行转出操作时不用检查天数等
    If lngTag = 0 Then
        If Not CheckDate(0) Then Exit Sub
    End If
    
    If CheckData = False Then Exit Sub
    
        
    If bytSpeedMode = 1 Then
        strMsg = "你选择了离线模式，数据转出期间将会禁用部分索引和约束，这将导致本系统的所有客户端不可用。" & vbCrLf & _
                "你确定要继续吗？"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If lngTag > 0 Then
            strMsg = "将转出所有已标记的数据," & vbCrLf & "如果数据较多，可能需要较长时间。" & vbCrLf & "你确定要继续吗？"
        Else
            strMsg = "如果转出数据较多，可能需要较长时间。" & vbCrLf & "你确定要继续吗？"
        End If
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    
    
    '禁用历史空间的索引与约束（在线库的处理在zl1_DataMoveOut1中进行）,即使在线模式也建议禁用，提高插入性能，以及避免产生大量日志
    blnNoData = cboBakspace(0).ListCount = 0
    If blnNoData = False Then
        gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys, cboBakspace(0).ItemData(cboBakspace(0).ListIndex))
        blnNoData = rsTmp.RecordCount = 0
    End If
    If blnNoData = True Then
        MsgBox "在表zlBakSpaces中未找到当前历史空间！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strBakUser = rsTmp!所有者
    
    If chkBakTbsDisable.value = 1 Then
        Call zlDatabase.UserIdentify(Me, "历史空间用户验证", glngSys, 0, strBakUser, cnBakDB, True)
        If cnBakDB Is Nothing Then
            MsgBox "转出前需要先禁用历史空间的约束和索引，必须先连接历史空间！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '体检子系统的判断
    If mlngPeisSys > 0 Then
        blnNoData = cboBakspace(1).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPeisSys, cboBakspace(1).ItemData(cboBakspace(1).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "在表zlBakSpaces中未找到当前体检子系统历史空间！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strPeisBakUser = rsTmp!所有者
        
        If chkBakTbsDisable.value = 1 Then
            If strBakUser = strPeisBakUser Then
                Set cnPeisBakDB = cnBakDB
            Else
                Call zlDatabase.UserIdentify(Me, "体检子系统历史空间用户验证", mlngPeisSys, 0, strPeisBakUser, cnPeisBakDB, True)
                If cnPeisBakDB Is Nothing Then
                    MsgBox "转出前需要先禁用体检子系统历史空间的约束和索引，必须先连接体检子系统历史空间！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '手麻子系统的判断
    If mlngOperSys > 0 Then
        blnNoData = cboBakspace(2).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOperSys, cboBakspace(2).ItemData(cboBakspace(2).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
         If blnNoData = True Then
            MsgBox "在表zlBakSpaces中未找到当前手麻子系统历史空间！", vbInformation, gstrSysName
            Exit Sub
        End If
        strOperBakUser = rsTmp!所有者
        
        If chkBakTbsDisable.value = 1 Then
            If strBakUser = strOperBakUser Then
                Set cnOperBakDB = cnBakDB
            Else
                Call zlDatabase.UserIdentify(Me, "手麻子系统历史空间用户验证", mlngOperSys, 0, strOperBakUser, cnOperBakDB, True)
                If cnOperBakDB Is Nothing Then
                    MsgBox "转出前需要先禁用手麻子系统历史空间的约束和索引，必须先连接手麻子系统历史空间！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
         
    '检查并授予系统所有者用户创建和删除触发器的权限，以便在线库禁用约束时，为级联删除外键引用的主表创建临时触发器
    '由于是通过动态SQL创建触发器，所以，即使系统所有者是DBA角色，也需要显式授权（即使所属的角色RESOURCE有创建触发器的权限）
    If lngTag = 0 And bytSpeedMode = 0 Then
        If GrantPrivilegeOfTrigger = False Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    datBegin = zlDatabase.Currentdate
           
    If chkBakTbsDisable.value = 1 Then
        lblStatus.Caption = "正在禁用历史空间的约束与索引，请耐心等待 … …"
        Call SetConstraintStatus(glngSys, cnBakDB, False)
        If mlngPeisSys > 0 Then Call SetConstraintStatus(mlngPeisSys, cnPeisBakDB, False)
        If mlngOperSys > 0 Then Call SetConstraintStatus(mlngOperSys, cnOperBakDB, False)
        
        Call SetIndexStatus(glngSys, cnBakDB, False)
        If mlngPeisSys > 0 Then Call SetIndexStatus(mlngPeisSys, cnPeisBakDB, False)
        If mlngOperSys > 0 Then Call SetIndexStatus(mlngOperSys, cnOperBakDB, False)
    End If

        
    TIMStatus.Enabled = True    '启用定时刷新以显示时度
    TIMStatus.Tag = "转出"
    SetCommandEnable False

    
    blnRollBack = True
    If lngTag > 0 Then  'a.根据标记转出进行转出
        For i = 1 To mrsMovelog.RecordCount
            lblStatus.Caption = "正在转出" & Format(mrsMovelog!截止时间, "yyyy-MM-dd") & "之前的数据(" & i & "/" & mrsMovelog.RecordCount & ")，请耐心等待 … …"
            Me.Refresh
            gstrSQL = "zl1_DataMoveOut1(0,2," & i & "," & lngTag & "," & bytSpeedMode & "," & chkTrigger.value & "," & chkjob.value & ",0,'" & strBakUser & "','" & strPeisBakUser & "','" & strOperBakUser & "')"
            DoEvents
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    Else                'b.直接转出
        lngEndDays = DateDiff("d", CDate(txtDateThis.Text), datBegin)
        lngBeginDays = mlngMaxDays
        
        If (lngBeginDays - lngEndDays) Mod lngSplit = 0 Then
            lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit
        Else
            lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit + 1
        End If
        
        For i = 1 To lngTotaltime
            datCurr = zlDatabase.Currentdate
            lngAddDay = DateDiff("d", datBegin, datCurr)    '转出期间可能跨天
            
            lngBeginDays = lngBeginDays - lngSplit
            lngCurrDays = IIF(lngBeginDays > lngEndDays, lngBeginDays, lngEndDays) + lngAddDay
            
            
            lblStatus.Caption = "正在转出" & Format(DateAdd("d", -lngCurrDays, datCurr), "yyyy-MM-dd") & "前的数据(" & i & "/" & lngTotaltime & ")，请耐心等待 … …"
            lblStatus.Refresh
            gstrSQL = "zl1_DataMoveOut1(" & lngCurrDays & ",0," & i & "," & lngTotaltime & "," & bytSpeedMode & "," & chkTrigger.value & "," & chkjob.value & ",0,'" & strBakUser & "','" & strPeisBakUser & "','" & strOperBakUser & "')"
            DoEvents
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    End If
    blnRollBack = False
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    mblnOffLineMoved = True
               
    If chkBakTbsDisable.value = 1 Then
        cnBakDB.Close
        If mlngPeisSys > 0 And cnPeisBakDB.State = adStateOpen Then cnPeisBakDB.Close
        If mlngOperSys > 0 And cnOperBakDB.State = adStateOpen Then cnOperBakDB.Close
    End If
    
    strTime = GetTimeString(datBegin, zlDatabase.Currentdate)
    Screen.MousePointer = 0
    MsgBox "数据转出执行完成，本次共耗时：" & strTime & "。", vbInformation, gstrSysName
    
    Call RefreshMove
    
    Exit Sub
    
    
errH:
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    If blnRollBack And chkBakTbsDisable.value = 1 Then
        cnBakDB.Close
        If mlngPeisSys > 0 And cnPeisBakDB.State = adStateOpen Then cnPeisBakDB.Close
        If mlngOperSys > 0 And cnOperBakDB.State = adStateOpen Then cnOperBakDB.Close
    End If
    
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        'Screen.MousePointer = 11
        'Resume
    End If
    Call SaveErrLog
    Call RefreshMove
End Sub

Private Function GetTimeString(ByVal datBegin As Date, ByVal datEnd As Date) As String
'功能：获取两个时间值差的格式字符串
'   datBegin=起始时间
'   datEnd=中止时间
    Dim intH As Integer, intM As Integer, intS As Integer
    Dim datTmp As Date

    intH = DateDiff("h", datBegin, datEnd)
    datTmp = DateAdd("h", intH, datBegin)
    intM = DateDiff("n", datTmp, datEnd)
    datTmp = DateAdd("n", intM, datTmp)
    intS = DateDiff("s", datTmp, datEnd)
    
    If intS < 0 Then
        intM = intM - 1
        intS = 60 + intS
    End If
    
    If intM < 0 Then
        intH = intH - 1
        intM = 60 + intM
    End If
    GetTimeString = IIF(intH <> 0, intH & "小时", "") & IIF(intM <> 0, intM & "分", "") & intS & "秒"
End Function

Private Sub cmdPrompt_Click()
    If txtPrompt.Visible = False Then
        txtPrompt.Top = cmdPrompt.Top + cmdPrompt.Height + 30
        txtPrompt.Left = cmdPrompt.Left
        txtPrompt.Height = Me.Height - (fraFunc(0).Top + cmdPrompt.Height + 120) - (PicBottom.Height + 120) - 240 '(滚动条)
        txtPrompt.Width = Me.Width - 120 - 240
        
        txtPrompt.ZOrder
        chkAffirm.value = 0
        txtPrompt.Visible = True
    End If
End Sub

Private Sub cmdRebBakSpace_Click()
'功能：恢复历史空间被禁用的约束和索引
    Dim strBakUser As String
    Dim cnBakDB As ADODB.Connection
    Dim strPeisBakUser As String
    Dim cnPeisBakDB As New ADODB.Connection
    Dim strOperBakUser As String
    Dim cnOperBakDB As New ADODB.Connection
    Dim rsTmp As ADODB.Recordset
    Dim strParallel As String, strTime As String
    Dim datCurr  As Date
    Dim blnNoData As Boolean
    
    
    If MsgBox("该操作非常耗时，你确定要恢复历史空间被禁用的约束及索引吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo errH
    strParallel = Val(txtParallel.Text)
    blnNoData = cboBakspace(0).ListCount = 0
    If blnNoData = False Then
        gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys, cboBakspace(0).ItemData(cboBakspace(0).ListIndex))
        blnNoData = rsTmp.RecordCount = 0
    End If
    If blnNoData = True Then
        MsgBox "在表zlBakSpaces中未找到当前历史空间！", vbInformation, gstrSysName
        Exit Sub
    End If
    strBakUser = rsTmp!所有者
    Call zlDatabase.UserIdentify(Me, "历史空间用户验证", glngSys, 0, strBakUser, cnBakDB, True)
    If cnBakDB Is Nothing Then
        MsgBox "离线模式恢复历史空间的约束和索引，必须先连接历史空间！", vbInformation, gstrSysName
        Exit Sub
    Else
        If Val(strParallel) > 1 Then
            cnBakDB.Execute "ALTER Session FORCE PARALLEL DDL PARALLEL " & strParallel
        Else
            cnBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        End If
    End If
    
    '体检子系统的判断
    If mlngPeisSys > 0 Then
        blnNoData = cboBakspace(1).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPeisSys, cboBakspace(1).ItemData(cboBakspace(1).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "在表zlBakSpaces中未找到当前体检子系统历史空间！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strPeisBakUser = rsTmp!所有者
        
        If strBakUser = strPeisBakUser Then
            Set cnPeisBakDB = cnBakDB
        Else
            Call zlDatabase.UserIdentify(Me, "体检子系统历史空间用户验证", mlngPeisSys, 0, strPeisBakUser, cnPeisBakDB, True)
            If cnPeisBakDB Is Nothing Then
                MsgBox "离线模式恢复体检子系统历史空间的约束和索引，必须先连接体检子系统历史空间！", vbInformation, gstrSysName
                Exit Sub
            Else
                If Val(strParallel) > 1 Then
                    cnPeisBakDB.Execute "ALTER Session FORCE PARALLEL DDL PARALLEL " & strParallel
                Else
                    cnPeisBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
                End If
            End If
        End If
    End If
    
    '手麻子系统的判断
    If mlngOperSys > 0 Then
        blnNoData = cboBakspace(2).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select 所有者 From zlBakSpaces Where 系统 = [1] And 编号 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOperSys, cboBakspace(2).ItemData(cboBakspace(2).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "在表zlBakSpaces中未找到当前手麻子系统历史空间！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strOperBakUser = rsTmp!所有者
        
        If strBakUser = strOperBakUser Then
            Set cnOperBakDB = cnBakDB
        Else
            Call zlDatabase.UserIdentify(Me, "手麻子系统历史空间用户验证", mlngOperSys, 0, strOperBakUser, cnOperBakDB, True)
            If cnOperBakDB Is Nothing Then
                MsgBox "离线模式恢复手麻子系统历史空间的约束和索引，必须先连接手麻子系统历史空间！", vbInformation, gstrSysName
                Exit Sub
            Else
                If Val(strParallel) > 1 Then
                    cnPeisBakDB.Execute "ALTER Session FORCE PARALLEL DDL PARALLEL " & strParallel
                Else
                    cnPeisBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
                End If
            End If
        End If
    End If
    
    
    lblPrompt.Caption = "正在恢复历史空间被禁用的约束与索引，请耐心等待 … …"
    Me.Refresh
    cmdRebBakSpace.Enabled = False
    Me.Enabled = False  '为了doevents
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    
    '启用历史空间的索引与约束（先启用索引，防止非主唯一约束字段与索引字段相同，顺序不同而引发错误）
    Call SetIndexStatus(glngSys, cnBakDB, True)
    If mlngPeisSys > 0 Then Call SetIndexStatus(mlngPeisSys, cnPeisBakDB, True)
    If mlngOperSys > 0 Then Call SetIndexStatus(mlngOperSys, cnOperBakDB, True)
    
    Call SetConstraintStatus(glngSys, cnBakDB, True)
    If mlngPeisSys > 0 Then Call SetConstraintStatus(mlngPeisSys, cnPeisBakDB, True)
    If mlngOperSys > 0 Then Call SetConstraintStatus(mlngOperSys, cnOperBakDB, True)
    
    
    '执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢)
    '取消之前设置的强制并行DDL
    If Val(strParallel) > 1 Then
        Call SetNOParallel(cnBakDB, 0)
        cnBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        
        If mlngPeisSys > 0 Then
            Call SetNOParallel(cnPeisBakDB, 0)
            cnPeisBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        End If
        If mlngOperSys > 0 Then
            Call SetNOParallel(cnOperBakDB, 0)
            cnOperBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        End If
    End If
    
    cnBakDB.Close
    If mlngPeisSys > 0 And cnPeisBakDB.State = adStateOpen Then cnPeisBakDB.Close
    If mlngOperSys > 0 And cnOperBakDB.State = adStateOpen Then cnOperBakDB.Close
    
    Screen.MousePointer = 0
    lblPrompt.Caption = ""
    Me.Enabled = True
    cmdRebBakSpace.Enabled = True
    
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "恢复操作完成，共耗时：" & strTime & "。", vbInformation, gstrSysName
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    Me.Enabled = True
    cmdRebBakSpace.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    lblPrompt.Caption = ""
End Sub

Private Sub cmdRebIndexForTag_Click()
'功能：重建标记转出查询所需的索引
    Dim bytRebScope As Byte
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String
    Dim i As Double
            
    If MsgBox("该操作非常耗时，你确定要重建“标记转出查询”所需的" & IIF(optRebScope_Manual(0).value, optRebScope_Manual(0).Caption, _
        IIF(optRebScope_Manual(1).value, optRebScope_Manual(1).Caption, optRebScope_Manual(2).Caption)) & "索引吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
            
    bytSpeedMode = IIF(optmode_Index(0).value, 0, 1)
    bytRebScope = IIF(optRebScope_Manual(0).value, 0, IIF(optRebScope_Manual(1).value, 1, 2))
    strParallel = Val(txtParallel.Text)
        
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    cmdRebIndexForTag.Enabled = False
    
    On Error GoTo errH
    lblPrompt.Caption = "正在重建“标记转出查询所需”索引。"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 6,1," & strParallel & "," & bytRebScope & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重建标记索引")
    
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "正在重建体检系统“标记转出查询所需”索引。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 6,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建标记索引")
    End If
    
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "正在重建手麻系统“标记转出查询所需”索引。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 6,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建标记索引")
    End If
        
    cmdRebIndexForTag.Enabled = True
    lblPrompt.Caption = ""
    Screen.MousePointer = 0
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "重建操作完成，共耗时：" & strTime & "。", vbInformation, gstrSysName
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    cmdRebIndexForTag.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRebIndexOther_Click()
'功能：重建所有历史数据转出表上，除了标记转出所需索引以外的其他索引
'       用于转出部分或全部数据后，收回这些索引中已删除数据的空闲空间，有利于提高删除数据的效率
    Dim bytRebScope As Byte
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String
    Dim i As Double
            
    If MsgBox("该操作非常耗时，你确定要重建“标记转出查询”所需以外的其他索引吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
            
    bytSpeedMode = IIF(optmode_Index(0).value, 0, 1)
    bytRebScope = IIF(optRebScope_Manual(0).value, 0, IIF(optRebScope_Manual(1).value, 1, 2))
    strParallel = Val(txtParallel.Text)
        
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    cmdRebIndexOther.Enabled = False
    
    On Error GoTo errH
    lblPrompt.Caption = "正在重建“标记转出查询所需”以外的其他索引。"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 8,1," & strParallel & "," & bytRebScope & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重建其他索引")
    
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "正在重建体检系统“标记转出查询所需”以外的其他索引。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 8,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建其他索引")
    End If
    
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "正在重建手麻系统“标记转出查询所需”以外的其他索引。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 8,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建其他索引")
    End If
        
    cmdRebIndexOther.Enabled = True
    lblPrompt.Caption = ""
    Screen.MousePointer = 0
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "重建操作完成，共耗时：" & strTime & "。", vbInformation, gstrSysName
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    cmdRebIndexOther.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRebJobTrigger_Click()
'功能：恢复转出前禁用的后台作业和触发器
    On Error GoTo errH
    
    gstrSQL = "Zl1_Datamove_Reb(100, 0, 1, 1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "恢复触发器")
    
    gstrSQL = "Zl1_Datamove_Reb(100, 0, 2, 1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "恢复自动作业")
       
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRebOnline_Click()
'功能：恢复在线空间的约束及索引
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String
        
    If MsgBox("该操作非常耗时，你确定要恢复在线空间被禁用的约束及索引吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    bytSpeedMode = IIF(optmode(0).value, 0, 1)  '如果转出后，重新选择了模式，可能不准确(因为没有记录上次转出的模式，多次转出的模式可能不一样)
    strParallel = Val(txtParallel.Text)
    
    Screen.MousePointer = 11
    cmdRebOnline.Enabled = False
    datCurr = zlDatabase.Currentdate
    
    On Error GoTo errH
    
    lblPrompt.Caption = "正在重建“非唯一”索引。"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 4,1," & strParallel & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重建索引")
        
    lblPrompt.Caption = "正在恢复“主键和唯一键、外键”"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 3,1," & strParallel & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重建约束")
    
    mblnOffLineMoved = False
        
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "正在重建体检系统的“非唯一”索引。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 4,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建索引")
        
        
        lblPrompt.Caption = "正在恢复体检系统的“主键和唯一键、外键”"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 3,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建约束")
    End If
     
        
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "正在重建手麻系统的“非唯一”索引。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngOperSys & ", " & bytSpeedMode & ", 4,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建索引")
        
        
        lblPrompt.Caption = "正在恢复手麻系统的“主键和唯一键、外键”"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngOperSys & ", " & bytSpeedMode & ", 3,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重建约束")
    End If
        
    cmdRebOnline.Enabled = True
    lblPrompt.Caption = ""
    Screen.MousePointer = 0
    
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "恢复操作完成，共耗时：" & strTime & "。", vbInformation, gstrSysName
        
    Exit Sub
errH:
    Screen.MousePointer = 0
    cmdRebOnline.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdShrink_Click()
'功能：收缩数据文件
    Dim strErr As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, rsSize As ADODB.Recordset
    Dim cnsys As ADODB.Connection, lngBlockSize As Long, lngSumSize As Long
    Dim cmdTmp As New ADODB.Command
        
    On Error GoTo errH
    
    '收缩数据文件（要求以DBA身份执行）
    If mblnDBA = False Then
        Call zlDatabase.UserIdentify(Me, "数据文件收缩", glngSys, 0, "sys", cnsys, True)
        If cnsys Is Nothing Then
            MsgBox "数据文件收缩要求以sys用户连接，请先以该用户连接，或者以其他方式进行收缩！", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Set cnsys = gcnOracle
    End If
    
    cmdShrink.Enabled = False
    Screen.MousePointer = 11
    
    gstrSQL = "select value from v$parameter where name = 'db_block_size'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open gstrSQL, cnsys, adOpenKeyset, adLockReadOnly
    lngBlockSize = Val("" & rsTmp!value)
        
    lblPrompt.Caption = "正在查询待收缩的数据文件。"
    Me.Refresh
    gstrSQL = "Select File_Name,'alter database datafile ''' || Trim(File_Name) || ''' resize ' || Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) || 'm' Cmd" & vbNewLine & _
            "From Dba_Data_Files A, (Select File_Id, Max(Block_Id + Blocks ) Hwm From Dba_Extents Group By File_Id) B" & vbNewLine & _
            "Where a.File_Id = b.File_Id(+) And a.Tablespace_Name Like 'ZL%' And" & vbNewLine & _
            "      Ceil(Blocks * " & lngBlockSize & " / 1024 / 1024) - Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) > 0"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open gstrSQL, cnsys, adOpenKeyset, adLockReadOnly
    
    Screen.MousePointer = 0
    lblPrompt.Caption = ""
    If rsTmp.RecordCount = 0 Then
        Call MsgBox("没有要收缩数据文件！", vbInformation, gstrSysName)
        
        cmdShrink.Enabled = True
        Exit Sub
    Else
        Set cmdTmp.ActiveConnection = cnsys
        cmdTmp.CommandType = adCmdText
        
        If MsgBox("共有" & rsTmp.RecordCount & "个待收缩的数据文件，你确定要收缩这些数据文件吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            If mblnDBA = False Then cnsys.Close
            
            cmdShrink.Enabled = True
            Exit Sub
        End If
    End If
    
    '记录收缩前的总大小
    strSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files Where Tablespace_Name Like 'ZL%'"
    Set rsSize = New ADODB.Recordset
    rsSize.Open strSQL, cnsys, adOpenKeyset, adLockReadOnly
    lngSumSize = rsSize!Mb_Size
    
    On Error Resume Next
    Screen.MousePointer = 11
    strErr = ""
    While Not rsTmp.EOF
        lblPrompt.Caption = "正在收缩：" & rsTmp!File_Name
        Me.Refresh
        DoEvents
        gstrSQL = rsTmp!cmd
        cmdTmp.CommandText = gstrSQL
        cmdTmp.Execute
        If Err.Number <> 0 Then
            strErr = strErr & vbCrLf & rsTmp!cmd & "，错误：" & Err.Description
            Err.Clear
        End If
        
        rsTmp.MoveNext
    Wend
    
    
    '记录收缩后的总大小
    strSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files Where Tablespace_Name Like 'ZL%'"
    Set rsSize = New ADODB.Recordset
    rsSize.Open strSQL, cnsys, adOpenKeyset, adLockReadOnly
    lngSumSize = lngSumSize - rsSize!Mb_Size
    
    If mblnDBA = False Then cnsys.Close
    
    cmdShrink.Enabled = True
    Screen.MousePointer = 0
    lblPrompt.Caption = ""
        
    If strErr <> "" Then
        MsgBox "错误信息：" & strErr, vbInformation, gstrSysName
    Else
        MsgBox "操作完成，共收缩了" & lngSumSize & "M的空间。", vbInformation, gstrSysName
    End If
        
    Exit Sub
errH:
    cmdShrink.Enabled = True
    Screen.MousePointer = 0
    
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdMoveTable_Click()
'功能：重整转出表，并恢复被禁用的索引，及收缩数据文件
'      Move操作必须在离线模式进行
    Dim strMoveScope As String
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String, strErr As String
    Dim rsTmp As ADODB.Recordset
    
    
    If MsgBox("该操作非常耗时，需要中断业务，要求在所有客户端停用的情况下进行。" & vbCrLf & _
        "你确定要重整" & IIF(optMove(0).value, optMove(0).Caption, optMove(1).Caption) & "历史转出表吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    strMoveScope = IIF(optMove(0).value, 0, 1)
    bytSpeedMode = IIF(optmode(0).value, 0, 1) '如果转出后，重新选择了模式，可能不准确(因为没有记录上次转出的模式，多次转出的模式可能不一样)
    strParallel = Val(txtParallel.Text)
        
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    cmdMoveTable.Enabled = False
    
    On Error GoTo errH
    '先move表，并恢复被禁用的索引
    lblPrompt.Caption = "正在重整历史转出表。"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 7,1," & strParallel & "," & strMoveScope & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重整转出表")
    
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "正在重整体检系统的历史转出表。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 7,1," & strParallel & "," & strMoveScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重整转出表")
    End If
        
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "正在重整手麻系统的历史转出表。"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngOperSys & ", " & bytSpeedMode & ", 7,1," & strParallel & "," & strMoveScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重整转出表")
    End If
   
    lblPrompt.Caption = ""
    cmdMoveTable.Enabled = True
    Screen.MousePointer = 0
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "重整操作完成，共耗时：" & strTime & "。", vbInformation, gstrSysName
           
            
    Exit Sub
errH:
    cmdMoveTable.Enabled = True
    Screen.MousePointer = 0
    
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cmdHelp_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    mblnOffLineMoved = False
    mblnDBA = False
    'Dba_Role_Privs是在安装和创建用户时自动进行了授权的
    gstrSQL = "Select Nvl(Count(*), 0) cnt From Sys.Dba_Role_Privs Where Grantee = User And Granted_Role = 'DBA'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        mblnDBA = rsTmp!cnt > 0
    End If
        
    mstrPrivs = gstrPrivs
        
    If InStr(mstrPrivs, "数据转移") = 0 Then
        tabFunc.Tabs.Remove "数据转移"
    End If
    If InStr(mstrPrivs, "数据抽选") = 0 Then
        tabFunc.Tabs.Remove "抽选返回"
    End If
    For i = 1 To tabFunc.Tabs.Count
        tabFunc.Tabs(i).Caption = tabFunc.Tabs(i).Key & "(&" & i & ")"
    Next
    
    mstrPeisPrivs = ""
    mlngPeisSys = 0
    gstrSQL = "Select 编号 From zlSystems Where 编号 Like '21%'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.BOF = False Then
        mlngPeisSys = rsTmp("编号").value
        mstrPeisPrivs = ";" & GetPrivFunc(mlngPeisSys, 2139) & ";"
    End If
    
    cmdData(4).Visible = (InStr(mstrPeisPrivs, "未转数据查询") > 0)
    lblData(4).Visible = (InStr(mstrPeisPrivs, "未转数据查询") > 0)
    Line2(4).Visible = (InStr(mstrPeisPrivs, "未转数据查询") > 0)
    
    mlngOperSys = 0
    gstrSQL = "Select 编号 From zlSystems Where 编号 Like '24%'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.BOF = False Then
        mlngOperSys = rsTmp("编号").value
    End If
    
    
    If Not RefreshMove Then
        Unload Me: Exit Sub
    End If
    
    
    If InStr(mstrPrivs, "数据抽选") > 0 Then
        cboBillType.AddItem "1-收费单据"
        cboBillType.ItemData(cboBillType.NewIndex) = 1
        
        cboBillType.AddItem "2-记帐单据"
        cboBillType.ItemData(cboBillType.NewIndex) = 2
        
        cboBillType.AddItem "3-自动记帐"
        cboBillType.ItemData(cboBillType.NewIndex) = 3
        
        cboBillType.AddItem "4-挂号单据"
        cboBillType.ItemData(cboBillType.NewIndex) = 4
        
        cboBillType.AddItem "5-就诊卡"
        cboBillType.ItemData(cboBillType.NewIndex) = 5
        
        cboBillType.AddItem "6-预交单据"
        cboBillType.ItemData(cboBillType.NewIndex) = 6
        
        cboBillType.AddItem "7-结帐单据"
        cboBillType.ItemData(cboBillType.NewIndex) = 7
        
        If InStr(mstrPeisPrivs, "数据抽选") > 0 Then
            cboBillType.AddItem "8-体检任务"
            cboBillType.ItemData(cboBillType.NewIndex) = 8
        End If
                
        cboBillType.ListIndex = 0
        
        cboPatiType.AddItem "1-门诊病人"
        cboPatiType.ItemData(cboPatiType.NewIndex) = 0
        cboPatiType.AddItem "2-住院病人"
        cboPatiType.ItemData(cboPatiType.NewIndex) = 1
        If InStr(mstrPeisPrivs, "数据抽选") > 0 Then
            cboPatiType.AddItem "3-受检人员"
            cboPatiType.ItemData(cboPatiType.NewIndex) = 2
            cboPatiType.AddItem "4-受检团体"
            cboPatiType.ItemData(cboPatiType.NewIndex) = 3
        End If
        
        cboPatiType.ListIndex = 0
    End If
    
    Call InitLogTable
    
    Call tabFunc_Click
    Exit Sub
errH:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckDate(ByVal bytMode As Byte) As Boolean
'功能：检查转出日期的有效性
'参数：bytMode=0-转出,1-标记转出
'      对标记转出和转出的保留数据天数区别限制，
'      是为了避免将近期的数据标记转出后，又进行业务回退处理，然后已标记的数据执行转出后，影响业务再次处理的正确性。
'      因为涉及的范围太广，应用程序中没有对已标记转出的数据进行操作限制，所以这里通过最小保留时间为3年进行限制
    Dim lngLimitDays As Long, lngDays As Long
    Dim dateCur As Date
    
    If txtDateLast.Enabled Then
        If IsNull(txtDateLast.Text) Then
            MsgBox "请输入有效的日期。", vbInformation, gstrSysName
            txtDateLast.SetFocus: Exit Function
        ElseIf IsDate(txtDateLast.Text) = False Then
            MsgBox "请输入有效的日期。", vbInformation, gstrSysName
            txtDateLast.SetFocus: Exit Function
        End If
    End If
    
    If IsNull(txtDateThis.Text) Then
        MsgBox "请输入有效的日期。", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    ElseIf IsDate(txtDateThis.Text) = False Then
        MsgBox "请输入有效的日期。", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    End If
    
    If CDate(txtDateLast.Text) < CDate(txtDateThis.Text) Then
        MsgBox "最终截止日期不能小于本次截止日期", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    End If
    
    If CDate(txtDateThis.Text) <= CDate(txtDatePre.Text) Then
        MsgBox "本次截止日期应大于上次转出日期", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    End If
    
    dateCur = zlDatabase.Currentdate
    lngDays = DateDiff("d", CDate(txtDateThis.Text), dateCur)
        
    If lngDays < mlngMinDays Then
        MsgBox "本次截止日期不能小于最小日期 " & Format(dateCur - mlngMinDays, "yyyy-mm-dd"), vbInformation, gstrSysName
        If txtDateThis.Enabled Then txtDateThis.SetFocus
        Exit Function
    End If
    
    If lngDays > mlngMaxDays Then
        MsgBox "本次截止日期不能大于最大日期 " & Format(dateCur - mlngMaxDays, "yyyy-mm-dd"), vbInformation, gstrSysName
        If txtDateThis.Enabled Then txtDateThis.SetFocus
        Exit Function
    End If
    
    
    lngLimitDays = IIF(bytMode = 0, 365, 365 * 2)   '标记转出太近的时间，如果不实际转出，容易导致这些数据在标记后被改变
    If lngDays < lngLimitDays Then
        MsgBox IIF(bytMode = 0, "转出操作，", "标记转出操作，") & "要求在线库必须保留至少" & lngLimitDays & " 天的数据。" & vbCrLf & _
                "保留天数不足，不能进行" & IIF(bytMode = 0, "转出操作。", "标记转出操作。"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckDate = True
End Function

Private Function CheckData() As Boolean
'功能：进行相关的数据逻辑检查
    Dim strMsg As String, i As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    CheckData = True
    
    '1.检查索引的并行度
    strMsg = ""
    strSQL = "Select Index_Name, Degree" & vbNewLine & _
            "From All_Indexes" & vbNewLine & _
            "Where Degree Not In ('0', '1') And Owner = Zl_Owner And Table_Name In (Select 表名 From zlBakTables)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        strMsg = strMsg & "," & rsTmp!Index_Name & "(" & Trim(rsTmp!degree) & ")"
        rsTmp.MoveNext
    Next
    If strMsg <> "" Then
        If MsgBox("以下索引设置了并行度：" & vbCrLf & Mid(strMsg, 2) & _
            vbCrLf & vbCrLf & "这样可能导致执行计划评估错误，将会严重影响标记转出操作的性能，强烈建议取消这些索引的并行度属性。" & _
            vbCrLf & "你确定要继续吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckData = False
            Exit Function
        End If
    End If
    
    '2.检查表的并行度
    strMsg = ""
    strSQL = "Select Table_Name, Degree" & vbNewLine & _
            "From All_Tables" & vbNewLine & _
            "Where Degree != ('         1') And Owner = Zl_Owner And Table_Name In (Select 表名 From zlBakTables)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        strMsg = strMsg & "," & rsTmp!Table_name & "(" & Trim(rsTmp!degree) & ")"
        rsTmp.MoveNext
    Next
    If strMsg <> "" Then
        If MsgBox("以下表设置了并行度：" & vbCrLf & Mid(strMsg, 2) & _
            vbCrLf & vbCrLf & "这样可能导致执行计划评估错误，将会严重影响标记转出操作的性能，强烈建议取消这些表的并行度属性。" & _
            vbCrLf & "你确定要继续吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckData = False
            Exit Function
        End If
    End If
    
    '3.检查索引存储表空间的规范性
    strMsg = ""
    strSQL = "Select a.Index_Name, a.Tablespace_Name" & vbNewLine & _
            "From All_Indexes A" & vbNewLine & _
            "Where a.Owner = Zl_Owner And a.Tablespace_Name Not Like 'ZL%INDEX%' And" & vbNewLine & _
            "      a.Table_Name In (Select 表名 From Zltools.Zlbaktables)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        strMsg = strMsg & "," & rsTmp!Index_Name & "(" & Trim(rsTmp!Tablespace_Name) & ")"
        rsTmp.MoveNext
    Next
    If strMsg <> "" Then
        If MsgBox("以下索引没有按规范存储到ZL%INDEX表空间：" & vbCrLf & Mid(strMsg, 2) & _
            vbCrLf & vbCrLf & "这将会严重影响转出操作的性能（不能利用Nologging特性），强烈建议你先重建索引到正确的表空间。" & _
            vbCrLf & "你确定要继续吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckData = False
            Exit Function
        End If
    End If
    
    
    '4.以下检查只有当前用户是dba角色时才进行
    If mblnDBA Then
        '检查delete或update时允许跳过禁用的索引
        strSQL = "Select Value From V$parameter Where Name = 'skip_unusable_indexes'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp!value = "FALSE" Then
            If MsgBox("历史数据转出需要将Oracle初始化参数skip_unusable_indexes修改为TRUE才能正常进行，是否调整该参数？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                CheckData = False
                Exit Function
            Else
                gstrSQL = "alter system set skip_unusable_indexes=true"
                Call gcnOracle.Execute(gstrSQL)
            End If
        End If
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnOffLineMoved And optmode(1).value Then
        If MsgBox("离线模式转出历史数据后没有恢复在线空间的约束和索引，将导致客户端的业务无法正常使用，并且下次进入此模块时也会非常慢，你确定要继续吗？", _
            vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsMovelog = Nothing
End Sub

Private Sub cmdDateLast_Click()
'功能：打开日期选择器
        If IsDate(txtDateLast.Text) Then monSel.value = CDate(txtDateLast.Text)
        
        monSel.Tag = "txtDateLast"
        monSel.Left = Me.ScaleLeft + Me.ScaleWidth - monSel.Width - 120
        monSel.Top = txtDateLast.Top + txtDateLast.Height + 30
        monSel.ZOrder
        monSel.Visible = True
        monSel.SetFocus
End Sub


Private Sub monSel_LostFocus()
    monSel.Visible = False
End Sub

Private Sub optInType_Click(Index As Integer)
    cboBillType.Enabled = Index = 0
    txtNO.Enabled = Index = 0
    cboPatiType.Enabled = Index = 1
    txtPati.Enabled = Index = 1
    
    cboBillType.BackColor = IIF(cboBillType.Enabled, txtDateThis.BackColor, Me.BackColor)
    txtNO.BackColor = IIF(txtNO.Enabled, txtDateThis.BackColor, Me.BackColor)
    cboPatiType.BackColor = IIF(cboPatiType.Enabled, txtDateThis.BackColor, Me.BackColor)
    txtPati.BackColor = IIF(txtPati.Enabled, txtDateThis.BackColor, Me.BackColor)
    
    If cboBillType.Enabled Then cboBillType.SetFocus
    If cboPatiType.Enabled Then cboPatiType.SetFocus
    
End Sub

Private Sub optMode_Click(Index As Integer)
    If Index = 1 Then chkBakTbsDisable.value = 1    '离线模式时固定禁用
    chkBakTbsDisable.Enabled = Index = 0
    chkjob.value = Index
    chkTrigger.value = Index
End Sub

Private Sub optmode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If Index = 0 Then
        strTip = "为了提高性能，转出期间将禁用以下外键和索引：" & vbCrLf & _
                "1.引用转出表主键的所有外键(例如：药品签名明细_FK_收发ID)" & vbCrLf & _
                "2.引用非转出表主键,在转出表外键上的索引(例如：病人医嘱计价_IX_收费细目ID)" & vbCrLf & _
                "3.级联删除的外键停用期间会自动创建触发器来保障正常业务操作时自动删除子表数据。"
    Else
        strTip = "为了提高性能，转出期间除了禁用上述约束和索引，还要禁用：" & vbCrLf & _
                "1.历史数据空间的所有约束和索引;" & vbCrLf & _
                "2.转出表的主键和唯一键(及索引,但保留标记转出查询所需的索引)" & vbCrLf & _
                "3.转出表的所有索引（保留标记转出查询所需的索引）"
    End If

    Call zlCommFun.ShowTipInfo(optmode(Index).hwnd, strTip, True)
End Sub

Private Sub txtDateLast_GotFocus()
    Call zlControl.TxtSelAll(txtDateLast)
End Sub
Private Sub txtDateThis_GotFocus()
    Call zlControl.TxtSelAll(txtDateThis)
End Sub

Private Sub txtDateLast_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txtDateThis_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDateLast_Validate(Cancel As Boolean)
    If IsNull(txtDateLast.Text) Then
        Cancel = True
        Exit Sub
    ElseIf IsDate(txtDateLast.Text) = False Then
        Cancel = True
        Exit Sub
    Else
        If CheckLessBegin(txtDateLast) Or CheckLessThis(CDate(txtDateLast.Text), CDate(txtDateThis.Text)) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDateThis_Validate(Cancel As Boolean)
    If IsNull(txtDateThis.Text) Then
        Cancel = True
        Exit Sub
    ElseIf IsDate(txtDateThis.Text) = False Then
        Cancel = True
        Exit Sub
    Else
        If CheckLessBegin(txtDateThis) Or CheckLessThis(CDate(txtDateLast.Text), CDate(txtDateThis.Text)) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub


Private Sub monSel_DateDblClick(ByVal DateDblClicked As Date)
    monSel.Visible = False
    
    If monSel.Tag = "txtDateLast" Then
        If CheckLessBegin(txtDateLast, monSel.value) = False And CheckLessThis(monSel.value, CDate(txtDateThis.Text)) = False Then
            txtDateLast.Text = Format(monSel.value, "YYYY-MM-DD")
        End If
        
        If txtDateLast.Enabled And txtDateLast.Visible Then txtDateLast.SetFocus
    Else
        If CheckLessBegin(txtDateThis, monSel.value) = False And CheckLessThis(CDate(txtDateLast.Text), monSel.value) = False Then
            txtDateThis.Text = Format(monSel.value, "YYYY-MM-DD")
        End If
        
        If txtDateThis.Enabled And txtDateThis.Visible Then txtDateThis.SetFocus
    End If
End Sub

Private Function CheckLessBegin(objText As TextBox, Optional ByVal dateTemp As Date) As Boolean
'功能：检查指定控件的日期是否小于上次终止日期
        
    If dateTemp = CDate(0) Then dateTemp = CDate(objText.Text)
        
    If dateTemp < mdatBegin Then
        Call FS.ShowTipInfo(objText.hwnd, "不能小于上次转出的终止日期:" & Format(mdatBegin, "YYYY-MM-DD"))
        CheckLessBegin = True
    Else
        Call FS.ShowTipInfo(objText.hwnd, "")
    End If
End Function

Private Function CheckLessThis(dateLast As Date, dateThis As Date) As Boolean
'功能：检查最终截止日期是否小于本次截止日期
        
    If dateLast < dateThis Then
        Call FS.ShowTipInfo(txtDateThis.hwnd, "最终截止日期不能小于本次截止日期")
        CheckLessThis = True
    Else
        Call FS.ShowTipInfo(txtDateThis.hwnd, "")
    End If
End Function

Private Sub txtParallel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "在索引及约束重建、表数据重组时采用的并行度。 " & vbCrLf & _
            "请根据CPU个数及存储设备性能试验后指定，通常由于单一存储设备的性能所限，并非越大越好。"
    Call zlCommFun.ShowTipInfo(txtParallel.hwnd, strTip, True)
End Sub

Private Sub txtSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "转出的天数越多，查询的数据范围就越大，排序所需的内存也越大，插入和删除的记录数也越多，对Undo和Temp表空间及日志缓存的需求就越大。" & vbCrLf & _
            "请根据业务量分析并测试确定，不同时间采用不同天数。"
    Call zlCommFun.ShowTipInfo(txtSplit.hwnd, strTip, True)
End Sub


Private Sub tabFunc_Click()
    Dim i As Long
    
    For i = 0 To fraFunc.UBound
        If fraFunc(i).Tag = tabFunc.SelectedItem.Key Then
            fraFunc(i).Visible = True
        Else
            fraFunc(i).Visible = False
        End If
    Next
    Set imgInfo.Picture = img48.ListImages(tabFunc.SelectedItem.Key).Picture
    
    If tabFunc.SelectedItem.Key = "数据转移" Then
        lblInfo.Caption = "数据转移"
        lblNote.Caption = "    为保持系统高效运行、减少备份数据量、缩短重建索引和统计信息收集等在线空间维护的时间，建议定期将历史数据转移到历史空间中。"
        If Visible And txtDateThis.Enabled Then txtDateThis.SetFocus
    ElseIf tabFunc.SelectedItem.Key = "抽选返回" Then
        lblInfo.Caption = "抽选返回"
        lblNote.Caption = "    抽选某些特殊的数据返回在线数据表，以便实施必要的操作"
        If Visible Then
            If optInType(0).value Then
                optInType(0).SetFocus
            Else
                optInType(1).SetFocus
            End If
        End If
    ElseIf tabFunc.SelectedItem.Key = "未转查询" Then
        lblInfo.Caption = "无法转移的数据原因查询"
        lblNote.Caption = "    列举符合转移时间条件，但未转出的数据记录和不能转移的原因"
        If Visible And dtpBegin.Enabled Then
            dtpBegin.SetFocus
        End If
        If Not dtpBegin.Enabled Then
            MsgBox "由于还未执行过数据转移，不能对无法转移的数据原因进行查询。", vbInformation, gstrSysName
        End If
    ElseIf tabFunc.SelectedItem.Key = "转移日志" Then
        lblInfo.Caption = "转移操作日志"
        lblNote.Caption = "    查看每次转出的数据时间段，以及转出操作的耗时(单位：分钟)"
    
        Call RefreshMoveLog
    
    ElseIf tabFunc.SelectedItem.Key = "转后处理" Then
        lblInfo.Caption = "转后处理"
        lblNote.Caption = "    转出全部完成后,需要人工执行的操作：恢复历史空间被禁用的约束及索引，恢复在线空间被禁用的约束及索引，恢复转出前禁用的后台作业和触发器"
    
        If txtParallel.Enabled And txtParallel.Visible Then txtParallel.SetFocus
    End If
End Sub

Private Sub TIMStatus_Timer()
'刷新进度
    Dim strStatus As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select To_Char(截止时间, 'yyyy-mm-dd') 截止时间, 当前进度" & vbNewLine & _
            "From Zldatamovelog" & vbNewLine & _
            "Where 批次 = (Select Max(批次) From Zldatamovelog)"
    On Error Resume Next
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        If TIMStatus.Tag = "转出" Then
            strStatus = "正在转出" & rsTmp!截止时间 & "之前的数据，当前进度：" & rsTmp!当前进度
        ElseIf TIMStatus.Tag = "标记转出" Then
            strStatus = "正在标记" & rsTmp!截止时间 & "之前的数据，当前进度：" & rsTmp!当前进度
        Else
            strStatus = "当前进度：" & rsTmp!当前进度
        End If
        lblStatus.Caption = strStatus
        lblStatus.Refresh
    End If
    If Err.Number > 0 Then
        lblStatus.Caption = "刷新进度出错:" & Err.Description
        Err.Clear
    End If
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    Select Case cboBillType.ItemData(cboBillType.ListIndex)
    Case 1 '收费单据
        txtNO.Text = GetFullNO(txtNO.Text, 13)
    Case 2, 3 '记帐单据,自动记帐
        txtNO.Text = GetFullNO(txtNO.Text, 14)
    Case 4 '挂号单据
        txtNO.Text = GetFullNO(txtNO.Text, 12)
    Case 5 '就诊卡
        txtNO.Text = GetFullNO(txtNO.Text, 16)
    Case 6 '预交单据
        txtNO.Text = GetFullNO(txtNO.Text, 11)
    Case 7 '结帐单据
        txtNO.Text = GetFullNO(txtNO.Text, 15)
    Case 8 '任务编号
        txtNO.Text = GetFullNO(txtNO.Text, 78)
    End Select
End Sub

Private Sub txtParallel_GotFocus()
    Call zlControl.TxtSelAll(txtParallel)
End Sub

Private Sub txtParallel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtParallel_Validate(Cancel As Boolean)
    If Val(txtParallel.Tag) <> 0 Then
        If Val(txtParallel.Text) > Val(txtParallel.Tag) Then
            MsgBox "并行度不能超过cpu个数" & txtParallel.Tag, vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtPati_Change()
    If txtPati.Text = "" Then
        txtPati.Tag = ""
        cboPatiType.Tag = ""
    End If
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    If Left(txtPati.Text, 1) = "." Then
        If txtPati.SelLength = 0 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
        
    If KeyAscii = 13 And txtPati.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPati.Text = txtPati.Text & Chr(KeyAscii)
            txtPati.SelStart = Len(txtPati.Text)
        End If
        KeyAscii = 0
            
        gstrSQL = Trim(txtPati.Text)
        
        Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
        Case 0, 1
            
            If Left(gstrSQL, 1) = "-" And IsNumeric(Mid(gstrSQL, 2)) Then '病人ID
                gstrSQL = " And A.病人ID=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "*" And IsNumeric(Mid(gstrSQL, 2)) Then '门诊号
                gstrSQL = " And A.门诊号=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "+" And IsNumeric(Mid(gstrSQL, 2)) Then '住院号
                gstrSQL = " And A.住院号=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "." Then '挂号单
                If cboPatiType.ListIndex = 0 Then
                    gstrSQL = " And B.NO='" & Mid(gstrSQL, 2) & "'"
                Else
                    gstrSQL = " And A.病人ID=-1"
                End If
            Else
                gstrSQL = " And A.姓名 Like '" & gstrSQL & "%' And Rownum<=100"
            End If
        
            
            Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
            Case 0
                gstrSQL = _
                    " Select Rownum as ID,A.病人ID,A.门诊号,A.姓名,B.NO as 挂号单," & _
                    " To_Char(B.登记时间,'YYYY-MM-DD') as 就诊时间,C.名称 as 就诊科室,B.执行人 as 医生" & _
                    " From 病人信息 A,H病人挂号记录 B,部门表 C" & _
                    " Where '%'='%' And A.病人ID=B.病人ID And B.执行部门ID=C.ID" & gstrSQL & _
                    " Order by B.登记时间 Desc"
            Case 1
                gstrSQL = _
                    " Select Rownum as ID,A.病人ID,A.住院号,A.姓名,B.主页ID as 住院次数," & _
                    " C.名称 as 住院科室,To_Char(B.入院日期,'YYYY-MM-DD')||'至'||To_Char(B.出院日期,'YYYY-MM-DD') as 住院期间" & _
                    " From 病人信息 A,病案主页 B,部门表 C" & _
                    " Where '%'='%' And B.数据转出=1 And A.病人ID=B.病人ID And B.出院科室ID=C.ID" & gstrSQL & _
                    " Order by B.入院日期 Desc"
            End Select
        Case 2
            If Left(gstrSQL, 1) = "-" And IsNumeric(Mid(gstrSQL, 2)) Then '病人ID
                gstrSQL = " And A.病人ID=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "*" And IsNumeric(Mid(gstrSQL, 2)) Then '门诊号
                gstrSQL = " And A.门诊号=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "+" And IsNumeric(Mid(gstrSQL, 2)) Then '住院号
                gstrSQL = " And A.健康号='" & Trim(Mid(gstrSQL, 2)) & "'"
            Else
                gstrSQL = " And A.姓名 Like '" & gstrSQL & "%' And Rownum<=100"
            End If
            gstrSQL = _
                " Select Rownum as ID,A.病人ID,A.门诊号,A.姓名,A.健康号" & _
                " From 病人信息 A,体检人员目录 B" & _
                " Where '%'='%' And A.病人ID=B.病人ID " & gstrSQL & _
                " Order by B.建档时间 Desc"
        Case 3
            
            If Left(gstrSQL, 1) = "-" And IsNumeric(Mid(gstrSQL, 2)) Then '团体id
                gstrSQL = " And A.ID=" & Val(Mid(gstrSQL, 2))
            Else
                gstrSQL = " And A.名称 Like '%" & gstrSQL & "%' And Rownum<=100"
            End If
            
            gstrSQL = "Select A.ID,A.编码,A.名称,A.说明 From 体检团体目录 A Where '%'='%' " & gstrSQL & " Order by A.建档时间 Desc"
                    
        End Select
        
        'gstrSQL = Replace(gstrSQL, "H病人挂号记录", "病人挂号记录")
        'gstrSQL = Replace(gstrSQL, "B.数据转出=1", "Nvl(B.数据转出,0)=0")
        
        vPoint = zlControl.GetCoordPos(txtPati.Container.hwnd, txtPati.Left, txtPati.Top)
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "门诊病人", , , , , , True, vPoint.X, vPoint.Y, txtPati.Height, blnCancel)
        If Not rsTmp Is Nothing Then
            Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
            Case 0
                txtPati.Tag = rsTmp!病人ID & "," & rsTmp!挂号单
                txtPati.Text = rsTmp!姓名 & "," & rsTmp!就诊时间 & "日就诊"
            Case 1
                txtPati.Tag = rsTmp!病人ID & "," & rsTmp!住院次数
                txtPati.Text = rsTmp!姓名 & ",第" & rsTmp!住院次数 & "次住院"
            Case 2
                txtPati.Tag = rsTmp("病人ID").value
                txtPati.Text = rsTmp("姓名").value & "," & rsTmp("健康号").value
            Case 3
                txtPati.Tag = rsTmp("ID").value
                txtPati.Text = rsTmp("名称").value
            End Select
            
            cboPatiType.Tag = txtPati.Text
            Call zlControl.TxtSelAll(txtPati)
        Else
            If Not blnCancel Then
                MsgBox "没有找到符合条件的病人。", vbInformation, gstrSysName
            End If
            txtPati.Text = "": txtPati.Tag = "": cboPatiType.Tag = ""
        End If
    ElseIf KeyAscii = 13 And txtPati.Text = "" Then
        KeyAscii = 0
    Else
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    If txtPati.Text <> cboPatiType.Tag Then
        Call txtPati_KeyPress(13)
    End If
End Sub

Private Sub txtSplit_GotFocus()
     Call zlControl.TxtSelAll(txtSplit)
End Sub

Private Sub txtSplit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号(费用部份)。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim intType As Integer, curDate As Date
    
    If strNo = "" Then Exit Function
    
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    gstrSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intNum)
    
    If Not rsTmp.EOF Then
        intType = NVL(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        gstrSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & gstrSQL & Format(Right(strNo, 4), "0000")
    Else
        '按年编号
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByNO(ByVal strNo As String, ByVal strTable As String, Optional ByVal strWhere As String) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    If strTable = "病人费用记录" Then
        gstrSQL = "" & _
        "   Select NO From H门诊费用记录 Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "") & _
        "   Union ALL " & _
        "   Select NO From H住院费用记录 Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "")
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, cboBillType.ListIndex + 1)
    Else
        gstrSQL = "Select NO From H" & strTable & " Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "")
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, cboBillType.ListIndex + 1)
    End If
    If Not rsTmp.EOF Then
        MovedByNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByPati(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断指定病人的住院数据是否已经转出
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    gstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        MovedByPati = NVL(rsTmp!数据转出, 0) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByPeis(ByVal bytMode As Byte, ParamArray varParam() As Variant) As Boolean
    '功能：判断指定体检任务的数据是否已经转出
    '参数：bytMode  判断方式
    '       varParam  参数
    '返回：如果有数据转出，返回True,否则返回False
        
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    Select Case bytMode
    Case 1          '按任务编号
        strSQL = "Select 1 From H体检任务记录 Where 任务编号=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(varParam(0)))
    Case 2          '按受检人员
        strSQL = "Select 1 From H体检任务人员 Where 病人id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(varParam(0)))
    Case 3          '按受检团体
        strSQL = "Select 1 From H体检任务记录 Where 体检团体id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(varParam(0)))
    End Select
    
    MovedByPeis = (rsTmp.BOF = False)
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetIndexStatus(ByVal lngSys As Long, ByVal cnThis As ADODB.Connection, ByVal blnEnable As Boolean)
'功能:禁用或启用索引，禁用后提高历史空间的数据插入速度
'     启用时，该过程执行要先于SetConstraintStatus，否则主键或唯一键字段存在无效索引会引发错误,ORA-14063
'参数:lngSys-系统编号
'     cnThis-连接对象
'     blnEnable-索引可用性，true-启用索引 false -禁用索引

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
    Dim strErr As String, i As Long

    '基于规则优化加快SQL执行
    If blnEnable Then
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index ' || a.Index_Name || ' Rebuild Nologging' Sql,a.Index_Name" & vbNewLine & _
                "From User_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Table_Name = t.表名 And t.系统 = " & lngSys & " And t.直接转出 = 1 And a.Status = 'UNUSABLE' And a.Index_Type = 'NORMAL' And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From User_Constraints C Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U')) Order by a.Index_Name"
    Else
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index ' || a.Index_Name || ' unusable' Sql,a.Index_Name" & vbNewLine & _
                "From User_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Table_Name = t.表名 And t.系统 = " & lngSys & " And t.直接转出 = 1 And a.Status = 'VALID' And a.Index_Type = 'NORMAL' And Not Exists" & vbNewLine & _
                " (Select 1 From User_Constraints C Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U')) Order by a.Index_Name"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
       
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    On Error Resume Next
    While Not rsTmp.EOF
        i = i + 1
        If blnEnable Then
            DoEvents  '为了刷新显示进度
            lblPrompt.Caption = "正在启用：" & rsTmp!Index_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
            lblPrompt.Refresh
        Else
            lblStatus.Caption = "正在禁用：" & rsTmp!Index_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
        End If
        
        strSQL = rsTmp!SQL
        cmdTmp.CommandText = strSQL
        cmdTmp.Execute
        
        If Err.Number <> 0 And blnEnable Then
            '如果该索引正在使用，则只能在线重建，比较慢
            If InStr(Err.Description, "ORA-00054") > 0 Then
                Err.Clear
                strSQL = Replace(rsTmp!SQL, "Rebuild", "Rebuild Online")
                cmdTmp.CommandText = strSQL
                cmdTmp.Execute
            End If
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        End If
        
        rsTmp.MoveNext
    Wend
    
    If strErr <> "" Then
        If Len(strErr) > 1000 Then strErr = Mid(strErr, 1, 1000) & "......"
        Call MsgBox(IIF(blnEnable, "启用", "禁用") & "以下索引时发生错误：" & strErr, vbInformation, "索引状态设置")
    End If
End Sub

Private Sub SetConstraintStatus(ByVal lngSys As Long, ByVal cnThis As ADODB.Connection, ByVal blnEnable As Boolean)
'功能:禁用或启用的约束，禁用后提高历史空间的数据插入速度
'     禁用主键或唯一键则会删除对应的索引
'参数:lngSys-系统编号
'     cnThis-连接对象
'     blnEnable=true-启用约束,false-禁用约束

    Dim strSQL As String, strErr As String, i As Long, strTbs As String
    Dim rsTmp As ADODB.Recordset, rsTbs As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    '历史空间没有外键和其他约束，所以，这里全是主键或唯一键
    If blnEnable Then
        '先重建索引，再启用约束，以便重建索引时利用并行执行缩短时间，并且启用约束时也可以采用novalidate方式
         strSQL = "Select d.Table_Name, d.Constraint_Name, f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "From User_Cons_Columns D," & vbNewLine & _
                    "     (Select a.Table_Name, a.Constraint_Name" & vbNewLine & _
                    "       From User_Constraints A, zlBakTables T" & vbNewLine & _
                    "       Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = " & lngSys & " And a.Status = 'DISABLED' And" & vbNewLine & _
                    "             a.Constraint_Type In ('P', 'U')) A" & vbNewLine & _
                    "Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name" & vbNewLine & _
                    "Group By d.Table_Name, d.Constraint_Name" & vbNewLine & _
                    "Order By Constraint_Name"
    Else
        strSQL = "Select " & vbNewLine & _
                " 'ALTER TABLE ' || a.Table_Name || ' disable constraint ' || a.Constraint_Name || Decode(a.Constraint_Type,'P',' Cascade drop index','U',' Cascade drop index','') Sql,a.Constraint_Name" & vbNewLine & _
                "From User_Constraints A, Zltools.Zlbaktables T, User_Tables b" & vbNewLine & _
                "Where a.Table_Name = t.表名 And t.系统 = " & lngSys & " And t.直接转出 = 1 And a.Status = 'ENABLED' And a.Table_Name = b.Table_Name And b.Iot_Type Is Null  Order by Constraint_Name"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
    
    If blnEnable Then
        '优先使用含IDX关键字的索引表空间
        strSQL = "Select Tablespace_Name" & vbNewLine & _
                "From (Select Tablespace_Name" & vbNewLine & _
                "       From User_Indexes" & vbNewLine & _
                "       Where Rownum < 2" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Tablespace_Name" & vbNewLine & _
                "       From User_Indexes" & vbNewLine & _
                "       Where Tablespace_Name Like '%IDX%' And Rownum < 2)" & vbNewLine & _
                "Order By 1 Desc"
        
        Set rsTbs = New ADODB.Recordset
        rsTbs.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
        If rsTbs.RecordCount > 0 Then
            strTbs = " Tablespace " & rsTbs!Tablespace_Name
        End If
    End If
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    On Error Resume Next
    While Not rsTmp.EOF
        i = i + 1
        If blnEnable Then
            DoEvents  '为了刷新显示进度
            lblPrompt.Caption = "正在启用：" & rsTmp!Constraint_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
            lblPrompt.Refresh
        Else
            lblStatus.Caption = "正在禁用：" & rsTmp!Constraint_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
        End If
        
        If blnEnable Then
            '禁用主键或唯一键时，索引是被删除了的，所以这里要用Create
            strSQL = "Create Unique Index " & rsTmp!Constraint_Name & " On " & rsTmp!Table_name & "(" & rsTmp!Colstr & ") Nologging" & strTbs
            cmdTmp.CommandText = strSQL
            cmdTmp.Execute
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        
            '会自动建立约束与索引的关联
            strSQL = "Alter Table " & rsTmp!Table_name & " Enable Novalidate Constraint " & rsTmp!Constraint_Name
            cmdTmp.CommandText = strSQL
            cmdTmp.Execute
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        Else
            strSQL = rsTmp!SQL
            cmdTmp.CommandText = strSQL
            cmdTmp.Execute
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        End If
        
        rsTmp.MoveNext
    Wend
    
    If strErr <> "" Then
        If Len(strErr) > 1000 Then strErr = Mid(strErr, 1, 1000) & "......"
        Call MsgBox(IIF(blnEnable, "启用", "禁用") & "以下约束时发生错误：" & strErr, vbInformation, "约束状态设置")
    End If
End Sub

Private Sub SetNOParallel(ByVal cnThis As ADODB.Connection, ByVal bytType As Byte)
'功能：并行执行后会自动为表名索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢)
'参数：bytType：0=索引，1=表

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    If bytType = 0 Then
        strSQL = "Select Index_Name From User_Indexes Where Degree Not In ('1', '0')"
    Else
        strSQL = "Select Table_name From User_Tables Where Degree !=('         1')"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText

    While Not rsTmp.EOF
        If bytType = 0 Then
            strSQL = "alter index " & rsTmp!Index_Name & " noparallel"
        Else
            strSQL = "alter table " & rsTmp!Table_name & " noparallel"
        End If
        cmdTmp.CommandText = strSQL
        cmdTmp.Execute
        
        rsTmp.MoveNext
    Wend
End Sub

Private Sub InitLogTable()
'功能：初始化表格
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "序列,450,7;批次,450,7;数据开始日期,1400,0;数据结束日期,1400,0;总耗时,850,7;标记耗时,850,7;转出耗时,850,7;重建耗时,850,7"
    arrHead = Split(strHead, ";")
    
    With vsflog
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            .ColKey(.FixedCols + i) = Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub RefreshMoveLog()
'功能：刷新转出日志
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, blnDo As Boolean
    Dim DatStart As Date, lngPre序列 As Long, lng批次 As Long

    On Error GoTo errH
    vsflog.Rows = vsflog.FixedRows
    
    If glngSys \ 100 = 1 Then
        gstrSQL = "Select Min(登记时间) 上次日期 From (Select Min(登记时间) 登记时间 From 门诊费用记录 Union All Select Min(登记时间) From 病人挂号记录)"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        If IsNull(rsTmp!上次日期) Then
            '未产生过业务数据
            Exit Sub
        Else
            DatStart = rsTmp!上次日期
        End If
    Else
        gstrSQL = "Select To_Date('2001-01-01', 'yyyy-mm-dd') as 上次日期 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    End If
 
    strSQL = "Select 序列, 截止时间, Nvl(标记耗时, 0) + Nvl(转出耗时, 0) + Nvl(重建耗时, 0) As 总耗时, 标记耗时, 转出耗时, 重建耗时" & vbNewLine & _
            "From (Select 批次, 序列, 截止时间, Round(To_Number(标记结束时间 - 标记开始时间) * 24 * 60) As 标记耗时," & vbNewLine & _
            "              Round(To_Number(转出结束时间 - Nvl(转出开始时间, 标记结束时间)) * 24 * 60) As 转出耗时," & vbNewLine & _
            "              Round(To_Number(重建结束时间 - 转出结束时间) * 24 * 60) As 重建耗时" & vbNewLine & _
            "       From Zldatamovelog" & vbNewLine & _
            "       Where 系统 = [1])" & vbNewLine & _
            "Order By 批次"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys)

    With vsflog
        .redraw = False
        .Rows = .FixedRows + rsTmp.RecordCount
        .MergeCells = flexMergeFree
        .MergeCol(.ColIndex("序列")) = True
        
        For i = .FixedRows To .Rows - 1
        
            If lngPre序列 = rsTmp!序列 Then
                lng批次 = lng批次 + 1
            Else
                lng批次 = 1
                lngPre序列 = rsTmp!序列
            End If
            
            .TextMatrix(i, .ColIndex("序列")) = rsTmp!序列
            .TextMatrix(i, .ColIndex("批次")) = lng批次
                        
            .TextMatrix(i, .ColIndex("数据开始日期")) = Format(DatStart, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("数据结束日期")) = Format(rsTmp!截止时间, "yyyy-MM-dd")
            DatStart = rsTmp!截止时间
            
            
            .TextMatrix(i, .ColIndex("总耗时")) = "" & rsTmp!总耗时
            .TextMatrix(i, .ColIndex("标记耗时")) = "" & rsTmp!标记耗时
            .TextMatrix(i, .ColIndex("转出耗时")) = "" & rsTmp!转出耗时
            .TextMatrix(i, .ColIndex("重建耗时")) = "" & rsTmp!重建耗时
            
            
            rsTmp.MoveNext
        Next
        
        .redraw = True
        If .Rows > .FixedRows Then
            .Row = .Rows - 1
            .TopRow = .Row
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCommandEnable(ByVal blnEnable As Boolean)
'功能：在耗时操作期间禁用界面主要功能的命令按钮
    
    cmdMoveMark.Enabled = blnEnable
    cmdMoveOut.Enabled = blnEnable
    
    cmdCancel.Enabled = blnEnable
End Sub




