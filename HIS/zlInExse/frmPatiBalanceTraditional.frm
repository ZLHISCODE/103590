VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiBalanceTraditional 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人结帐单(门诊结帐)"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14985
   Icon            =   "frmPatiBalanceTraditional.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   14985
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFormat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3825
      ScaleHeight     =   315
      ScaleWidth      =   2505
      TabIndex        =   70
      Top             =   210
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label lblFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据格式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   15
         TabIndex        =   71
         Top             =   30
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox picBalanceBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7590
      Left            =   7680
      ScaleHeight     =   7590
      ScaleWidth      =   6405
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1725
      Width           =   6405
      Begin VB.CommandButton cmdDelBalance 
         Caption         =   "结算作废(&D)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4725
         TabIndex        =   94
         Top             =   7035
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "继续结算(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   49
         Top             =   7050
         Width           =   1515
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "完成结算(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2970
         TabIndex        =   50
         Top             =   7050
         Width           =   1515
      End
      Begin zlIDKind.IDKindNew IDKindPaymentsType 
         Height          =   360
         Left            =   690
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3870
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         ShowSortName    =   0   'False
         Appearance      =   2
         IDKindStr       =   "现|现金|0|0|0|0|0|0;支|支票|0|0|0|0|0|"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         DefaultCardType =   "0"
         AllowAutoCommCard=   0   'False
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   "更多条件(&M)..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   150
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   90
         Width           =   2100
      End
      Begin VB.OptionButton opt出院 
         Appearance      =   0  'Flat
         Caption         =   "出院结帐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3750
         TabIndex        =   17
         Top             =   180
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton opt中途 
         Appearance      =   0  'Flat
         Caption         =   "中途结帐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2415
         TabIndex        =   16
         Top             =   180
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   -30
         TabIndex        =   73
         Top             =   2535
         Width           =   20000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消结算(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4725
         TabIndex        =   51
         Top             =   7035
         Width           =   1515
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   2745
         Left            =   60
         TabIndex        =   37
         Top             =   3180
         Width           =   6255
         _cx             =   11033
         _cy             =   4842
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPatiBalanceTraditional.frx":058A
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
         Begin VB.Image imgDel 
            Height          =   240
            Left            =   75
            Picture         =   "frmPatiBalanceTraditional.frx":06A0
            Top             =   45
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   1155
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin zl9InExse.txtEdit txtReceive 
         Height          =   405
         Left            =   885
         TabIndex        =   46
         Tag             =   "缴款"
         Top             =   6555
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         MaxLength       =   10
         InputMode       =   4
         Text            =   "99999.99"
      End
      Begin zl9InExse.txtEdit txtCaculated 
         Height          =   405
         Left            =   3945
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "缴款"
         Top             =   6555
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   714
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         MaxLength       =   10
         InputMode       =   2
         Text            =   "0.00"
      End
      Begin zl9InExse.txtEdit txtBalance 
         Height          =   360
         Index           =   3
         Left            =   1275
         TabIndex        =   35
         Tag             =   "冲预交"
         Top             =   2745
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         InputMode       =   2
         Text            =   ""
      End
      Begin VB.CheckBox chkDeposit 
         Caption         =   "退预交"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   2805
         Visible         =   0   'False
         Width           =   1110
      End
      Begin zl9InExse.txtEdit txt天数 
         Height          =   360
         Left            =   3690
         TabIndex        =   32
         Top             =   2055
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   635
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   0
         Text            =   "123"
      End
      Begin MSMask.MaskEdBox txtPatiEnd 
         Height          =   360
         Left            =   4110
         TabIndex        =   29
         Top             =   1590
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPatiBegin 
         Height          =   360
         Left            =   1155
         TabIndex        =   27
         Top             =   1590
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEnd 
         Height          =   360
         Left            =   4110
         TabIndex        =   25
         Top             =   1140
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBegin 
         Height          =   360
         Left            =   1170
         TabIndex        =   23
         Top             =   1140
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zl9InExse.ComboxExpend cboPatiNums 
         Height          =   360
         Left            =   1170
         TabIndex        =   21
         Top             =   675
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   635
         BorderStyle     =   1
         Text            =   "第1次,第2次"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   12
      End
      Begin zl9InExse.txtEdit txtOwe 
         Height          =   405
         Left            =   885
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "缴款"
         Top             =   6090
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   714
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         InputMode       =   2
         Text            =   "0.00"
      End
      Begin VB.CommandButton cmdYBBalance 
         Caption         =   "医保结算(&Y)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2970
         TabIndex        =   52
         Top             =   7050
         Width           =   1515
      End
      Begin VB.PictureBox picOwnerFee 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         FillColor       =   &H000000FF&
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   5160
         ScaleHeight     =   420
         ScaleWidth      =   1080
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   75
         Visible         =   0   'False
         Width           =   1110
         Begin VB.Label lblOwnerFee 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自费项目"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   45
            TabIndex        =   18
            Top             =   75
            Width           =   960
         End
      End
      Begin VB.Label lblOwe 
         AutoSize        =   -1  'True
         Caption         =   "差额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   255
         TabIndex        =   38
         Top             =   6150
         Width           =   600
      End
      Begin VB.Label lblBalanceType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中途结帐"
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
         Height          =   240
         Left            =   5205
         TabIndex        =   74
         Top             =   180
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblPatiNums 
         AutoSize        =   -1  'True
         Caption         =   "住院次数"
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
         TabIndex        =   20
         Top             =   735
         Width           =   960
      End
      Begin VB.Label lblFsTimeRange 
         AutoSize        =   -1  'True
         Caption         =   "至"
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
         Left            =   3390
         TabIndex        =   24
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblFsTime 
         AutoSize        =   -1  'True
         Caption         =   "费用时间"
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
         TabIndex        =   22
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblPatiTimeRange 
         AutoSize        =   -1  'True
         Caption         =   "至"
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
         Left            =   3390
         TabIndex        =   28
         Top             =   1650
         Width           =   240
      End
      Begin VB.Label lblPatiTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院期间"
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
         TabIndex        =   26
         Top             =   1650
         Width           =   960
      End
      Begin VB.Label lblDayName 
         AutoSize        =   -1  'True
         Caption         =   "天"
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
         Left            =   4695
         TabIndex        =   33
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "冲 预 交"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   72
         Top             =   2805
         Width           =   1035
      End
      Begin VB.Label lblCaculated 
         AutoSize        =   -1  'True
         Caption         =   "收款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3315
         TabIndex        =   47
         Top             =   6615
         Width           =   600
      End
      Begin VB.Label lblReceive 
         AutoSize        =   -1  'True
         Caption         =   "缴款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   255
         TabIndex        =   45
         Top             =   6615
         Width           =   600
      End
      Begin VB.Label lblPrevious 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上次自费9999.99"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3330
         TabIndex        =   40
         Top             =   6165
         Width           =   1965
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "结帐时间"
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
         TabIndex        =   30
         Top             =   2115
         Width           =   960
      End
      Begin VB.Label lbl预交余额 
         AutoSize        =   -1  'True
         Caption         =   "预交余额:0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3690
         TabIndex        =   36
         Top             =   2805
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape shpBalance 
         BackColor       =   &H8000000D&
         BorderColor     =   &H8000000D&
         BorderWidth     =   5
         Height          =   1515
         Left            =   210
         Top             =   7380
         Visible         =   0   'False
         Width           =   5925
      End
   End
   Begin VB.PictureBox picPati 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   14985
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   765
      Width           =   14985
      Begin VB.CommandButton cmdYB 
         Caption         =   "验证(&Y)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3420
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "医保病人身份验证,热键F6"
         Top             =   60
         Visible         =   0   'False
         Width           =   1100
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
         Left            =   1230
         TabIndex        =   2
         Top             =   60
         Width           =   2205
      End
      Begin zl9InExse.txtEdit txtSex 
         Height          =   345
         Left            =   4245
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   609
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   0
         Showline        =   1
         Text            =   "男"
      End
      Begin zl9InExse.txtEdit txtOld 
         Height          =   345
         Left            =   5505
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   0
         Showline        =   1
         Text            =   "23岁10天"
      End
      Begin zl9InExse.txtEdit txt费别 
         Height          =   345
         Left            =   7140
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   53
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   0
         Showline        =   1
         Text            =   "普通"
      End
      Begin zl9InExse.txtEdit txt标识号 
         Height          =   345
         Left            =   9585
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   0
         Showline        =   1
         Text            =   "123"
      End
      Begin zl9InExse.txtEdit txtBed 
         Height          =   345
         Left            =   11715
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   609
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   0
         Showline        =   1
         Text            =   "123"
      End
      Begin zl9InExse.txtEdit txt科室 
         Height          =   345
         Left            =   13185
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   609
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   0
         Showline        =   1
         Text            =   "门诊内科"
      End
      Begin zlIDKind.IDKindNew IDKIND 
         Height          =   345
         Left            =   600
         TabIndex        =   1
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         Appearance      =   2
         IDKindStr       =   $"frmPatiBalanceTraditional.frx":0C2A
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
         NotContainFastKey=   "F1;F2;CTRL+F4;F6;F8;F9;F11;F12;ESC"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Line lnPatiSplit 
         BorderColor     =   &H80000003&
         X1              =   -180
         X2              =   30000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
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
         Left            =   6630
         TabIndex        =   7
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   12585
         TabIndex        =   13
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lblBed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11190
         TabIndex        =   11
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbl标识号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
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
         Left            =   8805
         TabIndex        =   9
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
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
         Left            =   4965
         TabIndex        =   5
         Top             =   112
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
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
         Left            =   3720
         TabIndex        =   53
         Top             =   112
         Width           =   480
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   -60
         TabIndex        =   0
         Top             =   105
         Width           =   690
      End
   End
   Begin VB.PictureBox pic状态 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7170
      ScaleHeight     =   315
      ScaleWidth      =   3225
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lbl状态 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   65
         Top             =   30
         Width           =   960
      End
      Begin VB.Label lbl付款方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   645
         TabIndex        =   64
         Top             =   30
         Width           =   1920
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1275
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiBalanceTraditional.frx":0CC0
            Key             =   "Tools"
            Object.Tag             =   "Tools"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiBalanceTraditional.frx":125A
            Key             =   "Down"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiBalanceTraditional.frx":1394
            Key             =   "ColImg"
            Object.Tag             =   "ColImg"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSplitMenu 
      Height          =   45
      Left            =   -30
      TabIndex        =   60
      Top             =   735
      Width           =   30000
   End
   Begin VB.PictureBox picFeeList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8160
      Left            =   60
      ScaleHeight     =   8160
      ScaleWidth      =   7230
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2025
      Width           =   7230
      Begin VB.PictureBox picPatiType 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4620
         ScaleHeight     =   345
         ScaleWidth      =   2535
         TabIndex        =   76
         Top             =   45
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label lblPatiType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人类型:普通病人"
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
            Left            =   15
            TabIndex        =   77
            Top             =   30
            Visible         =   0   'False
            Width           =   2040
         End
      End
      Begin TabDlg.SSTab tabFeeList 
         Height          =   5775
         Left            =   135
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   555
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   10186
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   582
         TabMaxWidth     =   2646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "结算信息(&J)"
         TabPicture(0)   =   "frmPatiBalanceTraditional.frx":192E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "picFeeContain"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "费用明细(&L)"
         TabPicture(1)   =   "frmPatiBalanceTraditional.frx":194A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "picDetailContain"
         Tab(1).ControlCount=   1
         Begin VB.PictureBox picFeeContain 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4710
            Left            =   270
            ScaleHeight     =   4710
            ScaleWidth      =   6285
            TabIndex        =   82
            Top             =   855
            Width           =   6285
            Begin VB.PictureBox picDeposit 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   3825
               Left            =   -15
               ScaleHeight     =   3825
               ScaleWidth      =   5595
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   2010
               Width           =   5595
               Begin VB.CommandButton cmdDepositUp 
                  Caption         =   "↑"
                  Height          =   525
                  Left            =   3210
                  TabIndex        =   93
                  Top             =   600
                  Width           =   330
               End
               Begin VB.CommandButton cmdDepositDown 
                  Caption         =   "↓"
                  Height          =   525
                  Left            =   3210
                  TabIndex        =   92
                  Top             =   1470
                  Width           =   330
               End
               Begin zl9InExse.Command cmdTools 
                  Height          =   330
                  Left            =   4815
                  TabIndex        =   84
                  Top             =   45
                  Width           =   420
                  _ExtentX        =   741
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Picture         =   "frmPatiBalanceTraditional.frx":1966
               End
               Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
                  Height          =   1695
                  Left            =   90
                  TabIndex        =   85
                  Top             =   510
                  Width           =   4305
                  _cx             =   7594
                  _cy             =   2990
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
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483638
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   2
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   350
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
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
                  ExplorerBar     =   1
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   2
                  ShowComboButton =   0
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   0
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
                  Begin zl9InExse.Command cmdColSet 
                     Height          =   255
                     Left            =   45
                     TabIndex        =   86
                     Top             =   45
                     Width           =   195
                     _ExtentX        =   344
                     _ExtentY        =   450
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
               Begin VB.Label lblDeposit 
                  AutoSize        =   -1  'True
                  Caption         =   "预交情况"
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
                  TabIndex        =   88
                  Top             =   105
                  Width           =   960
               End
               Begin VB.Label lblTicketCount 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "预交款收据:"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   240
                  Left            =   1200
                  TabIndex        =   87
                  Top             =   105
                  Width           =   2400
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
               Height          =   1695
               Left            =   0
               TabIndex        =   89
               Top             =   0
               Width           =   4305
               _cx             =   7594
               _cy             =   2990
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
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483634
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   2
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   350
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
               ShowComboButton =   0
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   0
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
         Begin VB.PictureBox picDetailContain 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4230
            Left            =   -74700
            ScaleHeight     =   4230
            ScaleWidth      =   5685
            TabIndex        =   90
            Top             =   675
            Width           =   5685
            Begin VSFlex8Ctl.VSFlexGrid vsDetailList 
               Height          =   1140
               Left            =   150
               TabIndex        =   91
               Top             =   90
               Width           =   4305
               _cx             =   7594
               _cy             =   2011
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
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483634
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   2
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   350
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
               Editable        =   2
               ShowComboButton =   0
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   0
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
      End
      Begin VB.PictureBox picBalanceInfor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   60
         ScaleHeight     =   1050
         ScaleWidth      =   6960
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   6480
         Width           =   6960
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   1
            Left            =   4620
            TabIndex        =   44
            Tag             =   "本次结帐"
            Top             =   615
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            InputMode       =   4
            Text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   2
            Left            =   1110
            TabIndex        =   42
            Tag             =   "结帐说明"
            Top             =   165
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   0
            Left            =   1110
            TabIndex        =   79
            Tag             =   "本次未结"
            Top             =   615
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   635
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "本次未结"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   30
            TabIndex        =   80
            Top             =   690
            Width           =   1020
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "本次结帐"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   3570
            TabIndex        =   43
            Top             =   675
            Width           =   1020
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "结帐说明"
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
            Index           =   2
            Left            =   90
            TabIndex        =   41
            Top             =   225
            Width           =   960
         End
      End
      Begin VB.Line lnFeeSplit 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   30180
         Y1              =   420
         Y2              =   420
      End
   End
   Begin VB.PictureBox picNO 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   10875
      ScaleHeight     =   405
      ScaleWidth      =   2085
      TabIndex        =   58
      Top             =   195
      Width           =   2085
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   15
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   15
         Width           =   1515
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "废"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "热键：F8"
         Top             =   15
         Width           =   450
      End
      Begin VB.Label lblDelCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "废"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1665
         TabIndex        =   68
         Top             =   15
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.TextBox txtInvoice 
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
      IMEMode         =   3  'DISABLE
      Left            =   13125
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   225
      Width           =   1425
   End
   Begin VB.PictureBox pic死亡 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   6390
      ScaleHeight     =   420
      ScaleWidth      =   720
      TabIndex        =   61
      Top             =   165
      Visible         =   0   'False
      Width           =   750
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "死亡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   30
         TabIndex        =   62
         Top             =   45
         Width           =   660
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   2055
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   57
      Top             =   9855
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2884
            MinWidth        =   882
            Picture         =   "frmPatiBalanceTraditional.frx":1F00
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15743
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "上次结帐金额"
            Object.ToolTipText     =   "上次结帐金额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            Key             =   "险类"
            Object.ToolTipText     =   "险类"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "个人帐户余额"
            Object.ToolTipText     =   "个人帐户余额"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   270
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Image imgCol 
      Height          =   195
      Left            =   300
      ToolTipText     =   "选择需要显示的列(ALT+C)"
      Top             =   100
      Width           =   195
   End
   Begin VB.Label lblFact 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据号"
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
      Left            =   12450
      TabIndex        =   67
      Top             =   285
      Visible         =   0   'False
      Width           =   720
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatiBalanceTraditional.frx":2794
      Left            =   810
      Top             =   300
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   0
   End
End
Attribute VB_Name = "frmPatiBalanceTraditional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------
'1.程序入口参数
Private mEditType As gBalanceBill
Private mintPreEditType As Integer   '上次编辑类型
Private mstrPrivs As String, mlngModul As Long, mstrPrivsCard As String
Private mstrInNO As String  '结帐单号
Private mbln门诊转住院 As Boolean 'true:门诊转住院调用接口;False为其他
Private mstrPepositDate As String '指定特点的预交日期(主要是应用于门诊转住院费用时,使用转入的预交进行结帐)
Private mlngPatientID As Long        '当前要结帐的病人ID
Private mstr主页Id As String   '结某次费用:0-结门诊;1-结住院第几次费用;空为不处理
Private mblnNOMoved As Boolean       '操作的单据是否在后备数据表中
Private mobjInPati As Object
Private mblnViewCancel As Boolean
'----------------------------------------------------------------------
'2.菜单相关变量
Private mcbrControl As CommandBarControl, mcbrToolBar As CommandBar
Private mobjPopup As CommandBarPopup
Private mobjCommandBar As CommandBar
Private mobjControl As CommandBarControl
Private mblnNotChange As Boolean

Private Const M_VIEW_ICO = 102 '查询结算显示的图标
Private Const conMenu_View_Balance = 9000
Private Const conMenu_View_List = 9001
Private Const conMenu_View_ListItem = 9002
Private Const conMenu_View_SplitType = 9003
Private Const conMenu_View_SplitMonth = 9004
Private Const conMenu_View_DayBill = 9005
Private Const conMenu_View_DayFM = 9006

Private Const conMenu_View_LblFPH = 9010
Private Const conMenu_View_BillFPH = 9011
Private Const conMenu_View_LblNo = 9012
Private Const conMenu_View_BillNo = 9013
Private Const conMenu_View_CHKCancel = 9012
Private Const conMenu_Edit_NotUseDeposit = 9101 '不使用预交
Private Const conMenu_Edit_UseAllDeposit = 9102 '使用的所有预交
Private Const conMenu_Edit_MoneyUseDeposit = 9103  '按结帐金额使用预交


'3.本地模块变量
Private mobjPayCards As Cards  '结算方式集合
Private mblnFirst As Boolean, mblnInsure As Boolean
Private mblnUnload As Boolean, mblnInterUse As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnDateMoved As Boolean '病人的登记时间是否在转出数据之前
Private mblnCurMzBalanceNo As Boolean '当前为门诊结帐单(非结帐作时才生效)
Private mlngCardTypeID As Long '当前刷卡类型56615
Private mstrPassWord As String, mstrForceNote As String
Private mblnInvalidLoad As Boolean
Private mstr本次住院日期 As String
Private mblnChargeEnd As Boolean
Private mblnNotify As Boolean, mstrInvoice As String
Private mblnPrintInvoice As Boolean
Private mstrPatiBegin As String, mstrPatiEnd As String
Private mblnCurPatiInsure As Boolean
Private mblnReadByZYNo As Boolean
Private mstrBalanceLimit As String, mstrPayMode As String
Private mstrInputInNo As String, mblnBatchState As Boolean
Private mintSucces As Integer  '成功保成单据总数
Private mrsFeeList As ADODB.Recordset '病人未结病人明细
Private mrsDeposit As ADODB.Recordset  '病人预交信息
Private mrsBalance As ADODB.Recordset  '病人结算信息
Private mrsOldBalance As ADODB.Recordset  '病病人结帐信息
Private mbln连续结帐 As Boolean           '当前操作是否连续结帐操作
Private mrsClassMoney As ADODB.Recordset
Private mblnDepositBillPrint As Boolean '是否打印对交款票据
Private mrs结算方式 As ADODB.Recordset  '当前有效的结算方式
Private mstrDec As String       '本次结帐的费用小数位数
Private mblnNotClick As Boolean
Private mblnNotClearBill As Boolean '不清除结帐界面
Private mblnLockScreen As Boolean '当前是否刷屏
Private mstr退支票 As String
Private mstr缺省结算方式 As String  '缺省的支付方式
Private mblnConsChange As Boolean '是否行件发生了改变
Private mblnSecondLoadPati As Boolean
Private mfrmParent As Object, mstrCardPara As String
Private mblnManualEdit As Boolean
Private mstrNoSort As String
Private mblnNoTrigger As Boolean
Private mstrPatient As String

'3.1接口对象定义
Private mobjInPatient As Object
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'3.2 控件变量相关索引定义
Private Enum mInput_Idx
    Idx_本次未结 = 0
    Idx_本次结帐 = 1
    Idx_结帐说明 = 2
    Idx_冲预交 = 3
End Enum
Private Enum mCheck_Idx
    CK_Idx_普通 = 0
    CK_Idx_体检 = 1
End Enum
 
'3.3 模块参数定义
Private Type Ty_ModulePara
    int退款票据 As Integer  '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
    bln结帐后不清信息 As Boolean    ''刘兴洪 问题:27776 日期:2010-02-04 16:49:03
    bln结帐检查病历接收 As Boolean '30036
    byt缴款输入控制 As Byte  '
    bytMzDeposit As Byte    '门诊预交缺省使用方式:0-缺省不使用交;1-按结帐金额使用预交;2-使用所有预交
    bln结帐退款方式 As Boolean 'True-结帐退款默认按预交结算方式 False-结帐退款默认现金
    intPatientRange As Integer  '按姓名查找时,是否只显示未结费用的病人,0-含已结清,1-未结清,2-体检未结清,3-住院未结清
    blnZero  As Boolean '结帐时是否处理零费用
    strOwnerPayFeeType As String '自付费用类别
    int费用时间 As Integer '0-按登记时间,1-按发生时间
    byt结帐时输血费检查 As Byte   '34260
    bln仅用指定预交款 As Boolean  '仅使用指定住院次数的预交款
    bln中途结帐退预交 As Boolean '中途结帐缺省退预交款
    bytInvoiceKindZY As Byte     '0-住院医疗费收据,1-门诊医疗费收据
    bytInvoiceKindMZ As Byte
    int提醒剩余票据张数 As Integer
    blnNotPrintInvioce As Boolean '先结自费时不打印票据
    blnLedWelcome As Boolean
    intOutDay As Integer '结帐可选择出院病人天数
    blnAutoOut As Integer   '是否自动出院
    bytFeePrintSet As Byte      '0-不打印;1-打印提示;2-打印但不提示
    byt结帐检查代收款项 As Byte '出院结帐时检查病人的代收款项,0-禁止,1-提醒
    bln自费缺省使用预交 As Boolean '针对自费费用结帐时,控制是否允许缺省使用预交款进行结帐: 0-不使用预交款;1-使用预交款,缺省为不使用预交款
    byt刷卡缺省金额操作 As Byte '86853
    byt预交票据打印 As Byte
    str脱机医保结算方式 As String
    str自付合计色 As String
    str当前付款色 As String
    str缴款色 As String
    bln退款现金缺省金额 As Boolean
    bln结帐后弹出界面 As Boolean
    bln三方卡结帐退款控制 As Boolean
End Type
Private mty_ModulePara As Ty_ModulePara
Private mblnMC_TwoMode As Boolean '是否支持门诊和住院医保病人身份证验两种模式

'3.4 医保相关定义
Private Type TY_YBInfor
      bln个帐结算 As Boolean '本次是否返回了个帐结算
      cur个帐余额 As Currency '个人帐户余额
      cur个帐限额 As Currency '个人帐户最大限额
      cur个帐透支 As Currency '个人帐户允许透支金额
      cur个帐支付 As Currency   '当前个人帐户支付
      cur统筹支付 As Currency   '当前医保统筹支付
      strYBPati As String    '医保病人身份信息
      intInsure As Integer   '作废时,读取的单据中的险类,用来判断是否退现金,算误差等
      bln医保作废全退 As Boolean     '是否有不支持的作废结算方式
      bytMCMode As Byte '医保病人身份证验模式,包括1-门诊,2-住院两种模式,0-表示非医保
      strBalance As String '医保返回的各种结算金额:"结算方式;金额;是否允许修改|...
      blnAutoOut As Boolean '在院病人结帐后是否自动出院
End Type
Private mYBInFor As TY_YBInfor '医保相关信息
'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    '1.门诊，住院结算共用的参数
    分币处理 As Boolean
    医保接口打印票据 As Boolean
    '2.门诊结算用的参数
    门诊病人结算作废 As Boolean
    门诊必须传递明细 As Boolean
    门诊预结算 As Boolean
    门诊结算_结帐设置 As Boolean
    '3.住院结算用的参数
    未结清出院 As Boolean
    结算使用个人帐户 As Boolean
    出院结算必须出院 As Boolean
    出院病人结算作废 As Boolean
    中途结算仅处理已上传部分 As Boolean
    结帐设置后调用接口 As Boolean
    结帐作废后打印回单 As Boolean
    住院结算作废 As Boolean
    允许结多次住院费用 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

'-----------------------------------------------------------------
'3.4老版一卡通相关
Private Type TY_OneCard
      blnOneCard As Boolean      '是否启用了一卡通接口
      rsOneCard As ADODB.Recordset
      strOneCard As String       '读卡时所选择的一卡通接口对应的结算方式
End Type
Private mOldOneCard As TY_OneCard
'-----------------------------------------------------------------
'3.5 结帐条件相关
Private mobjBalanceAll As clsBalanceAllCon
Private mobjBalanceCon As clsBalanceCon

'当前结帐数据
Private Type TY_Balance_Infor
    dbl医保支付合计 As Currency  '医保支付合计
    dbl冲预交合计 As Double
    dbl本次未结 As Double
    dbl当前结帐 As Double
    dbl已付合计 As Double
    dbl未付合计 As Double
    dbl预结算总额 As Double
    bln预交刷卡 As Boolean '预交款是否已经刷卡
    blnSaveBill As Boolean '当前已经保存结帐单
    strNO As String   '当前保存的结帐单
    lng结帐ID As Long '当前保存的结帐ID
    dtBalanceDate As Date '当前结帐时间
    str病历原因 As String '病历原因
    dbl缴款 As Double
    dbl找补 As Double
    dbl退支票 As Double
    dbl误差额 As Double
    dbl现金 As Double
    lng预交ID As Long
    str预交No As String
    lng冲销ID As Long
End Type
Private mBalanceInfor As TY_Balance_Infor
'病人当前信息
Private Type ty_Pati_Infor
    lng病人ID  As Long
    lng主页ID As Long
    str姓名 As String
    str性别 As String
    str年龄 As String
    objCard As Card         '上次结算信息
    bln连续结帐 As Boolean  '是否连续结帐
    bln出院 As Boolean      '当前病人是否出院
    dbl预交余额 As Double   '本次预交余额
    dbl费用余额 As Double   '未结费用
    dbl剩余合计 As Double   '本次预交余额-未结费用
    dbl实际余额 As Double   '预交明细余额
    dbl未付累计 As Double  '上次未付累计金额
    bln退款标志 As Boolean
End Type
Private mPatiInfor As ty_Pati_Infor

'当前发票信息
Private mobjInvoice As clsInvoice
Private mobjFactProperty As clsFactProperty
Private mobjRedProperty As clsFactProperty
Private mobjDepositFactProperty As clsFactProperty
Private mstrDepositInvioce As String '当前预交发票号
Private mlng领用ID As Long
Private mlng预交领用ID As Long

'消费卡刷卡信息
Private mcllSquareBalance As Collection '消费卡结算信息
Private mcllCurSquareBalance As Collection '当前消费卡刷卡信息

'当前刷卡信息
Private Type TY_BrushCard    '刷卡类型
    str卡号 As String
    str密码 As String
    str交易流水号 As String    '交易流水号
    str交易说明  As String     '交易信息
    str扩展信息 As String    '交易的扩展信息
    dbl帐户余额 As Double
    str结算号码 As String
    str结算摘要 As String
    bln转帐 As Boolean '是否当前为转帐交易
End Type
Private Enum mConPans
    Pan_PatiCon = 1
    Pan_FeeList = 2
    Pan_Deposit = 3
    Pan_Balance = 4
End Enum
Private mbln已报价 As Boolean
'外挂评价器对象
Private mobjPlugIn As Object

Public Function ShowMe(ByVal frmMain As Object, ByVal EditType As gBalanceBill, _
    ByVal strPrivs As String, Optional lng病人ID As Long = 0, Optional str主页ID As String = "", _
    Optional ByVal strNO As String, Optional blnViewCancel As Boolean, Optional blnNOMoved As Boolean, _
    Optional bln门诊转住院 As Boolean, Optional strPepositDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐程序入口
    '入参:EditType-编辑类型
    '     strPrivs-权限串
    '     lng病人ID-当前要结帐的病人ID
    '     str主页Id As String   '结某次费用:0-结门诊;1-结住院第几次费用;空为不处理
    '     strNo-传入要操作的结帐单号,新结帐时,不传入
    '     blnViewCancel-是否查看的作废单据
    '     blnNOMoved-strNo是否已经转入后备表中
    '     bln门诊转住院-true:门诊转住院调用接口;False为其他
    '     strPepositDate-指定特点的预交日期(主要是应用于门诊转住院费用时,使用转入的预交进行结帐)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-12-29 15:24:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    mEditType = EditType: mstrPrivs = strPrivs
    mstrInNO = strNO: mbln门诊转住院 = bln门诊转住院
    mstrPepositDate = strPepositDate: mlngPatientID = lng病人ID
    mstr主页Id = str主页ID: mintSucces = 0: mblnNOMoved = blnNOMoved
    Set mfrmParent = frmMain
    mblnViewCancel = blnViewCancel
    mintPreEditType = -1 '上次编辑类型设置为负数,方便在保存数据后进行界面恢复处理
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    If mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then Exit Function
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    ShowMe = mintSucces > 0
End Function


Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    With mty_ModulePara
        '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
        .int退款票据 = Val(zlDatabase.GetPara("退款收据打印", glngSys, mlngModul))
        .bln结帐后不清信息 = IIf(Val(zlDatabase.GetPara("结帐后不清除信息", glngSys, mlngModul)) = 1, True, False)
        .bln结帐检查病历接收 = IIf(Val(zlDatabase.GetPara("结帐检查病历接收", glngSys, mlngModul)) = 1, True, False) '30036
        '问题:43153:0-不进行控制;1-存在收取现金时,必须输入缴款;2-结帐时按单病人累计
        .byt缴款输入控制 = Val(zlDatabase.GetPara("结帐缴款输入控制", glngSys, mlngModul, 0))
        .bytMzDeposit = Val(zlDatabase.GetPara("门诊预交缺省使用方式", glngSys, mlngModul, 2))
        .bln结帐退款方式 = IIf(Val(zlDatabase.GetPara("结帐退款缺省方式", glngSys, mlngModul)) = 1, True, False)
        .intPatientRange = Val(zlDatabase.GetPara("显示结清病人", glngSys, mlngModul, 0))
        .blnZero = zlDatabase.GetPara("处理零费用", glngSys, mlngModul) = "1"
        .strOwnerPayFeeType = zlDatabase.GetPara("结算前先结自费费用", glngSys, mlngModul, "")
        .int费用时间 = IIf(zlDatabase.GetPara("结帐费用时间", glngSys, mlngModul) = "1", 1, 0)
        .byt结帐时输血费检查 = Val(zlDatabase.GetPara("结帐时输血费检查", glngSys, mlngModul, "0"))
        .bln仅用指定预交款 = zlDatabase.GetPara("仅用指定预交款", glngSys, mlngModul) = "1"
        .bln中途结帐退预交 = zlDatabase.GetPara("中途结帐退预交", glngSys, mlngModul) = "1"
        .bytInvoiceKindZY = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
        .bytInvoiceKindMZ = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
        .blnLedWelcome = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, "1") = "1"
        .blnNotPrintInvioce = Val(zlDatabase.GetPara("先结自费费用不打印结帐票据", glngSys, mlngModul, "0")) = 1
        .blnAutoOut = zlDatabase.GetPara("在院病人结帐后自动出院", glngSys, mlngModul) = "1"
        .bytFeePrintSet = Val(zlDatabase.GetPara("结帐明细打印", glngSys, mlngModul, "0"))
        .byt结帐检查代收款项 = zlDatabase.GetPara("结帐检查代收款项", glngSys, mlngModul, , "0")
        .int提醒剩余票据张数 = 0 '暂时未有发票张数的参数控制
        .bln自费缺省使用预交 = Val(zlDatabase.GetPara("自费缺省使用预交", glngSys, mlngModul, "0")) = 1
        .byt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0")) '86853
        .byt预交票据打印 = Val(zlDatabase.GetPara("预交票据打印方式", glngSys, mlngModul, "0"))
        .str脱机医保结算方式 = zlDatabase.GetPara("脱机医保结算方式", glngSys)
        .str当前付款色 = zlDatabase.GetPara("当前付款栏字体色", glngSys, mlngModul, "255|255")
        .str缴款色 = zlDatabase.GetPara("缴款栏字体色", glngSys, mlngModul, "16711680|255")
        .str自付合计色 = zlDatabase.GetPara("自付合计栏字体色", glngSys, mlngModul, "16711680")
        .bln退款现金缺省金额 = zlDatabase.GetPara("退款现金结算缺省金额", glngSys, mlngModul) = "1"
        .bln结帐后弹出界面 = zlDatabase.GetPara("病人多次结帐弹出结帐条件窗体", glngSys, mlngModul) = "1"
        .bln三方卡结帐退款控制 = zlDatabase.GetPara("三方卡结帐退款控制", glngSys, mlngModul) = "1"
    End With
    
    txtReceive.ForeColor = Mid(mty_ModulePara.str缴款色, 1, InStr(mty_ModulePara.str缴款色, "|") - 1)
    lblReceive.ForeColor = Mid(mty_ModulePara.str缴款色, 1, InStr(mty_ModulePara.str缴款色, "|") - 1)
    IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str缴款色, 1, InStr(mty_ModulePara.str缴款色, "|") - 1)
    
    '成都老版医保支持门诊和住院两种身份验证模式
    mblnMC_TwoMode = InStr("," & GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "") & ",", ",20,") > 0
    
    mstrPrivsCard = ";" & GetPrivFunc(glngSys, 1151) & ";"
End Sub

Private Sub SetCurBalanceVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置结算信息是否显示
    '编制:刘兴洪
    '日期:2015-01-19 16:49:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    blnVisible = Not mEditType = g_Ed_单据查看
     
    If mEditType = g_Ed_重新作废 Then
        lblBalance(0).Visible = False
        txtBalance(Idx_本次未结).Visible = False
    Else
        lblBalance(0).Visible = blnVisible
        txtBalance(Idx_本次未结).Visible = blnVisible
        If blnVisible = False Then
            Set lblBalance(7).Container = picBalanceInfor
            lblBalance(7).Left = lblBalance(0).Left + 30
            lblBalance(7).Top = lblBalance(0).Top
            txtDate.Top = txtBalance(Idx_本次未结).Top + 15
            txtDate.Left = txtBalance(Idx_本次未结).Left - 15
            txtDate.Width = txtDate.Width - 15
            Set txtDate.Container = picBalanceInfor
        End If
    End If
    
    lblBalance(3).Visible = True
    If (mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1) Then
        If InStr(1, mstrPrivs, ";预交退现金;") > 0 Then
            chkDeposit.Visible = True
            cmdTools.Visible = True
            lblBalance(3).Visible = False
        Else
            cmdTools.Visible = False
            chkDeposit.Visible = False
        End If
    End If
    If mEditType = g_Ed_重新作废 Or mEditType = g_Ed_结帐作废 _
        Or mEditType = g_Ed_单据查看 Or chkDeposit.Visible Or chkCancel.Value = 1 Then
        cmdTools.Visible = False
        cmdDepositUp.Visible = False
        cmdDepositDown.Visible = False
    Else
        cmdTools.Visible = blnVisible
        cmdDepositUp.Visible = blnVisible
        cmdDepositDown.Visible = blnVisible
    End If
    Call picDeposit_Resize
    
End Sub

Private Sub cboPatiNums_NodeCheckValied(ByVal Node As MSComctlLib.Node, blnCancel As Boolean)
    Dim objNode As MSComctlLib.Node
    Dim varTemp As Variant, str主页Ids As String, lng病人ID As Long
    Dim int主页ID As Integer, intInsure As Integer, strInsureName As String
    Dim int主页ID1 As Integer, intInsure1 As Integer, strInsureName1 As String
    Dim blnFirst As Boolean
    If mrsInfo Is Nothing Then blnCancel = True: Exit Sub
    If mrsInfo.State <> 1 Then blnCancel = True: Exit Sub
    
    lng病人ID = Val(NVL(mrsInfo!病人ID))
    
    
    '选检查当前节点的有效性
    '主页ID|险类|险类名称
    str主页Ids = cboPatiNums.GetNodesCheckedDatas(False)
    
    If str主页Ids = "" Then '为空时，必须选择一个
        
        blnCancel = True: Exit Sub
    End If
    varTemp = Split(str主页Ids, ",")
    
    
    If zlGetTimeDataFromTimes(varTemp(0), int主页ID, intInsure, strInsureName) = False Then
         blnCancel = True: Exit Sub
    End If
    If intInsure <> 0 Then Call InitInsurePara(lng病人ID, intInsure)

    If Node.Key = "Root" Then '当前操作“所有住院”节点
        If Node.Checked Then     '选择所有
            blnFirst = True
            'If intInsure = 0 Then Exit Sub '非医保，则按自费进行结算
            For Each objNode In cboPatiNums.Nodes
                If objNode.Key <> "Root" Then
                    If zlGetTimeDataFromTimes(objNode.Tag, int主页ID1, intInsure1, strInsureName1) Then
                        If blnFirst Then
                            intInsure = int主页ID1: intInsure = intInsure1: strInsureName = strInsureName1
                            If intInsure <> 0 Then Call InitInsurePara(lng病人ID, intInsure)
                            If MCPAR.允许结多次住院费用 Then Exit Sub
                        Else
                            If intInsure <> 0 Then
                               Node.Checked = False
                               If int主页ID1 <> int主页ID Then objNode.Checked = False
                            End If
                        End If
                    End If
                    blnFirst = False
                End If
            Next
            Exit Sub
        End If
        blnCancel = True: Exit Sub '不能一个不勾选
    End If
    
    If zlGetTimeDataFromTimes(Node.Tag, int主页ID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
      
    If Node.Checked Then      '当前勾选
        If int主页ID1 = int主页ID Then   '当前选中的，就是第一个选择的
            If intInsure = 0 Or MCPAR.允许结多次住院费用 Then Exit Sub '非医保的，则允许自费或允许多次住院一次结，则全部按医保结算
            '只能选择第一次的住院的
            For Each objNode In cboPatiNums.Nodes
                If zlGetTimeDataFromTimes(objNode.Tag, int主页ID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
                If int主页ID <> int主页ID1 Or objNode.Key = "Root" Then objNode.Checked = False
            Next
            Exit Sub
        End If
        '肯定不是选择的第一个
        If intInsure = 0 Or MCPAR.允许结多次住院费用 Then Exit Sub '非医保的，则允许自费或允许多次住院一次结，则全部按医保结算
        
        If zlGetTimeDataFromTimes(Node.Tag, int主页ID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
        If intInsure <> 0 And intInsure1 = 0 Then
           '清除原来选择的，则以最后选择的为准
           int主页ID = int主页ID1: intInsure = intInsure1: strInsureName = strInsureName1
            For Each objNode In cboPatiNums.Nodes
                If zlGetTimeDataFromTimes(objNode.Tag, int主页ID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
                If int主页ID <> int主页ID1 Or objNode.Key = "Root" Then objNode.Checked = False
            Next
            Exit Sub
        
        End If
        
        For Each objNode In cboPatiNums.Nodes
            If zlGetTimeDataFromTimes(objNode.Tag, int主页ID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
            If int主页ID <> int主页ID1 Or objNode.Key = "Root" Then objNode.Checked = False
        Next
        Exit Sub
    Else
         If intInsure <> 0 And MCPAR.允许结多次住院费用 = False Then '第一个是医保，则需要作排拆处理
            For Each objNode In cboPatiNums.Nodes
                If objNode.Key <> Node.Key Then
                    If zlGetTimeDataFromTimes(objNode.Tag, int主页ID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
                    If int主页ID <> int主页ID1 Or objNode.Key = "Root" Then objNode.Checked = False
                End If
            Next
            Exit Sub
         End If
    End If
    
    '当前选择的只有一个，且就是当前这个
    If UBound(varTemp) = 0 Then
        If int主页ID = int主页ID1 Then blnCancel = True: Exit Sub   '不允许取消，必须选择一个
    End If
End Sub





Private Sub cmdDelBalance_Click()
    
    '结算作废
    If MsgBox("你真的要作废当前的结帐信息吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '重新提取数据
    If zlGetFromIDToBalanceData(mBalanceInfor.lng结帐ID, False, mrsBalance) = False Then Exit Sub
    
    If DeleteBalance(True) = False Then Exit Sub
    mintSucces = mintSucces + 1
    Call NewBill
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub cmdDepositUp_Click()
    If mEditType <> g_Ed_门诊结帐 And mEditType <> g_Ed_住院结帐 And mEditType <> g_Ed_重新结帐 Then Exit Sub
    With vsDeposit
        If .Row <= 1 Then Exit Sub
        .RowPosition(.Row) = .Row - 1
        .Select .Row - 1, 1
'        Call RecalcDepositMoney(2, Val(mBalanceInfor.dbl冲预交合计))
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor(True)
    End With
End Sub

Private Sub cmdDepositDown_Click()
    If mEditType <> g_Ed_门诊结帐 And mEditType <> g_Ed_住院结帐 And mEditType <> g_Ed_重新结帐 Then Exit Sub
    With vsDeposit
        If .Row >= .Rows - 1 Then Exit Sub
        .RowPosition(.Row) = .Row + 1
        .Select .Row + 1, 1
'        Call RecalcDepositMoney(2, Val(mBalanceInfor.dbl冲预交合计))
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor(True)
    End With
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2014-12-19 11:18:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnStartFactUseType As Boolean    '是否启用了多种使用类型票据

    If mEditType = g_Ed_门诊结帐 Then
        blnStartFactUseType = zlStartFactUseType(IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1))
    ElseIf mEditType = g_Ed_住院结帐 Then
        blnStartFactUseType = zlStartFactUseType(IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1))
    End If
    dkpMain.SetCommandBars Me.cbsThis
    
    Call InitPatiBalanceVariableCon
    
    Call InitVar    '初始化内部相关变量
        
    Set cmdColSet.Picture = imgCol.Picture
    Call initCardSquareData '初始化卡对象
    Call Load找补项(0, "找   补") '初始化找补项
    Call InitOldOneCardInfor '初始化老一卡通相关变量
    Call InitCombox_Cons '初始化结帐条件信息
    Call InitGrid
    
    Call SetCurBalanceVisible   '设置当前结算信息的显示
    Call InitPancel '初始化区域
     
    '问题号:112545,焦博,2017/08/25,发票不分类别时,进入结帐界面就显示当前操作员的票据号
    Call ReInitPatiInvoice(Not blnStartFactUseType)
    
    Set cmdColSet.Picture = imgList.ListImages("ColImg").Picture
    Call SetOperatonCommandCaption
    
    If mblnMC_TwoMode Then
        cmdYB.Caption = "刷"
        cmdYB.Width = 400
    End If
    Call NewBill
    cmdMore.Visible = InStr(mstrPrivs, ";结帐设置;") > 0
    txtBalance(Idx_本次结帐).Enabled = InStr(mstrPrivs, ";结帐设置;") > 0
    txtBalance(Idx_本次结帐).Locked = InStr(mstrPrivs, ";结帐设置;") = 0
    
    cboPatiNums.Enabled = InStr(mstrPrivs, ";结帐设置;") > 0
    txtBalance(Idx_本次结帐).BackColor = IIf(txtBalance(Idx_本次结帐).Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
    
    vsFeeList_LostFocus
    vsDeposit_LostFocus
    vsBlance_LostFocus
    
End Sub
Private Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关内部变量
    '编制:刘兴洪
    '日期:2015-01-14 11:27:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    mstr缺省结算方式 = ""
'    mstrDec = gstrDec
    mstrDec = "0.00"
    
    Set mobjFactProperty = New clsFactProperty
    Set mobjRedProperty = New clsFactProperty
    Set mobjDepositFactProperty = New clsFactProperty
    
    If mEditType = g_Ed_门诊结帐 Then
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1), 0, 0, 0, mobjFactProperty, , , 1)
    Else
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), 0, 0, 0, mobjFactProperty, , , 2)
    End If
    
   mstr退支票 = ""
   strSQL = " " & _
    " Select B.名称 " & _
    " From 结算方式应用 A, 结算方式 B " & _
    " Where A.应用场合 = '结帐' And B.名称 = A.结算方式 " & _
    "       And Nvl(B.应付款, 0) = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstr退支票 = NVL(rsTemp!名称)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitRedInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化红票信息
    '入参:blnFact-是否刷新发票号
    '编制:刘兴洪
    '日期:2015-01-07 16:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng病人ID As Long, lng主页ID As Long, intInsure As Integer
    
    intInsure = mYBInFor.intInsure
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(NVL(mrsInfo!病人ID)): lng主页ID = Val(NVL(mrsInfo!主页ID))
            intInsure = mYBInFor.intInsure
        End If
    End If
    Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 13, 11), lng病人ID, lng主页ID, intInsure, mobjRedProperty)
    If mobjRedProperty.启用使用类别 Then mlng领用ID = 0
    If blnFact Then Call RefreshRed
End Sub

Private Sub RefreshRed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新结帐的票据号
    '编制:刘兴洪
    '日期:2015-01-07 17:16:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjRedProperty Is Nothing Then Exit Sub
    If mobjRedProperty.打印方式 = 0 Then Exit Sub
      
    If Not mobjRedProperty.严格控制 Then
        '非严格控制下
        '松散：取下一个号码
        mstrInvoice = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
        Exit Sub
    End If
    
    If zlGetRedGroupUseID(mlng领用ID, 1, "") = False Then
        mstrInvoice = ""
        Exit Sub
    End If
    
    '严格：取下一个号码
    If mobjInvoice.zlGetNextBill(mlngModul, mlng领用ID, strFactNO) = False Then strFactNO = ""
    mstrInvoice = strFactNO
    
    If mobjRedProperty.启用使用类别 Then Call zlCheckFactIsEnough
End Sub

Private Sub InitCombox_Cons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化下拉数据
    '编制:刘兴洪
    '日期:2015-01-05 14:31:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cboPatiNums.Clear
    If mEditType = g_Ed_门诊结帐 Then
        cboPatiNums.AddItem "R", "所有门诊", True, True, True, , "0"
    Else
        cboPatiNums.AddItem "R", "所有住院", True, True, True, , "0"
    End If
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
      
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
        Set .Font = cboPatiNums.Font
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.DeleteAll
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = True
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_View_Balance, "结帐表")
        objControl.IconId = M_VIEW_ICO
        With objControl.CommandBar.Controls
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_List, "明细表")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListItem, "项目明细")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_SplitType, "分类表")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_SplitMonth, "分月表")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_DayBill, "逐日单据")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_DayFM, "逐日费目")
            mcbrControl.IconId = M_VIEW_ICO
        End With
        If InStr(1, mstrPrivs, ";门诊费用转住院;") > 0 Then
            Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicToHos, "门诊转住院")
            mcbrControl.IconId = 3036
            mcbrControl.BeginGroup = True
        End If
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        mcbrControl.BeginGroup = True
    End With
        '主菜单右侧的查找
    With mcbrToolBar.Controls
         Set mcbrControl = .Add(xtpControlLabel, conMenu_View_LblFPH, "发票号")
         mcbrControl.flags = xtpFlagRightAlign
         
        Set objCustom = .Add(xtpControlCustom, conMenu_View_BillFPH, "")
        objCustom.Handle = txtInvoice.hWnd
        objCustom.flags = xtpFlagRightAlign
        

        Set mcbrControl = .Add(xtpControlLabel, conMenu_View_LblNo, " 单据号")
        mcbrControl.flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_BillNo, "")
        objCustom.Handle = picNO.hWnd
        objCustom.flags = xtpFlagRightAlign
  
'        Set objCustom = .Add(xtpControlCustom, conMenu_View_CHKCancel, "")
'        objCustom.Handle = picCancel.hWnd
'        objCustom.Flags = xtpFlagRightAlign
        
        'IDKind.BackColor = picBillNo.BackColor
    End With

    For Each mcbrControl In mcbrToolBar.Controls
        Select Case mcbrControl.ID
        Case conMenu_View_LblFPH, conMenu_View_LblNo
        Case Else
            mcbrControl.Style = xtpButtonIconAndCaption
        End Select
    Next
    zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub
Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date, bytFlag As Byte
    Dim lng病人ID  As Long
    
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 15)
        
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("病人结帐记录", cboNO.Text, , , Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 7, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        '单据权限
        If Not ReadBillInfo(2, cboNO.Text, -1, strOper, vDate) Then
            cboNO.Text = "": If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
            Exit Sub
        End If
        
        If Not BillOperCheck(7, strOper, vDate, "作废") Then
            cboNO.Text = "": If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
            Exit Sub
        End If
        
        'lng病人ID:49084
        mYBInFor.intInsure = BalanceExistInsure(cboNO.Text, bytFlag, lng病人ID)
        mYBInFor.bytMCMode = bytFlag
        If mYBInFor.intInsure <> 0 Then
            '保险结算权限判断
            If InStr(mstrPrivs, ";保险结算;") = 0 Then
                MsgBox "你没有权限作废保险病人的结帐单据。", vbInformation, gstrSysName
                Exit Sub
            End If
            MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, mYBInFor.intInsure)
            If mYBInFor.bytMCMode = 1 Then
                MCPAR.门诊病人结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, mYBInFor.intInsure)
            Else
                MCPAR.出院病人结算作废 = gclsInsure.GetCapability(support出院病人结算作废, lng病人ID, mYBInFor.intInsure)
            End If
            MCPAR.结帐作废后打印回单 = gclsInsure.GetCapability(support结帐作废后打印回单, lng病人ID, mYBInFor.intInsure)
        Else
            If InStr(mstrPrivs, ";普通病人结算;") = 0 Then
                MsgBox "你没有权限作废普通病人的结帐单据。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If PatiErrBillPay(0, cboNO.Text) = False Then
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
            Exit Sub
        End If
        If CheckExistsGathering(cboNO.Text) Then
            MsgBox "该结帐单据存在已缴款的应收款记录，请退款后再执行作废。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckBillBeforIN(cboNO.Text) Then
            If MsgBox("该结帐单是本次住院之前发生的，你确定要作废该单据吗?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If InStr(mstrPrivs, ";结帐作废;") = 0 Then
             MsgBox "你没有权限作废结帐单据。", vbInformation, gstrSysName
             Exit Sub
        End If
        
        '读取要作废的结帐单
        If Not ReadBalance(cboNO.Text, True) Then
            cboNO.Text = "": If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        Else
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        End If
    Else
           If InStr(mstrPrivs, ";普通病人结算;") = 0 Then
                MsgBox "你没有权限作废非保险病人的结帐单据。", vbInformation, gstrSysName
                Exit Sub
           End If
    End If
End Sub


Private Sub cboPatiNums_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboPatiNums_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
    Dim intMaxTime As Integer, intNum As Integer, arrNum  As Variant
    Dim strNodesChecked As String, strAllSelTime As String
    Dim intInsure As Integer, strInsureName As String
    Dim i As Integer
    
    
    strNodesChecked = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas) '所有的住院次数返回空
    strAllSelTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas(False))
    
    If strAllSelTime <> "" Then
        arrNum = Split(strAllSelTime, ",")
        intNum = Val(arrNum(0))
    End If
    
    
    If Not mblnNotChange Then
        mblnNotChange = True
        Call RecalcFeeTotalDate
        mblnNotChange = False
    End If
     
    Call ClearVsBlance
    Call ClearListData(True)
    
    If strAllSelTime <> "" Then
        mobjBalanceCon.strTime = strAllSelTime
        If Not mrsInfo Is Nothing Then
            Call SetPatiNums
        End If
        
        If mobjBalanceCon.strTime = "" Then
            intMaxTime = mPatiInfor.lng主页ID
        Else
            arrNum = Split(mobjBalanceCon.strTime, ",")
            For i = 0 To UBound(arrNum)
                If Val(arrNum(i)) > intMaxTime Then intMaxTime = Val(arrNum(i))
            Next i
        End If
        
        Call LoadDefaultOutStatu(mPatiInfor.lng病人ID, intMaxTime, True)
        
        If Not ShowBalance() Then
            cmdOK.Enabled = False
            MsgBox "在当前条件下,病人不存需要结帐的费用！", vbInformation, gstrSysName
            mbln连续结帐 = False
            Exit Sub
        End If
    End If
     
End Sub

Private Sub SetPatiNums()
    Dim blnFirst As Boolean, i As Integer
    Dim varType As Variant, blnSelfFee As Boolean
    
    On Error GoTo errH
    If mEditType <> g_Ed_门诊结帐 Then
        '设置病人医保状态
        mobjBalanceAll.rsAllTime.Filter = "主页ID=" & Val(Split(mobjBalanceCon.strTime, ",")(0))
        If Not mobjBalanceAll.rsAllTime.EOF Then
            If mYBInFor.intInsure <> Val(mobjBalanceAll.rsAllTime!险类) Then
                Call InitInsurePara(Val(NVL(mobjBalanceAll.rsAllTime!病人ID)), Val(mobjBalanceAll.rsAllTime!险类))
            End If
            mYBInFor.intInsure = Val(mobjBalanceAll.rsAllTime!险类)
            mYBInFor.strBalance = ""
            mobjBalanceAll.rsAllTime.Filter = ""
        End If
    End If
    
    '设置病人自费状态
    blnSelfFee = True
    If mobjBalanceCon.strChargeType = "" Then
        varType = Split(Replace(mobjBalanceAll.strAllChargeType, "'", ""), ",")
    Else
        varType = Split(Replace(mobjBalanceCon.strChargeType, "'", ""), ",")
    End If
    For i = 0 To UBound(varType)
        If InStr("," & mty_ModulePara.strOwnerPayFeeType & ",", "," & varType(i) & ",") = 0 Then
            blnSelfFee = False
            Exit For
        End If
    Next i
    mobjBalanceCon.blnCurBalanceOwnerFee = blnSelfFee
    picOwnerFee.Visible = blnSelfFee
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearListData(Optional ByVal blnForceDel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除列表数据
    '编制:刘兴洪
    '日期:2015-02-05 18:13:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyBalance As TY_Balance_Infor
    
    If Not mblnConsChange And blnForceDel = False Then Exit Sub
    
    txtBalance(Idx_本次结帐).Text = ""
    txtBalance(Idx_本次未结).Text = ""
    txtBalance(Idx_结帐说明).Text = ""
    txtBalance(Idx_冲预交).Text = ""
    
    Set mrsFeeList = Nothing
    Set mrsBalance = Nothing
    mBalanceInfor = tyBalance
    Call ClearFeeList   '清除费列表
    Call ClearAdjustBalance '清除结算列表
    Call ClearAdjustDeposit  '清除预交列表
    Call InitPatiBalanceVariableCon
    Call SetOperationCtrl(3)
    Call LoadCurOwnerPayInfor
End Sub

Private Sub ExecuteFeeQuery(ByVal lngControlID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行费用查询
    '入参:lngControlID-菜单控件的ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-02-12 10:33:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long
    Dim objCon As clsBalanceCon
    Dim EditType As gBalanceBill
    
    If (mblnConsChange Or mrsInfo Is Nothing) And (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) And chkCancel.Value = 0 Then
        MsgBox "当前不存在结帐费用,请检查!", vbInformation, gstrSysName
        Exit Sub
    End If
    Set objCon = New clsBalanceCon
    With objCon
        .blnCurBalanceOwnerFee = mobjBalanceCon.blnCurBalanceOwnerFee
        .strBaby = mobjBalanceCon.strBaby
        .strChargeType = mobjBalanceCon.strChargeType
        .lng病人ID = IIf(mobjBalanceCon.lng病人ID = 0, mPatiInfor.lng病人ID, mobjBalanceCon.lng病人ID)
        .bytKind = mobjBalanceCon.bytKind
        .dtBeginDate = mobjBalanceCon.dtBeginDate
        .dtEndDate = mobjBalanceCon.dtEndDate
        .strClass = mobjBalanceCon.strClass
        .strDeptIDs = mobjBalanceCon.strDeptIDs
        .strItem = mobjBalanceCon.strItem
        .strDiag = mobjBalanceCon.strDiag
        .strTime = mobjBalanceCon.strTime
    End With
    lng结帐ID = IIf(mBalanceInfor.lng冲销ID <> 0, mBalanceInfor.lng冲销ID, mBalanceInfor.lng结帐ID)
    
    If chkCancel.Value = 1 Then
        EditType = g_Ed_结帐作废
    Else
        EditType = mEditType
    End If
    
    Select Case lngControlID
    Case conMenu_View_List ' "明细表"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_明细表)
    Case conMenu_View_ListItem ' "项目明细"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_项目明细)
    Case conMenu_View_SplitType ' "分类表"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_分类表)
    Case conMenu_View_SplitMonth ' "分月表"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_分月表)
    Case conMenu_View_DayBill ' "逐日单据"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_逐日单据)
    Case conMenu_View_DayFM ' "逐日费目"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_逐日费用)
    Case conMenu_View_Balance '结帐表
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng结帐ID, mlngModul, mstrPrivs, g_Ed_结帐表)
    End Select
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim dblMoney As Double
    Dim lngCash As Long
    Dim i As Long
    Dim bytSetFocus As Byte '1-预交;0-缴款
    '执行操作
    Select Case Control.ID
    Case conMenu_View_List ' "明细表"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_ListItem ' "项目明细"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_SplitType ' "分类表"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_SplitMonth ' "分月表"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_DayBill ' "触日单据"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_DayFM ' "触日费目"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_Balance '明细帐
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_Edit_ClinicToHos
        If InStr(1, mstrPrivs, ";门诊费用转住院;") = 0 Then Exit Sub
        If mobjInPati Is Nothing Then
            Err = 0: On Error Resume Next
            Set mobjInPati = CreateObject("zl9InPatient.clsInPatient")
            
            If Err <> 0 Then
                MsgBox "注意:" & vbCrLf & "   住院病人部件(zl9InPatient)创建失败,请与系统管理员联系!"
                Exit Sub
            End If
        End If
        Call mobjInPati.zlOutFeeToInFee(Me, gcnOracle, glngSys, mlngModul, mstrPrivs, gstrDBUser, mobjBalanceCon.lng病人ID, 0)
    Case conMenu_Edit_NotUseDeposit   '不使用预交款(C)
        '0-清除所有冲预交;1-按缺省使用预交款;2-按指定金额来冲预交(按时间先后来分摊）;3-全冲
        Call RecalcDepositMoney(0): mbln已报价 = False: GoTo GoFullDeposit:
        bytSetFocus = 0
    Case conMenu_Edit_MoneyUseDeposit   '按结帐金额使用预交(L)
        Call RecalcDepositMoney(0)
        bytSetFocus = 0
        mblnNotChange = True
        txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
        txtBalance(Idx_冲预交).BackColor = &H80000005
        mBalanceInfor.bln预交刷卡 = False
        mblnNotChange = False
        Call LoadCurOwnerPayInfor
        If bytSetFocus = 1 Then
            If txtBalance(Idx_冲预交).Enabled And txtBalance(Idx_冲预交).Visible Then txtBalance(Idx_冲预交).SetFocus
        Else
            Call txtBalance_Validate(Idx_冲预交, False)
            If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
        End If
        For i = 1 To vsBlance.Rows - 1
            If Val(vsBlance.RowData(i)) = 999 Then
                lngCash = i
                Exit For
            End If
        Next i
        If lngCash > 0 Then
        
            mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - Val(vsBlance.TextMatrix(lngCash, vsBlance.ColIndex("结算金额"))), 5)
            mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + Val(vsBlance.TextMatrix(lngCash, vsBlance.ColIndex("结算金额"))), 5)
            vsBlance.TextMatrix(lngCash, vsBlance.ColIndex("结算金额")) = Format(0, gstrDec)
        End If
        dblMoney = RoundEx(mBalanceInfor.dbl未付合计, 6)
        Call RecalcDepositMoney(2, dblMoney)
        mbln已报价 = False
        
        GoTo GoFullDeposit:
    Case conMenu_Edit_UseAllDeposit   '使用所有预交款(A)
        bytSetFocus = 0
        Call RecalcDepositMoney(3): mbln已报价 = False: GoTo GoFullDeposit:
    Case conMenu_File_Exit: Unload Me '退出
    Case Else
    End Select
    Exit Sub
GoFullDeposit:
    mblnNotChange = True
    txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
    txtBalance(Idx_冲预交).BackColor = &H80000005
    mBalanceInfor.bln预交刷卡 = False
    mblnNotChange = False
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor(True)
    If bytSetFocus = 1 Then
        If txtBalance(Idx_冲预交).Enabled And txtBalance(Idx_冲预交).Visible Then txtBalance(Idx_冲预交).SetFocus
    Else
        Call txtBalance_Validate(Idx_冲预交, False)
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    Top = txtPatient.Top - 60
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
    staThis.Top = Me.ScaleHeight - Me.staThis.Height
End Sub

Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
'    If CommandBar.Title = "结帐表" Then
'        With CommandBar.Controls
'            .DeleteAll
'            Set mcbrControl = .Add(xtpControlButton, conMenu_View_List, "明细表")
'            mcbrControl.IconId = M_VIEW_ICO
'            Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListItem, "项目明细")
'            mcbrControl.IconId = M_VIEW_ICO
'            Set mcbrControl = .Add(xtpControlButton, conMenu_View_SplitType, "分类表")
'            mcbrControl.IconId = M_VIEW_ICO
'            Set mcbrControl = .Add(xtpControlButton, conMenu_View_SplitMonth, "分月表")
'            mcbrControl.IconId = M_VIEW_ICO
'            Set mcbrControl = .Add(xtpControlButton, conMenu_View_DayBill, "触日单据")
'            mcbrControl.IconId = M_VIEW_ICO
'            Set mcbrControl = .Add(xtpControlButton, conMenu_View_DayFM, "触日费目")
'            mcbrControl.IconId = M_VIEW_ICO
'        End With
'    End If
End Sub

 

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '执行操作
    Select Case Control.ID
    Case conMenu_View_Balance   '结帐表
        Control.Enabled = Not mblnLockScreen
    Case conMenu_View_List      '明细表
    Case conMenu_View_SplitType '分类表
    Case conMenu_View_SplitMonth   '分月表
    Case conMenu_View_DayBill   '触日单据
    Case conMenu_View_DayBill   '触日单据
    Case conMenu_View_DayFM '触日费目
    Case conMenu_File_Exit '退出
        If mEditType <> g_Ed_单据查看 Then
            Control.Visible = Not mBalanceInfor.blnSaveBill
        End If
    Case conMenu_Edit_ClinicToHos
        Control.Visible = mEditType = g_Ed_住院结帐
    Case Else
    End Select
End Sub

Private Sub chkCancel_Click()
    If mblnNotChange Then Exit Sub
    
    If mBalanceInfor.blnSaveBill = True Then
        MsgBox "已经保存了结帐单据,请先完成当前结帐再切换作废模式!", vbInformation, gstrSysName
        mblnNotChange = True
        chkCancel.Value = 0
        mblnNotChange = False
        Exit Sub
    End If
    
    Call frmPatiBalanceSplit.ShowMe(Me, mEditType, mstrPrivs, , , , , , , , True)
    mblnNotChange = True
    chkCancel.Value = 0
    mblnNotChange = False
End Sub

Private Sub chkDeposit_Click()
    If mblnNotChange Then Exit Sub
    If Not (mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1) Then Exit Sub
    
    If chkDeposit.Value = 1 Then
        txtBalance(Idx_冲预交).Text = Format(Val(chkDeposit.Tag), "0.00")
        mBalanceInfor.dbl冲预交合计 = Val(chkDeposit.Tag)
     Else
        txtBalance(Idx_冲预交).Text = "0.00"
        mBalanceInfor.dbl冲预交合计 = 0
    End If
    
    Call LoadCurOwnerPayInfor
    If txtReceive.Enabled And txtReceive.Visible Then
        txtReceive.SetFocus
        zlControl.TxtSelAll txtReceive
    End If
    
End Sub

Private Sub cmdCancel_Click()
 
    '取消操作
    If mintPreEditType <> -1 Then
        mEditType = mintPreEditType '恢复上次操作
        Call NewBill
        If picPati.Enabled And txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        zlControl.TxtSelAll txtPatient
        mintPreEditType = -1
        Exit Sub
    End If
    
    If mEditType = g_Ed_单据查看 _
        Or mEditType = g_Ed_取消结帐 _
        Or mEditType = g_Ed_结帐作废 _
        Or mEditType = g_Ed_重新作废 _
        Or mEditType = g_Ed_重新结帐 Then
        '退出
        Unload Me: Exit Sub
    End If
    
    If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Then
        If mblnNotClearBill And mty_ModulePara.bln结帐后不清信息 Then
            '当前为结帐不清除票据,则取消时,清除
             If mrsInfo Is Nothing Then
                Call NewBill: mblnNotClearBill = False: Exit Sub
             End If
             If mrsInfo.State <> 1 Then
                Call NewBill: mblnNotClearBill = False: Exit Sub
             End If
        End If
        
        If chkCancel.Value = Checked And txtPatient.Text <> "" Then
           '当前操作为作废 ,则提示是否要退出
            If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Unload Me: Exit Sub
        End If
        
        '已经验证医保的操作
        If mYBInFor.bytMCMode = 1 Then
            If MsgBox("确实要取消当前病人身份验证吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            If YBIdentifyCancel Then Call NewBill
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                
            Exit Sub
            '不退出窗体,以便选择其它病人进行身份验证
        End If
        If Not mrsInfo Is Nothing Then
            If Val(txtBalance(Idx_本次结帐).Text) <> 0 And mrsInfo.State = adStateOpen Then
                If MsgBox("该病人尚未确定结帐,确实取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                Call NewBill
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                Exit Sub
            End If
        End If
        If txtPatient.Text <> "" Then
            If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdColSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    If Button <> 1 Then Exit Sub
    vRect = zlControl.GetControlRect(cmdColSet.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + cmdColSet.Height
    Call Grid.SetColVisible(Me, Me.Caption, vsDeposit, lngLeft, lngTop, cmdColSet.Height)
    zl_vsGrid_Para_Save mlngModul, vsDeposit, Me.Name, "预交列表"
End Sub

Private Function SaveDeposit(ByRef bln预交 As Boolean, Optional ByVal blnNoRecal As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存预交款
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-27 11:00:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnChargeEnd As Boolean, objSetFocus As Object
    Dim tyBrushCard As TY_BrushCard
       
    On Error GoTo errHandle
    
    If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 _
    Or chkCancel.Value = 1 And chkCancel.Visible Then Exit Function
    
    If mBalanceInfor.bln预交刷卡 Then
        If txtReceive.Enabled And txtReceive.Visible Then
            txtReceive.SetFocus
            zlControl.TxtSelAll txtReceive
        End If
        Exit Function
    End If
    
    Screen.MousePointer = 99
    mblnNotChange = True
    LockedScreen True
    mblnNotChange = False
 

    '先判断是否存在预交款刷卡的，则先处理预交款
    If Not CheckDepositValied(bln预交) Then
        LockedScreen False
        Set objSetFocus = txtBalance(Idx_冲预交)
        If Not objSetFocus Is Nothing Then
            If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
        End If
        zlControl.TxtSelAll objSetFocus
        Screen.MousePointer = 0
        Exit Function
    End If
    If Not bln预交 Then
        Screen.MousePointer = 0
        LockedScreen False
        SaveDeposit = True: Exit Function
    End If
    
    If Not SaveBalaceCharge(True, tyBrushCard, blnChargeEnd, objSetFocus) Then
        LockedScreen False
        If Not objSetFocus Is Nothing Then
            If objSetFocus.Enabled And objSetFocus.Visible Then
                objSetFocus.SetFocus
            End If
        End If
        Screen.MousePointer = 0
        Exit Function
    End If

    If blnChargeEnd And mEditType = g_Ed_重新结帐 Then Unload Me: Exit Function
    
    LockedScreen False
 
    '0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
    If mBalanceInfor.blnSaveBill Then
       Call SetOperationCtrl(2)
    Else
       Call SetOperationCtrl(0)
    End If
    
    If blnChargeEnd Then
        Call NewBill
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        SaveDeposit = True
        Exit Function
    End If
    
    If Not blnNoRecal Then Call LoadIntendBalance
    Call LoadCurOwnerPayInfor(True)
    If txtReceive.Enabled And txtReceive.Visible Then
        txtReceive.SetFocus
    End If
    
    SaveDeposit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockedScreen False
End Function

Private Function CheckInputValied() As Boolean
    On Error GoTo errHandle
    
    If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐 Then
        If Val(txtOwe.Text) <> 0 Then
            If txtOwe.ForeColor = vbRed Then
                MsgBox "结算缴款金额过多,请调整金额!", vbInformation, gstrSysName
                If vsBlance.Enabled And vsBlance.Visible Then vsBlance.SetFocus
                Exit Function
            Else
                MsgBox "结算缴款金额不足,请补足金额!", vbInformation, gstrSysName
                If vsBlance.Enabled And vsBlance.Visible Then vsBlance.SetFocus
                Exit Function
            End If
        End If
    End If
    
    CheckInputValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function SaveBalanceData(Optional objInCard As Card, Optional lngRow As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结帐数据
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-30 09:44:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln预交 As Boolean, tyBrushCard As TY_BrushCard
    Dim blnChargeEnd As Boolean, blnFind As Boolean
    Dim objSetFocus As Object, blnSaved As Boolean
    Dim strErrMsg As String, i As Long
    Dim blnNotClearPati As Boolean, strTime() As String
    Dim blnHaveFee As Boolean, intMaxTime As Integer
    Dim objCard As Card, strBlank As String
    Dim bln门诊留观病人 As Boolean, lng病人ID As Long, str姓名 As String, dbl费用余额 As Double
    
    On Error GoTo errHandle
    If mEditType = g_Ed_取消结帐 Then Exit Function
    
    
    If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1 And chkCancel.Visible Then
        
        If objInCard Is Nothing Then
            If CheckInputValied = False Then Exit Function
        End If
        If ExecuteBalaceCancel(GetCard(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算方式")))) = False Then Exit Function
        mintSucces = mintSucces + 1
        mbln已报价 = False
        SaveBalanceData = True: Exit Function
    Else
       
        If mBalanceInfor.bln预交刷卡 = False And Val(txtBalance(Idx_冲预交).Text) <> 0 Then
             If DepositMonyVerfy(True) = False Then Screen.MousePointer = 0: Exit Function
    '        MsgBox "预交金额未验证,请重新输入缴款金额!", vbInformation + vbOKOnly, gstrSysName
    '        Call txtBalance_Validate(Idx_冲预交, False)
    '        If txtReceive.Visible And txtReceive.Enabled Then txtReceive.SetFocus
    '        Exit Function
        End If
        
        If objInCard Is Nothing Then
            If CheckInputValied = False Then Exit Function
        End If
        
    End If
    
    If CheckChargeAudit(mPatiInfor.lng病人ID, True, mobjBalanceCon.strTime) = False Then Exit Function
    
    Screen.MousePointer = 99
    
    If mPatiInfor.lng病人ID = 0 Or Trim(txtPatient.Text) = "" Then
        Screen.MousePointer = 0
        MsgBox "未输入本次结帐的病人,不允许结帐操作", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    blnHaveFee = False
    If Not mrsFeeList Is Nothing Then
       If mrsFeeList.State = 1 Then
            If mrsFeeList.RecordCount <> 0 Then blnHaveFee = True
       End If
    End If
    
    If blnHaveFee = False And mEditType <> g_Ed_重新结帐 Then
        Screen.MousePointer = 0
        MsgBox "病人不存在需要结帐的费用,请调整结帐条件后再试", vbInformation + vbOKOnly, gstrSysName
        If cmdMore.Enabled And cmdMore.Visible Then cmdMore.SetFocus
        Exit Function
    End If
    
    If objInCard Is Nothing Then
        Set objCard = IDKindPaymentsType.GetCurCard
    Else
        Set objCard = objInCard
    End If
    
    If Not objCard Is Nothing Then
        If (objCard.消费卡 Or objCard.是否存在帐户) And (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐) Then
            If mPatiInfor.bln退款标志 And objCard.是否转帐及代扣 = False Then
                With vsBlance
                    blnFind = False
                    For i = 1 To .Rows - 1
                        If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 1 Then
                            If Not objCard.消费卡 And objCard.接口序号 <> 0 Then '消费卡,已经检查,不用再处理
                                If .TextMatrix(i, .ColIndex("结算方式")) = objCard.结算方式 Then blnFind = True
                            End If
                        End If
                    Next
                    If blnFind Then
                        Screen.MousePointer = 0
                        MsgBox objCard.结算方式 & " 已经支付了,不能再用" & objCard.结算方式 & "进行支付", vbOKOnly + vbDefaultButton1, gstrSysName
                        Exit Function
                    End If
                End With
            End If
        End If
    End If
  
    
    '先判断是否存在预交款刷卡的，则先处理预交款
    If Not CheckDepositValied(bln预交) Then
        If txtBalance(Idx_冲预交).Enabled And txtBalance(Idx_冲预交).Visible Then txtBalance(Idx_冲预交).SetFocus
        zlControl.TxtSelAll txtBalance(Idx_冲预交):
        Exit Function
    End If
    
    Call LedVoiceSpeak(False)
    
    If Not bln预交 Then
        If CheckCurBalanceIsValied(tyBrushCard, , objSetFocus, _
                                    objCard, IIf(lngRow <> 0, Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("结算金额"))), 0)) = False Then
            If Not objSetFocus Is Nothing Then
                If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
                If UCase(TypeName(objSetFocus)) = UCase("txtEdit") Then
                    zlControl.TxtSelAll objSetFocus
                End If
            End If
            Exit Function
        End If
        If CheckDepositFactValied = False Then Exit Function
        If Not objCard Is Nothing Then
             If objCard.消费卡 Then SaveBalanceData = True: Exit Function
        End If
    End If
    
    If mblnNotify = False Then
        If MsgBox("你确认要对该病人进行结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        
        mblnPrintInvoice = False
        If Not mobjBalanceCon.blnCurBalanceOwnerFee Then   '非自费费用时,要打印发票
            If Not (mYBInFor.intInsure <> 0 And MCPAR.医保接口打印票据) Then
                '保险病人根据使用类别来进行确认了
                Select Case mobjFactProperty.打印方式
                Case 0  '不打印
                Case 1
                    mblnPrintInvoice = True '自动打印
                Case 2  '提示打印
                    If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        Else
            If Not mty_ModulePara.blnNotPrintInvioce Then
                Select Case mobjFactProperty.打印方式
                Case 0  '不打印
                Case 1
                    mblnPrintInvoice = True '自动打印
                Case 2  '提示打印
                    If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        End If
        mblnNotify = True
    End If
    LockedScreen True
    blnSaved = SaveBalaceCharge(bln预交, tyBrushCard, blnChargeEnd, objSetFocus, objCard, lngRow)
    LockedScreen False
    
    mbln已报价 = False
    If blnChargeEnd Then
        If mEditType = g_Ed_重新结帐 Then Unload Me: Exit Function
    End If
    
    '0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
    If mBalanceInfor.blnSaveBill Then
       Call SetOperationCtrl(2)
    Else
       Call SetOperationCtrl(0)
    End If
     
    If blnChargeEnd Then
        mblnNotify = False
        Call AddNoToCombox  '加载单据号
        Call SetOperationCtrl(0)
        mintSucces = mintSucces + 1
        If mintPreEditType <> -1 Then mEditType = mintPreEditType
        mlngPatientID = 0
        mBalanceInfor.blnSaveBill = False
        picOwnerFee.Visible = False
        staThis.Panels(3).Text = "上次结帐:" & Format(mBalanceInfor.dbl当前结帐, "0.00")
        
        bln门诊留观病人 = Val(NVL(mrsInfo!病人性质)) = 1
        lng病人ID = Val(NVL(mrsInfo!病人ID)): str姓名 = NVL(mrsInfo!姓名)
        
        If mbln连续结帐 Or mobjBalanceCon.blnCurBalanceOwnerFee Then
            '不清除病人信息
            If mobjBalanceCon.blnCurBalanceOwnerFee Then
                lblPrevious.Visible = True
                strBlank = ""
                For i = 1 To (12 - Len(Format(mBalanceInfor.dbl当前结帐, "0.00"))) / 2
                    strBlank = strBlank & " "
                Next i
                lblPrevious.Caption = "上次自费结帐:" & Format(mBalanceInfor.dbl当前结帐, "0.00")
                lblPrevious.Left = lblCaculated.Left
                lblPrevious.Top = lblOwe.Top + 30
                txtReceive.Text = ""
            End If
           If ShowBalance(True, strErrMsg, blnNotClearPati) = False Then
                cmdOK.Enabled = False
                MsgBox "在当前条件下,病人不存在需要结帐的费用！", vbInformation, gstrSysName
                If cmdMore.Visible And cmdMore.Enabled Then cmdMore.SetFocus
                Call SetBatchControl(False)
                SaveBalanceData = blnSaved
                Exit Function
           End If
           
           If mobjBalanceCon.strTime = "" Then
                intMaxTime = mPatiInfor.lng主页ID
            Else
                strTime = Split(mobjBalanceCon.strTime, ",")
                For i = 0 To UBound(strTime)
                    If Val(strTime(i)) > intMaxTime Then intMaxTime = Val(strTime(i))
                Next i
            End If
            
            Call LoadDefaultOutStatu(mPatiInfor.lng病人ID, intMaxTime)
            Call Load余额信息(Val(NVL(mrsInfo!病人ID)), IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2))
            Call ReInitPatiInvoice
            
            mblnChargeEnd = True
        Else
            '刘兴洪:27503
            If mty_ModulePara.bln结帐后不清信息 Then
                Set mrsInfo = New ADODB.Recordset
                If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '主要是要保留信息,在确定后需要减判刑断
                Dim strTemp As String
                strTemp = txtInvoice.Text
                Call ReInitPatiInvoice
                txtInvoice.Text = strTemp   '主要是不要清空上次的发票,新的发票放在.tag中,在改变病人时,直接从这个地方读取
                mblnNotClearBill = True
                Call SetBatchControl(False)
            Else
                Call LoadBalanceBill
                Call ReInitPatiInvoice(Not mobjFactProperty.启用使用类别)
            End If
        End If
        
        '139063，门诊留观病人如果存在未结的住院费用，则进行提示
        If bln门诊留观病人 And mEditType = g_Ed_门诊结帐 Then
            dbl费用余额 = GetRemainderMoney(lng病人ID, 2)
            If dbl费用余额 > 0 Then
                MsgBox "注意：" & vbCrLf & _
                       "    当前病人『" & str姓名 & "』还存在未结清的住院费用，注意对其进行结账！", vbInformation, gstrSysName
            End If
        End If
        
        staThis.Panels(2) = "操作完毕，请输入其它病人标识！"
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    ElseIf blnSaved Then
        If bln预交 And txtBalance(Idx_冲预交).Enabled And txtBalance(Idx_冲预交).Visible Then
            txtBalance(Idx_冲预交).SetFocus
            zlControl.TxtSelAll txtBalance(Idx_冲预交)
        ElseIf txtReceive.Enabled And txtReceive.Visible Then
            txtReceive.SetFocus
            zlControl.TxtSelAll txtReceive
        End If
    ElseIf Not objSetFocus Is Nothing Then
        If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
        zlControl.TxtSelAll objSetFocus
    End If
    
    SaveBalanceData = blnSaved
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockedScreen False
    '0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
    If mBalanceInfor.blnSaveBill Then
       Call SetOperationCtrl(2)
    Else
       Call SetOperationCtrl(0)
    End If
End Function

Private Sub SetBatchControl(ByVal blnState As Boolean)
    mblnBatchState = Not blnState
    cmdOK.Enabled = blnState
    cmdCancel.Enabled = blnState
    cmdMore.Enabled = blnState And InStr(mstrPrivs, ";结帐设置;") > 0
    cmdNext.Enabled = blnState
    '大二院用作计算器使用，不进行屏蔽
'    TXTRECEIVE.Enabled = blnState
    txtBalance(Idx_结帐说明).Enabled = blnState
    txtBalance(Idx_冲预交).Enabled = blnState
    txtBalance(Idx_本次结帐).Enabled = blnState
    txtBalance(Idx_本次结帐).Locked = InStr(mstrPrivs, ";结帐设置;") = 0
    txtBegin.Enabled = False '不允许修改日期(118827,在结帐设置中更改)
    txtEnd.Enabled = False
    txtPatiBegin.Enabled = blnState
    txtPatiEnd.Enabled = blnState
    cboPatiNums.Enabled = blnState And InStr(mstrPrivs, ";结帐设置;") > 0
    opt中途.Enabled = blnState
    opt出院.Enabled = blnState
    If blnState Then
        txtBalance(Idx_结帐说明).BackColor = &H80000005
        txtBalance(Idx_冲预交).BackColor = &H80000005
    Else
        txtBalance(Idx_结帐说明).BackColor = &H8000000F
        txtBalance(Idx_冲预交).BackColor = &H8000000F
    End If
    txtBalance(Idx_本次结帐).BackColor = IIf(txtBalance(Idx_本次结帐).Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
        
End Sub

Private Sub AddNoToCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载单据民给Combox控件中
    '编制:刘兴洪
    '日期:2015-02-11 17:30:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    '加入单据历史记录(所有类型单据)
    On Error GoTo errHandle
    strTmp = mBalanceInfor.strNO
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdMore_Click()
    Dim blnNotPati As Boolean, intMaxTime As Integer
    Dim i As Long, j As Integer, objCard As Card
    Dim arrTime() As String, strTime() As String
    Dim dbl未付累计 As Double
    
    blnNotPati = False
    If mrsInfo Is Nothing Then blnNotPati = True
    If blnNotPati = False Then
        If mrsInfo.State = 0 Then blnNotPati = True
    End If
    
    If blnNotPati Then
        MsgBox "没有确定结帐病人,不能进行结帐设置！", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    
    If frmSetBalance.ShowMe(Me, IIf(mEditType = g_Ed_门诊结帐, 0, 1), mPatiInfor.lng病人ID, mYBInFor.intInsure, mobjBalanceAll, mobjBalanceCon) = False Then
        Exit Sub
    End If
    
    mblnNotChange = True
    txtBegin.Text = Format(mobjBalanceCon.dtBeginDate, "yyyy-mm-dd")
    txtEnd.Text = Format(mobjBalanceCon.dtEndDate, "yyyy-mm-dd")
    mblnNotChange = False
    
    '完成设置，重新读取
    If mPatiInfor.bln连续结帐 Then
        mbln连续结帐 = True
        dbl未付累计 = mPatiInfor.dbl未付累计
    End If
    
    cboPatiNums.Text = ""
    For i = 1 To cboPatiNums.ListCount
        If InStr("," & mobjBalanceCon.strTime & ",", "," & Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) & ",") > 0 Or mobjBalanceCon.strTime = "" Then
            cboPatiNums.Nodes.Item(i).Checked = True
            If cboPatiNums.Nodes.Item(i).Key <> "Root" Then
                cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
            End If
        Else
            cboPatiNums.Nodes.Item(i).Checked = False
        End If
    Next i
    If cboPatiNums.Text <> "" Then cboPatiNums.Text = Mid(cboPatiNums.Text, 2)
    
    If Not mrsInfo Is Nothing Then
        Call SetPatiNums
    End If
    
    If mbln连续结帐 Then
        mPatiInfor.bln连续结帐 = mbln连续结帐
        mPatiInfor.dbl未付累计 = dbl未付累计
    End If
    
    If mobjBalanceCon.strTime = "" Then
        intMaxTime = mPatiInfor.lng主页ID
    Else
        strTime = Split(mobjBalanceCon.strTime, ",")
        For i = 0 To UBound(strTime)
            If Val(strTime(i)) > intMaxTime Then intMaxTime = Val(strTime(i))
        Next i
    End If
    
    Call LoadDefaultOutStatu(mPatiInfor.lng病人ID, intMaxTime, True)
    
    If Not ShowBalance() Then
        If mrsInfo.State <> 1 Then
            txtPatient.Locked = False: txtPatient.Text = ""
           Call NewBill
           If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
           Exit Sub
        End If
        cmdOK.Enabled = False
        MsgBox "在当前条件下,病人不存需要结帐的费用！", vbInformation, gstrSysName
        Call cmdMore_Click
        mbln连续结帐 = False
        Exit Sub
    End If
    cmdOK.Enabled = True
    mbln连续结帐 = False
    '确定结帐顺序
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    
    If vsBlance.Enabled And vsBlance.Visible Then
        vsBlance.SetFocus
    End If
    
    If Val(txtBalance(Idx_冲预交).Text) <> 0 And txtBalance(Idx_冲预交).Enabled And txtBalance(Idx_冲预交).Visible Then
        txtBalance(Idx_冲预交).SetFocus
        zlControl.TxtSelAll txtBalance(Idx_冲预交)
    End If
    
    If cmdYBBalance.Visible And cmdYBBalance.Enabled Then cmdYBBalance.SetFocus
    
    mblnConsChange = False
    mbln已报价 = False
End Sub

Private Sub cmdNext_Click()
   If chkCancel.Value = 1 Then Exit Sub
   If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) Then Exit Sub
   mbln连续结帐 = True
   If SaveBalanceData = False Then Exit Sub
   mbln连续结帐 = False
End Sub

Private Sub cmdOK_Click()
    mbln连续结帐 = False
    If mEditType = g_Ed_单据查看 Then Unload Me: Exit Sub
    If mEditType = g_Ed_取消结帐 Then
        If DeleteBalance = False Then Exit Sub
        mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If
    If SaveBalanceData = False Then Exit Sub
End Sub

Private Sub cmdTools_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Call AddPopu
End Sub

Private Sub cmdYB_Click()
    '门诊病人结帐前的身份验证(成都医保还支持住院病人医保身份验证)
    Dim lng病人ID As Long, bytMode As Byte
    Dim strMessage As String, intInsure As Integer
    Dim strPatiName As String
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng病人ID = Val(NVL(mrsInfo!病人ID))
    End If
    strPatiName = Trim(txtPatient.Text)
    Call NewBill
    txtPatient.Text = strPatiName
    
    bytMode = 0
    If mblnMC_TwoMode Then
        If InStr(mstrPrivs, ";门诊费用结帐;") = 0 Then
            bytMode = 4
        Else
            If zlCommFun.ShowMsgbox("医保身证验证", "请选择病人身份验证模式。", "!住院医保(&Z),门诊医保(&M)", Me, vbInformation) = "住院医保" Then
                bytMode = 4
            End If
        End If
    End If
        
    '刘兴洪:门诊转住院费用时加入
    mYBInFor.strYBPati = gclsInsure.Identify(bytMode, lng病人ID, intInsure)
    mYBInFor.intInsure = intInsure
    
    If mYBInFor.strYBPati = "" Then GoTo ExceptionHand
    cmdOK.Enabled = False   '问题:43776
    
    mYBInFor.bytMCMode = IIf(bytMode = 0, 1, 2) '必须在LoadPatientInfo之前
    
    If mYBInFor.bytMCMode = 1 Then
        'lng病人ID:49084
        If Not gclsInsure.GetCapability(support门诊结帐, lng病人ID, intInsure) Then
            strMessage = "病人当前险类不支持门诊医保结帐。": GoTo ExceptionHand
        End If
    End If
    
    'New:空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mYBInFor.strYBPati, ";")) >= 8 Then lng病人ID = Val(Split(mYBInFor.strYBPati, ";")(8))
    If lng病人ID <> 0 Then
        txtPatient.Text = "-" & lng病人ID
        Call LoadPatientInfo(IDKind.GetCurCard, False, intInsure)
        If mrsInfo.State = 0 Then GoTo ExceptionHand
    Else
        strMessage = "病人身份验证成功,但未发现病人的帐户信息!" & vbCrLf & "可能是病人入院时没有进行验证,不能进行保险结算！"
        GoTo ExceptionHand
    End If
    Exit Sub
ExceptionHand:
    If strMessage <> "" Then Call MsgBox(strMessage, vbInformation, gstrSysName)
    Set mrsInfo = New ADODB.Recordset
    mYBInFor.strYBPati = "": mYBInFor.bytMCMode = 0
    txtPatient.Text = "": txtPatient.SetFocus
    cmdOK.Enabled = True
    Call NewBill
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Function ExcuteInsureSwapInteface(ByVal lng结帐ID As Long, ByVal cllSaveBill As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保交易接口
    '入参:cllSaveBill-保存单据的sql
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-13 15:11:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, i As Long, str结算方式 As String
    Dim blnTrans As Boolean, intInsure As Integer, strAdvance As String
    Dim blnTransMC As Boolean, blnMark As Boolean
    Dim cur个人帐户 As Currency, cur医保基金 As Currency
    Dim blnInsureCheck As Boolean
    On Error GoTo errHandle
    
    intInsure = mYBInFor.intInsure
    '非医保或结自费费用,不能执行
    If intInsure = 0 Or mobjBalanceCon.blnCurBalanceOwnerFee Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllSaveBill.Count
        zlAddArray cllPro, cllSaveBill(i)
    Next
    
    str结算方式 = GetMedicareStr(cur个人帐户, cur医保基金)
    If 医保数据更正(Val(NVL(mrsInfo!病人ID)), lng结帐ID, str结算方式, False, cllPro) = False Then Exit Function
    
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    '调用医保接口
    blnTransMC = False
    If mYBInFor.bytMCMode = 1 Then
        '门诊医保结算
        strAdvance = ""
        If cur个人帐户 <> 0 Or cur医保基金 <> 0 Or MCPAR.门诊必须传递明细 Then
            Call SetCmdStatus(False)
            If Not gclsInsure.ClinicSwap(lng结帐ID, cur个人帐户, cur医保基金, 0, 0, intInsure, strAdvance) Then
                Call SetCmdStatus(True)
                gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
            End If
            Call SetCmdStatus(True)
            blnTransMC = True
        End If
        GoTo SaveEnd:
    End If
    '住院医保结算
    Call SetCmdStatus(False)
    If Not gclsInsure.SettleSwap(lng结帐ID, intInsure, strAdvance) Then
        Call SetCmdStatus(True)
        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
    Else
        Call SetCmdStatus(True)
        blnTransMC = True
    End If

SaveEnd:
    If strAdvance <> "" Then
        If zlInsure_Check(str结算方式, strAdvance) Then
            blnInsureCheck = True
            Call 医保数据更正(Val(NVL(mrsInfo!病人ID)), lng结帐ID, strAdvance, False, Nothing)
CheckAgain:
            blnMark = False
            For i = 1 To vsBlance.Rows - 1
                If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("类型"))) = 2 Then
                    Call DeletePayInfor(i, True)
                    blnMark = True
                    Exit For
                End If
            Next i
            mbln已报价 = False '主要预结算和结算的不一致，需要进行再次报
            If blnMark = True Then GoTo CheckAgain
        End If
    End If
    mBalanceInfor.blnSaveBill = True
    gcnOracle.CommitTrans: blnTrans = False
    If blnTransMC Then
        Call gclsInsure.BusinessAffirm(IIf(mYBInFor.bytMCMode = 1, 交易Enum.Busi_ClinicSwap, 交易Enum.Busi_SettleSwap), True, intInsure)
    End If
    Set cllSaveBill = New Collection
    Screen.MousePointer = 0
    ExcuteInsureSwapInteface = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    Call SetCmdStatus(True)
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mYBInFor.bytMCMode = 1, 交易Enum.Busi_ClinicSwap, 交易Enum.Busi_SettleSwap), False, intInsure)
    End If
End Function

Private Function 医保数据更正(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal str医保结算 As String, ByVal bln作废 As Boolean, _
    ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保数据校对更正
    '返回:校对成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 17:45:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    If bln作废 Then
        'Zl_门诊退费结算_Modify
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 3 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str医保结算 & "')"
        '  退预交_In     病人预交记录.冲预交%Type := Null,
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成退费_In Number:=0
        ') As
        '  ------------------------------------------------------------------------------------------------------------------------------
        '  --功能:收费结算时,修改结算的相关信息
        '  --操作类型_In:
        '  --   1-普通退费方式:
        '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
        '  --   2.三方卡退费结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退预交_In: 传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  --   4-消费卡结算:
        '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  -- 误差金额_In:存在误差费时,传入
        '  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
        '  ------------------------------------------------------------------------------------------------------------------------------
     Else
  
        'Zl_病人结帐结算_Modify
        strSQL = "Zl_病人结帐结算_Modify("
        '  操作类型_In     Number,
        '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        strSQL = strSQL & "" & 2 & ","
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "" & IIf(str医保结算 = "", "NULL", "'" & str医保结算 & "'") & ","
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "NULL,"
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  结帐类型_In     Number := 2,(1-门诊结帐;2-住院结帐)
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0: (1-完成收费;0-未完成收费)
        strSQL = strSQL & "0)"
     End If
     
    If cllPro Is Nothing Then
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        zlAddArray cllPro, strSQL
    End If
    医保数据更正 = True
End Function
Public Function zlInsure_Check(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")
    
    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetMedicareStr(ByRef cur个人帐户 As Currency, cur医保基金 As Currency) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回保险结算方式串,"结算方式|金额||...."
    '出参:cur个人帐户-个人帐户
    '     cur医保基金-医保基金
    '返回:返回保险结算方式串,"结算方式|金额||...."
    '编制:刘兴洪
    '日期:2015-01-13 15:16:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    Dim curMoney As Currency, int类型 As Integer
    strTemp = ""
    cur个人帐户 = 0: cur医保基金 = 0
    With vsBlance
        For i = 1 To .Rows - 1
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int类型 = Val(.TextMatrix(i, .ColIndex("类型")))
            curMoney = Val(.TextMatrix(i, .ColIndex("结算金额")))
            
            If int类型 = 2 And .TextMatrix(i, .ColIndex("结算方式")) <> "" Then
                strTemp = strTemp & "||" & .TextMatrix(i, .ColIndex("结算方式")) & "|" & Format(curMoney, gstrDec)
                If Val(.TextMatrix(i, .ColIndex("结算性质"))) = 3 Then cur个人帐户 = cur个人帐户 + curMoney
                If Val(.TextMatrix(i, .ColIndex("结算性质"))) = 4 Then cur医保基金 = cur医保基金 + curMoney
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 3)
    GetMedicareStr = strTemp
End Function


Private Sub cmdYBBalance_Click()
    Dim cllPro As Collection
    Dim objFocus As Object
    Dim lng病人ID As Long
    
    '数据有效性检查
    If CheckInputConsValied(objFocus) = False Then
        If objFocus Is Nothing Then Exit Sub
        If objFocus.Enabled And objFocus.Visible Then objFocus.SetFocus
        If UCase(TypeName(objFocus)) = UCase("txtEdit") Then
            zlControl.TxtSelAll objFocus
        End If
        Exit Sub
    End If
    
    If mblnNotify = False Then
        If MsgBox("你确认要对该病人进行结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        mblnPrintInvoice = False
        If Not mobjBalanceCon.blnCurBalanceOwnerFee Then   '非自费费用时,要打印发票
            If Not (mYBInFor.intInsure <> 0 And MCPAR.医保接口打印票据) Then
                '保险病人根据使用类别来进行确认了
                Select Case mobjFactProperty.打印方式
                Case 0  '不打印
                Case 1
                    mblnPrintInvoice = True '自动打印
                Case 2  '提示打印
                    If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        Else
            If Not mty_ModulePara.blnNotPrintInvioce Then
                Select Case mobjFactProperty.打印方式
                Case 0  '不打印
                Case 1
                    mblnPrintInvoice = True '自动打印
                Case 2  '提示打印
                    If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        End If
        mblnNotify = True
    End If
    
    Call LockedScreen(True)
    If GetSaveBalanceSQL(cllPro) = False Then
        Call LockedScreen(False)      '解锁
        Exit Sub
    End If
    
    If ExcuteInsureSwapInteface(mBalanceInfor.lng结帐ID, cllPro) = False Then
        Call LockedScreen(False)      '解锁
        Exit Sub
    End If
    
    Call LockedScreen(False)      '解锁
    '加载结帐信息
    lng病人ID = Val(NVL(mrsInfo!病人ID))
    mblnInsure = True
    Call LoadBalancePayData(lng病人ID, mBalanceInfor.lng结帐ID)
    mblnInsure = False
'    Call RecalcDepositMoney(1)  '重新按缺省计算预交
    Call LoadIntendBalance
    mblnNotChange = True
    txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
    mblnNotChange = False
    '0-医保预算信息显示;1-显示费用信息
    Call ShowLedDisplayBank(1)
    
    Call LoadCurOwnerPayInfor(True) '加载当前支付信息
    'bytFun-0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
    Call SetOperationCtrl(2)
    If mBalanceInfor.dbl冲预交合计 <> 0 Then
        '光标定位到缴款处
        If txtBalance(Idx_冲预交).Enabled And txtBalance(Idx_冲预交).Visible Then txtBalance(Idx_冲预交).SetFocus
    Else
        '光标定位到缴款处
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
    End If
    If mBalanceInfor.dbl冲预交合计 = 0 And RoundEx(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl医保支付合计, 5) = 0 Then cmdOK_Click
End Sub

Private Sub LockedScreen(ByVal blnLocked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷屏，不允许在执行过程中点击相关控件
    '入参:blnLocked-true,锁屏,False-解锁
    '编制:刘兴洪
    '日期:2015-01-13 16:41:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUnLocked As Boolean
    
    Screen.MousePointer = IIf(blnLocked, 99, 0)
    
    mblnLockScreen = blnLocked
    blnUnLocked = Not blnLocked
    mblnInvalidLoad = True
    picPati.Enabled = blnUnLocked
    mblnInvalidLoad = False
    picBalanceInfor.Enabled = blnUnLocked
    cmdCancel.Enabled = blnUnLocked
    vsBlance.Enabled = blnUnLocked
    cmdOK.Enabled = blnUnLocked
    cmdYB.Enabled = blnUnLocked
    txtInvoice.Enabled = blnUnLocked
    picNO.Enabled = blnUnLocked
    picFeeList.Enabled = blnUnLocked
    picDeposit.Enabled = blnUnLocked
    
    
End Sub

Private Sub Form_Activate()
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If mblnUnload = True Then Unload Me: Exit Sub
    If Not mblnFirst Then Exit Sub
    
    
    mblnFirst = False
    Call Led_ClearDisplayPatient
    
    If mstrInNO <> "" And mEditType = g_Ed_单据查看 Then
        '作废时
        If txtPatient.Text = "" Then Unload Me: Exit Sub
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
    ElseIf mEditType = g_Ed_结帐作废 Then
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
'    Else
'        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    
    If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            '取消按钮
            If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus: Call cmdCancel_Click
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If cmdYBBalance.Enabled And cmdYBBalance.Visible Then cmdYBBalance.SetFocus: cmdYBBalance_Click: Exit Sub
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Call cmdOK_Click
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKind.Enabled Then
                    Dim intIndex As Integer
                    intIndex = IDKind.GetKindIndex("IC卡号")
                    If intIndex <= 0 Then Exit Sub
                    IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
                End If
                Exit Sub
            End If
            If Me.ActiveControl Is txtPatient Then
                If IDKind.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
                    Else
                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
                    End If
                End If
            End If
        Case vbKeyF6
            If cmdYB.Enabled And cmdYB.Visible Then cmdYB.SetFocus: Call cmdYB_Click
        Case vbKeyF8 '退号快捷
            chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
        Case vbKeyF9 '结帐设置
            If cmdMore.Enabled And cmdMore.Visible Then cmdMore.SetFocus: Call cmdMore_Click
        Case vbKeyF11 '定位到病人输入框
            If Not txtPatient.Locked And txtPatient.Enabled Then txtPatient.SetFocus
        Case vbKeyF12 '定位到单号框和强制报价
            If Shift = vbCtrlMask Then
                '强制性LED报价,(合计)
                mbln已报价 = False
                Call LedVoiceSpeak(True)
            Else
                If Not cboNO.Locked And cboNO.Enabled Then cboNO.SetFocus
            End If
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
    If mblnInvalidLoad = True Then
        mintSucces = 1
        mblnUnload = True: Exit Sub
    End If
    mlngModul = 1137
    mblnFirst = True: mblnUnload = False
    Call RestoreWinState(Me, App.ProductName)
    Call InitGrid_PayList
    Call zlInitModulePara
    If Init结算方式 = False Then Exit Sub
    '初始化界面
    Call InitFace
    '初始化菜单或工具栏
    Call zlDefCommandBars
    
    Call InitLed '初始化Led
    
    '81697:李南春,2015/6/8,评价器
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    If LoadBalanceBill = False Then mblnUnload = True: Exit Sub
End Sub
Private Sub SetDefaultPayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的支付方式
    '编制:刘兴洪
    '日期:2015-01-28 10:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim emEditType As gBalanceBill
    Dim strDefaultType As String, lng卡类别ID As Long
    Dim strBalance As String, strCash As String
    Dim dbl剩余金额 As Double, intKindIdx As Integer
    Dim i As Long, objCard As Card
    Dim blnFind As Boolean
    Dim dblMoney As Double
    On Error GoTo errHandle
    emEditType = mEditType
    If chkCancel.Value = 1 Then emEditType = g_Ed_结帐作废
    
    Select Case emEditType
    Case g_Ed_门诊结帐, g_Ed_住院结帐, g_Ed_重新结帐
        strBalance = mstr缺省结算方式
        With mBalanceInfor
            dbl剩余金额 = RoundEx(.dbl未付合计 - .dbl冲预交合计, 5)
        End With
        If dbl剩余金额 >= 0 Then GoTo GoLocal:
        If mPatiInfor.dbl未付累计 <> 0 Then Exit Sub
        '退款的缺省方式
        If mrsDeposit Is Nothing Then GoTo GoLocal:
        If mrsDeposit.State <> 1 Then GoTo GoLocal:
        If mrsDeposit.RecordCount = 0 Then GoTo GoLocal:
        If mty_ModulePara.bln结帐退款方式 Then
            mrsDeposit.Sort = "卡类别ID Desc,转帐及代扣,结算性质"
            With mrsDeposit
                .MoveFirst
                Do While Not .EOF
                    '1.三方卡时，只有代扣的才能缺省退款(主要是缴预交，可能存在多交次易，现简单处理)
                    If Val(NVL(!卡类别ID)) > 0 And NVL(NVL(!转帐及代扣, 0)) = 1 Then
                        '检查当前是否支持方结算卡
                        If Not GetLocalePayCard(Val(NVL(!卡类别ID)), False, intKindIdx) Is Nothing Then
                            IDKindPaymentsType.IDKind = intKindIdx
                            Exit Sub
                        End If
                    End If
                    '2.结算方式为XX卡的,则缺省为该方式
                    If Val(NVL(!结算性质)) = 2 And NVL(!结算方式) Like "*卡" Then
                        strBalance = NVL(!结算方式): GoTo GoLocal:
                    End If
                    If Val(NVL(!结算性质)) = 1 Then strCash = NVL(!结算方式)
                    If Val(NVL(!结算性质)) = 2 And NVL(!结算方式) Like "*支票" Then strBalance = NVL(!结算方式)
                    If strBalance = "" And Val(NVL(!结算性质)) = 2 Then
                        strBalance = NVL(!结算方式)
                    End If
                    .MoveNext
                Loop
                If strCash <> "" Then strBalance = strCash
                GoTo GoLocal:
            End With
        Else
            mrs结算方式.Filter = "缺省 = 1"
            If Not mrs结算方式.EOF Then
                strBalance = NVL(mrs结算方式!名称)
            End If
            mrs结算方式.Filter = 0
        End If
    Case g_Ed_结帐作废, g_Ed_重新作废
        With vsBlance
            
        End With
    Case Else
    End Select
GoLocal:
    '定位
    blnFind = False
    For i = 1 To IDKindPaymentsType.ListCount
        '缺省定位在现金上
        Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
        If strBalance = "" And objCard.结算性质 = 1 Then IDKindPaymentsType.IDKind = i: blnFind = True: Exit For
        If objCard.结算方式 = strBalance Then IDKindPaymentsType.IDKind = i: blnFind = True: Exit For
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mBalanceInfor.blnSaveBill And (mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_结帐作废) Then
        MsgBox "已经保存了结帐单据,不能退出!", vbInformation, gstrSysName
        Cancel = 1: Exit Sub
    End If
    If (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) And mstrInNO = "" And mYBInFor.strYBPati <> "" And Not mobjBalanceCon.blnCurBalanceOwnerFee Then
        
        If MsgBox("当前正在对医保病人结帐，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        '取消医保病人身份验证,返回假时不退出
            Cancel = 1: Exit Sub
        End If
    End If
    '清除入口参数
    mlngPatientID = 0: mblnViewCancel = False: mstrInNO = ""
    mblnNOMoved = False: mstrPrivs = ""
    mlng领用ID = 0: mbln门诊转住院 = False
    mstr主页Id = "": mstrPepositDate = ""
    mblnNotify = False
    
    Call ClearCustomType '清除自定义类型相关变量
 
    Call InitBalanceCondition
     
    Set mrsBalance = Nothing
    Set mrsFeeList = Nothing
    Set mrsDeposit = Nothing
    Set mobjPlugIn = Nothing
    Set mrsOldBalance = Nothing
    Set mrsInfo = New ADODB.Recordset
    mstrPatient = ""
    If mEditType <> g_Ed_单据查看 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    If mEditType <> g_Ed_单据查看 Then
        Call SaveRegInFor(g私有模块, Me.Name, "IDKIND", IDKind.IDKind)
    End If
    mblnBatchState = False
    Me.Visible = False
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, objPatiInfor.卡号)
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call FindPati(objCard, True, txtPatient.Text)
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub LoadDefaultMoney(Optional blnForceDefault As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缺省的缴款或退款金额
    '编制:刘兴洪
    '日期:2015-01-30 17:38:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lngCash As Long
    Dim i As Long, blnHave As Boolean
    On Error GoTo errHandle
        
    If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Then
        blnHave = False
        If vsBlance.Rows >= 2 Then
            For i = 2 To vsBlance.Rows
                If objCard.结算方式 = vsBlance.TextMatrix(i - 1, 0) Then
                    blnHave = True
                End If
            Next i
        End If
        If Not blnHave Then
            If objCard.结算性质 <> 1 Then
               txtReceive.Text = Format(Val(mBalanceInfor.dbl未付合计), "0.00")
            End If
       End If
       Exit Sub
    ElseIf mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐 Then
        With vsBlance
            For i = 1 To .Rows - 1
                If Val(.RowData(i)) = 999 Then
                    lngCash = i
                    Exit For
                End If
            Next i
        
            If lngCash <> 0 Then
                mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - Val(.TextMatrix(lngCash, .ColIndex("结算金额"))), 5)
                mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + Val(.TextMatrix(lngCash, .ColIndex("结算金额"))), 5)
                .TextMatrix(lngCash, .ColIndex("结算金额")) = Format(Val(mBalanceInfor.dbl未付合计), mstrDec)
                mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 + Val(.TextMatrix(lngCash, .ColIndex("结算金额"))), 5)
                mBalanceInfor.dbl未付合计 = 0
                txtOwe.Text = "0.00"
                Call SetCaculated
            End If
        End With
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetCaculated()
    Dim dblMoney As Double
    dblMoney = GetCashSum - Val(txtReceive.Text)
    If dblMoney < 0 Then
        lblCaculated.Caption = "找补"
        lblCaculated.ForeColor = vbRed
        txtCaculated.ForeColor = vbRed
    Else
        lblCaculated.Caption = "收款"
        lblCaculated.ForeColor = vbBlack
        txtCaculated.ForeColor = vbBlack
    End If
    txtCaculated.Text = Format(Abs(dblMoney), "0.00")
End Sub


Private Sub IDKindPaymentsType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub IDKindPaymentsType_KeyPress(KeyAscii As Integer)
    Call MoveIDKindItem(IDKindPaymentsType, KeyAscii)
End Sub


Private Sub opt出院_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt中途_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub picBalanceBack_Resize()
    Dim lngStep As Long
    Err = 0: On Error Resume Next
    lngStep = 100
    With picBalanceBack
        shpBalance.Left = .ScaleLeft
        shpBalance.Top = .ScaleTop
        shpBalance.Width = .ScaleWidth
        shpBalance.Height = .ScaleHeight
        
        cmdCancel.Top = picBalanceBack.ScaleHeight - cmdCancel.Height - 60
        
        cmdDelBalance.Left = cmdCancel.Left
        cmdDelBalance.Top = cmdCancel.Top
               
        cmdOK.Left = IIf(cmdCancel.Visible Or cmdDelBalance.Visible, cmdCancel.Left, .ScaleWidth) - cmdOK.Width - 60
        cmdOK.Top = cmdCancel.Top
        
        cmdYBBalance.Left = cmdOK.Left '- cmdYBBalance.Width - 60
        cmdYBBalance.Top = cmdCancel.Top
        
        cmdNext.Left = cmdOK.Left - cmdNext.Width - 60
        cmdNext.Top = cmdCancel.Top
        
        txtReceive.Top = cmdCancel.Top - txtReceive.Height - 90
        txtCaculated.Top = txtReceive.Top
        lblReceive.Top = txtReceive.Top + 60
        lblCaculated.Top = lblReceive.Top
        
        txtOwe.Top = txtReceive.Top - txtOwe.Height - 90
        lblOwe.Top = txtOwe.Top + 60
        
        lblBalance(3).Top = Frame3.Top + 200
        chkDeposit.Top = Frame3.Top + 200
        txtBalance(3).Top = chkDeposit.Top - 60
        lbl预交余额.Top = chkDeposit.Top
        
        vsBlance.Top = chkDeposit.Top + chkDeposit.Height + 120
        vsBlance.Height = txtOwe.Top - 60 - vsBlance.Top

    End With
    Call picBalanceInfor_Resize
End Sub


Private Sub picBalanceInfor_Resize()
    Err = 0: On Error Resume Next
    With picBalanceInfor
        txtBalance(Idx_结帐说明).Width = .ScaleWidth - txtBalance(Idx_结帐说明).Left - 100
        txtBalance(Idx_本次未结).Width = .ScaleWidth / 2 - txtBalance(Idx_本次未结).Left - 100
        lblBalance(Idx_本次结帐).Left = txtBalance(Idx_本次未结).Left + txtBalance(Idx_本次未结).Width + 100
        txtBalance(Idx_本次结帐).Left = lblBalance(Idx_本次结帐).Left + lblBalance(Idx_本次结帐).Width + 45
        txtBalance(Idx_本次结帐).Width = .ScaleWidth - txtBalance(Idx_本次结帐).Left - 100
    End With
End Sub

Private Sub picDetailContain_Resize()
    On Error Resume Next
    With vsDetailList
        .Top = 0
        .Left = 0
        .Height = picDetailContain.ScaleHeight
        .Width = picDetailContain.ScaleWidth
    End With
End Sub

Private Sub picFeeContain_Resize()
    On Error Resume Next
    With vsFeeList
        .Top = 0
        .Left = 0
        .Height = 3000
        .Width = picFeeContain.ScaleWidth
    End With
    With picDeposit
        .Top = vsFeeList.Top + vsFeeList.Height + 60
        .Left = 0
        .Width = picFeeContain.ScaleWidth
        .Height = picFeeContain.ScaleHeight - .Top - 30
    End With
End Sub

Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        picBalanceInfor.Left = 15
        picBalanceInfor.Top = .ScaleHeight - picBalanceInfor.Height
        picBalanceInfor.Width = .ScaleWidth - 30
        
        tabFeelist.Left = 15
        tabFeelist.Top = 15
        tabFeelist.Height = picBalanceInfor.Top - 30
        tabFeelist.Width = .ScaleWidth - 30
        
        picFeeContain.Left = 15
        picFeeContain.Top = 330
        picFeeContain.Width = .ScaleWidth - 30
        picFeeContain.Height = .ScaleHeight - 1400
        
        picDetailContain.Left = 15
        picDetailContain.Top = 330
        picDetailContain.Width = .ScaleWidth - 30
        picDetailContain.Height = .ScaleHeight - 1400
        
        If tabFeelist.Tab = 1 Then
            picDetailContain.Visible = True
            picFeeContain.Visible = False
        Else
            picDetailContain.Visible = False
            picFeeContain.Visible = True
        End If
        
        lnFeeSplit.X1 = .ScaleWidth - 15
        lnFeeSplit.X2 = .ScaleWidth - 15
        lnFeeSplit.Y1 = -30
        lnFeeSplit.Y2 = .ScaleHeight
    End With
End Sub

Private Sub tabFeelist_Click(PreviousTab As Integer)
    If tabFeelist.Tab = 1 Then
        picDetailContain.Visible = True
        picFeeContain.Visible = False
        If vsDetailList.Enabled And vsDetailList.Visible Then vsDetailList.SetFocus
    Else
        picDetailContain.Visible = False
        picFeeContain.Visible = True
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
    End If
End Sub

Private Sub picDeposit_Resize()
    Err = 0: On Error Resume Next
    With picDeposit
        vsDeposit.Left = 15
        vsDeposit.Top = lblDeposit.Top + lblDeposit.Height + 50
        vsDeposit.Height = .ScaleHeight - vsDeposit.Top - 30
        
        cmdDepositUp.Top = vsDeposit.Top + vsDeposit.Height / 4
        cmdDepositDown.Top = cmdDepositUp.Top + cmdDepositUp.Height + 250
        cmdDepositUp.Left = .ScaleWidth - cmdDepositUp.Width - 100
        cmdDepositDown.Left = cmdDepositUp.Left
        
        If cmdDepositUp.Visible Then
            vsDeposit.Width = cmdDepositUp.Left - vsDeposit.Left - 60
        Else
            vsDeposit.Width = .ScaleWidth - vsDeposit.Left - 100
        End If
        
        cmdTools.Left = .ScaleWidth - cmdTools.Width - 100
    End With
End Sub

Private Sub SetUpDown()
    With vsDeposit
        cmdDepositUp.Enabled = True
        cmdDepositDown.Enabled = True
        If .Row = 1 Then cmdDepositUp.Enabled = False
        If .Row = .Rows - 1 Then cmdDepositDown.Enabled = False
    End With
End Sub


Private Sub picNO_Resize()
    Err = 0: On Error Resume Next
    With picNO
        chkCancel.Left = .ScaleWidth - chkCancel.Width
        chkCancel.Top = .ScaleTop
        lblDelCaption.Left = .ScaleWidth - lblDelCaption.Width
        lblDelCaption.Top = .ScaleTop
        
        cboNO.Left = .ScaleLeft
        cboNO.Top = .ScaleTop
        If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) Then
            If mblnViewCancel Or mEditType = g_Ed_取消结帐 Or mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Then
                cboNO.Width = lblDelCaption.Left - cboNO.Left - 30
            Else
                cboNO.Width = .ScaleWidth
            End If
        Else
            cboNO.Width = chkCancel.Left - cboNO.Left - 30
        End If
        cboNO.Height = .ScaleHeight
    End With
End Sub
 
Private Sub AddPopu()
    Dim vRect As RECT
    vRect = zlControl.GetControlRect(cmdTools.hWnd)
    vRect.Left = vRect.Left + 10
    vRect.Top = vRect.Top + 50
    Call CreatePopuMenu
    If Not mobjCommandBar Is Nothing Then Call mobjCommandBar.ShowPopup(, vRect.Left, vRect.Top + cmdTools.Height)
End Sub

Private Sub CreatePopuMenu()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建临时菜单
    '编制:刘兴洪
    '日期:2012-11-21 09:49:35
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim objCustom As CommandBarControlCustom
   
    Set mobjCommandBar = cbsThis.Add("PopupPati", xtpBarPopup)
    With mobjCommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NotUseDeposit, "不使用预交款(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_UseAllDeposit, "使用所有预交款(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MoneyUseDeposit, "按结帐金额使用预交(&J)")
    End With
End Sub

Private Function InitGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-12-29 15:08:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsDeposit
        .Clear
        .Cols = 10: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "ID":  i = i + 1
        .TextMatrix(0, i) = "单据号": i = i + 1
        .TextMatrix(0, i) = "票据号": i = i + 1
        .TextMatrix(0, i) = "收款日期": i = i + 1
        .TextMatrix(0, i) = "结算方式": i = i + 1
        .TextMatrix(0, i) = "余额": i = i + 1
        .TextMatrix(0, i) = "冲预交": i = i + 1
        .TextMatrix(0, i) = "金额": i = i + 1
        .TextMatrix(0, i) = "预交ID": i = i + 1
        .TextMatrix(0, i) = "编辑状态": i = i + 1
        
        
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedCols = 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            ''ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "编辑状态" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*额" Or .ColKey(i) Like "*冲预交" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If

            Select Case .ColKey(i)
            Case "单据号"
                .ColData(i) = "1|0"
                .FixedAlignment(i) = flexAlignRightCenter
            Case "余额"
                 If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 _
                    Or mEditType = g_Ed_重新结帐 Then
                    .ColData(i) = "0|0"
                    .ColHidden(i) = False
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|1"
                 End If
            Case "冲预交"
                    .ColData(i) = "1|0"
                    .ColHidden(i) = False
            Case "金额"
                 If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐 Then
                     .ColHidden(i) = True: .ColData(i) = "0|1"
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|0"
                 End If
            Case Else
                If Not .ColKey(i) Like "*ID" Then
                    .ColData(i) = "0|0"
                End If
            End Select
        Next
        .ExtendLastCol = False
        .ColHidden(.ColIndex("票据号")) = True
        .ColWidth(.ColIndex("票据号")) = 1100
        .ColWidth(.ColIndex("收款日期")) = 1200
        .ColWidth(.ColIndex("单据号")) = 1100
        .ColWidth(.ColIndex("结算方式")) = 1400
        .ColWidth(.ColIndex("余额")) = 1100
        .ColWidth(.ColIndex("冲预交")) = 1100
        zl_vsGrid_Para_Restore mlngModul, vsDeposit, Me.Name, "预交列表"
        
        If mEditType = g_Ed_单据查看 Then
             .ColHidden(.ColIndex("余额")) = True: .ColData(.ColIndex("余额")) = "-1|1"
        End If
    End With
    With vsDetailList
        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .AllowBigSelection = False
        .HighLight = flexHighlightWithFocus
    End With
    Call InitTride_FeeList
    
    '结算信息
'    Call InitGrid_PayList
    InitGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitTride_FeeList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化费用列表
    '编制:刘兴洪
    '日期:2015-01-23 17:23:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsFeeList
        .Clear
        .Cols = 5: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "费目": i = i + 1
        .TextMatrix(0, i) = "应收金额": i = i + 1
        .TextMatrix(0, i) = "实收金额": i = i + 1
        .TextMatrix(0, i) = "未结金额": i = i + 1
        .TextMatrix(0, i) = "结帐金额": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            ''ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*额" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If
        Next
        
        .ColWidth(.ColIndex("费目")) = 2000
        .ColWidth(.ColIndex("应收金额")) = 1400
        .ColWidth(.ColIndex("实收金额")) = 1400
        .ColWidth(.ColIndex("未结金额")) = 1400
        .ColWidth(.ColIndex("结帐金额")) = 1400
    End With
    zl_vsGrid_Para_Restore mlngModul, vsFeeList, Me.Name, "费用列表"
    
    Call SetFeeListColumnShow
    With vsDetailList
        .Clear
        .Cols = 10: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "日期": i = i + 1
        .TextMatrix(0, i) = "单据": i = i + 1
        .TextMatrix(0, i) = "项目": i = i + 1
        .TextMatrix(0, i) = "未结金额": i = i + 1
        .TextMatrix(0, i) = "结帐金额": i = i + 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "记录性质": i = i + 1
        .TextMatrix(0, i) = "记录状态": i = i + 1
        .TextMatrix(0, i) = "执行状态": i = i + 1
        .TextMatrix(0, i) = "序号": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            ''ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*额" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "记录性质" Or .ColKey(i) = "记录状态" Or .ColKey(i) = "执行状态" Or .ColKey(i) = "序号" Or .ColKey(i) = "费目" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            End If
            If InStr(",日期,单据,", "," & .ColKey(i) & ",") > 0 Then .ColAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .ColWidth(.ColIndex("日期")) = 1400
        .ColWidth(.ColIndex("单据")) = 1100
        .ColWidth(.ColIndex("项目")) = 2800
        .ColWidth(.ColIndex("未结金额")) = 1400
        .ColWidth(.ColIndex("结帐金额")) = 1400
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDetailList, Me.Name, "明细列表"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetFeeListColumnShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置费用表显示信息
    '编制:刘兴洪
    '日期:2015-01-23 17:29:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsFeeList
        If (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) And chkCancel.Value = 0 Then
            .ColHidden(.ColIndex("结帐金额")) = True: .ColWidth(.ColIndex("结帐金额")) = 0
        Else
            .ColHidden(.ColIndex("未结金额")) = True: .ColWidth(.ColIndex("未结金额")) = 0
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitGrid_PayList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化支付列表
    '编制:刘兴洪
    '日期:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
        .Clear: .Rows = 2: i = 0: .Cols = 20
        .TextMatrix(0, i) = "卡类别ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "消费卡ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算性质": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "编辑状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "类型": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否全退": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "校对标志": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否密文": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "结算方式": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "结算金额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "结算号码": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "备注": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "卡号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易流水号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易说明": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "卡类别名称": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "组合信息": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否转账": .ColWidth(i) = 0: i = i + 1
        
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If InStr(",结算方式,结算金额,结算号码,备注,", "," & .ColKey(i) & ",") > 0 Then
                .ColData(i) = "-1||0"
            Else
                .ColData(i) = "-1||1"
            End If
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            Select Case .ColKey(i)
            Case "是否转账", "组合信息", "结算性质", "类型", "是否保存", "是否密文", "校对标志", "编辑状态", "是否退现", "是否全退", "卡类别名称", "结算状态", "是否验证"
                .ColHidden(i) = True
            Case "结算金额"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
        zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "结算列表"
        If Not mEditType = g_Ed_单据查看 Then
            .Editable = flexEDKbdMouse
        End If
        .Row = 1: .Col = .ColIndex("结算方式")
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
 
Private Sub picPati_Resize()
    Err = 0: On Error Resume Next
    With picPati
        lnPatiSplit.Y1 = .ScaleHeight - 10
        lnPatiSplit.Y2 = .ScaleHeight - 10
        txtSex.Width = 600 * (.ScaleWidth / 14000)
        lblOld.Left = txtSex.Left + txtSex.Width + 100
        txtOld.Left = lblOld.Left + lblOld.Width + 30
        txtOld.Width = 1000 * (.ScaleWidth / 14000)
        lbl费别.Left = txtOld.Left + txtOld.Width + 100
        txt费别.Left = lbl费别.Left + lbl费别.Width + 30
        txt费别.Width = 1560 * (.ScaleWidth / 14000)
        lbl标识号.Left = txt费别.Left + txt费别.Width + 100
        txt标识号.Left = lbl标识号.Left + lbl标识号.Width + 30
        txt标识号.Width = 1500 * (.ScaleWidth / 14000)
        lblBed.Left = txt标识号.Left + txt标识号.Width + 100
        txtBed.Left = lblBed.Left + lblBed.Width + 30
        txtBed.Width = 780 * (.ScaleWidth / 14000)
        lbl科室.Left = txtBed.Left + txtBed.Width + 100
        txt科室.Left = lbl科室.Left + lbl科室.Width + 30
        txt科室.Width = 1440 * (.ScaleWidth / 14000)
    End With
End Sub

Private Sub txtBalance_Change(Index As Integer)
    If mblnNotChange Then Exit Sub
    If mEditType = g_Ed_单据查看 Then Exit Sub
    
    Select Case Index
    Case Idx_冲预交
        If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_取消结帐 Or chkCancel.Value = 1 Or mblnManualEdit Then Exit Sub
        
        mBalanceInfor.bln预交刷卡 = False
        
        If mBalanceInfor.dbl冲预交合计 <> 0 Then
            mBalanceInfor.dbl冲预交合计 = 0
            If mEditType <> g_Ed_重新作废 Then Call RecalcDepositMoney(0)
            Call LoadCurOwnerPayInfor(True)
        End If
        
        txtBalance(Idx_冲预交).BackColor = IIf(txtBalance(Idx_冲预交).Enabled, &H80000005, &H8000000F)
        mbln已报价 = False
    Case Idx_本次结帐
        mbln已报价 = False
    Case Idx_结帐说明
    Case Else
    End Select
End Sub


Private Sub txtBalance_GotFocus(Index As Integer)
    Select Case Index
    Case Idx_冲预交
    Case Idx_结帐说明
        zlCommFun.OpenIme True
    End Select
    zlControl.TxtSelAll txtBalance(Index)
End Sub
Private Sub LedVoiceSpeak(ByVal blnGotFocus As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:语音报价
    '入参: blnGotFocus-是否进入缴款控件,True是进入时,False-离开时
    '编制:刘兴洪
    '日期:2015-01-28 14:10:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curTotal As Currency, dbl剩余 As Double
    Dim intSign As Integer
    Dim blnSign As Boolean
    Dim intMark As Integer
    Dim dbl找补 As Double
    If Not gblnLED Then Exit Sub
    '#21 1234.56   --请您付款一千二百三十四点五六元  J
    '#22 1234.56   --预收一千二百三十四点五六元 Y
    '#23 1234.56   --找零一千二百三十四点五六元 Z
    intSign = IIf(mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废, -1, 1)
    
    curTotal = Get应缴
    With mBalanceInfor
        dbl剩余 = RoundEx(intSign * (.dbl当前结帐 + mPatiInfor.dbl未付累计 - .dbl冲预交合计 - .dbl医保支付合计), 5)
    End With
    
    If blnGotFocus Then
        If mbln已报价 Then Exit Sub
        zl9LedVoice.DisplayBank (" ")
        If curTotal >= 0 Then
            zl9LedVoice.Speak "#21 " & curTotal
        Else
            zl9LedVoice.Speak "#23 " & Abs(curTotal)
        End If
        mbln已报价 = True
        Exit Sub
    End If
    curTotal = Abs(curTotal)
    intMark = IIf(dbl剩余 >= 0, 1, -1)
    '问题号:112948,焦博,2018/08/16,提取病人信息结账保存时报错
    dbl找补 = Val(txtCaculated.Text) 'Val(IIf(lblCaculated.Caption = "找补", txtCaculated.Text, 0))
    If intMark = 1 Then
        dbl找补 = Val(IIf(lblCaculated.Caption = "找补", txtCaculated.Text, 0))
        zl9LedVoice.DispCharge Format(curTotal, "0.00"), Val(txtReceive.Text), dbl找补
        zl9LedVoice.Speak "#22 " & Val(txtReceive.Text)
        zl9LedVoice.Speak "#23 " & dbl找补
        zl9LedVoice.Speak "#3"   '#3  --请当面点清, 谢谢!
    Else    '补119009
    
    
        If Val(txtReceive.Text) > 0 Then
            zl9LedVoice.DispCharge Format(intMark * curTotal, "0.00"), Abs(Val(txtReceive.Text)), dbl找补
                zl9LedVoice.Speak "#22 " & Abs(Val(txtReceive.Text))
                zl9LedVoice.Speak "#23 " & dbl找补
                zl9LedVoice.Speak "#3"   '#3  --请当面点清, 谢谢!
        ElseIf Abs(Val(txtReceive.Text)) > Val(curTotal) Then
                zl9LedVoice.DispCharge Format(intMark * curTotal, "0.00"), dbl找补, Abs(Val(txtReceive.Text))
                zl9LedVoice.Speak "#22 " & dbl找补
                zl9LedVoice.Speak "#23 " & Abs(Val(txtReceive.Text))
                zl9LedVoice.Speak "#3"   '#3  --请当面点清, 谢谢!
        Else
            zl9LedVoice.DispCharge Format(intMark * curTotal, "0.00"), 0, dbl找补 + Abs(Val(txtReceive.Text))
            zl9LedVoice.Speak "#22 " & 0
            zl9LedVoice.Speak "#23 " & dbl找补 + Abs(Val(txtReceive.Text))
            zl9LedVoice.Speak "#3"   '#3  --请当面点清, 谢谢!
        End If
    End If
End Sub
Private Sub MoveIDKindItem(ByVal objKind As IDKindNew, ByVal KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移动IDKind项目
    '入参:objKind-移动的IDKind对象
    '     Keyascii-键值
    '编制:刘兴洪
    '日期:2015-01-29 15:22:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If objKind Is Nothing Then Exit Sub
    If Not (KeyAscii = Asc("+") Or KeyAscii = Asc("-")) Then Exit Sub
    If objKind.ListCount = 1 Then Exit Sub
    
    If KeyAscii = Asc("+") Then
        '下移一项
        If objKind.IDKind + 1 > objKind.ListCount Then
            objKind.IDKind = 1
        Else
            objKind.IDKind = objKind.IDKind + 1
        End If
        Exit Sub
    End If
    If KeyAscii = Asc("-") Then '上移一项
        If objKind.IDKind - 1 <= 0 Then
            objKind.IDKind = objKind.ListCount
        Else
            objKind.IDKind = objKind.IDKind - 1
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtBalance_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim dblMoney As Double, blnChargeEnd As Boolean
    Dim objCard As Card, objKind As IDKindNew
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnNoRecal As Boolean
    
    If KeyAscii <> 13 Then
        If mPatiInfor.dbl未付累计 <> 0 Then Exit Sub
        If Index = Idx_本次结帐 Then
            If mYBInFor.intInsure <> 0 Then
                KeyAscii = 0
            End If
        End If
        Exit Sub
    End If
    
    KeyAscii = 0
    Select Case Index
    Case Idx_冲预交
        If chkDeposit.Visible Then Exit Sub
        If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 _
        Or chkCancel.Value = 1 And chkCancel.Visible Then Exit Sub
        dblMoney = RoundEx(Val(txtBalance(Index).Text), 6)
        If DepositMonyVerfy(False) = False Then Exit Sub
        If dblMoney = 0 Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
        Call SaveDeposit(True, blnNoRecal)
    Case Idx_结帐说明
        Call SkipSetFocus(1)
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtBalance_LostFocus(Index As Integer)
    Select Case Index
    Case Idx_冲预交
    Case Idx_结帐说明
        zlCommFun.OpenIme False
    End Select
End Sub



Private Sub txtBalance_Validate(Index As Integer, Cancel As Boolean)
    Dim dblMoney As Double, dbl找补 As Double
    Dim intSign As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnNoRecal As Boolean
    
    On Error GoTo errH
    
    Select Case Index
    Case Idx_冲预交
         If DepositMonyVerfy = False Then Cancel = True: Exit Sub
        
    Case Idx_本次结帐
        If chkCancel.Value = 1 Then Exit Sub
        If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) Then Exit Sub
        
        If RoundEx(Val(txtBalance(Idx_本次结帐).Text), 6) = 0 Then
            txtBalance(Idx_本次结帐).Text = Format(mBalanceInfor.dbl本次未结, gstrDec)
        Else
            txtBalance(Idx_本次结帐).Text = Format(Val(txtBalance(Idx_本次结帐).Text), gstrDec)
        End If
        
        If RoundEx(Val(txtBalance(Idx_本次结帐).Text), 6) > RoundEx(Val(txtBalance(Idx_本次未结).Text), 6) Then
            MsgBox "当前结帐金额大于了本次结帐的总额,不允许结帐!", vbInformation + vbOKOnly, gstrSysName
            zlControl.TxtSelAll txtBalance(Index)
            Cancel = True: Exit Sub
        End If
 
        
        If mblnNotClick Then Exit Sub
        mblnNotClick = True
        
        Call RelocateMoney
        mBalanceInfor.dbl当前结帐 = Val(txtBalance(Idx_本次结帐).Text)
        mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl已付合计, 5)
        
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor(True)
        mblnNotClick = False
    Case Else
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RelocateMoney()
    '分配金额
    Dim dblMoney As Double, i As Long
    Dim blnAll As Boolean
    dblMoney = Val(txtBalance(Idx_本次结帐).Text)
    blnAll = Val(txtBalance(Idx_本次结帐).Text) = Val(txtBalance(Idx_本次未结).Text)
    
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据")) <> "" Then
                If dblMoney >= Val(.Cell(flexcpData, i, .ColIndex("未结金额"))) And dblMoney <> 0 Or blnAll Then
                    .Cell(flexcpData, i, .ColIndex("结帐金额")) = Val(.Cell(flexcpData, i, .ColIndex("未结金额")))
                    dblMoney = dblMoney - Val(.Cell(flexcpData, i, .ColIndex("结帐金额")))
                Else
                    If dblMoney = 0 Then
                        .Cell(flexcpData, i, .ColIndex("结帐金额")) = ""
                    Else
                        .Cell(flexcpData, i, .ColIndex("结帐金额")) = dblMoney
                    End If
                    dblMoney = 0
                End If
                .TextMatrix(i, .ColIndex("结帐金额")) = Format(Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))), gstrDec)
            End If
        Next i
    End With
End Sub

 

Private Sub txtBegin_GotFocus()
    zlControl.TxtSelAll txtBegin
End Sub

Private Sub txtBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEnd_GotFocus()
    zlControl.TxtSelAll txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtPatiBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub txtPatiEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
 

Private Sub txtReceive_Change()
    Call SetCaculated
    SetNextBalanceCmdVisible
End Sub

Private Function GetCashSum() As Double
    Dim i As Long
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.RowData(i)) = 999 Then
                GetCashSum = RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))) + mPatiInfor.dbl未付累计, 5)
                Exit Function
            End If
        Next i
    End With
End Function

Private Sub txtReceive_GotFocus()
    If txtReceive.Locked Then Exit Sub
    Call LedVoiceSpeak(True)
    zlControl.TxtSelAll txtReceive
End Sub

Private Sub txtReceive_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub txtReceive_Validate(Cancel As Boolean)
    txtReceive.Text = Format(txtReceive.Text, "0.00")
End Sub

Private Sub vsBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblCurMoney As Double, dbl个人帐户 As Double, dbl医保基金 As Double
    Dim objCard As Card
    Dim i As Long
    With vsBlance
        Select Case Col
        Case .ColIndex("结算方式")
            '结算方式处理
            For i = 1 To .Rows - 1
                If .TextMatrix(Row, .ColIndex("结算方式")) = .TextMatrix(i, .ColIndex("结算方式")) And Row <> i And .TextMatrix(Row, .ColIndex("结算方式")) <> "" Then
                    MsgBox "结算方式<" & .TextMatrix(Row, .ColIndex("结算方式")) & ">已经被选择,不能重复添加！", vbInformation, gstrSysName
                    .TextMatrix(Row, .ColIndex("结算方式")) = ""
                    Exit Sub
                End If
            Next i
            Set objCard = GetCard(.TextMatrix(Row, .ColIndex("结算方式")))
            If objCard Is Nothing Then .TextMatrix(Row, .ColIndex("结算方式")) = "": mbln已报价 = False: Exit Sub
            
            '0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等
            Select Case objCard.结算性质
            Case 1  '现金
                .TextMatrix(Row, .ColIndex("类型")) = 0
                .TextMatrix(Row, .ColIndex("编辑状态")) = 1
            Case 2
                .TextMatrix(Row, .ColIndex("类型")) = 0
                .TextMatrix(Row, .ColIndex("编辑状态")) = 3
            Case 7, 8
                .TextMatrix(Row, .ColIndex("类型")) = IIf(objCard.消费卡, 5, 3)
                .TextMatrix(Row, .ColIndex("卡类别ID")) = objCard.接口序号
                .TextMatrix(Row, .ColIndex("编辑状态")) = 1
            End Select
            
            .TextMatrix(Row, .ColIndex("结算金额")) = "0.00"
            .TextMatrix(Row, .ColIndex("结算性质")) = objCard.结算性质
            
            .Rows = .Rows + 1
            Exit Sub
            
        Case .ColIndex("结算金额")
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            
            If InStr(",3,5,", "," & Val(.TextMatrix(Row, .ColIndex("类型"))) & ",") > 0 And Val(.TextMatrix(Row, .ColIndex("结算状态"))) = 0 Then
                Set objCard = GetCard(.TextMatrix(Row, .ColIndex("结算方式")))
                If objCard Is Nothing Then .TextMatrix(Row, .ColIndex("结算金额")) = "0.00": Exit Sub
                
                If chkCancel.Value = 1 Then
                    If ExecuteBalaceCancel(objCard) = False Then
                        Call DeletePayInfor(Row, True)
                        Exit Sub
                    Else
                        .TextMatrix(Row, .ColIndex("编辑状态")) = 0
                        .TextMatrix(Row, .ColIndex("结算状态")) = 1
                    End If
                Else
                    If Val(.TextMatrix(Row, .ColIndex("结算金额"))) <> 0 Then
                        Call LoadCurOwnerPayInfor(False)
                        
                        If SaveBalanceData(objCard, Row) = False Then
                            mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - Val(.TextMatrix(Row, .ColIndex("结算金额"))), 6)
                            mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + Val(.TextMatrix(Row, .ColIndex("结算金额"))), 6)
                            .TextMatrix(Row, .ColIndex("结算金额")) = "0.00"
                        Else
                            If mblnChargeEnd Then
                                mblnChargeEnd = False
                            ElseIf Not objCard.消费卡 Then
                                
                                If Row > .Rows - 1 Then Row = .Rows - 1: .Row = .Rows - 1
                                .TextMatrix(Row, .ColIndex("编辑状态")) = 0
                                .TextMatrix(Row, .ColIndex("结算状态")) = 1
                            End If
                        End If
                    End If
                End If
            End If
            If Row > .Rows - 1 Then Exit Sub
            Call LoadCurOwnerPayInfor(Val(.RowData(Row)) <> 999)
            mbln已报价 = False
            
        Case Else
        End Select
    End With
End Sub
Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "结算列表"
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Then Exit Sub
    If OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBlance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "结算列表"
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str结算方式 As String, int编辑状态 As Integer
    
    If mEditType = g_Ed_单据查看 Then Cancel = True: Exit Sub
    
    If mblnBatchState Then Cancel = True: Exit Sub
    
    
    With vsBlance
        If Val(.TextMatrix(Row, .ColIndex("结算状态"))) = 1 Then '已经结算的，不允许编辑
            Cancel = True: Exit Sub
        End If
        
        .ComboList = ""
        str结算方式 = .TextMatrix(Row, .ColIndex("结算方式"))
        Select Case Col
        Case .ColIndex("结算方式")
            If cmdYBBalance.Visible And cmdYBBalance.Enabled Then Cancel = True: Exit Sub
   
            If .RowData(Row) = "999" Then Cancel = True: Exit Sub   '缺省结算方式
            
            If str结算方式 = "" Then
                .ColComboList(.ColIndex("结算方式")) = mstrPayMode
                Exit Sub
            End If
            
            int编辑状态 = Val(.TextMatrix(Row, .ColIndex("编辑状态")))
            If InStr("12", .ColIndex("结算性质")) > 0 Then int编辑状态 = 2
            If int编辑状态 <> 2 And Val(.TextMatrix(Row, .ColIndex("结算金额"))) = 0 Then int编辑状态 = 2
            If int编辑状态 = 4 Then Cancel = True: Exit Sub
            '编辑状态: '0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等
            If int编辑状态 = 2 Then
                .ColComboList(.ColIndex("结算方式")) = ""
                .ComboList = "..."
                .CellButtonPicture = imgDel
            End If
            Exit Sub
        Case .ColIndex("结算金额")
            If Val(.TextMatrix(Row, .ColIndex("类型"))) = 9 Then Cancel = True: Exit Sub
            '0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等
            int编辑状态 = Val(.TextMatrix(Row, .ColIndex("编辑状态")))
            If int编辑状态 = 4 Then Cancel = True: Exit Sub

            If InStr("12", .TextMatrix(Row, .ColIndex("结算性质"))) > 0 And int编辑状态 <> 1 Then int编辑状态 = 1
            If int编辑状态 = 2 Then Cancel = True: Exit Sub
        Case .ColIndex("结算号码")
            int编辑状态 = Val(.TextMatrix(Row, .ColIndex("编辑状态")))
            If int编辑状态 = 4 Then Cancel = True: Exit Sub
            If cmdYBBalance.Visible And cmdYBBalance.Enabled Then Cancel = True: Exit Sub
            If Val(.TextMatrix(Row, .ColIndex("结算性质"))) = 2 Then Exit Sub
            Cancel = True
        Case Else
            int编辑状态 = Val(.TextMatrix(Row, .ColIndex("编辑状态")))
            If int编辑状态 = 4 Then Cancel = True: Exit Sub
            If cmdYBBalance.Visible And cmdYBBalance.Enabled Then Cancel = True: Exit Sub
        End Select
    End With
    
    
 
End Sub



Private Sub DeletePayInfor(ByVal lngDelRow As Long, Optional ByVal blnForceDel As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除支付信息
    '编制:刘兴洪
    '日期:2015-01-28 15:18:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngRow As Long
    Dim str操作员姓名 As String, strDBUser As String
    Dim strPrivs As String, i As Long
    Dim lng卡类别ID As Long, str卡号 As String, str交易说明 As String, str交易流水号 As String
    Dim dblCheckMoney As Double, strBalanceIDs As String
    Dim strArray() As String
    Dim intEdit As Integer
    
    
    On Error GoTo errHandle
    With vsBlance
        If lngDelRow > .Rows - 1 Or lngDelRow < 1 Then Exit Sub
        If Val(.TextMatrix(lngDelRow, .ColIndex("类型"))) = 3 And Val(.TextMatrix(lngDelRow, .ColIndex("结算金额"))) <> 0 Then
            
            lng卡类别ID = Val(.TextMatrix(lngDelRow, .ColIndex("卡类别ID")))
            str卡号 = .Cell(flexcpData, lngDelRow, .ColIndex("卡号"))
            str交易说明 = .TextMatrix(lngDelRow, .ColIndex("交易说明"))
            str交易流水号 = .TextMatrix(lngDelRow, .ColIndex("交易流水号"))
            dblCheckMoney = -1 * Val(.TextMatrix(lngDelRow, .ColIndex("结算金额")))
            
            If .TextMatrix(lngDelRow, .ColIndex("组合信息")) = "" Then
                If mBalanceInfor.lng结帐ID <> 0 Then
                    strBalanceIDs = "2|" & mBalanceInfor.lng结帐ID
                End If
            Else
                If Val(.Cell(flexcpData, lngDelRow, .ColIndex("组合信息"))) = 1 Then
                    strBalanceIDs = "1|" & .TextMatrix(lngDelRow, .ColIndex("组合信息"))
                Else
                    strArray = Split(.TextMatrix(lngDelRow, .ColIndex("组合信息")), "|")
                    For i = 0 To UBound(strArray)
                        strBalanceIDs = strBalanceIDs & "," & Split(strArray(i), ",")(4)
                    Next i
                    If strBalanceIDs <> "" Then
                        strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
                    End If
                End If
            End If
            If zlCallReturnCashCheckInterface(Me, mlngModul, lng卡类别ID, str卡号, strBalanceIDs, dblCheckMoney, str交易流水号, str交易说明) = False Then Exit Sub
 
        End If
        
    
        dblMoney = RoundEx(Val(.TextMatrix(lngDelRow, .ColIndex("结算金额"))), 5)
        If Val(.TextMatrix(lngDelRow, .ColIndex("是否退现"))) = 0 And Val(.TextMatrix(lngDelRow, .ColIndex("类型"))) = 3 And blnForceDel = False And dblMoney <> 0 Then
            '卡不支持退现的情况
            If InStr(";" & mstrPrivsCard & ";", ";三方退款强制退现;") = 0 Then
                If mstrForceNote = "" Then
                    '已经验证过的，不再验证
                    str操作员姓名 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
                    If str操作员姓名 = "" Then
                        MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    mstrForceNote = str操作员姓名 & "强制退现:" & .TextMatrix(lngDelRow, .ColIndex("卡类别名称")) & Format(Abs(dblMoney), gstrDec) & "元" & ";"
                Else
                    mstrForceNote = mstrForceNote & .TextMatrix(lngDelRow, .ColIndex("卡类别名称")) & Format(Abs(dblMoney), gstrDec) & "元" & ";"
                End If
            Else
                If MsgBox(.TextMatrix(lngDelRow, .ColIndex("结算方式")) & "不支持退现,是否强制将其退现？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
                mstrForceNote = mstrForceNote & IIf(mstrForceNote = "", UserInfo.姓名 & "强制退现:", ";") & .TextMatrix(lngDelRow, .ColIndex("卡类别名称")) & Format(Abs(dblMoney), gstrDec) & "元"
            End If
        End If

        
        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
        intEdit = Val(.TextMatrix(lngDelRow, .ColIndex("编辑状态")))
        
        If intEdit <> 2 And dblMoney = 0 Then intEdit = 2
        If InStr(1, "23", CStr(intEdit)) = 0 And blnForceDel = False Then Exit Sub
        
        lngRow = lngDelRow
        If Val(.TextMatrix(lngRow, .ColIndex("类型"))) <> 9 And dblMoney <> 0 Then
            mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + dblMoney, 6)
            mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - dblMoney, 6)
            Call LoadCurOwnerPayInfor
        End If
        
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            vsBlance.RemoveItem lngDelRow
        End If
        
        If lngRow <= 1 Then
            lngRow = 1
        ElseIf lngRow >= .Rows - 1 Then
            lngRow = .Rows - 1
        Else
            lngRow = lngDelRow + 1
        End If
        If lngRow > .Rows - 1 Or lngRow <= 1 Then lngRow = 1
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then .ShowCell .Row, .Col
        Call LoadCurOwnerPayInfor
    End With
    mbln已报价 = False
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If mEditType = g_Ed_单据查看 Then Exit Sub
    
    Call DeletePayInfor(Row)
    Call LoadDefaultMoney
    
End Sub

Private Sub vsBlance_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Select Case Col
    Case vsBlance.ColIndex("结算方式")
         
    Case Else
    End Select
    
End Sub

Private Sub vsBlance_DblClick()
    If mEditType = g_Ed_单据查看 Then Exit Sub
    With vsBlance
        
        If .Col <> .ColIndex("结算金额") Then Exit Sub
        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
        If Val(.TextMatrix(.Row, .ColIndex("编辑状态"))) <> 1 Then Exit Sub
        .EditCell
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub
 

Private Sub vsBlance_EnterCell()
    If mEditType = g_Ed_单据查看 Then Exit Sub
    With vsBlance
        Select Case .Col
        Case .ColIndex("结算方式")
        Case Else
        End Select
        If .Row < 0 Then Exit Sub
        Select Case Val(.TextMatrix(.Row, .ColIndex("结算性质")))
        Case 2
            .ColData(.ColIndex("结算号码")) = "0||0"
            .ColData(.ColIndex("备注")) = "0||0"
        Case Else
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            .ColData(.ColIndex("结算号码")) = "0||2"
            .ColData(.ColIndex("备注")) = "0||2"
        End Select
    End With
End Sub

Private Sub vsBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    Dim intEdit As Byte
    If mEditType = g_Ed_单据查看 Then Exit Sub
    With vsBlance
        If KeyCode <> vbKeyReturn And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                vsBlance_CellButtonClick .Row, .Col
            Else
                Select Case .Col
                Case .ColIndex("结算方式")
                    .ColComboList(.Col) = ""
                Case Else
                End Select
            End If
        End If
        '删除
        If KeyCode = vbKeyDelete Then
            '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
            intEdit = Val(.TextMatrix(.Row, .ColIndex("编辑状态")))
            If ((intEdit = 2 Or intEdit = 3) Or Val(.TextMatrix(.Row, .ColIndex("结算金额"))) = 0) And Val(.RowData(.Row)) <> 999 Then
                Call DeletePayInfor(.Row)
                Call LoadDefaultMoney
                 
            End If
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBlance
        Select Case .Col
        Case .ColIndex("结算方式")
            If Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case .ColIndex("结算金额")
            If (Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("结算金额"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case Else
            If (Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("结算金额"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("结算方式"), , IIf(mEditType = g_Ed_单据查看 Or mEditType = g_Ed_结帐作废, False, True), lngRow)
    End With
    
End Sub
 


Private Sub vsBlance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    If mEditType = g_Ed_单据查看 Or mEditType = g_Ed_结帐作废 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("结算方式")
        Case .ColIndex("结算金额")
           
        Case Else
            
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("结算方式"), -1, True, lngRow)
    End With
    'If lngRow >= 0 Then AfterAddRow  lngRow
    
End Sub

Private Sub vsBlance_KeyPress(KeyAscii As Integer)

    If mEditType = g_Ed_单据查看 Then Exit Sub
    If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    
'    With vsBlance
'        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
'        If Val(.TextMatrix(.Row, .ColIndex("编辑状态"))) <> 1 Then KeyAscii = 0: Exit Sub
'        If .Col <> .ColIndex("结算金额") Then KeyAscii = 0: Exit Sub
'    End With
'    Call VsFlxGridCheckKeyPress(vsBlance, vsBlance.Row, vsBlance.Col, KeyAscii, m金额式)
End Sub

Private Sub vsBlance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If mEditType = g_Ed_单据查看 Then Exit Sub
    If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) Then Exit Sub
    
    With vsBlance
        Select Case .Col
        Case .ColIndex("结算金额")
            Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m负金额式)
        Case .ColIndex("结算号码"), .ColIndex("备注")
            Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m文本式)
            Exit Sub
        Case Else
            KeyAscii = 0: Exit Sub
        End Select
    End With
End Sub

Private Function GetCard(str结算方式 As String) As Card
    Dim i As Long
    For i = 1 To mobjPayCards.Count
        If str结算方式 = mobjPayCards.Item(i).结算方式 Or str结算方式 = mobjPayCards.Item(i).名称 Or str结算方式 = CStr(mobjPayCards.Item(i).接口序号) Then
            Set GetCard = mobjPayCards.Item(i)
            Exit Function
        End If
    Next i
End Function

Private Sub vsBlance_LeaveCell()
    If mEditType = g_Ed_单据查看 Then Exit Sub
    OS.OpenIme False
End Sub

Private Sub vsBlance_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mEditType = g_Ed_单据查看 Then Exit Sub
    '设置单元格的编辑长度
    With vsBlance
       Select Case .Col
           Case .ColIndex("结算方式")
               .EditMaxLength = 50
           Case .ColIndex("结算金额")
               .EditMaxLength = 16
           Case .ColIndex("结算号码")
               .EditMaxLength = 30
           Case .ColIndex("备注")
               .EditMaxLength = 50
           Case Else
               .EditMaxLength = 100
       End Select
    End With
End Sub

Private Sub vsBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim objCard As Card, dbl原始金额 As Double
    Dim i As Long, str结算方式 As String
    Dim dblMoney As Double, blnYB As Boolean
    Dim strInput As String
    
    With vsBlance
        If Row <= 0 Then Exit Sub
        
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        
        Select Case Col
        Case .ColIndex("结算方式")
            If strInput = "" Then Exit Sub
            For i = 1 To .Rows - 1
                If strInput = .TextMatrix(i, .ColIndex("结算方式")) And Row <> i Then
              
                    MsgBox "结算方式<" & strInput & ">已经被选择,不能重复添加！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            Next
        Case .ColIndex("结算金额")
            If Not IsNumeric(strInput) And strInput <> "" Then
                MsgBox "输入的金额必须为数字！", vbInformation, gstrSysName
                .EditCell: .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
            If zlDblIsValid(strInput, 10, False, False, 0, .ColKey(Col)) = False Then
                Cancel = True: Exit Sub
            End If
            str结算方式 = Trim(.TextMatrix(.Row, .ColIndex("结算方式")))
            If str结算方式 = "" Then Exit Sub
            '结算金额不允许超过返回的原始金额(个人帐户允许透支时再判断)
            dbl原始金额 = Val(.Cell(flexcpData, .Row, .ColIndex("结算金额")))
            
            Select Case Val(.TextMatrix(.Row, .ColIndex("结算性质")))
            Case 3 '个人帐户
                If Val(strInput) > dbl原始金额 And Val(strInput) <> 0 And dbl原始金额 <> 0 Then
                    MsgBox "输入的""" & str结算方式 & """结算金额不能超过 " & Format(dbl原始金额, "0.00") & " ！", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                 '不允许超过允许透支金额
                If mYBInFor.cur个帐余额 + mYBInFor.cur个帐透支 - Val(strInput) < 0 Then
                    MsgBox "帐户余额:" & Format(mYBInFor.cur个帐余额, "0.00") & _
                        IIf(mYBInFor.cur个帐透支 = 0, "", "(" & "允许透支:" & Format(mYBInFor.cur个帐透支, "0.00") & ")") & _
                        "不足要结算的金额。", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                blnYB = True
            Case 4 '医保基金
                If Val(strInput) > dbl原始金额 And Val(strInput) <> 0 And dbl原始金额 <> 0 Then
                    MsgBox "输入的""" & str结算方式 & """结算金额不能超过 " & Format(dbl原始金额, "0.00") & " ！", vbInformation, gstrSysName
                    .EditCell
                    .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                blnYB = True
            End Select
            '重新计算医保结算金额
            Call ReCalcYBMoney
            
            dbl原始金额 = Val(.TextMatrix(Row, Col))
            strInput = Format(Val(strInput), "0.00")
            .EditText = strInput
            mPatiInfor.bln退款标志 = IIf(Val(strInput) > 0, False, True)
            
            dblMoney = RoundEx(Val(strInput) - dbl原始金额, 6)
            mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 + dblMoney, 6)
            mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 - dblMoney, 6)
            .TextMatrix(Row, Col) = strInput
            Call SetNextBalanceCmdVisible
        Case .ColIndex("结算号码"), .ColIndex("备注")
            If strInput = "" Then Exit Sub
            If zlCommFun.StrIsValid(strInput, .EditMaxLength, , .ColKey(Col)) = False Then Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub ReCalcYBMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算医保金额
    '编制:刘兴洪
    '日期:2015-01-21 15:41:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long
    Dim dbl个人帐户 As Double, dbl医保基金 As Double, dblMoney As Double
    Dim str结算方式 As String
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            If str结算方式 <> "" Then
                 varData = Split(.RowData(i) & "|||", "|")
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                 dblMoney = Val(.TextMatrix(i, .ColIndex("结算金额")))
                 Select Case Val(.TextMatrix(i, .ColIndex("结算性质")))
                 Case 3 '个人帐户
                    dbl个人帐户 = dbl个人帐户 + dblMoney
                 Case 4 '医保基金
                    dbl医保基金 = dbl医保基金 + dblMoney
                 End Select
            End If
        Next
    End With
        
    mBalanceInfor.dbl医保支付合计 = RoundEx(dbl个人帐户 + dbl医保基金, 5)
    mYBInFor.cur个帐支付 = dbl个人帐户
    mYBInFor.cur统筹支付 = dbl医保基金
    
    staThis.Panels(5).Text = Format(mYBInFor.cur个帐余额, "0.00")
    staThis.Panels(5).Visible = True
 
    txtBalance(Idx_本次结帐).Enabled = False

    'bytFun-0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
    Call SetOperationCtrl(IIf(mBalanceInfor.blnSaveBill, 2, 0))
    '显示医保虚算信息:bytFun-0-医保预算信息显示
    Call ShowLedDisplayBank(0)
    Call LoadCurOwnerPayInfor    '加载支付合计
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub


Private Function GetYBTotal(ByVal lngRow As Long, _
    Optional ByRef dbl个人帐户 As Double, _
    Optional ByRef dbl医保基金 As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保支付总额
    '入参:lngRow-不包含的行
    '返回:医保支付总额
    '编制:刘兴洪
    '日期:2015-01-21 15:41:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblMoney As Double, str结算方式 As String
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(.Row, .ColIndex("结算方式")))
            If str结算方式 <> "" And i <> lngRow Then
                '结算性质:结算方式.性质
                 dblMoney = Val(.TextMatrix(i, .ColIndex("结算金额")))
                 Select Case Val(.TextMatrix(i, .ColIndex("结算性质")))
                 Case 3 '个人帐户
                    dbl个人帐户 = dbl个人帐户 + dblMoney
                 Case 4 '医保基金
                    dbl医保基金 = dbl医保基金 + dblMoney
                 End Select
            End If
        Next
    End With
    
    GetYBTotal = dbl个人帐户 + dbl医保基金
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Sub vsDeposit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dbl预交余额 As Double, dbl冲预交 As Double
    Dim i As Long
    Dim dblMoney As Double
    
    If mblnNoTrigger Then
        mblnNoTrigger = False
        Exit Sub
    End If
    
    With vsDeposit
        If IsNumeric(.TextMatrix(Row, .ColIndex("冲预交"))) = False And .TextMatrix(Row, .ColIndex("冲预交")) <> "" Then
            MsgBox "请输入正确的冲预交金额!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("冲预交")) = ""
            If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        If Val(.TextMatrix(Row, .ColIndex("余额"))) < Val(.TextMatrix(Row, .ColIndex("冲预交"))) Then
            MsgBox "输入的冲预交金额过大,请重新输入!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("冲预交")) = ""
            If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        If Val(.TextMatrix(Row, .ColIndex("余额"))) < 0 And Val(.TextMatrix(Row, .ColIndex("冲预交"))) > 0 Then
            MsgBox "请输入正确的冲预交金额!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("冲预交")) = ""
            If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        If Val(.TextMatrix(Row, .ColIndex("余额"))) > 0 And Val(.TextMatrix(Row, .ColIndex("冲预交"))) < 0 Then
            MsgBox "请输入正确的冲预交金额!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("冲预交")) = ""
            If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            dbl预交余额 = RoundEx(dbl预交余额 + Val(.TextMatrix(i, .ColIndex("余额"))), 5)
            dbl冲预交 = RoundEx(dbl冲预交 + Val(.TextMatrix(i, .ColIndex("冲预交"))), 5)
        Next i
        If Val(dbl预交余额) < Val(dbl冲预交) Then
            MsgBox "输入的冲预交金额过大,请重新输入!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("冲预交")) = ""
            If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        .TextMatrix(Row, .ColIndex("冲预交")) = Format(.TextMatrix(Row, .ColIndex("冲预交")), "0.00")
        mblnManualEdit = True
        txtBalance(Idx_冲预交).Text = Format(dbl冲预交, "0.00")
        mBalanceInfor.dbl冲预交合计 = dbl冲预交
        
        If chkDeposit.Visible Then Exit Sub
        dblMoney = RoundEx(Val(txtBalance(Idx_冲预交).Text), 6)
        
        If mblnNotChange = False Then
            If Val(dblMoney) > Val(mPatiInfor.dbl实际余额) Then
                MsgBox "当前输入的冲预交大于预交余额,不能继续!" & vbCrLf & "实际余额:" & Format(mPatiInfor.dbl实际余额, "0.00") & vbCrLf & "冲预交:" & Format(Val(txtBalance(Idx_冲预交).Text), "0.00")
                .TextMatrix(Row, .ColIndex("冲预交")) = ""
                If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
                mblnManualEdit = False
                Exit Sub
            End If
        End If
        
        If Val(.TextMatrix(Row, .ColIndex("冲预交"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
        
        If Not mBalanceInfor.bln预交刷卡 Then
            If CheckDepositValied(True) = False Then mblnManualEdit = False: Exit Sub
        End If
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor(True)
        mbln已报价 = False
        mblnManualEdit = False
    End With
End Sub

Private Sub vsDeposit_AfterMoveColumn(ByVal Col As Long, Position As Long)
     zl_vsGrid_Para_Save mlngModul, vsDeposit, Me.Name, "预交列表"
End Sub

Private Sub vsDeposit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsDeposit, OldRow, NewRow, OldCol, NewCol
    Call SetUpDown
End Sub

Private Sub vsDeposit_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Long
    If mstrNoSort <> "" Then
        With vsDeposit
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据号")) = mstrNoSort Then
                    .Select i, Col
                    Exit For
                End If
            Next i
        End With
    End If
    Call RecalcDepositMoney(2, Val(mBalanceInfor.dbl冲预交合计))
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor(True)
End Sub

Private Sub vsDeposit_BeforeSort(ByVal Col As Long, Order As Integer)
    mstrNoSort = vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("单据号"))
End Sub

Private Sub vsDeposit_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
     zl_vsGrid_Para_Save mlngModul, vsDeposit, Me.Name, "预交列表"
End Sub

Private Sub vsDeposit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnBatchState Then Cancel = True: Exit Sub
    If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐) Then Cancel = True
    If chkCancel.Value = 1 Then Cancel = True
    If Val(vsDeposit.TextMatrix(Row, vsDeposit.ColIndex("编辑状态"))) <> 0 Then Cancel = True
    If Col <> vsDeposit.ColIndex("冲预交") Then Cancel = True
End Sub

Private Sub vsDeposit_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDeposit
        If Col = .ColIndex("单据号") Then Cancel = True: Exit Sub
    End With
End Sub
Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub
Private Sub txtPatient_LostFocus()
    IDKind.SetAutoReadCard (False)
End Sub
Private Sub txtPatient_Validate(Cancel As Boolean)
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If txtPatient.Text <> mrsInfo!姓名 Then txtPatient.Text = mrsInfo!姓名
End Sub
Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    
    If txtPatient.Locked Then Exit Sub
    If KeyAscii = 13 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If mrsInfo!姓名 = txtPatient.Text Then
                    If vsBlance.Enabled And vsBlance.Enabled Then
                        vsBlance.SetFocus
                        vsBlance.ShowCell vsBlance.Row, vsBlance.Col
                    Else
                        zlCommFun.PressKey vbKeyTab
                    End If
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '病人选择器
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        If mEditType = g_Ed_门诊结帐 Then
            Call cmdYB_Click
            Exit Sub
        Else
            With frmPatiSelect
                .mstrPrivs = mstrPrivs
                .mbytUseType = 3
                Set .mfrmParent = Me
                .Show 1, Me
                mty_ModulePara.intPatientRange = Val(zlDatabase.GetPara("显示结清病人", glngSys, mlngModul, 0))
            End With
        End If
    Else
        If IDKind.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        strInput = txtPatient.Text
        mstrPatient = txtPatient.Text
        Call FindPati(IDKind.GetCurCard, blnCard, strInput)
    End If
End Sub
Private Sub Led_ClearDisplayPatient()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除Led显示屏
    '编制:刘兴洪
    '日期:2014-12-31 10:38:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mstrInNO <> "" Or Not gblnLED Then Exit Sub
    If mEditType = g_Ed_单据查看 Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub


Private Sub HideYBMoneyInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:隐藏统筹支付信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-12-31 11:39:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    staThis.Panels(5).Text = ""
    staThis.Panels(5).Visible = False
'    lbl个人帐户.Visible = False
End Sub

Private Sub NewBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结帐界面
    '编制:刘兴洪
    '日期:2014-12-31 10:05:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearCustomType '清除自定义类型相关变量
    Call SetBatchControl(True)
    Call Led_ClearDisplayPatient '清除Led屏的显示
    Set mrsInfo = New ADODB.Recordset '清除病人信息
    mblnCurMzBalanceNo = False
    mbln已报价 = False
    
    '清除费用及预交信息
    Call InitGrid
    '清除结算信息
    Call ClearBalance '清除结算信息
    Call HideYBMoneyInfo    '隐藏统筹支付及余额
    Call InitBalanceCondition   '初始化结帐条件相关变量
    Call InitPatiBalanceVariableCon     '清除病人结帐相关变量
     
    Call SetControlEnabled(True) '设置控件的相关状态
    
    txtPatient.ForeColor = Me.ForeColor
   
    pic状态.Visible = False: lbl状态.Caption = "":  lbl付款方式.Caption = ""
    txtPatient.Text = "":    txtSex.Text = "":      txtOld.Text = ""
    txt费别.Text = "":       txt标识号.Text = "":   txtBed.Text = "": txt科室.Text = ""
    
    txtBegin.Text = "____-__-__": txtEnd.Text = "____-__-__"
    txtPatiBegin.Text = "____-__-__": txtPatiEnd.Text = "____-__-__":    txtPatiEnd.Tag = "____-__-__"
    txtDate.Text = "____-__-__ __:__:__": txt天数.Text = ""
    txtBalance(Idx_结帐说明).Text = ""
    lblBed.Visible = False:     txtBed.Visible = False
    lbl标识号.Visible = True:  txt标识号.Visible = True
    lbl科室.Visible = False:    txt科室.Visible = False
    picOwnerFee.Visible = False
    mblnNotify = False
    mstrBalanceLimit = ""
    mstrForceNote = ""
    mstrCardPara = ""
        
    lblPrevious.Visible = False
    lblPrevious.Caption = ""
    
    lblTicketCount.Caption = "预交款收据:"
    staThis.Panels(2) = ""
    staThis.Panels(3) = ""
    staThis.Panels(4) = ""
    staThis.Panels(4).Visible = False
    lblBalanceType.Visible = False
    Call SetOperationCtrl(0)
    Call SetFeeListColumnShow
    Call SetPatiConsControlVisible
    Call SetOperatonCommandCaption
    Call SetDefaultPayType
End Sub

Private Sub SetPatiConsControlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置病人条件控件的显示
    '编制:刘兴洪
    '日期:2014-12-31 14:26:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMzBalance As Boolean, blnVisible As Boolean
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then
        blnMzBalance = True
    ElseIf mEditType = g_Ed_住院结帐 Then
        blnMzBalance = False
    End If
    lblBed.Visible = Not blnMzBalance
    lbl科室.Visible = Not blnMzBalance
    txt科室.Visible = Not blnMzBalance
    blnVisible = mEditType = g_Ed_门诊结帐 And InStr(mstrPrivs, ";保险结算;") > 0
    cmdYB.Visible = blnVisible
    If blnVisible And Not mblnMC_TwoMode And InStr(mstrPrivs, ";门诊费用结帐;") = 0 Then
       cmdYB.Visible = False
    End If
    
    lblPatiTime.Visible = Not blnMzBalance
    lblPatiTimeRange.Visible = Not blnMzBalance
    txtPatiBegin.Visible = Not blnMzBalance
    txtPatiEnd.Visible = Not blnMzBalance
    txt天数.Visible = Not blnMzBalance
    lblDayName.Visible = Not blnMzBalance
    
    lblPatiNums.Caption = IIf(blnMzBalance, "门诊次数", "住院次数")
    lblPatiNums.Visible = True
    cboPatiNums.Visible = True
     
    opt中途.Visible = Not blnMzBalance
    opt出院.Visible = Not blnMzBalance
    
    txtBed.Visible = Not blnMzBalance
    lblBed.Visible = Not blnMzBalance
    
    chkCancel.Visible = (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐)
    lblDelCaption.Visible = mblnViewCancel Or mEditType = g_Ed_取消结帐 Or mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废
    
    Call picNO_Resize
    If (mEditType <> g_Ed_门诊结帐 And mEditType <> g_Ed_住院结帐) _
        Or chkCancel.Value Or mEditType = g_Ed_单据查看 Then
        '非结帐时，不存在以下条件
        opt中途.Visible = False: opt出院.Visible = False
        lblPatiNums.Visible = False
        cboPatiNums.Visible = False
        cmdMore.Visible = False
        cmdYB.Visible = False
    Else
        cmdMore.Visible = True
    End If
    
    If blnMzBalance Then
        lbl标识号.Caption = "门诊号"
    End If
    
    Call MovePatiConsControl
End Sub


Private Sub MovePatiConsControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整控件位置
    '编制:刘兴洪
    '日期:2014-12-31 15:03:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMzBalance As Boolean
    Dim lngStep As Long, sngLeft As Single
    Dim objPan As Pane
    
    '1.住院就是原界面
    Set objPan = dkpMain.FindPane(Pan_PatiCon)
    If objPan Is Nothing Then Exit Sub
    
    If mEditType = g_Ed_住院结帐 And chkCancel.Value <> 1 Then
        objPan.MaxTrackSize.Height = 550 \ Screen.TwipsPerPixelY
        objPan.MinTrackSize.Height = 550 \ Screen.TwipsPerPixelY
        dkpMain.RecalcLayout
        Exit Sub
    End If
    
    If mEditType = g_Ed_门诊结帐 And chkCancel.Value <> 1 Then
        '2.门诊结帐界面
        objPan.MaxTrackSize.Height = 550 \ Screen.TwipsPerPixelY
        objPan.MinTrackSize.Height = 550 \ Screen.TwipsPerPixelY
        dkpMain.RecalcLayout
        
        lblSex.Left = cmdYB.Left + cmdYB.Width + 120
        txtSex.Left = lblSex.Left + lblSex.Width + 30
        
        lblOld.Left = txtSex.Left + txtSex.Width + 120
        txtOld.Left = lblOld.Left + lblOld.Width + 30
        
        lblOld.Left = txtSex.Left + txtSex.Width + 120
        txtOld.Left = lblOld.Left + lblOld.Width + 30
        
        lbl费别.Left = txtOld.Left + txtOld.Width + 120
        txt费别.Left = lbl费别.Left + lbl费别.Width + 30
        
        lbl标识号.Left = txt费别.Left + txt费别.Width + 120
        txt标识号.Left = lbl标识号.Left + lbl标识号.Width + 30
        
        lblPatiNums.Top = 200
        cboPatiNums.Top = lblPatiNums.Top - 60
        cboPatiNums.Width = picOwnerFee.Left - cboPatiNums.Left - 60
        
        lblFsTime.Top = lblPatiNums.Top + 500
        lblFsTimeRange.Top = lblFsTime.Top
        txtBegin.Top = lblFsTime.Top - 60
        txtEnd.Top = lblFsTime.Top - 60
        
        lblDate.Top = lblFsTime.Top + 500
        txtDate.Top = lblDate.Top - 60
        
        cmdMore.Top = lblDate.Top - 90
        cmdMore.Left = txtEnd.Left + txtEnd.Width - cmdMore.Width
        
        Frame3.Top = cmdMore.Top + cmdMore.Height + 200
        
        lblBalance(3).Top = Frame3.Top + 200
        chkDeposit.Top = Frame3.Top + 200
        txtBalance(3).Top = chkDeposit.Top - 60
        lbl预交余额.Top = chkDeposit.Top
        
        vsBlance.Top = chkDeposit.Top + chkDeposit.Height + 120
        vsBlance.Height = txtOwe.Top - 60 - vsBlance.Top
        
        Exit Sub
    End If
    
    '3.其他界面(作废,重结,查阅等)
    If mEditType = g_Ed_重新结帐 Then
        lblFsTime.Top = 200
        lblFsTimeRange.Top = lblFsTime.Top
        txtBegin.Top = lblFsTime.Top - 60
        txtEnd.Top = lblFsTime.Top - 60
        
        If lbl标识号.Caption = "门诊号" Then
            lblPatiTime.Top = lblFsTime.Top
            txtPatiBegin.Top = lblPatiTime.Top - 60
            txtPatiEnd.Top = txtPatiBegin.Top
            lblPatiTimeRange.Top = lblPatiTime.Top
            lblDate.Top = lblFsTime.Top + 500
        Else
            lblPatiTime.Top = lblFsTime.Top + 500
            txtPatiBegin.Top = lblPatiTime.Top - 60
            txtPatiEnd.Top = txtPatiBegin.Top
            lblPatiTimeRange.Top = lblPatiTime.Top
            lblDate.Top = lblPatiTime.Top + 500
        End If
        
        
        txtDate.Top = lblDate.Top - 60
        txt天数.Top = txtDate.Top
        lblDayName.Top = lblDate.Top
        
        
        Frame3.Top = txtDate.Top + txtDate.Height + 200
        
        lblBalance(3).Top = Frame3.Top + 200
        chkDeposit.Top = Frame3.Top + 200
        txtBalance(3).Top = chkDeposit.Top - 60
        lbl预交余额.Top = chkDeposit.Top
        
        vsBlance.Top = chkDeposit.Top + chkDeposit.Height + 120
        vsBlance.Height = txtOwe.Top - 60 - vsBlance.Top
    End If
    
    If chkCancel.Value = 1 Or mEditType = g_Ed_重新作废 Or mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_取消结帐 Then
        lblFsTime.Top = 200
        lblFsTimeRange.Top = lblFsTime.Top
        txtBegin.Top = lblFsTime.Top - 60
        txtEnd.Top = lblFsTime.Top - 60
        
        lblPatiTime.Top = lblFsTime.Top + 500
        txtPatiBegin.Top = lblPatiTime.Top - 60
        txtPatiEnd.Top = txtPatiBegin.Top
        lblPatiTimeRange.Top = lblPatiTime.Top
        
        lblDate.Top = lblPatiTime.Top + 500
        txtDate.Top = lblDate.Top - 60
        txt天数.Top = txtDate.Top
        lblDayName.Top = lblDate.Top
        
        Frame3.Top = txtDate.Top + txtDate.Height + 200
        
        lblBalance(3).Top = Frame3.Top + 200
        chkDeposit.Top = Frame3.Top + 200
        txtBalance(3).Top = chkDeposit.Top - 60
        lbl预交余额.Top = chkDeposit.Top
        
        vsBlance.Top = chkDeposit.Top + chkDeposit.Height + 120
        vsBlance.Height = txtOwe.Top - 60 - vsBlance.Top
    End If
End Sub

Private Sub SetPatiEnabled(blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置病人相关的编辑属性
    '编制:刘兴洪
    '日期:2015-01-04 16:39:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    chkCancel.Enabled = blnEnabled And Not mPatiInfor.bln连续结帐
    cmdYB.Enabled = blnEnabled
    txtPatient.Locked = Not blnEnabled
    txtPatiBegin.Enabled = blnEnabled
    txtPatiEnd.Enabled = blnEnabled
    txtBalance(Idx_本次结帐).Locked = (InStr(mstrPrivs, ";结帐设置;") = 0)
    
    If mEditType = g_Ed_门诊结帐 Then
        opt中途.Enabled = False
        opt出院.Enabled = False
    Else
        opt中途.Enabled = blnEnabled
        opt出院.Enabled = blnEnabled
    End If
End Sub

Private Sub SetControlEnabled(blnEanbled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:控制结帐状态
    '入参:blnEanbled-是否有效
    '编制:刘兴洪
    '日期:2014-12-31 12:01:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim EditType As gBalanceBill
    
    EditType = mEditType
    If chkCancel.Value = 1 And (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) Then
        EditType = g_Ed_结帐作废
    End If
    
    Select Case EditType
    Case g_Ed_门诊结帐
        txtPatient.Locked = Not blnEanbled
        chkCancel.Enabled = blnEanbled And Not mPatiInfor.bln连续结帐
        cmdYBBalance.Enabled = blnEanbled
        cmdYB.Enabled = blnEanbled
        txtPatient.Locked = Not blnEanbled
        txtBalance(Idx_本次结帐).Locked = (InStr(mstrPrivs, ";结帐设置;") = 0)
        txtBalance(Idx_本次结帐).Enabled = Not txtBalance(Idx_本次结帐).Locked
        txtBalance(Idx_结帐说明).Enabled = blnEanbled
        
        txtInvoice.Enabled = blnEanbled
        cboPatiNums.Enabled = blnEanbled And InStr(mstrPrivs, ";结帐设置;") > 0
        txtBegin.Enabled = False    '不允许修改日期(118827,在结帐设置中更改)
        txtEnd.Enabled = False
        txtPatiBegin.Enabled = False
        txtPatiEnd.Enabled = False
        opt中途.Enabled = False
        opt出院.Enabled = False
    Case g_Ed_住院结帐
        txtPatient.Locked = Not blnEanbled
        chkCancel.Enabled = blnEanbled And Not mPatiInfor.bln连续结帐
        cmdYBBalance.Enabled = blnEanbled
        txtPatient.Locked = Not blnEanbled
        txtPatiBegin.Enabled = blnEanbled
        txtPatiEnd.Enabled = blnEanbled
        
        cboPatiNums.Enabled = blnEanbled And InStr(mstrPrivs, ";结帐设置;") > 0
        txtInvoice.Enabled = blnEanbled
        opt中途.Enabled = blnEanbled
        opt出院.Enabled = blnEanbled
        opt出院.Caption = "出院结帐"
        txtBalance(Idx_本次结帐).Locked = (InStr(mstrPrivs, ";结帐设置;") = 0)
        txtBalance(Idx_本次结帐).Enabled = Not txtBalance(Idx_本次结帐).Locked
        txtBalance(Idx_结帐说明).Enabled = blnEanbled
        cboPatiNums.Enabled = blnEanbled And InStr(mstrPrivs, ";结帐设置;") > 0
    Case Else  'g_Ed_取消结帐, g_Ed_单据查看, g_Ed_结帐作废, g_Ed_重新结帐, g_Ed_重新作废
        IDKind.Enabled = False
        txtPatient.Locked = True
        chkCancel.Enabled = Not mPatiInfor.bln连续结帐 And IIf(mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐, True, False)
        cmdYBBalance.Enabled = False
        txtPatiBegin.Enabled = False
        txtPatiEnd.Enabled = False
        opt中途.Enabled = False
        opt出院.Enabled = False
        cboPatiNums.Enabled = False
        
        txtBegin.Enabled = False
        txtEnd.Enabled = False
        txtInvoice.Enabled = IIf(mEditType = g_Ed_重新结帐, True, False)
        
        txtBalance(Idx_本次结帐).Enabled = False
        txtBalance(Idx_结帐说明).Enabled = blnEanbled And mEditType = g_Ed_重新结帐
        If mEditType = g_Ed_单据查看 Or mEditType = g_Ed_取消结帐 Or mEditType = g_Ed_重新作废 Then
            txtInvoice.Enabled = False
            cboNO.Enabled = False
        End If
    End Select
    
    txtBegin.BackColor = IIf(txtBegin.Enabled, &H80000005, &H8000000F)
    txtEnd.BackColor = IIf(txtEnd.Enabled, &H80000005, &H8000000F)
          
    txtPatiBegin.BackColor = IIf(txtPatiBegin.Enabled, &H80000005, &H8000000F)
    txtPatiEnd.BackColor = IIf(txtPatiEnd.Enabled, &H80000005, &H8000000F)
    txtBalance(Idx_本次结帐).BackColor = IIf(txtBalance(Idx_本次结帐).Enabled, &H80000005, &H8000000F)
    txtBalance(Idx_结帐说明).BackColor = IIf(txtBalance(Idx_结帐说明).Enabled, &H80000005, &H8000000F)
    cboNO.BackColor = IIf(cboNO.Enabled, &H80000005, &H8000000F)
    txtInvoice.BackColor = IIf(txtInvoice.Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
                
End Sub


Private Sub InitBalanceCondition()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结帐条件相关变量
    '编制:刘兴洪
    '日期:2014-12-31 11:46:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjBalanceAll = New clsBalanceAllCon
    With mobjBalanceAll
        .strAllTime = ""
        .strAllDeptIDs = ""
        .strAllItem = ""
        .strAllDiag = ""
        .strAllClass = ""
        .strUnAuditTime = ""
        .strAllChargeType = ""  '34260
        .MinDate = #1/1/1900#
        .MaxDate = #1/1/1900#
        Set .rsAllTime = Nothing
        .strAllFullTims = ""
    End With
End Sub

Private Sub InitPatiBalanceVariableCon()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初台化病人相关结帐条件变量
    '编制:刘兴洪
    '日期:2014-12-31 11:56:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjBalanceCon = New clsBalanceCon
    With mobjBalanceCon
        .strTime = ""
        .strDeptIDs = ""
        .strClass = ""
        .strBaby = ""
        .strItem = ""
        .strDiag = ""
        .bytKind = 0
        .dtBeginDate = CDate("0:00:00"):
        .dtEndDate = CDate("0:00:00")
        .strChargeType = ""
        .blnCurBalanceOwnerFee = False
    End With
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call NewBill
    txtPatient.Text = strInput
    '刘兴洪:27503
    If mty_ModulePara.bln结帐后不清信息 Then
        If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '主要是要保留信息,在确定后需要减判刑断
    End If
    
    If mOldOneCard.blnOneCard And Not mobjICCard Is Nothing And objCard.名称 Like "IC卡*" And objCard.系统 Then
        Call SetOldOneCardBalance  '显示老一卡通余额
    End If
    Call LoadPatientInfo(objCard, blnCard)
End Sub
Private Sub SetOldOneCardBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置一卡通结算方式
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 09:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curOneCard As Currency, strName As String
    If mOldOneCard.blnOneCard = False Or mobjICCard Is Nothing Then Exit Sub
    curOneCard = mobjICCard.GetSpare(strName)
    If curOneCard <> 0 Then
       mOldOneCard.rsOneCard.Filter = "名称='" & strName & "'"
       If mOldOneCard.rsOneCard.RecordCount > 0 Then mOldOneCard.strOneCard = mOldOneCard.rsOneCard!结算方式
    End If
    staThis.Panels(2).Text = "卡余额:" & Format(curOneCard, "0.00") & "元"
End Sub





 
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    ByVal blnCard As Boolean, Optional ByVal lng主页ID As Long, _
    Optional blnOnlyReadPati As Boolean, Optional ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '     lng主页ID=读取指定住院次数的病人信息
    '     intInsure-险类(主要是重结或作废时传入)
    '     blnOnlyReadPati-只读取病人信息，不作相关检查(主要是重结或作废时传入)
    '出参:
    '返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close,strInput返回是用来判断是否已提示过,避免再次提示没有找到病人
    '编制:刘兴洪
    '日期:2015-01-04 12:15:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strWhere As String, strField As String, bytMzMode As Byte
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strRange As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    Dim str住院号 As String
    
    mstrPassWord = "": mstrInputInNo = "": strWhere = ""
    mblnReadByZYNo = False: mlngCardTypeID = 0
    
    On Error GoTo errH
    
    strField = ",A.当前科室ID"
    
    bytMzMode = mYBInFor.bytMCMode
    
    
    If mEditType = g_Ed_住院结帐 Then
        If Not (blnCard = True And objCard.名称 Like "姓名*") Then    '非刷卡部分
            If Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
                str住院号 = Val(Mid(strInput, 2))
            ElseIf objCard.名称 = "住院号" Then
                str住院号 = Val(strInput)
            End If
        End If
    End If
    
    If mEditType = g_Ed_门诊结帐 Then   '门诊
        strWhere = strWhere & " And   A.主页ID=B.主页ID(+)"
        '问题:43730
        bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as 当前科室ID"
    Else
        If lng主页ID <> 0 Then
            strField = ",Decode(A.主页ID,[3],A.当前科室ID,NULL) as 当前科室ID"
            strWhere = " And B.主页ID=[3]"
        ElseIf str住院号 <> "" Then '按住院号查找病人
            strWhere = "And (B.病人ID,B.主页ID)=(Select max(病人ID)as 病人ID, Max(主页ID) As 主页ID From 病案主页 Where 住院号=[2])"
        Else
            strWhere = " And A.主页ID=B.主页ID(+)"
        End If
        bytMzMode = 2
    End If
    
    If intInsure <> 0 Then
        strField = strField & ",[4] as 险类"
    ElseIf bytMzMode = 0 Then
        strField = strField & ",NULL as 险类"
    ElseIf bytMzMode = 1 Then
        strField = strField & ",A.险类 as 险类"
    Else
        strField = strField & ",B.险类 as 险类"
    End If

    strSQL = _
    " Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号,nvl(B.住院号,A.住院号) as 住院号,B.入院病床,B.出院病床," & _
    "       nvl(B.姓名,A.姓名) as 姓名, nvl(B.性别,Nvl(A.性别,'未知')) as  性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
    "       Nvl(B.费别,A.费别) as 费别,C.名称 as 入院科室" & strField & ",D.名称 as 出院科室,B.出院科室ID," & _
    "       E.卡号,E.医保号,E.密码," & _
    "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志, " & _
    "       B.入院日期,B.出院日期,B.病人性质,B.病人类型,Decode(B.病人ID,Null,A.在院,Decode(B.出院日期,Null,1,0)) As 在院" & _
    " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
    " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+)   " & strWhere & _
    "   And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
    "   And B.入院科室ID=C.ID(+) And B.出院科室ID=D.ID(+)"
        
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        mlngCardTypeID = lng卡类别ID
        strSQL = strSQL & " And A.病人ID=[1] "
        
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strInput = Mid(strInput, 2)
        If mEditType <> g_Ed_住院结帐 Then
            strSQL = strSQL & " And A.病人ID=(Select nvl(Max(病人ID),0) As 病人ID From 病案主页   Where  住院号=[2])"
        Else
           mblnReadByZYNo = True
           mstrInputInNo = mobjBalanceAll.zlGetNumsFromZyNo(Val(strInput))
           If InStr(mstrInputInNo, ",") > 0 Then mstrInputInNo = "": mblnReadByZYNo = False
        End If
    
    Else '当作姓名
        mlngCardTypeID = objCard.接口序号
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If mrsInfo!姓名 = Trim(txtPatient.Text) Then
                        GetPatient = True
                        Exit Function
                    End If
                End If
                
                If mty_ModulePara.intPatientRange > 0 Then
                    Select Case mty_ModulePara.intPatientRange
                        Case 1  '任何费用未结清病人
                            strRange = ""
                        Case 2  '体检未结清的病人
                            strRange = " And C.来源途径 = 4"
                        Case 3  '住院未结清的病人
                            strRange = " And C.来源途径 = 2"
                        Case 4  '门诊未结清的病人
                            strRange = " And C.来源途径 = 1"
                    End Select
                    strPati = " And Exists(Select 1 From 病人未结费用 C Where C.病人id=A.病人ID And Nvl(C.主页ID,0)=A.主页ID" & strRange & ")"
                End If
                
                 '通过姓名查找
                strPati = "" & _
                " Select A.病人ID as ID,A.病人ID,A.姓名,A.住院号, A.门诊号, nvl(B.性别,Nvl(A.性别,'未知')) as  性别, A.年龄, A.住院次数, A.家庭地址, A.工作单位," & vbNewLine & _
                "   To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,  To_Char(B.入院日期,'YYYY-MM-DD') as 入院日期, To_Char(B.出院日期,'YYYY-MM-DD') as 出院日期" & vbNewLine & _
                " From 病人信息 A, 病案主页 B" & vbNewLine & _
                " Where A.病人id = B.病人id(+) And A.主页ID = B.主页id(+) And A.停用时间 Is Null And A.姓名 = [1] " & vbNewLine & strPati & vbNewLine & _
                " Order By Decode(住院号, Null, 1, 0), 入院日期 Desc"
                        
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!病人ID)
                    strSQL = strSQL & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
                End If
                
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "身份证号", "二代身份证", "身份证"
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng病人ID
                blnHavePassWord = True
                strSQL = strSQL & " And A.病人ID=[1] "
            Case "IC卡号"
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng病人ID
                blnHavePassWord = True
                strSQL = strSQL & " And A.病人ID=[1] "
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                
                If mEditType <> g_Ed_住院结帐 Then
                    strSQL = strSQL & " And A.病人ID=(Select nvl(Max(病人ID),0) As 病人ID From 病案主页   Where  住院号=[2])"
                Else
                   mblnReadByZYNo = True
                   mstrInputInNo = mobjBalanceAll.zlGetNumsFromZyNo(Val(strInput))
                   If InStr(mstrInputInNo, ",") > 0 Then mstrInputInNo = "": mblnReadByZYNo = False
                   
                End If
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
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, lng主页ID, intInsure)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    If mstr主页Id <> "" Then mstrInputInNo = mstr主页Id: mblnReadByZYNo = True: mstr主页Id = ""
    mYBInFor.intInsure = Val(NVL(mrsInfo!险类))
    
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then
        mstrPassWord = NVL(mrsInfo!卡验证码)
    End If
    
    If blnOnlyReadPati Then GetPatient = True: Exit Function
    
    '检查死亡情况:如果死亡则提示
    '34681:35686
    If zlCheckPatiIsDeath(Val(NVL(mrsInfo!病人ID))) = True Then
        pic死亡.Visible = True
        If MsgBox("注意:" & vbCrLf & "    该病人已经死亡,是否继续结帐?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            pic死亡.Visible = False
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
        End If
    Else
        pic死亡.Visible = False
    End If
    
    '需要再次检查,以防结帐期间已审核的病人被取消审核
    '36209
    If (InStr(mstrPrivs, ";未审核病人中途结帐;") = 0 And opt中途.Value _
        Or InStr(mstrPrivs, ";未审核病人出院结帐;") = 0 And opt出院.Value) _
        And mEditType = g_Ed_住院结帐 Then
        If Not Chk病人审核(mrsInfo!病人ID, Val(NVL(mrsInfo!主页ID))) Then
            If MsgBox("待结帐费用中包含病人第" & Val(NVL(mrsInfo!主页ID)) & "次住院未审核的费用记录。" & vbCrLf & _
                " 是否继续结帐?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
            End If
        End If
    End If
    
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

Private Sub LoadPatientInfo(ByVal objCard As Card, ByVal blnCard As Boolean, _
    Optional ByVal intInsure As Integer, _
    Optional ByVal lng主页ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '入参:objCard-当前刷处理的卡对象
    '     blnCard-是否刷卡
    '     intInsure-当前的险类
    '     lng主页ID-读取指定住院次数的病人信息
    '编制:刘兴洪
    '日期:2015-01-04 12:12:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, strSQL As String
    Dim tyPatiInfor As ty_Pati_Infor
    Dim blnICCard As Boolean, curDue As Currency, blnIDCard As Boolean
    Dim blnNotClearPati As Boolean
    Dim lngPageID As Long
    Dim strPage() As String
    
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset

    txtPatient.ForeColor = Me.ForeColor
    
    mPatiInfor = tyPatiInfor '清空病人信息
    If objCard.名称 Like "IC卡*" And objCard.系统 = True Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 = True Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    staThis.Panels(2).Text = ""
    
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCard, lng主页ID, , intInsure) Then
        If txtPatient.Text = "" Then MsgBox "没有找到该病人,请检查输入内容是否正确！", vbInformation, gstrSysName
        txtPatient.PasswordChar = "": txtPatient.Text = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        mstr本次住院日期 = ""
        Call ReInitPatiInvoice

        Exit Sub
    End If
    
    mstr本次住院日期 = ""
    '就诊卡密码检查
    If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
    
    If gTy_System_Para.TY_Balance.bln刷卡输入密码 _
        And (blnCard Or ((blnICCard Or blnIDCard Or IDKind.GetCurCard.接口序号 <> 0) And mstrPassWord <> "")) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            GoTo ExitHandle
        End If
    End If
    
    '102236,调用外挂部件接口
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
        '    ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
        '    ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
        ''功能：检查当前病人是否是指定的特殊病人
        ''返回：true时允许继续操作，False时不允许操作
        ''参数：
        ''      lngSys,lngModual=当前调用接口的主程序系统号及模块号
        ''      lngType 操作类型：1－门诊挂号，2－住院入院，3－门诊收费，4－住院结帐，5－门诊结帐。
        ''      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
        ''      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
        ''      strPatiInforXML-病人信息:针对未建档病人传入，"姓名，性别，年龄，出生日期，医保号，身份证号"，出生日期 格式:2016-11-11 12:12:12
        ''                      固定格式：<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
        ''      strReserve=保留参数,用于扩展使用
        Dim blnChecked As Boolean
        blnChecked = gobjPlugIn.PatiValiedCheck(glngSys, mlngModul, IIf(mEditType = g_Ed_门诊结帐, 5, 4), Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID)), "")
        If Err <> 0 Then
            Call zlPlugInErrH(Err, "PatiValiedCheck"): Err.Clear
        Else
            If blnChecked = False Then GoTo ExitHandle
        End If
        On Error GoTo errHandle
    End If
        
    '问题:27690
    If mYBInFor.intInsure = 0 Then
        If InStr(1, mstrPrivs, ";普通病人结算;") = 0 Then
            MsgBox "你没有权限对非保险病人进行结算。", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
    End If
    
    '医保相关判断
    If mYBInFor.intInsure <> 0 Then
        If InStr(mstrPrivs, ";保险结算;") = 0 Then
            MsgBox "你没有权限对保险病人进行结算。", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
        
        If mYBInFor.strYBPati <> "" And intInsure <> mYBInFor.intInsure Then
            MsgBox "病人登记的险类与医保身份验证的险类不符。", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
        
        If mYBInFor.bytMCMode = 1 And Not IsNull(mrsInfo!当前科室id) Then
            MsgBox "在院病人不能进行门诊医保身份验证。", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
        Call InitInsurePara(Val(NVL(mrsInfo!病人ID)), mYBInFor.intInsure)
    ElseIf mYBInFor.strYBPati <> "" Then
        MsgBox "病人身份验证成功,但病人登记的险类为空！", vbInformation, gstrSysName
        GoTo ExitHandle
    End If
    
    If mblnReadByZYNo Then
        strPage = Split(mstrInputInNo, ",")
        For i = 0 To UBound(strPage)
            If Val(strPage(i)) > lngPageID Then lngPageID = Val(strPage(i))
        Next i
        '问题:34763 检查病人是否存在备注信息
        If zlCheckPatiIsMemo(Val(NVL(mrsInfo!病人ID)), lngPageID) = True Then
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, Val(NVL(mrsInfo!病人ID)), lngPageID, mobjInPatient)
        End If
        
        If lng主页ID = 0 Then
            '加载缺省出院状态
            If Not LoadDefaultOutStatu(mrsInfo!病人ID, lngPageID) Then GoTo ExitHandle
            '黑名单提醒
            If Not CheckPatiBlacklist(mrsInfo!病人ID) Then GoTo ExitHandle
                                                                                        
            '记帐未审核检查
            If Not CheckChargeAudit(mrsInfo!病人ID) Then GoTo ExitHandle
    
            '自动计算病人的床位费用和护级费用
            Call AutoCalcChareFee(Val(NVL(mrsInfo!病人ID)), lngPageID)
            
            '加载病人余额信息
            Call Load余额信息(Val(NVL(mrsInfo!病人ID)), IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2))
            
            '加载和检查应收款余额
            Call Load应收款信息(Val(NVL(mrsInfo!病人ID)))
            '88786,结帐不处理历史数据
            mblnDateMoved = False
        Else
            If Val(NVL(mrsInfo!在院)) = 1 And NVL(mrsInfo!状态, 0) <> 3 Then '在院病人()
                '状态:0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院
                If zlDatabase.GetPara("默认出院结帐", glngSys, mlngModul, "1") <> "0" Then
                    opt出院.Value = True
                    opt中途.Value = False
                Else
                    opt中途.Value = True
                    opt出院.Value = False
                End If
                If gbln在院不准结帐 Then opt中途.Value = True: opt出院.Enabled = False
            Else
                '出院病人(包含预出院的病人)
                 opt出院.Value = True
                 opt中途.Value = False
                 opt出院.Enabled = True
            End If
        End If
    Else
        '问题:34763 检查病人是否存在备注信息
        If zlCheckPatiIsMemo(Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID))) = True Then
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID)), mobjInPatient)
        End If
        
        If lng主页ID = 0 Then
            '加载缺省出院状态
            If Not LoadDefaultOutStatu(mrsInfo!病人ID, Val(NVL(mrsInfo!主页ID))) Then GoTo ExitHandle
            '黑名单提醒
            If Not CheckPatiBlacklist(mrsInfo!病人ID) Then GoTo ExitHandle
                                                                                        
            '记帐未审核检查
            If Not CheckChargeAudit(mrsInfo!病人ID) Then GoTo ExitHandle
    
            '自动计算病人的床位费用和护级费用
            Call AutoCalcChareFee(Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID)))
            
            '加载病人余额信息
            Call Load余额信息(Val(NVL(mrsInfo!病人ID)), IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2))
            
            '加载和检查应收款余额
            Call Load应收款信息(Val(NVL(mrsInfo!病人ID)))
            '88786,结帐不处理历史数据
            mblnDateMoved = False
        Else
            If Val(NVL(mrsInfo!在院)) = 1 And NVL(mrsInfo!状态, 0) <> 3 Then '在院病人()
                '状态:0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院
                If zlDatabase.GetPara("默认出院结帐", glngSys, mlngModul, "1") <> "0" Then
                    opt出院.Value = True
                    opt中途.Value = False
                Else
                    opt中途.Value = True
                    opt出院.Value = False
                End If
                If gbln在院不准结帐 Then opt中途.Value = True: opt出院.Enabled = False
            Else
                '出院病人(包含预出院的病人)
                 opt出院.Value = True
                 opt中途.Value = False
                 opt出院.Enabled = True
            End If
        End If
    End If

    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    txtPatient.Text = mrsInfo!姓名: txtSex.Text = NVL(mrsInfo!性别): txtOld.Text = NVL(mrsInfo!年龄)
    With mPatiInfor
        .lng病人ID = Val(NVL(mrsInfo!病人ID))
        .lng主页ID = Val(NVL(mrsInfo!主页ID))
        .str姓名 = NVL(mrsInfo!姓名)
        .str性别 = NVL(mrsInfo!性别)
        .str年龄 = NVL(mrsInfo!年龄)
        .bln出院 = Val(NVL((mrsInfo!在院))) <> 1
    End With
    '加载病人状态
    Call Load住院状态(Val(NVL(mrsInfo!病人ID)))
    
    cmdYB.Enabled = IIf(mEditType = g_Ed_门诊结帐, True, False)
    If mYBInFor.intInsure <> 0 Then
        staThis.Panels(4).Text = GetInsureName(mYBInFor.intInsure)
        staThis.Panels(4).Visible = True
        If mYBInFor.bytMCMode = 1 Then Call SetPatiEnabled(False)
        cmdOK.Enabled = False
    Else
        staThis.Panels(4).Visible = False
    End If
    If NVL(mrsInfo!病人类型) = "" And mYBInFor.intInsure <> 0 Then
        txtPatient.ForeColor = vbRed
    Else
        txtPatient.ForeColor = zlDatabase.GetPatiColor(NVL(mrsInfo!病人类型))
    End If
    
    lblPatiType.Caption = "病人类型:" & NVL(mrsInfo!病人类型)
    
    txt费别.Text = NVL(mrsInfo!费别)
    
    If mEditType = g_Ed_住院结帐 Then
        If Not IsNull(mrsInfo!住院号) Then
            txt标识号.Text = mrsInfo!住院号
            lbl标识号.Visible = True: txt标识号.Visible = True
            lbl标识号.Caption = "住院号"
        End If
        If Not IsNull(mrsInfo!入院科室) Then
            txtBed.Text = "" & NVL(mrsInfo!出院病床, mrsInfo!入院病床)
            txt科室.Text = NVL(mrsInfo!出院科室, mrsInfo!入院科室)
            lblBed.Visible = True: txtBed.Visible = True
            lbl科室.Visible = True: txt科室.Visible = True
        ElseIf Not IsNull(mrsInfo!出院科室) Then
            txtBed.Text = NVL(mrsInfo!出院病床)
            txt科室.Text = mrsInfo!出院科室
            lblBed.Visible = True: txtBed.Visible = True
            lbl科室.Visible = True: txt科室.Visible = True
        End If
    ElseIf mEditType = g_Ed_门诊结帐 Then
        If Not IsNull(mrsInfo!门诊号) Then
            txt标识号.Text = mrsInfo!门诊号
            lbl标识号.Visible = True: txt标识号.Visible = True
            lbl标识号.Caption = "门诊号"
        End If
    End If
    
    '异常单据处理
    If PatiErrBillPay(Val(NVL(mrsInfo!病人ID))) Then Exit Sub
    
    '显示病人要结帐内容,并初始化结算金额
    '-------------------------------------------------------------------------------------------
    If lng主页ID = 0 Then
        strTmp = ""
        If Not ShowBalance(True, strTmp, blnNotClearPati) Then
            If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            If blnNotClearPati = False Then GoTo ExitHandle:
            If cmdMore.Enabled And cmdMore.Visible Then cmdMore.SetFocus
            Exit Sub
        End If
        Call Led欢迎信息
    End If
    
    Call ReInitPatiInvoice  '重新刷新发票信息
    
    mblnNotChange = True
    Call txtBalance_Validate(Idx_冲预交, False)
    mblnNotChange = False
    
    If mobjBalanceAll.strAllTime <> "" Then
        '多次结帐弹出界面
        If UBound(Split(mobjBalanceAll.strAllTime, ",")) > 0 And mty_ModulePara.bln结帐后弹出界面 Then
            Call cmdMore_Click
        Else
            Call SkipSetFocus(0)
        End If
        Exit Sub
    End If
    Call SkipSetFocus(0)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
ExitHandle:
    Call NewBill
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub opt出院_Click()
    Dim dtBeginDate As Date, dtEndDate As Date
    If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) Then
        txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
        txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
        txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
        Call zlChangeDefaultTime
    End If
    If IsDate(txtPatiEnd.Text) = False Or IsDate(txtPatiBegin.Text) = False Then Exit Sub
    txt天数.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt中途.Value = True, 1, 0)
    If Val(txt天数.Text) = 0 Then txt天数.Text = 1
End Sub

Private Sub opt中途_Click()
    Dim dtBeginDate As Date, dtEndDate As Date
    If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) Then
        txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
        txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
        txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
        Call zlChangeDefaultTime
    End If
    If IsDate(txtPatiEnd.Text) = False Or IsDate(txtPatiBegin.Text) = False Then Exit Sub
    txt天数.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt中途.Value = True, 1, 0)
    If Val(txt天数.Text) = 0 Then txt天数.Text = 1
End Sub


Private Sub InitInsurePara(ByVal lng病人ID As Long, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2015-01-04 13:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, intInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
    MCPAR.结帐作废后打印回单 = gclsInsure.GetCapability(support结帐作废后打印回单, lng病人ID, intInsure)
    If mYBInFor.bytMCMode = 1 Then
        MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, intInsure)
        MCPAR.门诊必须传递明细 = gclsInsure.GetCapability(support门诊必须传递明细, lng病人ID, intInsure)
        MCPAR.门诊结算_结帐设置 = gclsInsure.GetCapability(support门诊结帐_结帐设置后调用接口, lng病人ID, intInsure)
        MCPAR.门诊病人结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure)
    Else
        MCPAR.未结清出院 = gclsInsure.GetCapability(support未结清出院, lng病人ID, intInsure)
        MCPAR.结算使用个人帐户 = gclsInsure.GetCapability(support结算使用个人帐户, lng病人ID, intInsure)
        MCPAR.出院结算必须出院 = gclsInsure.GetCapability(support出院结算必须出院, lng病人ID, intInsure)
        MCPAR.中途结算仅处理已上传部分 = gclsInsure.GetCapability(support中途结算仅处理已上传部分, lng病人ID, intInsure)
        MCPAR.结帐设置后调用接口 = gclsInsure.GetCapability(support结帐_结帐设置后调用接口, lng病人ID, intInsure)
        MCPAR.住院结算作废 = gclsInsure.GetCapability(support住院结算作废, lng病人ID, intInsure)
        MCPAR.门诊结算_结帐设置 = False
        MCPAR.出院病人结算作废 = gclsInsure.GetCapability(support出院病人结算作废, lng病人ID, intInsure)
        MCPAR.允许结多次住院费用 = gclsInsure.GetCapability(support允许一次结多次住院费用, lng病人ID, intInsure)
    End If
End Sub

Private Function LoadDefaultOutStatu(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnNoPromt As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载默认的出院状态
    '编制:刘兴洪
    '日期:2015-01-04 14:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    On Error GoTo errHandle
    
    If mYBInFor.bytMCMode = 1 Then LoadDefaultOutStatu = True: Exit Function
    If mEditType = g_Ed_门诊结帐 Then LoadDefaultOutStatu = True: Exit Function
    
    If lng主页ID = 0 Then
        opt出院.Value = True: opt出院.Enabled = False
        opt中途.Enabled = False: LoadDefaultOutStatu = True: Exit Function
    Else
        '默认结以前住院次数的,出院结帐
        If lng主页ID < Val(NVL(mrsInfo!主页ID)) Then
            opt出院.Enabled = True: opt出院.Value = True: LoadDefaultOutStatu = True: Exit Function
        End If
    End If
    
    '问题:30027:现在缺省的中途规则
    '       1.出院病人,默认为出院结帐 或者:没有"中途结帐"权限的,也默认为出院结帐
    '       2.在院病人(根据上次出院病人的选择的为准)
    '              默认出院结(即上次选择的中途结帐或住院结帐)参数为true,默认为出院结帐,否则默认为中途结帐
    If InStr(mstrPrivs, ";中途结帐;") = 0 Then
        opt出院.Value = True: opt中途.Enabled = False
    ElseIf Val(NVL(mrsInfo!在院)) = 1 And NVL(mrsInfo!状态, 0) <> 3 Then '在院病人()
        '状态:0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院
        If zlDatabase.GetPara("默认出院结帐", glngSys, mlngModul, "1") <> "0" Then
            opt出院.Value = True
        Else
            opt中途.Value = True
        End If
        If gbln在院不准结帐 Then opt中途.Value = True
    Else
        '出院病人(包含预出院的病人)
         opt出院.Value = True
    End If
    opt出院.Enabled = True
    
    If CheckOutBalanceIsvalied = False Then Exit Function
    
    If Not blnNoPromt Then
        If mEditType = g_Ed_门诊结帐 Then
            If Val(NVL(mrsInfo!在院)) = 1 And NVL(mrsInfo!状态, 0) <> 3 Then
                If MsgBox("当前病人在院，需要继续对该病人进行门诊结帐吗?", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        Else
            If Val(NVL(mrsInfo!在院)) = 1 And NVL(mrsInfo!状态, 0) <> 3 And gbln在院不准结帐 Then
                If MsgBox("当前病人在院，不允许出院结帐。 如果是出院结帐，请先将病人出院。" & _
                    vbCrLf & "需要对该病人进行中途结帐吗?", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Function
            End If
        End If
    End If
    
    If mblnFirst And mlngPatientID <> 0 Then
        If Val(NVL(mrsInfo!在院)) = 1 And NVL(mrsInfo!状态, 0) <> 3 Then '在院病人()
            '状态:0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院
            If zlDatabase.GetPara("默认出院结帐", glngSys, mlngModul, "1") <> "0" Then
                opt出院.Value = True
                opt中途.Value = False
            Else
                opt中途.Value = True
                opt出院.Value = False
            End If
            If gbln在院不准结帐 Then opt中途.Value = True: opt出院.Enabled = False
        Else
            '出院病人(包含预出院的病人)
             opt出院.Value = True
             opt中途.Value = False
        End If
        
        LoadDefaultOutStatu = True: Exit Function
    End If
    
    If opt中途.Value Then
        opt出院.Value = False
        If gbln在院不准结帐 Then opt出院.Enabled = False
        LoadDefaultOutStatu = True: Exit Function
    End If

    LoadDefaultOutStatu = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckOutBalanceIsvalied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:出院结帐检查
    '返回:出院结帐有效,返回成功,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 14:15:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(NVL(mrsInfo!主页ID)) = 0 Or Val(NVL(mrsInfo!在院)) <> 1 Then CheckOutBalanceIsvalied = True: Exit Function
    If Not gTy_System_Para.TY_Balance.bln在院不准结帐 Then CheckOutBalanceIsvalied = True: Exit Function
    If Not opt中途.Enabled Then
        MsgBox "在院病人不允许出院结帐,并且你没有中途结帐的权限,所以不能对该病人结帐!", vbInformation, gstrSysName
        Exit Function
    End If
    CheckOutBalanceIsvalied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckPatiBlacklist(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人黑名单
    '入参:lng病人ID-病人ID
    '返回:无黑名单或继续操作,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 14:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    '黑名单提醒
    On Error GoTo errHandle
    strTmp = inBlackList(mrsInfo!病人ID)
    If strTmp = "" Then CheckPatiBlacklist = True: Exit Function
    If MsgBox("病人""" & mrsInfo!姓名 & """在特殊病人名单中。" & vbCrLf & vbCrLf & "原因：" & vbCrLf & vbCrLf & "　　" & strTmp & vbCrLf & vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    CheckPatiBlacklist = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckChargeAudit(ByVal lng病人ID As Long, Optional blnSaveCheck As Boolean = False, Optional ByVal strTimes As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐审核检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 15:04:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'bytAuditing:0-不检查,1-检查并提示,2-检查并禁止
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If gTy_System_Para.TY_Balance.bytAuditing = 0 Then CheckChargeAudit = True: Exit Function
    '检查过了，退出
    If mblnNotify = True Then CheckChargeAudit = True: Exit Function
    If strTimes = "" Then
        strSQL = _
            "Select 1 From 住院费用记录 A" & _
                " Where 记帐费用=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And 病人ID=[1] And Not Exists" & _
                " (Select 1 From 药品收发记录 C Where A.ID = C.费用ID And Mod(C.记录状态, 3) = 1 And Nvl(C.摘要,'大一')='拒发' And instr( ',8,9,10,21,24,25,26,',','||C.单据||',')>0) And Not Exists" & _
                " (Select 1 From 病人医嘱发送 B Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱序号=B.医嘱ID And B.执行状态 = 2) And Rownum=1"
    Else
        strSQL = _
            "Select 1 From 住院费用记录 A" & _
                " Where 记帐费用=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And 病人ID=[1] And Not Exists" & _
                " (Select 1 From 药品收发记录 C Where A.ID = C.费用ID And Mod(C.记录状态, 3) = 1 And Nvl(C.摘要,'大一')='拒发' And instr( ',8,9,10,21,24,25,26,',','||C.单据||',')>0) And Not Exists" & _
                " (Select 1 From 病人医嘱发送 B Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱序号=B.医嘱ID And B.执行状态 = 2) And Rownum < 2 And a.主页ID In (Select Column_Value From Table(f_str2list([2]))) "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, strTimes)
    If rsTmp.RecordCount = 0 Then CheckChargeAudit = True: Exit Function
    Select Case gTy_System_Para.TY_Balance.bytAuditing
    Case 1
        If MsgBox("该病人还存在未审核的记帐费用，要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        If blnSaveCheck Then
            If opt出院.Value = True Then
                MsgBox "该病人还存在未审核的记帐费用,不能出院结帐！", vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("该病人还存在未审核的记帐费用，要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        Else
            If opt中途.Enabled Then opt中途.Value = True '使用中途结帐
        End If
    Case Else
    End Select
    CheckChargeAudit = True
End Function

Private Function AutoCalcChareFee(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动记帐计算
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '返回:计算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 15:13:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    'bytMCMode:1-门诊,2-住院两种模式,0-表示非医保
    If mYBInFor.bytMCMode = 1 Then AutoCalcChareFee = True: Exit Function
    If lng主页ID = 0 Then AutoCalcChareFee = True: Exit Function
    
    '自动计算病人的床位费用和护级费用
    strSQL = "ZL1_AUTOCPTPATI(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    AutoCalcChareFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Load余额信息(ByVal lng病人ID As Long, ByVal byt类型 As Byte) As Boolean
   
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人相关的余额信息
    '入参:lng病人ID=病人ID
    '     byt类型-0-所有;1-门诊;2-住院
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 15:30:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    '获取病人费用余额
    On Error GoTo errHandle
    If byt类型 = 0 Then
        strSQL = "Select sum(预交余额) As 预交余额,sum(费用余额) As 费用余额 From 病人余额 Where 病人ID= [1] And 性质=1"
    Else
        strSQL = "Select 预交余额,费用余额 From 病人余额 Where 病人ID= [1] And 性质=1 And 类型= [2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, byt类型)
    If rsTemp.RecordCount <> 0 Then
        mPatiInfor.dbl预交余额 = Format(Val(NVL(rsTemp!预交余额)), "0.00")
        mPatiInfor.dbl费用余额 = Format(Val(NVL(rsTemp!费用余额)), "0.00")
        mPatiInfor.dbl剩余合计 = Format(Val(NVL(rsTemp!预交余额)) - Val(NVL(rsTemp!费用余额)), "0.00")
        staThis.Panels(3).Text = "" & _
        "预交:" & Format(mPatiInfor.dbl预交余额, "0.00") & _
        "/费用:" & Format(mPatiInfor.dbl费用余额, "0.00") & _
        "/剩余:" & Format(mPatiInfor.dbl剩余合计, "0.00")
    End If
    
    Load余额信息 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Load应收款信息(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人应收款信息
    '入参:lng病人ID-病人ID
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 15:42:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curDue As Currency
    
    On Error GoTo errHandle
    If InStr(mstrPrivs, ";应收款管理;") = 0 Then Load应收款信息 = True: Exit Function
    curDue = GetPatientDue(lng病人ID)
    If curDue = 0 Then Load应收款信息 = True: Exit Function
    
    MsgBox mrsInfo!姓名 & ",应收款余额:" & Format(curDue, "0.00") & "元", vbInformation, gstrSysName
    staThis.Panels(2).Text = "病人应收款余额:" & Format(curDue, "0.00") & "元"
    Load应收款信息 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Load住院状态(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载住院状态
    '编制:刘兴洪
    '日期:2015-01-04 16:47:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lbl状态.Caption = GetPatiState(lng病人ID)
    lbl付款方式.Left = lbl状态.Left + lbl状态.Width + 60
    lbl付款方式.Caption = "" & mrsInfo!医疗付款方式
    pic状态.Width = lbl状态.Width + lbl付款方式.Width + 180
    If pic状态.Width >= 2500 Then
        pic状态.Width = 2500
    End If
    pic状态.Visible = True
    
    Load住院状态 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除网格信息
    '编制:刘兴洪
    '日期:2015-01-04 17:25:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBlance
        .Rows = 2
        .Clear 1
    End With
    With vsDeposit
        .Rows = 2
        .Clear 1
    End With
    With vsFeeList
        .Rows = 2
        .Clear 1
    End With
    With vsDetailList
        .Rows = 2
        .Clear 1
    End With
End Sub
Private Sub ClearBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除结算信息
    '编制:刘兴洪
    '日期:2014-12-31 11:17:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Call ClearVsBlance
    Call InitBalanceMoney  '清除变量信息
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    txtBalance(Idx_冲预交).Text = ""
    txtBalance(Idx_本次未结).Text = gstrDec
    txtBalance(Idx_本次未结).Tag = gstrDec
    mBalanceInfor.dbl未付合计 = "0.00"
    txtOwe.Text = "0.00"
    txtReceive.Text = ""
    txtCaculated.Text = "0.00"
End Sub

Private Sub ClearFeeList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除费用信息
    '编制:刘兴洪
    '日期:2015-01-04 17:29:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsFeeList
        .Redraw = False
        .Clear 1
        .Row = 1: .Col = .FixedCols
        .Redraw = True
    End With
    With vsDetailList
        .Redraw = False
        .Clear 1
        .Row = 1: .Col = .FixedCols
        .Redraw = True
    End With
End Sub

Private Sub ClearAdjustBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整结算项目列表
    '编制:刘兴洪
    '日期:2015-01-04 17:31:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intRedraw As RedrawSettings
    Dim i As Long
    With mYBInFor
        .bln个帐结算 = False
        .cur个帐余额 = 0
        .cur个帐限额 = 0
        .cur个帐透支 = 0
    End With
    Call ClearVsBlance
End Sub

Private Sub ClearVsBlance()
    Dim lngCurRow As Long, intRedraw As Integer
    Dim i As Long
    
    lngCurRow = 1
    With vsBlance
        intRedraw = .Redraw
        .Redraw = flexRDNone
        Do While Not lngCurRow > .Rows - 1
            If Val(.RowData(lngCurRow)) = 999 Then
                .TextMatrix(lngCurRow, .ColIndex("结算金额")) = "0.00"
                For i = .ColIndex("结算金额") + 1 To .Cols - 1
                    .TextMatrix(lngCurRow, .ColIndex("备注")) = ""
                Next
                lngCurRow = lngCurRow + 1
            Else
                .TextMatrix(lngCurRow, .ColIndex("结算状态")) = ""
                .TextMatrix(lngCurRow, .ColIndex("编辑状态")) = ""
                .RemoveItem lngCurRow
            End If
        Loop
        .Rows = .Rows + 1
        .Redraw = intRedraw
    End With
End Sub

Private Sub ClearAdjustDeposit()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除预交列表
    '编制:刘兴洪
    '日期:2015-01-04 17:35:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intRedraw As RedrawSettings
    With vsDeposit
        intRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = 2
        .Clear 1
        .Redraw = intRedraw
    End With
End Sub

Private Function ShowBalance(Optional ByVal blnInputPatiAfterID As Boolean, _
    Optional ByRef strMessage As String, Optional blnNotClearPati As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据设置,显示病人要结帐内容,并初始化结算金额
    '入参:blnInputPatiAfterID-病人身份确定时调用
    '出参:strMessage-返回提示信息
    '     blnNotClearPati-true:不清除病人，操作员重新选择条件
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 16:54:10
    '说明：
    '   该功能可能是上一个病人结帐完成后进行,也可能是当一个病人在结帐时另一病人中途进行
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUpload As Boolean, blnZero As Boolean
    Dim dtBeginDate As Date, dtEndDate As Date
    Dim str主页Ids As String, i As Long
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lngLast As Long, blnLastYb As Boolean
    Dim varTemp As Variant
    Dim intInsure As Integer, strInsureName As String
    Dim blnFind As Boolean
    
    On Error GoTo errHandle
        
    blnNotClearPati = False
    Call ClearFeeList   '清除费列表
    Call ClearAdjustBalance '清除结算列表
    Call ClearAdjustDeposit  '清除预交列表
    If mrsInfo.State <> 1 Then Exit Function
    
    Screen.MousePointer = 11
    If blnInputPatiAfterID Then
        Call InitPatiBalanceVariableCon
    End If
    
    blnZero = mty_ModulePara.blnZero
    If mYBInFor.intInsure <> 0 And mYBInFor.bytMCMode <> 1 Then
        If opt中途.Value And MCPAR.中途结算仅处理已上传部分 Then blnUpload = True
    End If
    
    If blnInputPatiAfterID Then mobjBalanceCon.bytKind = 2
    
    If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    If blnInputPatiAfterID Then Call LoadDefaultFilterCons
    
    If mbln连续结帐 Then
        mobjBalanceCon.strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas(False))
        Call SetPatiNums
    End If
    
    
    If mstrInputInNo <> "" Then
        varTemp = Split(mstrInputInNo, ",")
        blnFind = False
        For i = 0 To UBound(varTemp)
            If InStr("," & mobjBalanceAll.strAllTime & ",", "," & varTemp(i) & ",") > 0 Then
                blnFind = True: Exit For
            End If
        Next
        
        If blnFind = False Then
            mstrInputInNo = ""
            If mobjBalanceAll.strAllTime <> "" Then
                If MsgBox("病人:" & mrsInfo!姓名 & "的第" & mstrInputInNo & "次住院费用已经结清,但还存在其他住院费用未结，是否结其他费用？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                
            Else
                MsgBox "病人:" & mrsInfo!姓名 & "的第" & mstrInputInNo & "次住院费用已经结清，将重新读取该病人的未结费用!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            txtPatient.Text = "-" & mrsInfo!病人ID
            '显示最后一次未结费用的信息
            Call LoadPatientInfo(IDKind.GetCurCard, False, , Val(Split(mobjBalanceAll.strAllTime, ",")(0)))
            
        End If
    End If
    
    
    
    If mstrInputInNo <> "" Then
        cboPatiNums.Text = ""
        blnFind = False
        
        For i = 1 To cboPatiNums.ListCount
            If InStr("," & mstrInputInNo & ",", "," & Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) & ",") > 0 Then
                cboPatiNums.Nodes.Item(i).Checked = True
                cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
                blnFind = True
            Else
                cboPatiNums.Nodes.Item(i).Checked = False
            End If
        Next i
        
        If cboPatiNums.Text <> "" Then cboPatiNums.Text = Mid(cboPatiNums.Text, 2)
        
        If blnFind = False Then
            '全选未结部分
            MsgBox "病人:" & mrsInfo!姓名 & "的第" & mstrInputInNo & "次住院费用已经结清，将重新读取该病人的未结费用!", vbInformation + vbDefaultButton1, gstrSysName
            txtPatient.Text = "-" & Val(mrsInfo!病人ID)
            mstrInputInNo = "": mblnReadByZYNo = False
            Call LoadPatientInfo(IDKind.GetCurCard, False, , Split(mobjBalanceAll.strAllTime, ",")(0))
            mYBInFor.intInsure = Val(NVL(mrsInfo!险类))
            
            For i = 1 To cboPatiNums.ListCount
                If mYBInFor.intInsure > 0 Then
                    If Split(mobjBalanceAll.strAllTime, ",")(0) = Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) Then
                        cboPatiNums.Nodes.Item(i).Checked = True
                        cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
                    Else
                        cboPatiNums.Nodes.Item(i).Checked = False
                    End If
                Else
                    cboPatiNums.Nodes.Item(i).Checked = True
                    cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
                End If
            Next i
            If cboPatiNums.Text <> "" Then cboPatiNums.Text = Mid(cboPatiNums.Text, 2)
        Else
            mYBInFor.intInsure = Val(NVL(mrsInfo!险类))
        End If
        mobjBalanceCon.strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas())

'        If mEditType = g_Ed_住院结帐 Then
'            strSQL = "Select 险类 From 病案主页 Where 病人ID = [2] And 病人性质 <> 1 And 主页ID In (Select Column_Value From Table(f_str2list([1]))) Order By 主页ID Desc "
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBalanceCon.strTime, Val(nvl(mrsInfo!病人ID)))
'            If Not rsTemp.EOF Then
'                Do While Not rsTemp.EOF
'                    If Val(nvl(rsTemp!险类)) = 0 Then
'                        mYBInFor.intInsure = 0
'                        Exit Do
'                    Else
'                        mYBInFor.intInsure = Val(nvl(rsTemp!险类))
'                    End If
'                    rsTemp.MoveNext
'                Loop
'            End If
'        End If
                    
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    End If
    
    If blnInputPatiAfterID And mrsFeeList.RecordCount = 0 And mstrInputInNo <> "" Then
        mstrInputInNo = ""
        mobjBalanceCon.strTime = mstrInputInNo
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    End If
    
    If mrsFeeList Is Nothing Then Screen.MousePointer = 0: Exit Function
    If blnInputPatiAfterID And mrsFeeList.RecordCount = 0 And mEditType = g_Ed_门诊结帐 Then
        mobjBalanceCon.bytKind = 1 '缺省只取普通费用，如果没有再检查只有体检费用这种情况
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
        If mrsFeeList Is Nothing Then Screen.MousePointer = 0: Exit Function
        
        If mrsFeeList.RecordCount > 0 Then
            If MsgBox("该病人普通费用已结清,要对体检费用进行结帐吗?", vbInformation + vbYesNo, Me.Caption) = vbNo Then
                Set mrsFeeList = Nothing
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    If mrsFeeList.RecordCount = 0 Then
        Set mrsFeeList = Nothing
        If blnInputPatiAfterID Then strMessage = "该病人没有需要结帐的费用！"
        Screen.MousePointer = 0: Exit Function
    End If
    
    If blnInputPatiAfterID Then
         '加载缺省的过滤条件
        If mobjBalanceAll.strAllOwnerFeeType <> "" Then
            picOwnerFee.Visible = True
            blnNotClearPati = True
            '先缺省结自费项目
            mobjBalanceCon.strChargeType = mobjBalanceAll.strAllOwnerFeeType
            mobjBalanceCon.blnCurBalanceOwnerFee = True
            If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
            If mrsFeeList Is Nothing Then Screen.MousePointer = 0: Exit Function
            If mrsFeeList.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        End If
        
        '检查费用是否审核
        If CheckPatiIsVerfy(strMessage) = False Then Screen.MousePointer = 0: Exit Function
        '检查输血费
        If CheckInputBlood = False Then Screen.MousePointer = 0: Exit Function
        
        If mobjBalanceCon.blnCurBalanceOwnerFee = False _
            And (mYBInFor.intInsure <> 0 And MCPAR.结帐设置后调用接口) Or MCPAR.门诊结算_结帐设置 Then
            '----------------------------------------------------------------
            '获取住院日期范围和缺省的住院时间
            If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) = False Then Exit Function
            txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
            txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
            txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
            Call zlChangeDefaultTime
            mblnConsChange = True
            Call ClearListData
            Screen.MousePointer = 0
            mblnConsChange = False
            ShowBalance = True
            mblnInterUse = True
            Call ShowBalance(False)
            mblnInterUse = False
            mstrInputInNo = ""
            Exit Function
        End If
        mblnInterUse = True
        Call ShowBalance(False)
        mblnInterUse = False
        Call ResetTime
        mstrInputInNo = ""
        ShowBalance = True
        Exit Function
    End If
    
    '78317:医保病人默认只读取最后一次住院的数据
    If mEditType <> g_Ed_门诊结帐 And mobjBalanceCon.blnCurBalanceOwnerFee = False _
        And mYBInFor.intInsure <> 0 And (blnInputPatiAfterID Or mblnInterUse) And mstrInputInNo = "" Then
        lngLast = Val(Split(mobjBalanceAll.strAllTime & ",", ",")(0))
        If lngLast <> 0 And mEditType <> g_Ed_门诊结帐 Then
            Call CheckPatiFromZyNumIsYB(Val(NVL(mrsInfo!病人ID)), lngLast, intInsure, strInsureName)
            If intInsure <> 0 Then
                If mYBInFor.intInsure <> intInsure Then Call InitInsurePara(Val(NVL(mrsInfo!病人ID)), intInsure)
                mYBInFor.intInsure = intInsure
                If NVL(mrsInfo!病人类型) = "" Then
                    txtPatient.ForeColor = vbRed
                Else
                    txtPatient.ForeColor = zlDatabase.GetPatiColor(NVL(mrsInfo!病人类型))
                End If
                staThis.Panels(4).Text = strInsureName
                staThis.Panels(4).Visible = True
            Else
                    mYBInFor.intInsure = 0
                    mYBInFor.strBalance = ""
                    txtPatient.ForeColor = Me.ForeColor
                    staThis.Panels(4).Text = ""
                    staThis.Panels(4).Visible = False
            End If
        End If

        '最后一次不是医保入院,当成普通病人处理
        mobjBalanceCon.strTime = lngLast
          '最后一次不是医保入院,当成普通病人处理
        mobjBalanceCon.strTime = lngLast
        For i = 1 To cboPatiNums.ListCount
            blnFind = InStr("," & mobjBalanceCon.strTime & ",", "," & Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) & ",") > 0
            If Not blnFind And mYBInFor.intInsure <> 0 And MCPAR.允许结多次住院费用 Then blnFind = True
            cboPatiNums.Nodes.Item(i).Checked = blnFind
        Next i
        mobjBalanceCon.strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas())
        Call cboPatiNums.Refresh
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    End If
    
    '加载费用列表信息
    If LoadFeeList = False Then Screen.MousePointer = 0: Exit Function
    
    
    '加载交款信息
    str主页Ids = IIf(mty_ModulePara.bln仅用指定预交款 And mbln门诊转住院 = False, _
        IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime), "")
    
    
    If LoadDepositList(Val(NVL(mrsInfo!病人ID)), str主页Ids) = False Then Screen.MousePointer = 0: Exit Function
                                
    '----------------------------------------------------------------
    '获取住院日期范围和缺省的住院时间
    If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) = False Then Exit Function
    txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
    txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
    txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
    Call zlChangeDefaultTime
    
    '----------------------------------------------------------------
    '医保预结算(普通病人也调用，内部有处理，直接返回true
    If InsureBudgeting(blnUpload) = False Then
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If

    '重新分配预交款(bytOperationType-操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按结帐金额来冲预交(按时间先后来分摊）;3-全冲)
    If mobjBalanceCon.blnCurBalanceOwnerFee Then
        Call RecalcDepositMoney(IIf(mty_ModulePara.bln自费缺省使用预交, 2, 0))
    Else
        Call RecalcDepositMoney(1)
    End If

    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor(True)
    Call SetDefaultPayType '设置缺省的支付方式
    mblnNotChange = True
    txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
    mblnNotChange = False
    txtDate.Text = Format(zlDatabase.Currentdate, txtDate.Format)
    
    Screen.MousePointer = 0
    mblnConsChange = False
    ShowBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ResetTime()
    Dim dtDate As Date
    With mrsFeeList
        If .RecordCount <> 0 Then
            .MoveFirst
            If mty_ModulePara.int费用时间 = 0 Then
                dtDate = mrsFeeList!登记时间
            Else
                dtDate = mrsFeeList!时间
            End If
             mobjBalanceAll.MinDate = dtDate: mobjBalanceAll.MaxDate = dtDate
        End If
        
        Do While Not .EOF
            '比较取最大最小值
            If mty_ModulePara.int费用时间 = 0 Then
                dtDate = mrsFeeList!登记时间
            Else
                dtDate = mrsFeeList!时间
            End If
            If dtDate < mobjBalanceAll.MinDate Then mobjBalanceAll.MinDate = dtDate
            If dtDate > mobjBalanceAll.MaxDate Then mobjBalanceAll.MaxDate = dtDate
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        '显示结帐时间
        mblnNotChange = True
        Call RecalcFeeTotalDate
        If Format(mobjBalanceAll.MinDate, txtBegin.Format) < Format(txtBegin.Text, txtBegin.Format) Then txtBegin.Text = Format(mobjBalanceAll.MinDate, txtBegin.Format)
        If Format(mobjBalanceAll.MaxDate, txtEnd.Format) > Format(txtEnd.Text, txtEnd.Format) Then txtEnd.Text = Format(mobjBalanceAll.MaxDate, txtEnd.Format)
        mblnNotChange = False
    End With
End Sub

Private Sub LoadIntendBalance(Optional ByVal dblSum As Double = 0, Optional objCard As Card)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str结算方式 As String, intCount As Integer
    Dim dblBalanceSum As Double
    Dim i As Long, j As Long
    Dim lngRow As Long, blnThirdSingle As Boolean, strErrMsg As String
    Dim dblAdd As Double, blnAdd As Boolean
    Dim dblTotal As Double, blnDo As Boolean
    Dim dblAlr As Double, strArray() As String, intArray As Integer
    Dim dblMoney As Double
    
    On Error GoTo errHandle

    mstrBalanceLimit = ""
    If mstrForceNote <> "" Then
        mstrForceNote = Mid(mstrForceNote, 1, InStr(mstrForceNote, "强制退现") + 4)
    End If
    
    For i = 1 To vsBlance.Rows - 1
        If Val(vsBlance.RowData(i)) = 999 Then '现金
            dblMoney = Val(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算金额")))
            mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - dblMoney, 5)
            mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + dblMoney, 5)
            vsBlance.TextMatrix(i, vsBlance.ColIndex("结算金额")) = "0.00"
            Exit For
        End If
    Next
    If objCard Is Nothing Then
        With vsBlance
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("类型"))) <> 9 Then
                    If .TextMatrix(i, .ColIndex("组合信息")) <> "" Then
                        dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("结算金额")))
                        intCount = intCount + 1
                    Else
                        dblAlr = dblAlr + Val(.TextMatrix(i, .ColIndex("结算金额")))
                    End If
                End If
            Next i
            For i = 1 To intCount
                For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("组合信息")) <> "" Then
                        .RemoveItem j
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        mBalanceInfor.dbl已付合计 = dblAlr
        mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl当前结帐 - dblAlr - mBalanceInfor.dbl冲预交合计, 5)
    Else
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("组合信息")) <> "" And Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = objCard.接口序号 Then
                    dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("结算金额")))
                    intCount = intCount + 1
                Else
                    dblAlr = dblAlr + Val(.TextMatrix(i, .ColIndex("结算金额")))
                End If
            Next i
        End With
    End If
    
    If dblSum = 0 Then
        dblBalanceSum = RoundEx(mBalanceInfor.dbl冲预交合计 + dblAlr - mBalanceInfor.dbl当前结帐, 2)
    Else
        dblBalanceSum = RoundEx(mBalanceInfor.dbl冲预交合计 + dblAlr - mBalanceInfor.dbl当前结帐, 2)
        If dblSum <= dblBalanceSum Then
            dblBalanceSum = dblSum
        Else
            If MsgBox("输入的退款金额超过了允许的退款金额(" & dblBalanceSum & ")" & vbCrLf & "是否以允许的退款金额进行结算?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
        End If
    End If
    
    If dblBalanceSum <= 0 Then Exit Sub
    If mrsDeposit Is Nothing Then Exit Sub
    If mrsDeposit.RecordCount = 0 Then
        If objCard Is Nothing Then
            Exit Sub
        Else
            MsgBox objCard.名称 & "不支持接口退款!", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            If dblBalanceSum > 0 Then
                mrsDeposit.Filter = "预交ID=" & Val(.TextMatrix(i, .ColIndex("预交ID")))
                
                If Val(NVL(mrsDeposit!卡类别ID)) = 0 Then
                    dblAdd = Val(.TextMatrix(i, .ColIndex("冲预交")))
                    If dblBalanceSum > dblAdd Then
                        dblBalanceSum = RoundEx(dblBalanceSum - dblAdd, 2)
                    Else
                        dblBalanceSum = 0
                    End If
                    GoTo GoNext
                End If

                If mrsDeposit.RecordCount <> 0 Then
                    If Val(NVL(mrsDeposit!卡类别ID)) <> 0 Then
                        If Not objCard Is Nothing Then
                            If objCard.接口序号 <> Val(NVL(mrsDeposit!卡类别ID)) Then GoTo GoNext
                        End If
                        If Val(NVL(mrsDeposit!结算性质)) = 8 And Val(.TextMatrix(i, .ColIndex("冲预交"))) <> 0 Then
                            
                            strSQL = "Select 是否退现,是否全退,卡号密文,名称,是否缺省退现 From 医疗卡类别 Where ID= [1]"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(mrsDeposit!卡类别ID)))
                            If Val(NVL(rsTmp!是否退现)) = 1 And mty_ModulePara.bln三方卡结帐退款控制 And Val(NVL(rsTmp!是否缺省退现)) = 1 Then
                                dblAdd = Val(.TextMatrix(i, .ColIndex("冲预交")))
                                If dblBalanceSum > dblAdd Then
                                    dblBalanceSum = RoundEx(dblBalanceSum - dblAdd, 2)
                                Else
                                    dblBalanceSum = 0
                                End If
                                GoTo GoNext
                            End If
                            If mstrCardPara <> "" Then
                                strArray = Split(mstrCardPara, "|")
                                blnDo = True
                                For intArray = 0 To UBound(strArray)
                                    If Val(Split(strArray(intArray), ",")(0)) = Val(NVL(mrsDeposit!卡类别ID)) Then
                                        blnDo = False
                                        blnThirdSingle = Val(Split(strArray(intArray), ",")(1)) = 1
                                        Exit For
                                    End If
                                Next intArray
                            Else
                                blnDo = True
                            End If
                            
                            If blnDo And Val(NVL(mrsDeposit!转帐及代扣)) = 0 Then
                                blnThirdSingle = gobjSquare.objSquareCard.ZlGetParaConfig(Me, Val(NVL(mrsDeposit!卡类别ID)), False, 2, strErrMsg)
                                mstrCardPara = mstrCardPara & IIf(mstrCardPara = "", "", "|") & Val(NVL(mrsDeposit!卡类别ID)) & "," & IIf(blnThirdSingle, 1, 0)
                            End If
                            
                            dblAdd = Val(.TextMatrix(i, .ColIndex("冲预交")))
                            blnAdd = True
                            With vsBlance
                                If blnThirdSingle And Val(NVL(mrsDeposit!转帐及代扣)) = 0 Then
                                    If Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("编辑状态"))) <> 0 Then GoTo GoNext
                                    If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Then
                                        .Rows = .Rows + 1
                                    End If
                                    
                                    '新增结算
                                    .RowData(.Rows - 1) = ""
                                    .TextMatrix(.Rows - 1, .ColIndex("类型")) = 3
                                    .TextMatrix(.Rows - 1, .ColIndex("卡类别ID")) = Val(NVL(mrsDeposit!卡类别ID))
                                    .TextMatrix(.Rows - 1, .ColIndex("消费卡ID")) = 0
                                    .TextMatrix(.Rows - 1, .ColIndex("结算性质")) = Val(NVL(mrsDeposit!结算性质))
                                    .TextMatrix(.Rows - 1, .ColIndex("编辑状态")) = 2   '0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等
                                    .TextMatrix(.Rows - 1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                                    .TextMatrix(.Rows - 1, .ColIndex("是否退现")) = Val(NVL(rsTmp!是否退现))
                                    .TextMatrix(.Rows - 1, .ColIndex("是否全退")) = Val(NVL(rsTmp!是否全退))
                                    .TextMatrix(.Rows - 1, .ColIndex("校对标志")) = 0
                                    .TextMatrix(.Rows - 1, .ColIndex("是否转账")) = Val(NVL(mrsDeposit!转帐及代扣))
                                    .TextMatrix(.Rows - 1, .ColIndex("是否密文")) = Val(NVL(rsTmp!卡号密文))
                                    .TextMatrix(.Rows - 1, .ColIndex("卡类别名称")) = Trim(NVL(rsTmp!名称))
                                    .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = NVL(mrsDeposit!结算方式)
                                    If dblBalanceSum > dblAdd Then
                                        .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = Format(-1 * dblAdd, "0.00")
                                    Else
                                        .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = Format(-1 * dblBalanceSum, "0.00")
                                    End If
                                    .TextMatrix(.Rows - 1, .ColIndex("结算号码")) = ""
                                    .TextMatrix(.Rows - 1, .ColIndex("备注")) = ""
                                    .TextMatrix(.Rows - 1, .ColIndex("交易流水号")) = NVL(mrsDeposit!交易流水号)
                                    .TextMatrix(.Rows - 1, .ColIndex("交易说明")) = NVL(mrsDeposit!交易说明)
                                    .TextMatrix(.Rows - 1, .ColIndex("卡号")) = IIf(Val(NVL(rsTmp!卡号密文)) = 1, String(Len(NVL(mrsDeposit!卡号)), "*"), NVL(mrsDeposit!卡号))
                                    .TextMatrix(.Rows - 1, .ColIndex("组合信息")) = NVL(mrsDeposit!预交ID)
                                    .Cell(flexcpData, .Rows - 1, .ColIndex("组合信息")) = 1
                                    .Cell(flexcpData, .Rows - 1, .ColIndex("卡号")) = NVL(mrsDeposit!卡号)
                                Else
                                    For j = 1 To .Rows - 1
                                        If Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = Val(NVL(mrsDeposit!卡类别ID)) And Val(.TextMatrix(j, .ColIndex("结算状态"))) = 1 Then
                                            GoTo GoNext
                                        End If
                                    Next j
                                    lngRow = 0
                                    For j = 1 To .Rows - 1
                                        If .TextMatrix(j, .ColIndex("结算方式")) = NVL(mrsDeposit!结算方式) Then lngRow = j
                                    Next j
                                    If lngRow = 0 Then
                                        If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Then
                                            .Rows = .Rows + 1
                                        End If
                                        '新增结算
                                        .RowData(.Rows - 1) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("类型")) = 3
                                        .TextMatrix(.Rows - 1, .ColIndex("卡类别ID")) = Val(NVL(mrsDeposit!卡类别ID))
                                        .TextMatrix(.Rows - 1, .ColIndex("消费卡ID")) = 0
                                        .TextMatrix(.Rows - 1, .ColIndex("结算性质")) = Val(NVL(mrsDeposit!结算性质))
                                        .TextMatrix(.Rows - 1, .ColIndex("编辑状态")) = 2   ' '0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等
                                        .TextMatrix(.Rows - 1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                                        .TextMatrix(.Rows - 1, .ColIndex("是否退现")) = Val(NVL(rsTmp!是否退现))
                                        .TextMatrix(.Rows - 1, .ColIndex("是否全退")) = Val(NVL(rsTmp!是否全退))
                                        .TextMatrix(.Rows - 1, .ColIndex("校对标志")) = 0
                                        .TextMatrix(.Rows - 1, .ColIndex("是否转账")) = Val(NVL(mrsDeposit!转帐及代扣))
                                        .TextMatrix(.Rows - 1, .ColIndex("是否密文")) = Val(NVL(rsTmp!卡号密文))
                                        .TextMatrix(.Rows - 1, .ColIndex("卡类别名称")) = Trim(NVL(rsTmp!名称))
                                        .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = NVL(mrsDeposit!结算方式)
                                        If dblBalanceSum > dblAdd Then
                                            .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = Format(-1 * dblAdd, "0.00")
                                        Else
                                            .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = Format(-1 * dblBalanceSum, "0.00")
                                        End If
                                        .TextMatrix(.Rows - 1, .ColIndex("结算号码")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("备注")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("交易流水号")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("交易说明")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("卡号")) = IIf(Val(NVL(rsTmp!卡号密文)) = 1, String(Len(NVL(mrsDeposit!卡号)), "*"), NVL(mrsDeposit!卡号))
                                        If dblBalanceSum > dblAdd Then
                                            .TextMatrix(.Rows - 1, .ColIndex("组合信息")) = NVL(mrsDeposit!卡号) & "," & TruncStringEx(NVL(mrsDeposit!交易流水号)) & "," & TruncStringEx(NVL(mrsDeposit!交易说明)) & "," & RoundEx(-1 * dblAdd, 2) & "," & NVL(mrsDeposit!预交ID)
                                        Else
                                            .TextMatrix(.Rows - 1, .ColIndex("组合信息")) = NVL(mrsDeposit!卡号) & "," & TruncStringEx(NVL(mrsDeposit!交易流水号)) & "," & TruncStringEx(NVL(mrsDeposit!交易说明)) & "," & RoundEx(-1 * dblBalanceSum, 2) & "," & NVL(mrsDeposit!预交ID)
                                        End If
                                        .Cell(flexcpData, .Rows - 1, .ColIndex("卡号")) = NVL(mrsDeposit!卡号)
                                    Else
                                        '更新结算
                                        If dblBalanceSum > dblAdd Then
                                            .TextMatrix(lngRow, .ColIndex("结算金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("结算金额"))) - dblAdd, "0.00")
                                            .TextMatrix(lngRow, .ColIndex("组合信息")) = .TextMatrix(lngRow, .ColIndex("组合信息")) & "|" & NVL(mrsDeposit!卡号) & "," & TruncStringEx(NVL(mrsDeposit!交易流水号)) & "," & TruncStringEx(NVL(mrsDeposit!交易说明)) & "," & RoundEx(-1 * dblAdd, 2) & "," & NVL(mrsDeposit!预交ID)
                                        Else
                                            .TextMatrix(lngRow, .ColIndex("结算金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("结算金额"))) - dblBalanceSum, "0.00")
                                            .TextMatrix(lngRow, .ColIndex("组合信息")) = .TextMatrix(lngRow, .ColIndex("组合信息")) & "|" & NVL(mrsDeposit!卡号) & "," & TruncStringEx(NVL(mrsDeposit!交易流水号)) & "," & TruncStringEx(NVL(mrsDeposit!交易说明)) & "," & RoundEx(-1 * dblBalanceSum, 2) & "," & NVL(mrsDeposit!预交ID)
                                        End If
                                    End If
                                End If
                            End With
                            If dblBalanceSum > dblAdd Then
                                mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - RoundEx(dblAdd, 2), 5)
                                mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + RoundEx(dblAdd, 2), 5)
                                dblBalanceSum = RoundEx(dblBalanceSum - dblAdd, 5)
                            Else
                                mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - RoundEx(dblBalanceSum, 2), 5)
                                mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + RoundEx(dblBalanceSum, 2), 5)
                                dblBalanceSum = 0
                            End If
                        End If
                    End If
                End If
            End If
GoNext:
        Next i

    End With
    
    mrsDeposit.Filter = ""
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("结算性质"))) = 8 And .TextMatrix(i, .ColIndex("组合信息")) <> "" And Val(.Cell(flexcpData, i, .ColIndex("组合信息"))) = 0 Then
                mstrBalanceLimit = mstrBalanceLimit & "|" & .TextMatrix(i, .ColIndex("卡类别ID")) & "," & .TextMatrix(i, .ColIndex("结算金额"))
            End If
        Next i
        If mstrBalanceLimit <> "" Then mstrBalanceLimit = Mid(mstrBalanceLimit, 2)
    End With
    If blnAdd = False And Not objCard Is Nothing Then
        MsgBox "没有可以退款的金额,不能使用" & objCard.名称 & "退款", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not objCard Is Nothing Then
        IDKindPaymentsType.IDKind = 1
    End If
    If vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算方式")) <> "" Then
        vsBlance.Rows = vsBlance.Rows + 1
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InsureBudgeting(ByVal blnOnlyUpload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结算
    '入参: blnOnlyUpload-是否只处理已上传部份
    '返回:预算成功(含普通病人未行医保虚拟结算),返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-06 16:48:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln发生时间 As Boolean, intInsure As Integer, lng病人ID As Long, str医保号 As String
    Dim strBalance As String, varData As Variant, varTemp As Variant
    Dim strNotBalance As String '不存在的结算方式
    Dim lngRow As Long, blnOk As Boolean
    Dim cur个人帐户 As Currency, cur统筹支付 As Currency
    Dim curMoney As Currency
    Dim rsDetail As ADODB.Recordset
    Dim i As Long, byt状态 As Byte, bytEditSta As Byte
    On Error GoTo errHandle
    
    Call ClearVsBlance
    
    txtBalance(Idx_本次结帐).Enabled = True
    txtBalance(Idx_本次结帐).Locked = InStr(mstrPrivs, ";结帐设置;") = 0

    Call HideYBMoneyInfo '隐藏医保支付信息
    
    intInsure = mYBInFor.intInsure
    lng病人ID = Val(NVL(mrsInfo!病人ID))
    str医保号 = "" & mrsInfo!医保号
    
    cmdOK.Enabled = True
    If mobjBalanceCon.blnCurBalanceOwnerFee Or intInsure = 0 Then
        '当前正在结自费的,则不处理医保
        Call SetOperationCtrl(0)
        InsureBudgeting = True: Exit Function     '先结自费费用，不用先预结算
    End If
     
    bln发生时间 = mty_ModulePara.int费用时间 = 1 '0-按登记时间,1-按发生时间
    '医保预结算
    If mEditType = g_Ed_门诊结帐 Then
        With mobjBalanceCon
            Set rsDetail = GetMzBalance_Insure(intInsure, lng病人ID, _
                .dtBeginDate, .dtEndDate, blnOnlyUpload, mblnDateMoved, mYBInFor.bytMCMode = 1, .bytKind, .strItem, .strDeptIDs, .strClass, .strChargeType, bln发生时间)
        End With
    Else
        With mobjBalanceCon
            Set rsDetail = GetZYBalance_Insure(intInsure, lng病人ID, _
                .strTime, .dtBeginDate, .dtEndDate, blnOnlyUpload, mblnDateMoved, .strBaby, .strItem, .strDeptIDs, .strClass, .strChargeType, bln发生时间)
        End With
    End If
    
    mYBInFor.strBalance = ""
    '医保接口:返回各种报销金额
    If mYBInFor.bytMCMode = 1 Then
        If MCPAR.门诊预结算 Then
            If rsDetail.RecordCount = 0 Then
                Screen.MousePointer = 0:
                MsgBox "读取医保预结算数据失败!", vbInformation, gstrSysName
                Exit Function
            End If
        
            'strAdvance:
            '1.收费时传入空
            '2.退费时，如果发生重新收费时，传入1,表示重新收费调用
            '3. 医保二次结算时，传入2
            '4. 医保二次结算发生部分退费时，重新二次结算，传入3
            '5．门诊结帐传入4
            Call SetCmdStatus(False)
            If Not gclsInsure.ClinicPreSwap(rsDetail, strBalance, intInsure, "4") Then
                Call SetCmdStatus(True)
                Screen.MousePointer = 0
                MsgBox "门诊医保预结算失败!", vbInformation, gstrSysName
                Exit Function
            End If
            Call SetCmdStatus(True)
        End If
    Else
        Call SetCmdStatus(False)
        strBalance = gclsInsure.WipeoffMoney(rsDetail, lng病人ID, str医保号, "1", intInsure, "|" & IIf(opt中途.Value, 0, 1))
        Call SetCmdStatus(True)
    End If
    
    '显示个帐余额
    mYBInFor.cur个帐余额 = gclsInsure.SelfBalance(lng病人ID, str医保号, IIf(mYBInFor.bytMCMode = 1, 10, 40), _
        mYBInFor.cur个帐透支, intInsure)
    
    
    '结算方式;金额;是否允许修改|...
    mYBInFor.strBalance = strBalance
    varData = Split(mYBInFor.strBalance, "|")
    
    '显示各类统筹报销总额
    cur统筹支付 = 0: cur个人帐户 = 0
    strNotBalance = ""
    blnOk = True
    
    With vsBlance
        .Redraw = flexRDNone
        For i = 0 To UBound(varData)
            '结算方式;金额;是否允许修改|...
            varTemp = Split(varData(i) & ";;;;", ";")
            mrs结算方式.Filter = "名称 ='" & varTemp(0) & "'"
            curMoney = Val(varTemp(1))
            byt状态 = 0: bytEditSta = IIf(Val(varTemp(2)) = 1, "1", "4")
            
            If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "" Then
                lngRow = .Rows - 1
            Else
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
                        
            If mrs结算方式.EOF = False Then
                '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
                Select Case Val(NVL(mrs结算方式!性质))
                Case 3  '3-医保个人帐户
                    cur个人帐户 = cur个人帐户 + curMoney
                    If mYBInFor.cur个帐余额 - curMoney < -1 * mYBInFor.cur个帐透支 Then
                        curMoney = 0
                        MsgBox "个人帐户余额不足或未更新,不允许医保结算!", vbInformation, Me.Caption
                        blnOk = False
                        Exit Function
                    End If
                    byt状态 = 2
                Case 4  '4-医保各类统筹
                    cur统筹支付 = cur统筹支付 + curMoney
                    byt状态 = 2
                Case Else  '非医保类,需要提醒
                    strNotBalance = strNotBalance & "," & varTemp(0)
                End Select
                .TextMatrix(lngRow, .ColIndex("结算性质")) = Val(NVL(mrs结算方式!性质))
            Else
                strNotBalance = strNotBalance & "," & varTemp(0)
            End If
            
            .TextMatrix(lngRow, .ColIndex("类型")) = byt状态
            .TextMatrix(lngRow, .ColIndex("编辑状态")) = bytEditSta   '0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等
            If bytEditSta <> 0 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
            End If
            .TextMatrix(lngRow, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
            
            .TextMatrix(lngRow, .ColIndex("结算方式")) = varTemp(0)
            .TextMatrix(lngRow, .ColIndex("结算金额")) = Format(curMoney, gstrDec)
            .Cell(flexcpData, lngRow, .ColIndex("结算金额")) = Val(varTemp(1))
        Next
        
        If strNotBalance <> "" Then
            .Rows = 2: .Clear 1
            .Redraw = flexRDBuffered
            Screen.MousePointer = 0
            MsgBox "结帐场合的保险结算方式未设置完全,该病人还有以下保险结算方式可以报销：" & _
            vbCrLf & strNotBalance & vbCrLf & vbCrLf & "您可以到费用基础项目\结算方式管理中去设置这些结算方式！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Then .Rows = .Rows + 1
        .Redraw = flexRDBuffered
    End With
    mYBInFor.cur个帐支付 = cur个人帐户
    mYBInFor.cur统筹支付 = cur统筹支付
    
    mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl当前结帐 - (cur个人帐户 + cur统筹支付), 6)
    mBalanceInfor.dbl医保支付合计 = RoundEx(cur个人帐户 + cur统筹支付, 3)
    mBalanceInfor.dbl预结算总额 = mBalanceInfor.dbl医保支付合计
    staThis.Panels(5).Text = Format(mYBInFor.cur个帐余额, "0.00")
    staThis.Panels(5).Visible = True
    txtBalance(Idx_本次结帐).Enabled = False
    
    'bytFun-0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
    Call SetOperationCtrl(1)
    '显示医保虚算信息:bytFun-0-医保预算信息显示
    Call ShowLedDisplayBank(0)
    Call LoadCurOwnerPayInfor  '加载支付合计
    InsureBudgeting = True
    Exit Function
errHandle:
    vsBlance.Redraw = flexRDBuffered
     Screen.MousePointer = 0
    Call SetCmdStatus(True)
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetCmdStatus(blnStatus As Boolean)
    cmdMore.Enabled = blnStatus And InStr(mstrPrivs, ";结帐设置;") > 0
    cmdCancel.Enabled = blnStatus
    cmdOK.Enabled = blnStatus
    cmdNext.Enabled = blnStatus
    cmdYBBalance.Enabled = blnStatus
End Sub

Private Sub ShowLedDisplayBank(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:Led信息显示
    '入参:bytFun-0-医保预算信息显示;1-显示费用信息
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-07 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtTmpDate As Date, dblTemp As Double, strDepositName As String
    If Not gblnLED Then Exit Sub
    
    On Error GoTo errHandle
    
    If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_重新结帐 Or mEditType = g_Ed_住院结帐) Then Exit Sub
    
    strDepositName = "住院预交款"
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then strDepositName = "门诊预交"
    Select Case bytFun
    Case 0 '医保预算信息显示
        zl9LedVoice.DisplayBank "医保结算:", _
            "帐户余额" & Format(mYBInFor.cur个帐余额, "0.00"), _
            "帐户支付" & Format(mYBInFor.cur个帐支付, "0.00"), _
            "统筹支付" & Format(mYBInFor.cur统筹支付, "0.00")
    Case 1 '显示费用信息
        zl9LedVoice.DisplayBank _
            "总费用" & Format(mBalanceInfor.dbl当前结帐, "0.00"), _
             strDepositName & Format(mPatiInfor.dbl预交余额, "0.00"), _
            "冲预交" & Format(mBalanceInfor.dbl冲预交合计, "0.00"), _
            IIf(mBalanceInfor.dbl本次未结 < 0, "找补", "应缴") & Format(Abs(mBalanceInfor.dbl本次未结), "0.00")
    End Select
    
    '延迟时间
    dtTmpDate = Time
    Do While Time < DateAdd("s", 4, dtTmpDate)
    Loop
    
    Exit Sub
errHandle:
    
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetOperationCtrl(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置操作控件的相关属性
    '入参:bytFun-0-结算前;1-医保虚拟结算后;2-已保存了结帐单
    '            3-未设置任条条件
    '编制:刘兴洪
    '日期:2015-01-07 11:21:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    Dim lngColor As Long, objCard As Card
    Dim blnTemp As Boolean
    
    If mEditType = g_Ed_单据查看 Then cmdCancel.Visible = False: Exit Sub
    
    If mobjBalanceCon.blnCurBalanceOwnerFee Then
        cmdYBBalance.Visible = False
    Else
        cmdYBBalance.Visible = bytFun = 1 Or mYBInFor.intInsure <> 0 And mBalanceInfor.blnSaveBill = False
    End If
    
    If mYBInFor.bytMCMode = 1 Then
        cmdYBBalance.Enabled = True ' MCPAR.门诊预结算
    Else
        cmdYBBalance.Enabled = mYBInFor.strBalance <> ""
    End If
    
    If cmdYBBalance.Visible Then
        cmdYBBalance.Left = cmdOK.Left
    End If
    
    cmdOK.Visible = bytFun <> 1 And cmdYBBalance.Visible = False
    cmdCancel.Visible = bytFun <> 2 And chkCancel.Value = 0
    
    Call SetNextBalanceCmdVisible   '设置连续结帐按钮
  
    cmdYB.Enabled = True
    txtBalance(Idx_本次未结).Enabled = False
    IDKindPaymentsType.Locked = mPatiInfor.bln连续结帐
    Select Case bytFun
    Case 0  '结算前
    
        txtBalance(Idx_结帐说明).Enabled = True
        
        mblnNotChange = True
        txtBalance(Idx_冲预交).Enabled = mPatiInfor.dbl实际余额 <> 0
        mblnNotChange = False
        
        txtBalance(Idx_结帐说明).BackColor = &H80000005
        txtBalance(Idx_本次结帐).Enabled = True
        txtBalance(Idx_本次结帐).Locked = InStr(mstrPrivs, ";结帐设置;") = 0
        txtBalance(Idx_本次结帐).BackColor = &H80000005
        If Not mBalanceInfor.bln预交刷卡 Then
            txtBalance(Idx_冲预交).BackColor = IIf(txtBalance(Idx_冲预交).Enabled, &H80000005, &H8000000F)
        End If
        txtBalance(Idx_本次未结).Enabled = False
        txtBalance(Idx_本次未结).BackColor = &H8000000F
        txtReceive.Locked = False

        IDKindPaymentsType.Enabled = True
        
        cmdMore.Visible = chkCancel.Value = 0
        cmdMore.Enabled = True And InStr(mstrPrivs, ";结帐设置;") > 0
        cboPatiNums.Enabled = True And InStr(mstrPrivs, ";结帐设置;") > 0
        txtPatient.Locked = False
        IDKind.Enabled = True
        cmdDelBalance.Visible = False
        picPati.Enabled = True
        cmdOK.Left = IIf(cmdCancel.Visible, cmdCancel.Left, picBalanceBack.ScaleWidth) - cmdOK.Width - 60
    Case 1, 3 '医保虚拟结算后 或未设置过滤条件时
    
        If bytFun = 3 Then txtBalance(Idx_冲预交).Text = "0.00"
        
        txtBalance(Idx_结帐说明).Enabled = bytFun <> 3
        txtBalance(Idx_本次未结).Enabled = False
        txtBalance(Idx_本次结帐).Enabled = False
        
        txtBalance(Idx_冲预交).Enabled = False
        txtBalance(Idx_结帐说明).BackColor = IIf(bytFun <> 3, &H80000005, &H8000000F)
        txtBalance(Idx_本次结帐).BackColor = &H8000000F
        txtBalance(Idx_本次未结).BackColor = &H8000000F
        
        txtReceive.Locked = True '锁住不允许输入
        
        txtBalance(Idx_冲预交).BackColor = &H8000000F
        txtBalance(Idx_本次结帐).BackColor = IIf(txtBalance(Idx_本次结帐).Enabled, &H80000005, &H8000000F)
        txtBalance(Idx_本次未结).BackColor = IIf(txtBalance(Idx_本次未结).Enabled, &H80000005, &H8000000F)
        IDKindPaymentsType.Enabled = False
        cmdDelBalance.Visible = False
        
    Case Else   '已保存了结帐单
        blnEnabled = mEditType <> g_Ed_取消结帐
        lngColor = IIf(blnEnabled, &H80000005, &H8000000F)
        txtBalance(Idx_结帐说明).Enabled = False
        txtBalance(Idx_本次结帐).Enabled = False
        txtBalance(Idx_本次未结).Enabled = False
        txtBalance(Idx_冲预交).Enabled = mPatiInfor.dbl实际余额 <> 0 And blnEnabled
        
        txtBalance(Idx_结帐说明).BackColor = lngColor
        txtBalance(Idx_本次结帐).BackColor = IIf(txtBalance(Idx_本次结帐).Enabled, &H80000005, &H8000000F)
        txtBalance(Idx_本次未结).BackColor = IIf(txtBalance(Idx_本次未结).Enabled, &H80000005, &H8000000F)

        If mBalanceInfor.bln预交刷卡 Then
            txtBalance(Idx_冲预交).BackColor = IIf(txtBalance(Idx_冲预交).Enabled, &HE0E0E0, &H8000000F)
        Else
            txtBalance(Idx_冲预交).BackColor = IIf(txtBalance(Idx_冲预交).Enabled, &H80000005, &H8000000F)
        End If
        txtReceive.Locked = False '解锁

        IDKindPaymentsType.Enabled = blnEnabled
        
        cmdMore.Enabled = False
        cboPatiNums.Enabled = False
        txtPatient.Locked = True
        IDKind.Enabled = False
        picPati.Enabled = False
        cmdYBBalance.Visible = False
        cmdOK.Visible = True
        cmdOK.Enabled = True
    
        cmdDelBalance.Visible = chkCancel.Value = 0
        cmdDelBalance.Left = cmdCancel.Left
        cmdDelBalance.Top = cmdCancel.Top
        
        cmdCancel.Visible = IIf(mEditType = g_Ed_重新作废 Or mEditType = g_Ed_重新结帐 Or mEditType = g_Ed_取消结帐, True, False)
        cmdOK.Left = IIf(cmdCancel.Visible Or cmdDelBalance.Visible, cmdCancel.Left, picBalanceBack.ScaleWidth) - cmdOK.Width - 60
        cmdNext.Left = cmdOK.Left - cmdNext.Width - 50
    End Select
    txtBalance(Idx_本次结帐).BackColor = IIf(txtBalance(Idx_本次结帐).Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
End Sub

Private Sub SetNextBalanceCmdVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置连续结帐按钮
    '编制:刘兴洪
    '日期:2015-02-26 16:09:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean, objCard As Card
    Dim blnHave As Boolean, i As Long
    
    On Error GoTo errHandle
    
    If mty_ModulePara.byt缴款输入控制 <> 2 Then
        cmdNext.Visible = False: Exit Sub
    End If
    
    blnHave = False
    If Not mrsFeeList Is Nothing Then
        If mrsFeeList.State = 1 Then
            blnHave = mrsFeeList.RecordCount <> 0
        End If
    End If
    '普通收费或医保已经结算
    blnTemp = (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐) And chkCancel.Value = 0
    blnTemp = blnTemp And mPatiInfor.lng病人ID <> 0
    blnTemp = blnTemp And (mYBInFor.intInsure = 0 Or mobjBalanceCon.blnCurBalanceOwnerFee Or mYBInFor.intInsure <> 0 And Not cmdYBBalance.Visible)
    blnTemp = blnTemp And Val(txtReceive.Text) = 0
    blnTemp = blnTemp And blnHave
    
    cmdNext.Visible = blnTemp
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Function LoadDefaultFilterCons() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缺省的过滤条件
    '编制:刘兴洪
    '日期:2015-01-05 14:07:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyBalance As clsBalanceAllCon, dtDate As Date
    Dim cllOwnerFeeType As Collection, cllBalanceFeeType As Collection
    Dim blnCheck As Boolean, i As Long, bln体检 As Boolean, bln普通 As Boolean
    Dim int主页ID As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnAll As Boolean, rsAllTime As ADODB.Recordset
    Dim objNode As Node, intInsure As Integer, strInsureName As String
    Dim varTemp As Variant
    
    On Error GoTo errHandle
    
    '处理病人未结帐的所有范围变量
    Set mobjBalanceAll = New clsBalanceAllCon
    With mobjBalanceAll
        .MinDate = #1/1/1900#: .MaxDate = #1/1/1900#
    End With
    Set cllOwnerFeeType = New Collection
    Set cllBalanceFeeType = New Collection
    With mrsFeeList
        If .RecordCount <> 0 Then
            .MoveFirst
            If mty_ModulePara.int费用时间 = 0 Then
                dtDate = mrsFeeList!登记时间
            Else
                dtDate = mrsFeeList!时间
            End If
             mobjBalanceAll.MinDate = dtDate: mobjBalanceAll.MaxDate = dtDate
        End If
        
        Do While Not .EOF
            If mEditType <> g_Ed_门诊结帐 Then
                If InStr(mobjBalanceAll.strAllTime & ",", "," & Val(NVL(!主页ID)) & ",") = 0 And Val(NVL(!主页ID)) <> 0 Then
                    mobjBalanceAll.strAllTime = mobjBalanceAll.strAllTime & "," & Val(NVL(!主页ID))
                End If
            Else
                If InStr(mobjBalanceAll.strAllTime & ",", "," & Val(NVL(!主页ID)) & ",") = 0 Then
                    mobjBalanceAll.strAllTime = mobjBalanceAll.strAllTime & "," & Val(NVL(!主页ID))
                End If
            End If
            
            If Val(NVL(mrsFeeList!开单部门ID)) <> 0 Then
                If InStr(mobjBalanceAll.strAllDeptIDs & ",", "," & Val(NVL(!开单部门ID)) & ",") = 0 Then
                    mobjBalanceAll.strAllDeptIDs = mobjBalanceAll.strAllDeptIDs & "," & mrsFeeList!开单部门ID
                End If
            End If
            
            If Trim(NVL(!费目, "")) <> "" Then
                If InStr(mobjBalanceAll.strAllItem & ",", ",'" & !费目 & "',") = 0 Then
                     mobjBalanceAll.strAllItem = mobjBalanceAll.strAllItem & ",'" & !费目 & "'"
                End If
            End If
            
            If Trim(NVL(!诊断, "")) <> "" Then
                If InStr(mobjBalanceAll.strAllDiag & ",", ",'" & !诊断 & "',") = 0 Then
                     mobjBalanceAll.strAllDiag = mobjBalanceAll.strAllDiag & ",'" & !诊断 & "'"
                End If
            End If
            
            If Trim(NVL(!收费类别)) <> "" Then  '34260
                If InStr("," & mobjBalanceAll.strAllChargeType & ",", ",'" & !收费类别 & "',") = 0 Then
                    mobjBalanceAll.strAllChargeType = mobjBalanceAll.strAllChargeType & ",'" & !收费类别 & "'"
                    If InStr(1, "," & mty_ModulePara.strOwnerPayFeeType & ",", "," & !收费类别 & ",") > 0 Then
                        If InStr("," & mobjBalanceAll.strAllOwnerFeeType & ",", ",'" & !收费类别 & "',") = 0 Then
                            mobjBalanceAll.strAllOwnerFeeType = mobjBalanceAll.strAllOwnerFeeType & ",'" & !收费类别 & "'"
                        End If
                        cllOwnerFeeType.Add Array("'" & !收费类别 & "'", NVL(!收费类别名, "未知"))
                    Else
                        cllBalanceFeeType.Add Array("'" & !收费类别 & "'", NVL(!收费类别名, "未知"))
                    End If
                End If
             
            End If
            '如果为空,指没有设置费用类型
            If InStr("," & mobjBalanceAll.strAllClass & ",", ",'" & NVL(!类型, "无") & "',") = 0 Then
                mobjBalanceAll.strAllClass = mobjBalanceAll.strAllClass & ",'" & NVL(!类型, "无") & "'"
            End If
               
            If InStr("," & mobjBalanceAll.strAllBabys & ",", "," & Val(NVL(!婴儿费)) & ",") = 0 And Val(NVL(!婴儿费)) <> 0 Then
                mobjBalanceAll.strAllBabys = mobjBalanceAll.strAllBabys & "," & Val(NVL(!婴儿费)) & ""
            End If
            
            '比较取最大最小值
            If mty_ModulePara.int费用时间 = 0 Then
                dtDate = mrsFeeList!登记时间
            Else
                dtDate = mrsFeeList!时间
            End If
            If dtDate < mobjBalanceAll.MinDate Then mobjBalanceAll.MinDate = dtDate
            If dtDate > mobjBalanceAll.MaxDate Then mobjBalanceAll.MaxDate = dtDate
            If mEditType = g_Ed_门诊结帐 Then
                If Val(NVL(mrsFeeList!门诊标志)) = 4 Then
                    bln体检 = True
                Else
                    bln普通 = True
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '加载住院次数
    If LoadDataPatiNumsToComBox(Val(NVL(mrsInfo!病人ID)), Mid(mobjBalanceAll.strAllTime, 2), blnAll, rsAllTime, intInsure, strInsureName) = False Then Exit Function
    
    Set mobjBalanceAll.rsAllTime = rsAllTime
    With mobjBalanceAll
        .strAllTime = Mid(.strAllTime, 2)
        .strAllItem = Mid(.strAllItem, 2)
        .strAllDiag = Mid(.strAllDiag, 2)
        .strAllDeptIDs = Mid(.strAllDeptIDs, 2)
        .strAllChargeType = Mid(.strAllChargeType, 2)
        .strAllOwnerFeeType = Mid(.strAllOwnerFeeType, 2)
        .strAllClass = Mid(.strAllClass, 2)
        '显示结帐时间
        mblnNotChange = True
        txtBegin.Text = Format(.MinDate, txtBegin.Format)
        txtEnd.Text = Format(.MaxDate, txtEnd.Format)
        mblnNotChange = False
    End With
    Call SetPatiConsControlVisible
    LoadDefaultFilterCons = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPatiIsVerfy(Optional ByRef strMessage As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人是否审核
    '出参:strMessage-错误信息
    '编制:刘兴洪
    '日期:2015-01-05 14:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnAll As Boolean, lng主页ID As Long, i As Long
    

    On Error GoTo errHandle
    '门诊不进行检查
    If mEditType = g_Ed_门诊结帐 Or mrsInfo Is Nothing Then CheckPatiIsVerfy = True: Exit Function
    If InStr(mstrPrivs, ";未审核病人中途结帐;") > 0 Or InStr(mstrPrivs, ";未审核病人出院结帐;") > 0 Then CheckPatiIsVerfy = True: Exit Function
    If Val(NVL(mrsInfo!主页ID)) = 0 Then CheckPatiIsVerfy = True: Exit Function
    
    If CStr(mrsInfo!主页ID) = mobjBalanceAll.strAllTime Then  '只有最后一次未结
        If mrsInfo!审核标志 = 0 Then
            strMessage = "当前病人未审核，你不能对未审核的病人进行结帐。"
            Exit Function
        End If
        CheckPatiIsVerfy = True: Exit Function
    End If
    blnAll = True
    For i = 0 To UBound(Split(mobjBalanceAll.strAllTime, ","))
        lng主页ID = Val(Split(mobjBalanceAll.strAllTime, ",")(i))
        If lng主页ID <> 0 Then
            If Not Chk病人审核(mrsInfo!病人ID, lng主页ID) Then
                 mobjBalanceAll.strUnAuditTime = mobjBalanceAll.strUnAuditTime & "," & lng主页ID
            Else
                blnAll = False
            End If
        Else
            blnAll = False
        End If
    Next
    If mobjBalanceAll.strUnAuditTime <> "" Then mobjBalanceAll.strUnAuditTime = Mid(mobjBalanceAll.strUnAuditTime, 2)
    If blnAll Then
        strMessage = "该病人所有住院费用都没有审核，不能进行结帐！"
        Exit Function
    End If
    CheckPatiIsVerfy = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckInputBlood() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输血费检查
    '返回:血费检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 15:18:37
    '问题:34260:输血费检查
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '0:不检查;1-检查并提示
    If mty_ModulePara.byt结帐时输血费检查 <> 1 Then CheckInputBlood = True: Exit Function
    If InStr(1, "," & mobjBalanceAll.strAllChargeType & ",", ",'K',") = 0 Then CheckInputBlood = True: Exit Function
    If MsgBox("注意:" & vbCrLf & "    该病人未结费用中包含了输血费,本次是否只结输血费?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then CheckInputBlood = True: Exit Function
    
    mobjBalanceCon.strChargeType = "'K'"
    If ShowBalance(False) Then CheckInputBlood = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadBalanceData(ByRef rsBalance As ADODB.Recordset, ByVal blnUpload As Boolean, Optional ByVal blnInputAfterPati As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结帐数据
    '入参:blnUpload-是否只读上传数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 15:35:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng病人ID = Val(NVL(mrsInfo!病人ID))
    End If
    
        '读取结帐信息
    If mEditType = g_Ed_门诊结帐 Then
        With mobjBalanceCon
            If .strChargeType = "" And .blnCurBalanceOwnerFee = False And blnInputAfterPati = False Then
                Set rsBalance = GetMzBalanceData(lng病人ID, .strDeptIDs, _
                        .strClass, .dtBeginDate, .dtEndDate, .strItem, blnUpload, _
                       mty_ModulePara.blnZero, mblnDateMoved, .bytKind, .strChargeType, mty_ModulePara.int费用时间 = 1, .strTime, mty_ModulePara.strOwnerPayFeeType, .strDiag)
            Else
                Set rsBalance = GetMzBalanceData(lng病人ID, .strDeptIDs, _
                        .strClass, .dtBeginDate, .dtEndDate, .strItem, blnUpload, _
                       mty_ModulePara.blnZero, mblnDateMoved, .bytKind, .strChargeType, mty_ModulePara.int费用时间 = 1, .strTime, , .strDiag)
            End If
        End With
        ReadBalanceData = True
        Exit Function
    End If
    With mobjBalanceCon
        If .strChargeType = "" And .blnCurBalanceOwnerFee = False And blnInputAfterPati = False Then
            Set rsBalance = GetZYBalanceData(lng病人ID, .strTime, .strDeptIDs, .strClass, _
                .dtBeginDate, .dtEndDate, .strBaby, .strItem, blnUpload, mty_ModulePara.blnZero, _
                mblnDateMoved, .strChargeType, mty_ModulePara.int费用时间 = 1, mty_ModulePara.strOwnerPayFeeType, .strDiag)
        Else
            Set rsBalance = GetZYBalanceData(lng病人ID, .strTime, .strDeptIDs, .strClass, _
                .dtBeginDate, .dtEndDate, .strBaby, .strItem, blnUpload, mty_ModulePara.blnZero, _
                mblnDateMoved, .strChargeType, mty_ModulePara.int费用时间 = 1, , .strDiag)
        End If
    End With
    ReadBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBalanceMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结帐金额
    '编制:刘兴洪
    '日期:2015-01-12 14:11:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mBalanceInfor
        .dbl本次未结 = 0
        .dbl当前结帐 = 0
        .dbl已付合计 = 0
        .dbl未付合计 = 0
        .dbl医保支付合计 = 0
        .dbl冲预交合计 = 0
    End With
End Sub

Private Function LoadFeeListFromBalanceID(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来加载费目表数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 18:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str主页Ids As String
    Dim lngRow As Long, dblMoney(0 To 2) As Double
    Dim intSign As Integer
    On Error GoTo errHandle
    
    Call LoadDetailListFromBalanceID(lng结帐ID)
    intSign = IIf(mEditType = g_Ed_重新作废 Or (mEditType = g_Ed_单据查看 And mblnViewCancel), -1, 1)
    
    strSQL = _
    " Select Mod(B.记录性质,10) as 记录性质, B.NO,B.序号,B.收据费目, " & _
    "          Sum(B.应收金额) As 应收金额," & _
    "          Sum(B.实收金额) As 实收金额,0 as 结帐金额" & _
    " From 住院费用记录 A,住院费用记录 B " & _
    " Where A.结帐ID=[1] And  Mod(A.记录性质,10)=Mod(B.记录性质,10)  " & _
    "       And A.NO=B.NO And A.序号=B.序号 And A.记录状态 = B.记录状态 " & _
    " Group by Mod(B.记录性质,10), B.NO,B.序号,B.收据费目"
    strSQL = strSQL & " UNION ALL " & _
    "   Select Mod(A.记录性质,10) as 记录性质, A.NO,序号,A.收据费目, " & _
    "           0 as 应收金额,0 as 实收金额,sum(A.结帐金额) as 结帐金额 " & _
    "   From 住院费用记录 A " & _
    "   Where A.结帐ID= [1]  " & _
    "   Group by Mod(A.记录性质,10),A.NO,A.序号,A.收据费目 "
    
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "住院费用记录", "门诊费用记录")

    If mblnNOMoved Then
        strSQL = Replace(Replace(strSQL, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
    End If
    
    strSQL = "" & _
    "   Select 收据费目, sum(应收金额) as 应收金额,sum(实收金额) as 实收金额,sum(结帐金额) as 结帐金额 " & _
    "   From (" & strSQL & ")" & _
    "   Group by 收据费目" & _
    "   Order by 收据费目"
    

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    
    On Error GoTo errHandle
    With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        str主页Ids = "": lngRow = 1
        Do While Not rsTemp.EOF
          .TextMatrix(lngRow, .ColIndex("费目")) = NVL(rsTemp!收据费目, "未知")
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))) + Val(NVL(rsTemp!应收金额))
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("实收金额"))) + Val(NVL(rsTemp!实收金额))
          .TextMatrix(lngRow, .ColIndex("实收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("实收金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))) + Val(NVL(rsTemp!结帐金额))
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("结帐金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))) + RoundEx(intSign * Val(NVL(rsTemp!结帐金额)), 6)
          .TextMatrix(lngRow, .ColIndex("结帐金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))), gstrDec)
          dblMoney(0) = dblMoney(0) + Val(NVL(rsTemp!应收金额))
          dblMoney(1) = dblMoney(1) + Val(NVL(rsTemp!实收金额))
          dblMoney(2) = dblMoney(2) + RoundEx(intSign * Val(NVL(rsTemp!结帐金额)), 6)
          .Rows = .Rows + 1: lngRow = .Rows - 1
          rsTemp.MoveNext
        Loop
        If str主页Ids <> "" Then str主页Ids = Mid(str主页Ids, 2)
        
        If .TextMatrix(1, .ColIndex("费目")) <> "" Then
           lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("费目")) = "合计"
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("实收金额")) = Format(dblMoney(1), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = Val(dblMoney(2))
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(dblMoney(2), gstrDec)
         
          .Cell(flexcpData, lngRow, .ColIndex("结帐金额")) = Val(dblMoney(2))
          .TextMatrix(lngRow, .ColIndex("结帐金额")) = Format(dblMoney(2), gstrDec)
         
           .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True
        End If
        .Redraw = flexRDBuffered
    End With
    
    mBalanceInfor.dbl本次未结 = RoundEx(dblMoney(2), 6)
    mBalanceInfor.dbl当前结帐 = mBalanceInfor.dbl本次未结
    mBalanceInfor.dbl未付合计 = mBalanceInfor.dbl本次未结
    
    mblnNotChange = True
    txtBalance(Idx_本次未结).Text = Format(dblMoney(2), gstrDec)
    txtBalance(Idx_本次未结).Enabled = False
    txtBalance(Idx_本次结帐).Text = Format(dblMoney(2), gstrDec)
    mblnNotChange = False
    LoadFeeListFromBalanceID = True
    Exit Function
errHandle:
    vsFeeList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDetailListFromBalanceID(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来加载费目表数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 18:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str主页Ids As String
    Dim lngRow As Long, dblMoney(0 To 2) As Double
    Dim intSign As Integer
    On Error GoTo errHandle
    
    intSign = IIf(mEditType = g_Ed_重新作废 Or (mEditType = g_Ed_单据查看 And mblnViewCancel), -1, 1)
    
    strSQL = _
    " Select Mod(B.记录性质,10) as 记录性质, B.NO,B.序号,C.名称 As 项目,Max(B.登记时间) As 登记时间, " & _
    "          Avg(B.应收金额) As 应收金额," & _
    "          Avg(B.实收金额) As 实收金额,0 as 结帐金额,Decode(B.记录状态,2,2,1) As 记录状态,Max(a.门诊标志) As 门诊标志" & _
    " From 住院费用记录 A,住院费用记录 B,收费项目目录 C " & _
    " Where A.结帐ID=[1] And  Mod(A.记录性质,10)=Mod(B.记录性质,10) And A.记录状态 = B.记录状态  " & _
    "       And B.收费细目ID=C.ID And A.NO=B.NO And A.序号=B.序号 " & _
    " Group by Mod(B.记录性质,10), B.NO,B.序号,C.名称,Decode(B.记录状态,2,2,1)"
    strSQL = strSQL & " UNION ALL " & _
    "   Select Mod(A.记录性质,10) as 记录性质, A.NO,序号,B.名称 As 项目,Max(A.登记时间) As 登记时间, " & _
    "           0 as 应收金额,0 as 实收金额,Sum(A.结帐金额) as 结帐金额,Decode(A.记录状态,2,2,1) As 记录状态,Max(a.门诊标志) As 门诊标志 " & _
    "   From 住院费用记录 A,收费项目目录 B " & _
    "   Where A.结帐ID= [1] And A.收费细目ID=B.ID  " & _
    "   Group by Mod(A.记录性质,10),A.NO,A.序号,B.名称,Decode(A.记录状态,2,2,1) "
    
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "住院费用记录", "门诊费用记录")

    If mblnNOMoved Then
        strSQL = Replace(Replace(strSQL, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
    End If
    
    strSQL = "" & _
    "   Select Max(登记时间) As 登记时间,NO,序号,项目, sum(应收金额) as 应收金额,sum(实收金额) as 实收金额," & _
    "          sum(结帐金额) as 结帐金额,记录状态,Max(门诊标志) As 门诊标志 " & _
    "   From (" & strSQL & ")" & _
    "   Group by NO,序号,项目,记录状态" & _
    "   Order by NO,序号"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    
    On Error GoTo errHandle
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        str主页Ids = "": lngRow = 1
        .TextMatrix(0, .ColIndex("未结金额")) = "实收金额"
        If intSign = -1 Then
            .TextMatrix(0, .ColIndex("结帐金额")) = "作废金额"
        Else
            .TextMatrix(0, .ColIndex("结帐金额")) = "结帐金额"
        End If
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("日期")) = Format(NVL(rsTemp!登记时间), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("单据")) = NVL(rsTemp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = NVL(rsTemp!项目)
            .TextMatrix(.Rows - 1, .ColIndex("未结金额")) = Format(NVL(rsTemp!实收金额), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("未结金额")) = Val(NVL(rsTemp!实收金额))
            .TextMatrix(.Rows - 1, .ColIndex("结帐金额")) = Format(intSign * Val(NVL(rsTemp!结帐金额)), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("结帐金额")) = intSign * Val(NVL(rsTemp!结帐金额))
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = Val(NVL(rsTemp!序号))
            If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then
                .Cell(flexcpData, .Rows - 1, .ColIndex("序号")) = Val(NVL(rsTemp!门诊标志))
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .Cell(flexcpBackColor, 1, .ColIndex("结帐金额"), .Rows - 1, .ColIndex("结帐金额")) = .Cell(flexcpBackColor, 1, .ColIndex("日期"), 0.1, .ColIndex("日期"))
        .Redraw = flexRDBuffered
    End With
    
    LoadDetailListFromBalanceID = True
    Exit Function
errHandle:
    vsDetailList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCentMoney(ByVal dblMoney As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据分币处理规则,返回分币处理后的金额
    '入参:dblMoney-未处理的原始金额
    '返回:返回分币处理后的金额
    '编制:刘兴洪
    '日期:2015-01-26 10:57:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    Set objCard = GetCard(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算方式")))
    If objCard Is Nothing Then GetCentMoney = Format(dblMoney, "0.00"): Exit Function
    '非现金的,保留两位小数
    If objCard.结算性质 <> 1 Then GetCentMoney = Format(dblMoney, "0.00"): Exit Function
    
    If mYBInFor.intInsure = 0 Then
        GetCentMoney = CentMoney(CCur(dblMoney))
        Exit Function
    End If
    If MCPAR.分币处理 Then
        GetCentMoney = CentMoney(CCur(dblMoney))
    Else
        GetCentMoney = Format(dblMoney, "0.00")
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub LoadCurOwnerPayInfor(Optional ByVal blnDefault As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载当前结算信息
    '编制:刘兴洪
    '日期:2015-01-12 14:14:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long, objCard As Card
    Dim dbl剩余自付 As Double, dbl自付合计 As Double
    Dim i As Long
    Dim dblCashMoney As Double
    
    On Error GoTo errHandler
    With mBalanceInfor
        '取出误差费和现金一起进行分币处理
        For i = 1 To vsBlance.Rows - 1
            If Val(vsBlance.RowData(i)) = 999 Then      '现金
                dblCashMoney = Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("结算金额")))
                Exit For
            End If
        Next
        .dbl未付合计 = RoundEx(.dbl当前结帐 - .dbl已付合计 - .dbl冲预交合计, 5)
        .dbl未付合计 = RoundEx(GetCentMoney(.dbl未付合计 + dblCashMoney) - dblCashMoney, 5)
        
        Select Case mEditType
        Case g_Ed_取消结帐, g_Ed_结帐作废, g_Ed_重新作废
            mPatiInfor.bln退款标志 = .dbl未付合计 >= 0
        Case Else
            If chkCancel.Value = 1 Then
                mPatiInfor.bln退款标志 = .dbl未付合计 >= 0
            Else
                mPatiInfor.bln退款标志 = .dbl未付合计 < 0
            End If
        End Select
        '设置字体显示
        lngColor = IIf(mPatiInfor.bln退款标志, vbRed, vbBlue)
    End With
    
    txtOwe.ForeColor = IIf(blnDefault, vbBlue, lngColor)
    txtOwe.Text = Format(Abs(mBalanceInfor.dbl未付合计), mstrDec)
    If blnDefault Then Call LoadDefaultMoney
    Call SetCaculated
    Show误差金额 False
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function Get应缴() As Currency
    Dim i As Long
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("结算性质"))) = 1 Then
                Get应缴 = Val(.TextMatrix(i, .ColIndex("结算金额")))
                Exit Function
            End If
        Next
    End With
End Function


Private Function LoadFeeList() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费目表数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dblMoney(0 To 2) As Double
 
    On Error GoTo errHandle
    Call LoadDetailList
    With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If mrsFeeList.RecordCount <> 0 Then mrsFeeList.MoveFirst
        Do While Not mrsFeeList.EOF
           lngRow = .FindRow(NVL(mrsFeeList!费目, "未知"), "1", .ColIndex("费目"), , True)
           If lngRow < 0 Then
                If .TextMatrix(1, .ColIndex("费目")) = "" Then
                    lngRow = 1
                Else
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                End If
           End If
           
           If .TextMatrix(1, .ColIndex("费目")) = "" Then lngRow = 1
          .TextMatrix(lngRow, .ColIndex("费目")) = NVL(mrsFeeList!费目, "未知")
          
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))) + Val(NVL(mrsFeeList!应收金额))
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("实收金额"))) + Val(NVL(mrsFeeList!实收金额))
          .TextMatrix(lngRow, .ColIndex("实收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("实收金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("未结金额"))) + Val(NVL(mrsFeeList!未结金额))
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("未结金额"))), gstrDec)
            
          dblMoney(0) = dblMoney(0) + Val(NVL(mrsFeeList!应收金额))
          dblMoney(1) = dblMoney(1) + Val(NVL(mrsFeeList!实收金额))
          dblMoney(2) = dblMoney(2) + Val(NVL(mrsFeeList!未结金额))
            mrsFeeList.MoveNext
        Loop
        .ColSort(.ColIndex("费目")) = flexSortUseColSort
        If .TextMatrix(1, .ColIndex("费目")) <> "" Then
          .Rows = .Rows + 1: lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("费目")) = "合计"
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("实收金额")) = Format(dblMoney(1), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = Val(dblMoney(2))
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(dblMoney(2), gstrDec)
         
           .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True
        End If
         .Redraw = flexRDBuffered
    End With
    
'    zl_vsGrid_Para_Restore mlngModul, vsFeeList, Me.Name, "费用列表"
    mBalanceInfor.dbl本次未结 = dblMoney(2)
    mBalanceInfor.dbl当前结帐 = dblMoney(2)
    mBalanceInfor.dbl未付合计 = dblMoney(2)
    
    mblnNotChange = True
    txtBalance(Idx_本次未结).Text = Format(dblMoney(2), gstrDec)
    txtBalance(Idx_本次结帐).Text = Format(dblMoney(2), gstrDec)
    mblnNotChange = False
    
    Call LoadCurOwnerPayInfor '加载当前自付信息
    LoadFeeList = True
    Exit Function
errHandle:
    mblnNotChange = False
    vsFeeList.Redraw = flexRDNone
     Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDetailList() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费目表数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dblMoney(0 To 2) As Double
    Dim i As Long
 
    On Error GoTo errHandle
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If mrsFeeList.RecordCount <> 0 Then mrsFeeList.MoveFirst
        Do While Not mrsFeeList.EOF
            .TextMatrix(.Rows - 1, .ColIndex("日期")) = Format(NVL(mrsFeeList!时间), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("单据")) = NVL(mrsFeeList!单据号)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = NVL(mrsFeeList!项目)
            .TextMatrix(.Rows - 1, .ColIndex("未结金额")) = Format(NVL(mrsFeeList!未结金额), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("未结金额")) = Val(NVL(mrsFeeList!未结金额))
            .TextMatrix(.Rows - 1, .ColIndex("结帐金额")) = Format(NVL(mrsFeeList!结帐金额), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("结帐金额")) = Val(NVL(mrsFeeList!结帐金额))
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(mrsFeeList!ID, 0)
            .TextMatrix(.Rows - 1, .ColIndex("记录性质")) = Val(NVL(mrsFeeList!记录性质))
            .TextMatrix(.Rows - 1, .ColIndex("记录状态")) = IIf(Val(NVL(mrsFeeList!记录状态)) = 3, 1, Val(NVL(mrsFeeList!记录状态)))
            .TextMatrix(.Rows - 1, .ColIndex("执行状态")) = Val(NVL(mrsFeeList!执行状态))
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = Val(NVL(mrsFeeList!序号))
            If mEditType = g_Ed_门诊结帐 Then
                .Cell(flexcpData, .Rows - 1, .ColIndex("序号")) = Val(NVL(mrsFeeList!门诊标志))
            End If
            .Rows = .Rows + 1
            mrsFeeList.MoveNext
        Loop
        If mYBInFor.intInsure <> 0 Then
            .Cell(flexcpBackColor, 1, .ColIndex("结帐金额"), .Rows - 1, .ColIndex("结帐金额")) = .Cell(flexcpBackColor, 1, .ColIndex("单据"))
        Else
            .Cell(flexcpBackColor, 1, .ColIndex("结帐金额"), .Rows - 1, .ColIndex("结帐金额")) = &HFFFFC0
        End If
        If .TextMatrix(1, .ColIndex("单据")) <> "" Then .Rows = .Rows - 1
         .Redraw = flexRDBuffered
    End With
    
'    zl_vsGrid_Para_Restore mlngModul, vsDetailList, Me.Name, "明细列表"
    LoadDetailList = True
    Exit Function
errHandle:
     vsDetailList.Redraw = flexRDNone
     Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadBalanceDepositList(ByVal lng病人ID As Long, _
    ByVal lng结帐ID As Long, ByVal blnDateMoved As Boolean, _
    str主页Ids As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载指定结帐单的冲预交信息
    '入参:lng结帐ID-指定的结帐ID
    '     blnDateMoved-当前是否移动到后备表中
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 15:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim dblTotal As Double
    Dim lng原结帐ID As Long
    On Error GoTo errHandle
    
    Set rsTemp = GetBalanceDeposit(lng结帐ID, blnDateMoved)
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        i = 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        'ID,单据号,票据号,日期,结算方式, 金额
        Do While Not rsTemp.EOF
            .RowData(i) = ""
            .TextMatrix(i, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(i, .ColIndex("单据号")) = rsTemp!单据号
            .TextMatrix(i, .ColIndex("票据号")) = "" & rsTemp!票据号
            .TextMatrix(i, .ColIndex("收款日期")) = Format(rsTemp!日期, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("结算方式")) = NVL(rsTemp!结算方式)
            .TextMatrix(i, .ColIndex("冲预交")) = Format(rsTemp!金额, "0.00")
            .Rows = .Rows + 1: i = i + 1
            dblTotal = dblTotal + Val(NVL(rsTemp!金额))
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = .Cols - 1
        .ColWidth(.ColIndex("收款日期")) = 1305
        .ColWidth(.ColIndex("单据号")) = 1100
        .ColWidth(.ColIndex("结算方式")) = 1400
        .ColWidth(.ColIndex("余额")) = 1100
        .ColWidth(.ColIndex("冲预交")) = 1100
        
        .Redraw = flexRDBuffered
        If i > 1 Then .Rows = .Rows - 1
    End With
    
    txtBalance(Idx_冲预交).Text = Format(dblTotal, "0.00")
    chkDeposit.Tag = dblTotal
    mBalanceInfor.dbl冲预交合计 = dblTotal
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    If Not rsTemp.EOF Then
        lblTicketCount.Caption = "预交款收据:" & rsTemp.RecordCount & "张  合计:" & Format(dblTotal, "0.00") & "元"
    Else
        lblTicketCount.Caption = ""
    End If
    If rsTemp.RecordCount <> 0 Then LoadBalanceDepositList = True: Exit Function
    
    If mEditType = g_Ed_重新作废 Then
        '加载原结帐数据
        If mblnNotChange Then Exit Function
        mblnNotChange = True
        lng原结帐ID = zlGetFormerBalanceID(mBalanceInfor.strNO)
        LoadBalanceDepositList = LoadBalanceDepositList(lng病人ID, lng原结帐ID, blnDateMoved, str主页Ids)
      
        If mBalanceInfor.dbl冲预交合计 <> 0 Then chkDeposit.Value = 1
        mblnNotChange = False
        LoadBalanceDepositList = True
        Exit Function
    End If
    
    If mEditType <> g_Ed_单据查看 And mEditType <> g_Ed_结帐作废 And chkCancel.Value <> 1 Then
        '重新加载预交
        If LoadDepositList(lng病人ID, str主页Ids) = False Then Exit Function
    End If
    
    
    LoadBalanceDepositList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function LoadDepositList(ByVal lng病人ID As Long, _
    ByVal str主页Ids As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载预交数据
    '入参:lng病人ID-病人ID
    '     str主页IDs:多个用逗号分离
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 18:32:22
    '   mbln门诊转住院:36984
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str住院次数 As String, i As Long, str结算方式 As String
    Dim intTYPE As Integer, dblMoney As Double, dblTotal As Double
    On Error GoTo errHandle
    
    '显示预交明细
    str住院次数 = "": intTYPE = 1
    If mEditType = g_Ed_住院结帐 Or (mEditType <> g_Ed_门诊结帐 And mblnCurMzBalanceNo = False) Then
        str住院次数 = str主页Ids
        intTYPE = 2
    End If
    
    Set mrsDeposit = GetDeposit(lng病人ID, mblnDateMoved, str住院次数, mbln门诊转住院, mstrPepositDate, intTYPE, mrs结算方式)
    dblMoney = mBalanceInfor.dbl未付合计
    
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2
        .Clear 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        str结算方式 = ""
        If mrsDeposit.RecordCount <> 0 Then mrsDeposit.MoveFirst
        i = 1
        Do While Not mrsDeposit.EOF
            .RowData(i) = Val(NVL(mrsDeposit!记录状态))
            '.TextMatrix(i, .ColIndex("标志")) = i
            .TextMatrix(i, .ColIndex("ID")) = mrsDeposit!ID
            .TextMatrix(i, .ColIndex("单据号")) = mrsDeposit!NO
            .TextMatrix(i, .ColIndex("票据号")) = "" & mrsDeposit!票据号
            .TextMatrix(i, .ColIndex("收款日期")) = Format(mrsDeposit!日期, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("结算方式")) = NVL(mrsDeposit!结算方式)
            .TextMatrix(i, .ColIndex("余额")) = Format(mrsDeposit!金额, "0.00")
            .TextMatrix(i, .ColIndex("预交ID")) = NVL(mrsDeposit!预交ID)
            If mbln门诊转住院 Or _
                (mobjBalanceCon.blnCurBalanceOwnerFee And mty_ModulePara.bln自费缺省使用预交) Then
                If Val(NVL(mrsDeposit!金额)) <= dblMoney Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(mrsDeposit!金额, "0.00")
                    dblMoney = dblMoney - RoundEx(Val(NVL(mrsDeposit!金额)), 2)
                ElseIf dblMoney <> 0 Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(dblMoney, "0.00")
                    dblMoney = 0
                End If
            ElseIf Not mobjBalanceCon.blnCurBalanceOwnerFee Then
                .TextMatrix(i, .ColIndex("冲预交")) = Format(mrsDeposit!金额, "0.00")
            End If
            dblTotal = dblTotal + RoundEx(Val(NVL(mrsDeposit!金额)), 2)
            i = i + 1
            .Rows = .Rows + 1
            mrsDeposit.MoveNext
        Loop
        .Row = 1: .Col = .Cols - 1
        If i >= 2 And .Rows >= 2 Then .Rows = .Rows - 1
        .Redraw = flexRDBuffered
    End With
    
    
    '问题号113702,焦博,2017/08/30,格式化病人实际金额
    mPatiInfor.dbl实际余额 = RoundEx(dblTotal, 6)
    If mrsDeposit.RecordCount <> 0 Then mrsDeposit.MoveFirst
    If Not mrsDeposit.EOF Then
        lblTicketCount.Caption = "预交款收据:" & mrsDeposit.RecordCount & "张  合计:" & Format(dblTotal, "0.00") & "元"
    Else
        lblTicketCount.Caption = ""
    End If
    Call SetUpDown
    LoadDepositList = True
    Exit Function
errHandle:
    vsDeposit.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetDefaultHospitalizedDate(ByVal lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的住院日期
    '入参:lng病人ID-病人ID
    '返回:返回上次中途结帐的结束日期,无中途结帐时,返回空
    '编制:刘兴洪
    '日期:2015-01-06 15:25:02
    '说明:原问题号是30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select to_char( Max(结束日期) + 1,'yyyy-mm-dd') as 结束日期 " & _
    "   From 病人结帐记录 " & _
    "   Where  记录状态=1  And 病人iD=[1] and nvl(中途结帐,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsTemp.EOF Then Exit Function
    GetDefaultHospitalizedDate = NVL(rsTemp!结束日期)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function GetPatiHospitalzedDateRange(ByRef dtBeginDate As Date, ByRef dtEndDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 获取病人的入出院时间,门诊病人取最大和最小费用时间
    '出参:dtBeginDate-开始时间
    '     dtEndDate-结束时间
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-06 15:43:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strDefaultDate As String, lng主页ID As Long, lng病人ID As Long
    Dim strTime As String
    
    On Error GoTo errHandle
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State = 0 Then Exit Function
    If mrsInfo.RecordCount = 0 Then Exit Function
    
    lng病人ID = Val(NVL(mrsInfo!病人ID))
    
    strTime = mobjBalanceCon.strTime
    If strTime = "" Then strTime = mobjBalanceAll.strAllTime
        
    strDefaultDate = ""
    If mEditType <> g_Ed_门诊结帐 And strTime <> "" Then
        strDefaultDate = GetDefaultHospitalizedDate(lng病人ID)
    End If

    Call GetFeeDate(dtBeginDate, dtEndDate)
    If Val(NVL(mrsInfo!主页ID)) = 0 Then GetPatiHospitalzedDateRange = True: Exit Function
    
    lng主页ID = GetMinMaxTime(0)     '最小住院次数
    If lng主页ID = 0 Then GetPatiHospitalzedDateRange = True: Exit Function
    
    
    If lng主页ID = Val(NVL(mrsInfo!主页ID)) Then
        dtBeginDate = mrsInfo!入院日期
        
        If Not IsNull(mrsInfo!出院日期) Then
            dtEndDate = mrsInfo!出院日期
        Else
            dtEndDate = zlDatabase.Currentdate
        End If
        
        '入院时间比缺省的最后一次结帐时间还小,则开始时间以最后一次结帐时间为准
        If IsDate(strDefaultDate) Then    '问题:30043
            If Format(dtBeginDate, "yyyy-mm-dd") < strDefaultDate And Format(dtEndDate, "yyyy-mm-dd") > strDefaultDate Then dtBeginDate = CDate(strDefaultDate)
        End If

        GetPatiHospitalzedDateRange = True: Exit Function
    End If
    
    If CStr(lng主页ID) = strTime Then '可能是结以前某次住院的帐
        strSQL = "Select 入院日期,Nvl(出院日期,Sysdate) as 出院日期 From 病案主页" & _
                " Where 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        dtBeginDate = rsTmp!入院日期
        dtEndDate = rsTmp!出院日期

        If IsDate(strDefaultDate) Then
            If Format(dtBeginDate, "yyyy-mm-dd") < strDefaultDate And Format(dtEndDate, "yyyy-mm-dd") > strDefaultDate Then dtBeginDate = CDate(strDefaultDate)
        End If

    End If
    GetPatiHospitalzedDateRange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Function GetFeeDate(ByRef dtBeginDate As Date, ByRef dtEndDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的最小和最大费用时间
    '出参:dtBeginDate-开始时间
    '     dtEndDate-结束时间
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-06 15:54:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dtDate As Date
    On Error GoTo errHandle
    
    If mrsFeeList Is Nothing Then Exit Function
    If mrsFeeList.State <> 1 Then Exit Function
    If mrsFeeList.RecordCount = 0 Then GoTo GoEnd:
    mrsFeeList.MoveFirst
    If mty_ModulePara.int费用时间 = 0 Then
        dtDate = mrsFeeList!登记时间
    Else
        dtDate = mrsFeeList!时间
    End If
    
    dtBeginDate = dtDate: dtEndDate = dtDate
    With mrsFeeList
        Do While Not .EOF
            If mty_ModulePara.int费用时间 = 0 Then
                dtDate = mrsFeeList!登记时间
            Else
                dtDate = mrsFeeList!时间
            End If
            If dtDate < dtBeginDate Then dtBeginDate = dtDate
            If dtDate > dtEndDate Then dtEndDate = dtDate
            .MoveNext
        Loop
    End With
    mrsFeeList.MoveFirst
GoEnd:
    GetFeeDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetMinMaxTime(ByVal bytMode As Byte) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取未结费用中的最小或最大的住院次数,可能返回0
    '入参:bytMode,0-最小次数,1-最大次数
    '返回:住院次数
    '编制:刘兴洪
    '日期:2015-01-06 16:02:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTime As String, varData As Variant
    Dim i As Long, intTime As Integer
    
    On Error GoTo errHandle
        
    strTime = mobjBalanceCon.strTime
    If strTime = "" Then strTime = mobjBalanceAll.strAllTime
    
    varData = Split(strTime, ",")
    For i = 0 To UBound(varData)
        If i = 0 Then intTime = Val(varData(i))
        If bytMode = 0 Then
            If intTime > Val(varData(i)) Then intTime = Val(varData(i))
        Else
            If intTime < Val(varData(i)) Then intTime = Val(varData(i))
        End If
    Next
    GetMinMaxTime = intTime
Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlChangeDefaultTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:改变缺省的住院时间范围
    '编制:刘兴洪
    '日期:2015-01-06 16:42:36
    '说明：30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If opt出院.Value Then txtPatiEnd.Text = txtPatiEnd.Tag: Exit Sub

    txtPatiEnd.Text = Format(zlDatabase.Currentdate - 1, "yyyy-mm-dd")
    If txtPatiEnd.Text < txtPatiBegin.Text Then
        txtPatiEnd.Text = txtPatiEnd.Tag
    End If
    If txtPatiEnd.Text > txtPatiEnd.Tag Then
        txtPatiEnd.Text = txtPatiEnd.Tag
    End If
End Sub

Private Sub RecalcDepositMoney(ByVal bytOperationType As Byte, _
    Optional ByVal dblMoney As Double = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算冲预交金额
    '入参:bytOperationType-操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按指定金额来冲预交(按时间先后来分摊）;3-全冲
    '     dblMoneny-冲预交金额
    '编制:刘兴洪
    '日期:2015-01-07 14:49:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytCurFun As Byte  '0-全清预交款;1-按结帐金额来冲预交;2-使用所有预交款;
    Dim dblTotal As Double, i As Long
    
    On Error GoTo errHandle
    mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + mBalanceInfor.dbl冲预交合计, 6)
    mBalanceInfor.dbl冲预交合计 = 0
    
    Select Case bytOperationType
    Case 0  '0-清除所有冲预交
        bytCurFun = 0
    Case 1  '1-按缺省使用预交款
        bytCurFun = 1   '门诊结帐或中途结帐，缺省按结帐金额来使用
        If mEditType = g_Ed_住院结帐 And opt出院.Value Then bytCurFun = 2
        If mEditType = g_Ed_住院结帐 And mty_ModulePara.bln中途结帐退预交 Then bytCurFun = 2
        If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then
            Select Case mty_ModulePara.bytMzDeposit '门诊预交缺省使用方式
            Case 0 ' 0-缺省不使用交;1-按结帐金额使用预交;2-使用所有预交
                bytCurFun = 0
            Case 1 '1-按结帐金额使用预交
                bytCurFun = 1
            Case 2 '2-使用所有预交
                bytCurFun = 2
            End Select
        End If
        If mEditType = g_Ed_重新结帐 Then
            If InStr(lblBalanceType.Caption, "出院") > 0 Then
                bytCurFun = 2
            End If
        End If
        dblMoney = RoundEx(mBalanceInfor.dbl未付合计, 2)
    Case 2 '2-按指定金额来冲预交(按时间先后来分摊）
        bytCurFun = 1
        If dblMoney = 0 Then dblMoney = RoundEx(mBalanceInfor.dbl未付合计, 2)
    Case 3 '3-全冲
        bytCurFun = 2
    Case Else
         bytCurFun = 0
    End Select
    
    If dblMoney < 0 Then dblMoney = 0
    With vsDeposit
        dblTotal = 0

            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If Val(.TextMatrix(i, .ColIndex("编辑状态"))) = 0 Then
                        .Cell(flexcpText, i, .ColIndex("冲预交"), i, .ColIndex("冲预交")) = "0.00"
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                        Select Case bytCurFun
                            Case 1 '按结帐金额使用
                                If dblMoney = 0 Then GoTo NextDeposit
                                If Val(.TextMatrix(i, .ColIndex("余额"))) <= dblMoney Then
                                      .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                                Else
                                    .TextMatrix(i, .ColIndex("冲预交")) = Format(dblMoney, "0.00")
                                End If
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                                dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("冲预交")))
                            Case 2 '全冲
                                .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                            Case Else
                        End Select
                    Else
                        dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                        dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("冲预交")))
                    End If
                End If
NextDeposit:
            Next
    End With
    mBalanceInfor.dbl冲预交合计 = dblTotal
    mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 - mBalanceInfor.dbl冲预交合计, 6)
    '0-医保预算信息显示;1-显示费用信息
    Call ShowLedDisplayBank(1)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化发票信息
    '入参:blnFact-是否刷新发票号
    '编制:刘兴洪
    '日期:2015-01-07 16:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng病人ID As Long, lng主页ID As Long, intInsure As Integer
    
    intInsure = mYBInFor.intInsure
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(NVL(mrsInfo!病人ID)): lng主页ID = Val(NVL(mrsInfo!主页ID))
            intInsure = mYBInFor.intInsure
        End If
    End If
    If mEditType = g_Ed_门诊结帐 Then
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1), lng病人ID, lng主页ID, intInsure, mobjFactProperty, , , 1)
    Else
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), lng病人ID, lng主页ID, intInsure, mobjFactProperty, , , 2)
    End If
    If mobjFactProperty.启用使用类别 Then mlng领用ID = 0
    If blnFact Then Call RefreshFact
    
    If mEditType = g_Ed_门诊结帐 Then
        Call ZlShowBillFormat(mty_ModulePara.bytInvoiceKindMZ, lblFormat, mobjFactProperty.打印格式)
    Else
        Call ZlShowBillFormat(mty_ModulePara.bytInvoiceKindZY, lblFormat, mobjFactProperty.打印格式)
    End If
    picFormat.Visible = lblFormat.Visible
End Sub

Private Function CheckDepositFactValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取预交发票号
    '返回:正常获取,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-30 11:14:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng领用ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    
    On Error GoTo errHandle
    mlng预交领用ID = 0
    
    mstrDepositInvioce = "": mblnDepositBillPrint = False

    '不存在找补
    CheckDepositFactValied = True: Exit Function
    
    If mobjInvoice.zlGetInvoicePreperty(mlngModul, EM_预交收据, mPatiInfor.lng病人ID, mPatiInfor.lng主页ID, 0, mobjDepositFactProperty, , objCard.接口序号 = 2) = False Then Exit Function
    
    Select Case mty_ModulePara.byt预交票据打印
    Case 0 '不打印
        CheckDepositFactValied = True: Exit Function
    Case 1 '自动打印
        mblnDepositBillPrint = True
    Case 2 '选择是否打印
        If MsgBox("是否打印预交票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) <> vbYes Then CheckDepositFactValied = True: Exit Function
        mblnDepositBillPrint = True
    End Select
    
    If mobjDepositFactProperty.严格控制 = False Then
        '有可能是第一次使用
        Do
            blnInput = False
            '非严格控制时直接从本地读取
            strInvoice = UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModul, ""))
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("没有找到已用的预交票据的最大票据号码，无法确定将要使用的开始票据号。" & _
                                vbCrLf & "请输入将要使用的预交票据的开始票据号码：", gstrSysName, _
                                "", Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                strInvoice = UCase(InputBox("请确认使用的预交票据的开始票据号码：", gstrSysName, _
                                strInvoice, Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            End If
                
            '用户取消输入,允许打印
            If strInvoice = "" Then
                If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '检查输入有效性
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> mobjDepositFactProperty.票号长度 Then
                        MsgBox "输入预交的票据号码长度应该为 " & mobjDepositFactProperty.票号长度 & " 位！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        mstrDepositInvioce = strInvoice
        CheckDepositFactValied = True: Exit Function
    End If
    
    Do
        '根据票据领用读取
        blnInput = False
        If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.姓名, EM_预交收据, _
            mobjDepositFactProperty.使用类别, lng领用ID, mobjDepositFactProperty.共享批次ID, lng领用ID, 1, strInvoice) = False Then Exit Function
        If lng领用ID <= 0 Then
            Select Case lng领用ID
                Case 0 '操作失败
                Case -1
                    If Trim(mobjDepositFactProperty.使用类别) = "" Then
                        MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Else
                        MsgBox "你没有自用和共用的『" & mobjFactProperty.使用类别 & "』预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    End If
                    Exit Function
                Case -2
                    If Trim(mobjFactProperty.使用类别) = "" Then
                        MsgBox "本地的共用预交票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    Else
                        MsgBox "本地的共用预交票据的『" & mobjFactProperty.使用类别 & "』预交票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    End If
                    Exit Function
                    strInvoice = ""
            End Select
        End If
        If Not mobjInvoice.zlGetNextBill(mlngModul, lng领用ID, strInvoice) Then Exit Function
        
        If strInvoice = "" Then
            '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
            strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用预交票据的开始票据号，" & _
                            vbCrLf & "请你输入将要使用的票据号码：", gstrSysName, _
                            "", Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        Else
            strInvoice = UCase(InputBox("请确认使用使用预交票据的票据号码：", gstrSysName, _
                            strInvoice, Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        End If
        
        '用户取消输入,不打印
        If strInvoice = "" Then Exit Function
        
        '检查输入有效性
        If blnInput Then
            If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.姓名, EM_预交收据, _
                     mobjDepositFactProperty.使用类别, lng领用ID, mobjDepositFactProperty.共享批次ID, lng领用ID, 1, strInvoice) = False Then Exit Function
            If lng领用ID < 0 Then
                MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
            Else
                blnValid = True
            End If
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    mstrDepositInvioce = strInvoice
    mlng预交领用ID = lng领用ID
    CheckDepositFactValied = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新结帐的票据号
    '编制:刘兴洪
    '日期:2015-01-07 17:16:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjFactProperty Is Nothing Then Exit Sub
    If mobjFactProperty.打印方式 = 0 Then Exit Sub
      
    If Not mobjFactProperty.严格控制 Then
        '非严格控制下
        '松散：取下一个号码
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    
    If zlGetInvoiceGroupUseID(mlng领用ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '严格：取下一个号码
    If mobjInvoice.zlGetNextBill(mlngModul, mlng领用ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    
    'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
    '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
    '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    
    If mobjFactProperty.启用使用类别 Then Call zlCheckFactIsEnough
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Function zlGetRedGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.姓名, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), _
        mobjRedProperty.使用类别, lng领用ID, mobjFactProperty.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng领用ID > 0 Then zlGetRedGroupUseID = True: Exit Function
    
    Select Case lng领用ID
        Case 0 '操作失败
        Case -1
            If Trim(mobjRedProperty.使用类别) = "" Then
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & mobjRedProperty.使用类别 & "』结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjRedProperty.使用类别) = "" Then
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & mobjRedProperty.使用类别 & "』结帐票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If mEditType = g_Ed_门诊结帐 Then
        If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.姓名, IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1), _
            mobjFactProperty.使用类别, lng领用ID, mobjFactProperty.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    Else
        If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.姓名, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), _
            mobjFactProperty.使用类别, lng领用ID, mobjFactProperty.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    End If
    If lng领用ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng领用ID
        Case 0 '操作失败
        Case -1
            If Trim(mobjFactProperty.使用类别) = "" Then
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & mobjFactProperty.使用类别 & "』结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFactProperty.使用类别) = "" Then
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & mobjFactProperty.使用类别 & "』结帐票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 


Private Sub zlCheckFactIsEnough(Optional ByVal intInvoicePages As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    '入参:intInvoicePages-需要的发票张数,如果为0,按系统参数提醒
    '编制:刘兴洪
    '日期:2015-01-07 18:21:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long, lngNums As Long
    Dim bytKind As Byte
    If mEditType = g_Ed_单据查看 Or mEditType = g_Ed_取消结帐 Or mEditType = g_Ed_重新作废 Then Exit Sub
    If mEditType = g_Ed_门诊结帐 Then
        bytKind = IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1)
    Else
        bytKind = IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1)
    End If
    If intInvoicePages <> 0 Then
        If mobjInvoice.zlCheckInvoiceOverplusEnough(bytKind, intInvoicePages, lng剩余数量, mlng领用ID, mobjFactProperty.使用类别) = False Then
            MsgBox "注意:" & vbCrLf & _
                   "    当前剩余票据不足(" & lng剩余数量 & ") ,当前需要" & intInvoicePages & "张票据,请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    Else
        If mobjInvoice.zlCheckInvoiceOverplusEnough(bytKind, mty_ModulePara.int提醒剩余票据张数, lng剩余数量, mlng领用ID, mobjFactProperty.使用类别) = False Then
            MsgBox "注意:" & vbCrLf & _
                   "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & mty_ModulePara.int提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    End If
End Sub


Public Function Chk病人审核(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：判断病人是否已审核
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Nvl(审核标志,0) as 审核标志" & _
        " From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    '49501
    If gTy_System_Para.byt病人审核方式 = 0 Then
        Chk病人审核 = (rsTmp!审核标志 >= 1)
    Else
        Chk病人审核 = (rsTmp!审核标志 > 1)
    End If

    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Led欢迎信息()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:Led初始化
    '编制:刘兴洪
    '日期:2015-01-08 10:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mEditType = g_Ed_单据查看 Or Not gblnLED Then Exit Sub
    If mty_ModulePara.blnLedWelcome Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.编号 & "号 为您服务", mlngModul, gcnOracle
    End If
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    zl9LedVoice.DisplayPatient txtPatient.Text & " " & txtSex.Text & " " & txtOld.Text, Val("" & mrsInfo!病人ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitLed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Led
    '编制:刘兴洪
    '日期:2015-01-08 14:28:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mEditType = g_Ed_单据查看 Or Not gblnLED Then Exit Sub
    zl9LedVoice.Reset com
    zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
End Sub

Private Function GetPatiState(lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人状态说明
    '入参:lng病人ID-病人ID
    '出参:
    '返回:返回病人状态说明
    '     普通在院,留观在院,医保在院;普通出院,留观出院,医保出院;门诊普通,门诊留观,门诊医保
    '编制:刘兴洪
    '日期:2015-01-08 10:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng主页ID As Long, str说明 As String
    
    On Error GoTo errHandle
     
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    If mrsInfo.EOF Then Exit Function
    
    If mrsInfo!主页ID = 0 Or mYBInFor.bytMCMode = 1 Then
        If mYBInFor.intInsure = 0 Then
            GetPatiState = "门诊普通"
        Else
            GetPatiState = "门诊医保"
        End If
        Exit Function
    End If
    
    If NVL(mrsInfo!病人性质, 0) = 1 Then
        str说明 = "门诊留观"
        If NVL(mrsInfo!状态, 0) = 3 Then
            str说明 = str说明 & "(预出院)"
        End If
        GetPatiState = str说明
        Exit Function
    End If

    If mYBInFor.intInsure <> 0 Then
        str说明 = "医保"
    ElseIf NVL(mrsInfo!病人性质, 0) = 2 Then
        str说明 = "留观"
    Else
        str说明 = "普通"
    End If
    
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then
        If Is门诊留观(mrsInfo!病人ID, lng主页ID) Then
            str说明 = "门诊留观"
        Else
            str说明 = "门诊" & str说明
        End If
    Else
        If IsNull(mrsInfo!出院日期) Then
            str说明 = str说明 & "在院"
        Else
            str说明 = str说明 & "出院"
        End If
    End If
    
    If NVL(mrsInfo!状态, 0) = 3 Then
        str说明 = str说明 & "(预出院)"
    End If
    
    GetPatiState = str说明
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Function Is门诊留观(ByVal lng病人ID As Long, ByRef lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前费用是否在门诊留观病人费用期间
    '入参:lng病人ID
    '出参:lng主页ID-返回当前病人ID(第几次留观的)
    '返回:是门诊留观,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 11:23:41
    '问题:45302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String, dtStartDate As Date, dtEndDate As Date
    Dim str时间 As String, strCond As String, rsTemp As ADODB.Recordset
    str时间 = IIf(mty_ModulePara.int费用时间 = 0, "A.登记时间", "A.发生时间")
    strCond = "": dtStartDate = CDate("1901-01-01"): dtEndDate = dtStartDate
    If Not mobjBalanceCon.dtBeginDate = CDate("0:00:00") Then
        strCond = " " & str时间 & " Between [3] And [4]"
        dtStartDate = CDate(Format(mobjBalanceCon.dtBeginDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(mobjBalanceCon.dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
    
    gstrSQL = "" & _
    "Select A.主页id " & _
    "   From 病案主页 A, " & _
    "        (Select Min(" & str时间 & ") As 最小费用时间, Max(" & str时间 & " ) 最大费用时间 " & _
    "          From 门诊费用记录 A " & _
    "          Where  病人id = [1] " & strCond & ") B " & _
    "   Where A.病人id = [1] And A.病人性质 = 1  " & _
    "       And (B.最小费用时间 Between A.入院日期 And Nvl(A.出院日期, Sysdate) Or " & _
    "                B.最大费用时间 Between A.入院日期 And Nvl(A.出院日期, Sysdate) Or " & _
    "                A.入院日期 Between B.最小费用时间 And B.最大费用时间 Or " & _
    "                Nvl(A.出院日期, Sysdate) Between B.最小费用时间 And B.最大费用时间)" & _
    "   Order by 主页ID Desc"
    
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, dtStartDate, dtEndDate)
    If rsTemp.EOF Then rsTemp.Close: Set rsTemp = Nothing: Exit Function
    lng主页ID = Val(NVL(rsTemp!主页ID))
    rsTemp.Close: Set rsTemp = Nothing
    Is门诊留观 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitOldOneCardInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化老一卡通信息
    '编制:刘兴洪
    '日期:2015-01-08 12:02:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mEditType = g_Ed_单据查看 Then Exit Sub
    Set mOldOneCard.rsOneCard = GetOneCard
    With mOldOneCard
        .blnOneCard = .rsOneCard.RecordCount > 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function Init结算方式() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算信息
    '编制:刘兴洪
    '日期:2015-01-08 12:06:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim objCards As Cards, objCard As Card
    Dim objPayCards As Cards, i As Long
    Dim blnOnlyDeposit As Boolean
    
    On Error GoTo errHandle
    
    If mEditType = g_Ed_单据查看 Then Init结算方式 = True: Exit Function
    
    Set objCards = New Cards: Set objPayCards = New Cards
    '性质:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项, _
    '     6-费用折扣,7-一卡通结算,8-结算卡结算
    
    If InStr(1, mstrPrivs, ";费用打折结算;") = 0 Then
        strTmp = "1,2,3,4,5,9,7,8"
    Else
        strTmp = "1,2,3,4,5,6,9,7,8"
    End If
    
    If InStr(1, mstrPrivs, ";允许现金结帐;") = 0 Then
        blnOnlyDeposit = True
    End If
    
    Set mrs结算方式 = Get结算方式("结帐", strTmp)
    If mrs结算方式.RecordCount = 0 Then
        MsgBox "结帐场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrPayMode = ""
     
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare Is Nothing Then
        '0-所有医疗卡;1-启用的医疗卡,2-所有存在三方账户的三方卡 3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)   '获取有效的三方帐户会付
    End If

    If blnOnlyDeposit Then
        mrs结算方式.Filter = "性质=3 Or 性质=4"
    Else
        mrs结算方式.Filter = "性质<7"
    End If
    
    With mrs结算方式
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
        Do While Not .EOF
            If (InStr(",3,4,", "," & Val(NVL(!性质)) & ",") = 0) And Val(NVL(!应付款)) <> 1 Then
                Set objCard = New Card
                objCard.接口序号 = -1 * i
                objCard.接口编码 = !编码
                objCard.名称 = !名称
                objCard.结算方式 = !名称
                objCard.结算性质 = Val(NVL(!性质))
                objCard.启用 = True
                objCard.是否刷卡 = 1
                objCard.缺省标志 = Val(NVL(!缺省)) = 1
                objPayCards.Add objCard
                mstrPayMode = mstrPayMode & "|" & !名称
                If objCard.缺省标志 Then
                    If Val(!性质) = 1 Then
                        '新增现金结算
                        If vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算方式")) <> "" Then
                            vsBlance.Rows = vsBlance.Rows + 1
                        End If
                        
                        vsBlance.RowData(vsBlance.Rows - 1) = "999"
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("类型")) = 0
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算性质")) = 1
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("编辑状态")) = 1   '0-禁止删除;1-允许编辑金额;2-允许删除
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算方式")) = NVL(!名称)
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算金额")) = "0.00"
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算号码")) = ""
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("备注")) = ""
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("交易流水号")) = ""
                        vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("交易说明")) = ""
                        vsBlance.Cell(flexcpFontBold, vsBlance.Rows - 1, 0, vsBlance.Rows - 1, vsBlance.Cols - 1) = True
                    End If
                    mstr缺省结算方式 = objCard.结算方式
                End If
                i = i + 1
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    If vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算方式")) <> "" Then
        vsBlance.Rows = vsBlance.Rows + 1
    End If
    
    If InStr(";" & mstrPrivsCard & ";", ";三方接口消费;") > 0 Then
        mrs结算方式.Filter = "性质>=7 and 性质<9" '一卡通结算
        With mrs结算方式
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For Each objCard In objCards
                    If objCard.结算方式 = NVL(!名称) Then
                        '找到了,增加
                        '85565,李南春,2015/7/19:读卡性质
                        objCard.是否刷卡 = True
                        objCard.缺省标志 = Val(NVL(!缺省)) = 1
                        objCard.结算性质 = Val(NVL(!性质))
                        objPayCards.Add objCard
                        mstrPayMode = mstrPayMode & "|" & !名称
                        If objCard.缺省标志 Then
                            mstr缺省结算方式 = objCard.结算方式
                        End If
                        Exit For
                    End If
                Next
                .MoveNext
            Loop
            .Filter = 0
        End With
    End If
    
    mrs结算方式.Filter = 0
    mblnNotChange = True
    Set mobjPayCards = objPayCards
    If objPayCards.Count = 0 And blnOnlyDeposit = False Then
        mblnNotChange = True
        MsgBox "结帐场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnload = True: Exit Function
    End If
    mblnNotChange = False
    Init结算方式 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Load找补项(ByVal bytFun As Byte, ByVal str找补项名 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载找补项
    '入参:bytFun-0-只有找补;1-含存预交
    '编制:刘兴洪
    '日期:2015-01-09 15:13:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, objCard As Card
     
    On Error GoTo errHandle
    If mEditType = g_Ed_单据查看 Then Exit Sub
    
    On Error GoTo errHandle
    
    Set objCards = New Cards
    Set objCard = New Card
    objCard.接口序号 = 1
    objCard.接口编码 = 1
    objCard.名称 = IIf(str找补项名 = "", "找补", str找补项名)
    objCard.结算方式 = objCard.名称
    objCard.结算性质 = 0
    objCard.启用 = True
    '85565,李南春,2015/7/10:读卡性质
    objCard.是否刷卡 = True
    objCards.Add objCard
    If bytFun <> 0 Then
        Set objCard = New Card
        objCard.接口序号 = 2
        objCard.接口编码 = 2
        objCard.名称 = "门诊预交"
        objCard.结算方式 = objCard.名称
        objCard.结算性质 = 0
        objCard.启用 = True
        '85565,李南春,2015/7/10:读卡性质
        objCard.是否刷卡 = True
        objCards.Add objCard
        
        Set objCard = New Card
        objCard.接口序号 = 3
        objCard.接口编码 = 3
        objCard.名称 = "住院预交"
        objCard.结算方式 = objCard.名称
        objCard.结算性质 = 0
        objCard.启用 = True
        '85565,李南春,2015/7/10:读卡性质
        objCard.是否刷卡 = True
        
        objCards.Add objCard
    End If
    mblnNotChange = True
    
    mblnNotChange = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub

Private Function LoadBalanceBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结帐单的相关信息
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 14:30:45
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Then
        '执行结帐
        Call NewBill
        If mlngPatientID <> 0 Then
            txtPatient.Text = "-" & mlngPatientID
            mobjBalanceCon.strTime = mstr主页Id
            Call txtPatient_KeyPress(vbKeyReturn)
            If Val(mstr主页Id) = "0" Then cmdYB.Enabled = True
            If mrsInfo Is Nothing Then mblnUnload = True: Exit Function
            If mrsInfo.State = 0 Then mblnUnload = True: Exit Function
        End If
        Me.Caption = IIf(mEditType = g_Ed_门诊结帐, "门诊病人结帐单", "住院病人结帐单")
        LoadBalanceBill = True: Exit Function
    End If
    
    Select Case mEditType
    Case g_Ed_取消结帐, g_Ed_结帐作废, g_Ed_重新作废
        mblnNotChange = True
        chkCancel.Value = 1
        mblnNotChange = False
    Case Else
    End Select

    If Not ReadBalance(mstrInNO) Then mblnUnload = True: Exit Function
    LoadBalanceBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadBalancePayData(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    Optional ByVal blnNOMoved As Boolean = False, Optional bln原结帐 As Boolean, Optional blnInsure As Boolean, _
    Optional ByVal intCustomSign As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载已经支付的数据
    '入参:lng结帐ID-结帐ID
    '     blnNOMoved-是否已经转入后备表
    '     bln原结帐-读取的是原结帐数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 15:24:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strSQL As String
    Dim dblTotal As Double, strTable As String, blnYB As Boolean
    Dim strCardNo As String, cllBillPro As New Collection
    Dim objCard As Card, bytEdit As Byte
    Dim lng卡类别ID  As Long, dblMoney As Double
    Dim TyBrushCardInor As TY_BrushCard
    Dim blnAdd As Boolean, intYBpara As Integer
    Dim byt结算状态 As Byte
    Dim dbl医保基金 As Double
    Dim intSign As Integer
    Dim blnUnload As Boolean
    Dim blnNoPrepay As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim j As Long
    Dim blnCheck As Boolean
    
    On Error GoTo errHandle
     
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    If zlGetFromIDToBalanceData(lng结帐ID, blnNOMoved, mrsBalance) = False Then Exit Function
    
    With mrsBalance
        i = 1: blnYB = False
'        vsBlance.Clear 1: vsBlance.Rows = 2
        mBalanceInfor.dbl已付合计 = 0
        mBalanceInfor.dbl医保支付合计 = 0
        If Not mEditType = g_Ed_重新作废 And mblnInsure = False Then
            mBalanceInfor.dbl冲预交合计 = 0
        End If
        If intCustomSign <> 0 Then
            intSign = intCustomSign
        Else
            intSign = IIf(mEditType = g_Ed_重新作废, -1, 1)
        End If
        
        Do While Not .EOF
            dblMoney = RoundEx(intSign * Val(NVL(!冲预交)), 6)
            blnAdd = True
            Select Case NVL(!类型)
            Case 1 '预交款
                If Not mEditType = g_Ed_重新作废 Then
                    '重新作废在加载预交款时,已经赋值,原因是要找原始结帐时的冲预交
                    mBalanceInfor.dbl冲预交合计 = RoundEx(mBalanceInfor.dbl冲预交合计 + dblMoney, 6)
                End If
            Case 2, 3, 5 '医保,一卡通,消费卡
                blnAdd = True
                If NVL(!类型) = 2 Then
                    If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_取消结帐 Or bln原结帐 Or chkCancel.Value = 1 Then
                        Select Case Val(NVL(mrsBalance!性质))
                        Case 3   '个人帐户
                            If mYBInFor.bytMCMode = 1 And Not MCPAR.门诊病人结算作废 Then
                                blnAdd = False
                            Else
                                intYBpara = IIf(mYBInFor.bytMCMode = 1, support门诊结算作废, support住院结算作废)
                                blnAdd = gclsInsure.GetCapability(intYBpara, lng病人ID, mYBInFor.intInsure, NVL(mrsBalance!结算方式))
                            End If
                        Case 4  '医保基金
                            intYBpara = IIf(mYBInFor.bytMCMode = 1, support门诊结算作废, support住院结算作废)
                            blnAdd = gclsInsure.GetCapability(intYBpara, lng病人ID, mYBInFor.intInsure, NVL(mrsBalance!结算方式))
                        End Select
                    End If
                    
                    If blnAdd Then
                        mBalanceInfor.dbl医保支付合计 = RoundEx(mBalanceInfor.dbl医保支付合计 + dblMoney, 6)
                        If Val(NVL(mrsBalance!性质)) = 4 Then
                            dbl医保基金 = dbl医保基金 + dblMoney
                        End If
                        blnYB = True
                    End If
                End If
                
                If Not blnAdd Then GoTo GoAddEnd:
                
                With vsBlance
                    strCardNo = NVL(mrsBalance!卡号)
                    lng卡类别ID = IIf(Val(NVL(mrsBalance!类型)) = 5, Val(NVL(mrsBalance!结算卡序号)), Val(NVL(mrsBalance!卡类别ID)))
                    TyBrushCardInor.str卡号 = strCardNo
                    TyBrushCardInor.str结算号码 = NVL(mrsBalance!结算号码)
                    TyBrushCardInor.str结算摘要 = NVL(mrsBalance!摘要)
                    TyBrushCardInor.str交易流水号 = NVL(mrsBalance!交易流水号)
                    TyBrushCardInor.str交易说明 = NVL(mrsBalance!交易说明)
                    TyBrushCardInor.str扩展信息 = ""
                    If Val(NVL(mrsBalance!校对标志)) = 1 And mEditType <> g_Ed_单据查看 Then
                        Select Case Val(NVL(mrsBalance!类型))
                        Case 3 '3-一卡通
                            If MsgBox("警告:" & vbCrLf & _
                                       "     在使用『" & NVL(mrsBalance!卡类别名称, NVL(mrsBalance!结算方式, "")) & "』进行支付结算时失败,是否继续进行支付?" & vbCrLf & _
                                       "结算金额:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                '调用一卡通支付接口
                                Set objCard = IDKindPaymentsType.GetIDKindCard(lng卡类别ID, CardTypeID)
                                If objCard Is Nothing Then
                                    MsgBox "当前站点未启用:" & NVL(mrsBalance!卡类别名称, NVL(mrsBalance!结算方式, "")) & ",请在『结算方式管理』或本地参数的设备配置中设置启用!", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                End If
                                '先检查是否合法
                                If CheckThreeSwapValied(objCard, dblMoney, TyBrushCardInor) = False Then Exit Function
                                If ExecuteThreeSwapPayInterface(lng病人ID, lng结帐ID, objCard, dblMoney, cllBillPro, TyBrushCardInor) = False Then Exit Function
                                byt结算状态 = 1
                            Else
                                Exit Function
                            End If
                        Case 4 '4-一卡通(老)
                            
                            If MsgBox("警告:" & vbCrLf & _
                                       "     在使用『" & NVL(mrsBalance!卡类别名称, NVL(mrsBalance!结算方式, "")) & "』进行支付结算时失败,是否继续进行支付?" & vbCrLf & _
                                       "结算金额:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                '调用一卡通支付(老版)接口
                                Set objCard = GetOldCard(mrsBalance!结算方式)
                                If objCard Is Nothing Then
                                    MsgBox "当前站点未启用:" & NVL(mrsBalance!卡类别名称, NVL(mrsBalance!结算方式, "")) & ",请在『基础参数设置』中设置启用!", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                End If
                                '1.先检查是否合法
                                If CheckOldOneCardIsValied(objCard, dblMoney, TyBrushCardInor) = False Then Exit Function
                                '2.调用支付
                                If ExecuteOldOneCardPayInterface(lng病人ID, lng结帐ID, objCard, dblMoney, TyBrushCardInor, cllBillPro) = False Then Exit Function
                                byt结算状态 = 1
                            Else
                                Exit Function
                            End If
                        End Select
                    End If
                    
                    blnNoPrepay = False
                    blnUnload = False
                    If Val(NVL(mrsBalance!类型)) = 3 Then
                        If mEditType = g_Ed_结帐作废 Or chkCancel.Value = 1 Then
                            strSQL = "Select 1 From 三方退款信息 Where 结帐ID=[1]"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
                            If rsTmp.EOF Then
                                blnUnload = False
                                blnNoPrepay = False
                            Else
                                blnNoPrepay = True
                                blnUnload = True
                            End If
                        End If
                    End If
                    
                    If blnUnload = False Then
                        If blnYB Then
                            blnYB = False
                            For j = 1 To .Rows - 1
                                If .TextMatrix(j, .ColIndex("结算方式")) = NVL(mrsBalance!结算方式) Then
                                    i = j
                                    blnYB = True
                                End If
                            Next j
                        End If
                        If .TextMatrix(i, .ColIndex("结算方式")) <> "" And blnYB = False Then
                            If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Then
                                .Rows = .Rows + 1
                            End If
                            i = .Rows - 1
                        End If
                        bytEdit = 0
                        If (mEditType = g_Ed_结帐作废 Or chkCancel.Value = 1) And mEditType <> g_Ed_重新作废 And mEditType <> g_Ed_取消结帐 Then
                            If Val(NVL(mrsBalance!类型)) = 3 And Val(NVL(mrsBalance!是否退现)) = 1 Then    '一卡通
                                bytEdit = 2
                            End If
                            If Val(NVL(mrsBalance!类型)) = 5 And Val(NVL(mrsBalance!是否退现)) = 1 Then bytEdit = 2
                        End If
                        If byt结算状态 <> 1 Then
                            If mEditType = g_Ed_结帐作废 Or chkCancel.Value = 1 Then
                                If mEditType = g_Ed_重新作废 Then
                                    byt结算状态 = IIf(Val(NVL(mrsBalance!校对标志)) = 1, 0, 1)
                                Else
                                    byt结算状态 = 0
                                End If
                            Else
                                byt结算状态 = 1
                            End If
                        End If
                        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                        .TextMatrix(i, .ColIndex("类型")) = Val(NVL(mrsBalance!类型))
                        .TextMatrix(i, .ColIndex("卡类别ID")) = lng卡类别ID
                        .TextMatrix(i, .ColIndex("消费卡ID")) = Val(NVL(mrsBalance!消费卡ID))
                        .TextMatrix(i, .ColIndex("结算性质")) = Val(NVL(mrsBalance!性质))
                        .TextMatrix(i, .ColIndex("编辑状态")) = bytEdit   '0-禁止删除;1-允许编辑金额;2-允许删除
                        .TextMatrix(i, .ColIndex("结算状态")) = byt结算状态  '是否已结算:1-已结算;0-未结算
                        .TextMatrix(i, .ColIndex("是否退现")) = Val(NVL(mrsBalance!是否退现))
                        .TextMatrix(i, .ColIndex("是否全退")) = Val(NVL(mrsBalance!是否全退))
                        .TextMatrix(i, .ColIndex("校对标志")) = Val(NVL(mrsBalance!校对标志))
                        .TextMatrix(i, .ColIndex("是否密文")) = Val(NVL(mrsBalance!是否密文))
                        .TextMatrix(i, .ColIndex("卡类别名称")) = Trim(NVL(mrsBalance!卡类别名称))
                        .TextMatrix(i, .ColIndex("结算方式")) = NVL(mrsBalance!结算方式)
                        .TextMatrix(i, .ColIndex("结算金额")) = Format(dblMoney, gstrDec)
                        .TextMatrix(i, .ColIndex("结算号码")) = NVL(mrsBalance!结算号码)
                        .TextMatrix(i, .ColIndex("备注")) = NVL(mrsBalance!摘要)
                        .TextMatrix(i, .ColIndex("交易流水号")) = NVL(mrsBalance!交易流水号)
                        .TextMatrix(i, .ColIndex("交易说明")) = NVL(mrsBalance!交易说明)
                        .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(NVL(mrsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("卡号")) = NVL(mrsBalance!卡号)
                        
                        If mEditType = g_Ed_单据查看 Then
                            If Val(NVL(mrsBalance!校对标志)) = 1 Then    '未执行成功的
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                            ElseIf Val(NVL(mrsBalance!校对标志)) = 2 Then '执行成功且当前处于查看的
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                            Else
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = Me.ForeColor
                            End If
                        End If
                        
                        mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 + dblMoney, 6)
                    End If
                End With
GoAddEnd:
        Case Else '0-普通结算
            
            If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新结帐 Or chkCancel.Value = 1 Or bln原结帐 Then
                '只能缺省为收款
                If Val(NVL(!性质)) = 1 Then blnAdd = False
            End If
            With vsBlance
                If NVL(mrsBalance!结算方式) <> "" And (NVL(mrsBalance!类型) <> 6 Or mEditType = g_Ed_单据查看) And blnAdd Then
                    blnCheck = False
                    For j = 1 To .Rows - 1
                        If .TextMatrix(j, .ColIndex("结算方式")) = NVL(mrsBalance!结算方式) Then
                            i = j
                            blnCheck = True
                        End If
                    Next j
                     If .TextMatrix(i, .ColIndex("结算方式")) <> "" And NVL(mrsBalance!结算方式) <> "" And blnCheck = False Then
                         .Rows = .Rows + 1
                         i = .Rows - 1
                     End If
                     bytEdit = 0
                     If mEditType = g_Ed_取消结帐 Or mEditType = g_Ed_结帐作废 Or chkCancel.Value = 1 Then bytEdit = 2
                    
                     If mEditType = g_Ed_结帐作废 Or chkCancel.Value = 1 Then
                         byt结算状态 = 0
                     Else
                         byt结算状态 = 1
                     End If
                     '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                     .TextMatrix(i, .ColIndex("类型")) = Val(NVL(mrsBalance!类型))
                     .TextMatrix(i, .ColIndex("卡类别ID")) = lng卡类别ID
                     .TextMatrix(i, .ColIndex("消费卡ID")) = Val(NVL(mrsBalance!消费卡ID))
                     .TextMatrix(i, .ColIndex("结算性质")) = Val(NVL(mrsBalance!性质))
                     .TextMatrix(i, .ColIndex("编辑状态")) = bytEdit   '0-禁止删除;1-允许编辑金额;2-允许删除
                     .TextMatrix(i, .ColIndex("结算状态")) = byt结算状态  '是否已结算:1-已结算;0-未结算
                     .TextMatrix(i, .ColIndex("是否退现")) = Val(NVL(mrsBalance!是否退现))
                     .TextMatrix(i, .ColIndex("是否全退")) = Val(NVL(mrsBalance!是否全退))
                     .TextMatrix(i, .ColIndex("校对标志")) = Val(NVL(mrsBalance!校对标志))
                     .TextMatrix(i, .ColIndex("是否密文")) = Val(NVL(mrsBalance!是否密文))
                     .TextMatrix(i, .ColIndex("卡类别名称")) = Trim(NVL(mrsBalance!卡类别名称))
                     
                     .TextMatrix(i, .ColIndex("结算方式")) = NVL(mrsBalance!结算方式)
                     .TextMatrix(i, .ColIndex("结算金额")) = Format(intSign * Val(NVL(mrsBalance!冲预交)), gstrDec)
                     .TextMatrix(i, .ColIndex("结算号码")) = NVL(mrsBalance!结算号码)
                     .TextMatrix(i, .ColIndex("备注")) = NVL(mrsBalance!摘要)
                     .TextMatrix(i, .ColIndex("交易流水号")) = NVL(mrsBalance!交易流水号)
                     .TextMatrix(i, .ColIndex("交易说明")) = NVL(mrsBalance!交易说明)
                     .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(NVL(mrsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                     .Cell(flexcpData, i, .ColIndex("卡号")) = NVL(mrsBalance!卡号)
                     mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 + intSign * Val(NVL(mrsBalance!冲预交)), 6)
                 End If
            End With
        End Select
        .MoveNext
        Loop
    End With
    
    If mEditType = g_Ed_重新作废 Then
        strSQL = "Select 1 From 三方退款信息 Where 结帐ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBalanceInfor.lng结帐ID)
        If rsTmp.EOF Then
            blnNoPrepay = False
        Else
            blnNoPrepay = True
        End If
    End If
    
    If blnNoPrepay Then
        mBalanceInfor.dbl冲预交合计 = 0
        chkDeposit.Enabled = False
    End If
    
    If vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算方式")) = "" Then
        vsBlance.RemoveItem vsBlance.Rows - 1
    End If
    
    mrsBalance.Filter = "性质 = 3 Or 性质 = 4"
    If mrsBalance.EOF Then
        blnCheck = True
        Do While blnCheck = True
            blnCheck = False
            For i = 1 To vsBlance.Rows - 1
                If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("结算性质"))) = 3 Or Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("结算性质"))) = 4 Then
                    mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("结算金额"))), 6)
                    vsBlance.RemoveItem i
                    blnCheck = True
                    Exit For
                End If
            Next i
        Loop
    End If
    mrsBalance.Filter = ""

    mblnNotChange = True
    txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
    chkDeposit.Tag = mBalanceInfor.dbl冲预交合计
    chkDeposit.Value = 0
    If mBalanceInfor.dbl冲预交合计 <> 0 Then chkDeposit.Value = 1
    mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl已付合计, 5)
    
    If vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("结算方式")) <> "" Then vsBlance.Rows = vsBlance.Rows + 1
    
    mblnNotChange = False
    LoadBalancePayData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetOldCard(ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算方式,获取老一卡通的卡对象
    '编制:刘兴洪
    '日期:2015-01-08 18:05:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCards As Cards
    
    Set objCards = IDKindPaymentsType.Cards
    For Each objCard In objCards
        If objCard.结算方式 = str结算方式 And objCard.结算性质 = 7 Then
            GetOldCard = objCard: Exit Function
        End If
    Next
    Set GetOldCard = Nothing
End Function

Private Sub ClearCustomType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除自定义类型相关变量
    '编制:刘兴洪
    '日期:2015-01-26 17:16:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyBalanceCons As clsBalanceAllCon
    Dim tyBalanceInfor As TY_Balance_Infor
    Dim tyYBInFor As TY_YBInfor, tyPatiInfor As ty_Pati_Infor
    
    On Error GoTo errHandle
        
    mPatiInfor = tyPatiInfor
    Set mobjBalanceCon = New clsBalanceCon    '初始化条件
    Set mobjBalanceAll = New clsBalanceAllCon
    mBalanceInfor = tyBalanceInfor
    mYBInFor = tyYBInFor
    mPatiInfor = tyPatiInfor '清空病人信息
    '老版一卡通
    With mOldOneCard
        .strOneCard = ""
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ReadBalance(strNO As String, Optional blnInputNo As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查看或作废时,读取并显示结帐单
    '入参:strNo-结帐单号号
    '     blnInputNo-输入单据号进行作废
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 14:43:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strFullNO As String, lng结帐ID As Long
    Dim lngID As Long, i As Long, j As Long, lngDefault As Long
    Dim strSQL As String, dMax As Date, dMin As Date, blnUndo As Boolean
    Dim curTmp As Currency, curMoney As Currency, curDeposit As Currency
    Dim lngMaxLength As Long, lngP As Long, lng病人ID As Long
    Dim rsUnit As ADODB.Recordset, rsFee As New ADODB.Recordset
    Dim strTable As String, lng主页ID As Long
    Dim str主页Ids As String, rsTmp As ADODB.Recordset
    Dim strOper As String, vDate As Date
    
    On Error GoTo errH
    Call ClearCustomType
    
    '单据主体
    strFullNO = GetFullNO(strNO, 15)
     
    strSQL = "" & _
    "   Select A.ID,A.实际票号,A.病人ID,B.门诊号,B.住院号,b.当前床号,B.当前科室ID,B.费别,B.姓名,B.性别,B.年龄, " & _
    "          A.收费时间,A.开始日期,A.结束日期,A.备注,A.原因,A.结算状态,A.结帐类型,A.住院次数,A.结帐金额, " & _
    "          nvl(A.主页ID,nvl(B.主页ID,0)) as 主页ID,B.在院,nvl(A.中途结帐,0) as 中途结帐,A.记录状态" & _
    "   From 病人结帐记录 A,病人信息 B" & _
    "   Where A.病人ID=B.病人ID(+) " & _
    "       And A.NO=[1] And A.记录状态 " & IIf(mblnViewCancel, "= 2", "In(1,3)")
    
    If mblnNOMoved Then strSQL = Replace(strSQL, "病人结帐记录", "H病人结帐记录")
    
    strSQL = _
    "Select A.ID,A.实际票号 as 票据号,A.病人ID,A.门诊号, " & _
    "       nvl(D.住院号,A.住院号) as 住院号, Nvl(D.出院病床,A.当前床号)  as 当前床号, " & _
    "       Nvl(E.名称,C.名称) as 当前科室,A.在院," & _
    "       Nvl(D.费别,A.费别) as 费别,nvl(D.姓名,A.姓名) as 姓名,nvl(D.性别,A.性别) as 性别,nvl(D.年龄,A.年龄) as 年龄, " & _
    "       A.收费时间,A.开始日期,A.结束日期,A.备注,A.原因,A.结算状态,A.结帐类型,A.住院次数,A.结帐金额,A.主页ID,A.中途结帐,A.记录状态" & _
    " From (" & strSQL & ") A,部门表 C,病案主页 D,部门表 E" & _
    " Where  A.当前科室ID=C.ID(+) And D.出院科室ID=E.ID(+)" & _
    "       And A.病人ID=D.病人ID(+) And A.主页ID =D.主页ID(+) "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO)
    If rsTemp.EOF Then
        MsgBox "没有发现该结帐单据,可能已经作废！", vbInformation, gstrSysName
        Exit Function
    End If
    If blnInputNo = True And Val(NVL(rsTemp!记录状态)) <> 1 Then
        MsgBox "该结帐单据为已经结帐作废，不能再结帐作废操作！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not GetMinMaxDate(rsTemp!ID, dMin, dMax, mblnNOMoved) Then
        MsgBox "该结帐单据内容不正确，没有发现结帐的费用明细！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mEditType = g_Ed_重新作废 And Val(NVL(rsTemp!结算状态)) <> 1 Then
        MsgBox "该结帐单据不为异常单据，不能重新作废！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mEditType = g_Ed_重新结帐 And Val(NVL(rsTemp!结算状态)) <> 1 Then
        MsgBox "该结帐单据不为异常单据，不能重新结帐！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mEditType = g_Ed_取消结帐 And Val(NVL(rsTemp!结算状态)) <> 1 Then
        MsgBox "该结帐单据不为异常单据，不能取消结帐！", vbInformation, gstrSysName
        Exit Function
    End If
    If mEditType = g_Ed_结帐作废 And Val(NVL(rsTemp!结算状态)) = 1 Then
        MsgBox "该结帐单据为异常单据，不能结帐作废！", vbInformation, gstrSysName
        Exit Function
    End If
    If mEditType = g_Ed_结帐作废 And Val(NVL(rsTemp!记录状态)) <> 1 Then
        MsgBox "该结帐单据为已经结帐作废，不能再结帐作废操作！", vbInformation, gstrSysName
        Exit Function
    End If
       
    
    lng结帐ID = Val(NVL(rsTemp!ID))
    cboNO.Text = strFullNO
    
    If mEditType = g_Ed_结帐作废 Then
        If CheckExistsGathering(cboNO.Text) Then
            MsgBox "该结帐单据存在已缴款的应收款记录，请退款后再执行作废。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    txtInvoice.Text = NVL(rsTemp!票据号)

    lng病人ID = Val(NVL(rsTemp!病人ID))
    lng主页ID = Val(NVL(rsTemp!主页ID))
    
    If mEditType = g_Ed_结帐作废 Then
        '单据权限
        If Not ReadBillInfo(2, cboNO.Text, -1, strOper, vDate) Then
            Exit Function
        End If

        If Not BillOperCheck(7, strOper, vDate, "作废") Then
            Exit Function
        End If
    End If
    
    mobjBalanceAll.strAllTime = NVL(rsTemp!住院次数)
    mblnCurMzBalanceNo = False
    If Val(NVL(rsTemp!结帐类型)) = 0 Then
        Me.Caption = gstrUnitName & "病人结帐单"
        If mobjBalanceAll.strAllTime = "" Then mobjBalanceAll.strAllTime = GetFromalanceIDToPatiNum(lng结帐ID, lng主页ID)
    ElseIf Val(NVL(rsTemp!结帐类型)) = 1 Then
        Me.Caption = gstrUnitName & "门诊病人结帐单"
        mobjBalanceAll.strAllTime = "": lng主页ID = 0
        mblnCurMzBalanceNo = True
    Else
        Me.Caption = gstrUnitName & "住院病人结帐单"
        If mobjBalanceAll.strAllTime = "" Then mobjBalanceAll.strAllTime = GetFromalanceIDToPatiNum(lng结帐ID, lng主页ID)
    End If
    mobjBalanceCon.strTime = mobjBalanceAll.strAllTime
    mBalanceInfor.strNO = strFullNO
    With mPatiInfor
        .lng病人ID = lng病人ID
        .lng主页ID = lng主页ID
        .str姓名 = NVL(rsTemp!姓名)
        .str性别 = NVL(rsTemp!性别)
        .str年龄 = NVL(rsTemp!年龄)
        .bln出院 = Val(NVL((rsTemp!在院))) <> 1
    End With
    
    With mBalanceInfor
        .strNO = strFullNO
        .blnSaveBill = IIf(mEditType = g_Ed_结帐作废 Or blnInputNo, False, True)
        If mblnViewCancel And mEditType <> g_Ed_单据查看 Then
            .lng冲销ID = lng结帐ID
            .lng结帐ID = zlGetFormerBalanceID(mBalanceInfor.strNO)
        Else
            .lng冲销ID = 0
            .lng结帐ID = lng结帐ID
        End If
        .dtBalanceDate = CDate(Format(rsTemp!收费时间, "yyyy-mm-dd hh:MM:SS"))
    End With
    
    If mEditType <> g_Ed_单据查看 Then
        mYBInFor.intInsure = BalanceExistInsure(strNO, mYBInFor.bytMCMode)
        If mYBInFor.intInsure <> 0 Then
            Call InitInsurePara(mPatiInfor.lng病人ID, mYBInFor.intInsure)
        End If
    End If
    
    If mEditType = g_Ed_重新结帐 Or mEditType = g_Ed_取消结帐 Then
        If Val(NVL(rsTemp!结帐类型)) = 0 Then
            If zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") = False And zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") = False Then
                MsgBox "你没有结账权限，不能进行结帐操作！", vbInformation, gstrSysName
                Exit Function
            End If
        ElseIf Val(NVL(rsTemp!结帐类型)) = 1 Then
            If zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") = False Then
                MsgBox "你没有门诊费用结帐权限，不能进行结帐操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") = False Then
                MsgBox "你没有住院费用结帐权限，不能进行结帐操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If mYBInFor.intInsure <> 0 Then
            If zlStr.IsHavePrivs(mstrPrivs, "保险结算") = False Then
                MsgBox "你没有保险结算权限，不能进行结帐操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "普通病人结算") = False Then
                MsgBox "你没有普通病人结算权限，不能进行结帐操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mEditType = g_Ed_重新作废 Or mEditType = g_Ed_结帐作废 Then
        If zlStr.IsHavePrivs(mstrPrivs, "结帐作废") = False Then
            MsgBox "你没有结账作废权限，不能进行结帐作废操作！", vbInformation, gstrSysName
            Exit Function
        End If
        If mYBInFor.intInsure <> 0 Then
            If zlStr.IsHavePrivs(mstrPrivs, "保险结算") = False Then
                MsgBox "你没有保险结算权限，不能进行结帐作废操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "普通病人结算") = False Then
                MsgBox "你没有普通病人结算权限，不能进行结帐作废操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
 
    '加载余额信息
    Call Load余额信息(lng病人ID, Val(NVL(rsTemp!结帐类型)))

    '检查是否合约单位结帐:问题:35090
    If Val(NVL(rsTemp!病人ID)) = 0 Then
        If NVL(rsTemp!原因) <> "" Then
            txtPatient.Text = NVL(rsTemp!原因)
        Else
            strSQL = "" & _
            "   Select  D.名称 " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A, 病人信息 C, 合约单位 D " & _
            "   Where A.结帐ID=[1]  And A.病人ID=C.病人ID And C.合同单位id = D.ID(+) and Rownum=1 " & _
            "    Union ALL " & _
            "   Select  D.名称 " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A, 病人信息 C, 合约单位 D " & _
            "   Where A.结帐ID=[1] And C.合同单位id = D.ID(+) and Rownum=1 " & _
            "   "
            Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(rsTemp!ID)))
            If Not rsUnit.EOF Then
                txtPatient.Text = NVL(rsUnit!名称)
            Else
                txtPatient.Text = "未找到合约单位"
            End If
        End If
        txtPatient.Tag = "合约单位"
    Else
        txtPatient.Text = NVL(rsTemp!姓名)
        txtPatient.Tag = Val(NVL(rsTemp!病人ID))
    End If
     
    txtSex.Text = NVL(rsTemp!性别)
    txtOld.Text = NVL(rsTemp!年龄)
    txt费别.Text = NVL(rsTemp!费别)
    txtDate.Text = Format(rsTemp!收费时间, "yyyy-MM-dd HH:mm:ss")
    txtInvoice.Text = NVL(rsTemp!票据号)
    '问题65105,刘尔旋:结账查阅中新增门诊号码的显示
    mobjBalanceCon.blnCurBalanceOwnerFee = False
    lblBalanceType.Visible = False
    Select Case Val(NVL(rsTemp!结帐类型))
        '10.29以前的类型，不做处理
        Case 0
        Case 1
            txt标识号.Text = NVL(rsTemp!门诊号)
            txt标识号.Visible = True
            lbl标识号.Visible = True
            lbl标识号.Caption = "门诊号"
            lblPatiTime.Visible = False
            txtPatiBegin.Visible = False
            lblPatiTimeRange.Visible = False
            txtPatiEnd.Visible = False
            txt天数.Visible = False
            lblDayName.Visible = False
        Case 2
            txt标识号.Text = NVL(rsTemp!住院号)
            txt标识号.Visible = True
            lbl标识号.Visible = True
            lbl标识号.Caption = "住院号"

            If Not IsNull(rsTemp!当前床号) Then
                txtBed.Text = rsTemp!当前床号
                txtBed.Visible = True
                lblBed.Visible = True
            End If

            If Not IsNull(rsTemp!当前科室) Then
                txt科室.Text = rsTemp!当前科室
                txt科室.Visible = True
                lbl科室.Visible = True
            End If
            opt出院.Value = IIf(Val(NVL(rsTemp!中途结帐)) = 1, False, True)
            opt中途.Value = IIf(Val(NVL(rsTemp!中途结帐)) = 1, True, False)
'           lblBalanceType.Visible = True
            lblBalanceType.Caption = IIf(Val(NVL(rsTemp!中途结帐)) = 1, "中途结帐", "出院结帐")
    End Select

    txtBegin.Text = Format(dMin, txtBegin.Format)
    txtEnd.Text = Format(dMax, txtEnd.Format)
    txtBalance(Idx_结帐说明).Text = NVL(rsTemp!备注)

    If mobjBalanceCon.blnCurBalanceOwnerFee = False Then
        '非门诊结帐时
        If Not IsNull(rsTemp!开始日期) Then
            txtPatiBegin.Text = Format(rsTemp!开始日期, "yyyy-MM-dd")
        End If

        If Not IsNull(rsTemp!结束日期) Then
            txtPatiEnd.Text = Format(rsTemp!结束日期, "yyyy-MM-dd")
        End If
    End If

    lngID = rsTemp!ID
    
    
    str主页Ids = IIf(mty_ModulePara.bln仅用指定预交款 And mbln门诊转住院 = False, _
    IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime), "")
    If Not LoadFeeListFromBalanceID(lngID) Then Exit Function    '加载费用明细
    If Not LoadBalanceDepositList(lng病人ID, lngID, mblnNOMoved, str主页Ids) Then Exit Function  '加载冲预交款
    
    If Not LoadBalancePayData(lng病人ID, lngID, mblnNOMoved) Then Exit Function  '加载已经支付数据
    If mEditType = g_Ed_重新作废 Then
        Dim blnReadOldBalan As Boolean
        
        mrsBalance.Filter = 0
        blnReadOldBalan = mrsBalance.RecordCount = 0
        If mrsBalance.RecordCount = 1 Then
            blnReadOldBalan = NVL(mrsBalance!结算方式) = ""
        End If
        If blnReadOldBalan Then
            If Not LoadBalancePayData(lng病人ID, mBalanceInfor.lng结帐ID, mblnNOMoved, True) Then Exit Function     '加载已经支付数据
        End If
        If zlGetFromIDToBalanceData(mBalanceInfor.lng结帐ID, mblnNOMoved, mrsOldBalance) = False Then Exit Function
           
    End If
    If mEditType = g_Ed_重新结帐 Then
        '  dblMoney = RoundEX(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl医保支付合计, 2)
        '操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按指定金额来冲预交(按时间先后来分摊）;3-全冲
        strSQL = "Select 1 From 三方退款信息 Where 结帐ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
        If rsTmp.EOF Then
            Call RecalcDepositMoney(1)
            mblnNotChange = True
            txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
            Call LoadIntendBalance
            mblnNotChange = False
        Else
            Call RecalcDepositMoney(3)
            mblnNotChange = True
            txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
            mblnNotChange = False
            txtBalance(Idx_冲预交).Enabled = False
            chkDeposit.Enabled = False
        End If
    End If

    If mEditType <> g_Ed_单据查看 Then
        mblnNotChange = True
        Call LoadCurOwnerPayInfor(mEditType = g_Ed_重新结帐)
        '0-医保预算信息显示;1-显示费用信息
        Call ShowLedDisplayBank(1)
        Call SetOperationCtrl(2)     'bytFun-0-结算前;1-医保虚拟结算后;2-已保存了结帐单;
        If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1 Then
            ReInitPatiInvoice False
            InitRedInvoice True
        Else
            ReInitPatiInvoice True
        End If
        mblnNotChange = False
    End If
    
    Call SetCurBalanceVisible
    If mEditType = g_Ed_重新结帐 Then
        Call txtBalance_Validate(Idx_冲预交, False)
        SetDefaultPayType
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
    End If
    ReadBalance = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetFromalanceIDToPatiNum(ByVal lng结帐ID As Long, Optional ByVal lngMax As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来获取本次结帐的住院次数
    '出参:lngMax-最大的住院次数
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-16 11:10:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTime As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct 主页ID " & _
    "   From 住院费用记录  " & _
    "   Where 结帐ID= [1]  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    lngMax = 0
    With rsTemp
        Do While Not .EOF
            If lngMax < Val(NVL(!主页ID)) Then lngMax = Val(NVL(!主页ID))
            strTime = strTime & "," & Val(NVL(!主页ID))
            .MoveNext
        Loop
    End With
    If strTime <> "" Then strTime = Mid(strTime, 2)
    GetFromalanceIDToPatiNum = strTime
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function ExecuteOldOneCardPayInterface(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal objCard As Card, ByVal dblMoney As Double, tyBrushCardInfor As TY_BrushCard, _
    ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(老版本)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次结算金额
    '     TYBrushCardInfor-当前刷卡信息
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 16:14:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl余额 As Double, str医院编码 As String
    Dim i As Long, strSQL As String, str结算方式 As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intCardType As Integer, strSwapNO As String
    '非一卡通支付,直接返回
    If objCard.结算性质 <> 7 Then ExecuteOldOneCardPayInterface = True: Exit Function

    mOldOneCard.rsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mOldOneCard.rsOneCard.EOF Then
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        ExecuteOldOneCardPayInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '调用之前,先处理数据
    'Zl_病人结帐结算_Modify
    strSQL = "Zl_病人结帐结算_Modify("
    '  操作类型_In     Number,
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    strSQL = strSQL & "1,"
    '  病人id_In       病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  结帐id_In       病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算方式_In     Varchar2,
    str结算方式 = objCard.结算方式
    str结算方式 = str结算方式 & "|" & dblMoney
    str结算方式 = str结算方式 & "|" & IIf(tyBrushCardInfor.str结算号码 = "", " ", tyBrushCardInfor.str结算号码)
    str结算方式 = str结算方式 & "|" & IIf(tyBrushCardInfor.str结算摘要 = "", " ", tyBrushCardInfor.str结算摘要)
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  冲预交_In       病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  退支票额_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In         病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In     病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In         病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL,"
    '  找补_In         病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In     门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "NULL,"
    '  结帐类型_In     Number := 2,
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
    '  缺省结算方式_In 结算方式.名称%Type := Null,
    strSQL = strSQL & "NULL,"
    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "NULL,"
    '  完成结算_In Number:=0
    strSQL = strSQL & "0)"
    zlAddArray cllPro, strSQL
    
    '一卡通结算
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If Not mobjICCard.PaymentSwap(dblMoney, dbl余额, intCardType, Val("" & mOldOneCard.rsOneCard!医院编码), tyBrushCardInfor.str卡号, tyBrushCardInfor.str交易流水号, lng结帐ID, lng病人ID) Then
        gcnOracle.RollbackTrans
        MsgBox objCard.结算方式 & "结算失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    strSQL = "Zl_一卡通结算_Update(" & 0 & ",'" & objCard.结算方式 & "','" & tyBrushCardInfor.str卡号 & "','" & intCardType & "','" & strSwapNO & "'," & dbl余额 & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set cllBillPro = New Collection
    blnTrans = False
    ExecuteOldOneCardPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
 End Function
 
Private Function CheckOldOneCardIsValied(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef tyBrushCard As TY_BrushCard, _
    Optional bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '入参:objCard-当前卡对象
    '     bln退款-是否退款
    '出参:tyBrushCard-返回刷卡信息
    '返回:一卡通验证正确或非一卡通,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 17:19:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl未付金额 As Double, strCardNo As String
    Dim dblTemp As Double, strXmlIn As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then CheckOldOneCardIsValied = True: Exit Function
    
    If objCard.结算性质 <> 7 Then CheckOldOneCardIsValied = True: Exit Function
    
    mOldOneCard.rsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mOldOneCard.rsOneCard.EOF Then
        Screen.MousePointer = 0
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        CheckOldOneCardIsValied = False: Exit Function
    End If

    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "一卡通接口创建失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblMoney = 0 Then dblMoney = Val(txtReceive.Text)
     
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "收款") & "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    dbl未付金额 = RoundEx(mBalanceInfor.dbl未付合计 - mBalanceInfor.dbl冲预交合计, 6)
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "收款") & "金额不能大于本次" & IIf(bln退款, "未退", "未收") & "金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Not bln退款 Then
       
       '弹出刷卡界面
       'zlBrushCard(frmMain As Object, _
       '    ByVal lngModule As Long, _
       '    ByVal rsClassMoney As ADODB.Recordset, _
       '    ByVal lngCardTypeID As Long, _
       '    ByVal bln消费卡 As Boolean, _
       '    ByVal strPatiName As String, ByVal strSex As String, _
       '    ByVal strOld As String, ByVal dbl金额 As Double, _
       '    Optional ByRef strCardNo As String, _
       '    Optional ByRef strPassWord As String, _
       '    Optional ByRef bln退费 As Boolean = False, _
       '    Optional ByRef blnShowPatiInfor As Boolean = False, _
       '    Optional ByRef bln退现 As Boolean = False, _
       '    Optional ByVal bln余额不足禁止 As Boolean = True) As Boolean
       '---------------------------------------------------------------------------------------------------------------------------------------------
       '功能:根据指定支付类别,弹出刷卡窗口
       '入参:rsClassMoney:收费类别,金额
       '        lngCardTypeID-为零时,为老一卡通刷卡
       '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
        
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, 0, False, _
        mrsInfo!姓名, NVL(mrsInfo!性别), NVL(mrsInfo!年龄), IIf(mPatiInfor.bln退款标志, -1, 1) * dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
        False, True, False, False, Nothing, False, False, strXmlIn) = False Then Exit Function
        
        tyBrushCard.dbl帐户余额 = mobjICCard.GetSpare
        If tyBrushCard.dbl帐户余额 < dblMoney Then
            Screen.MousePointer = 0
            MsgBox "卡余额不够支付,请检查!" & vbCrLf & vbCrLf & _
            "   卡 余  额" & Format(tyBrushCard.dbl帐户余额, "0.00") & vbCrLf & _
            "   本次支付" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
            Exit Function
        End If
        staThis.Panels(2).Text = Format(tyBrushCard.dbl帐户余额, "0.00")
        staThis.Panels(2).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(tyBrushCard.dbl帐户余额, "0.00")
       
        CheckOldOneCardIsValied = True
        Exit Function
    End If
    '退款检查
    If mrsBalance Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    mrsBalance.Filter = "类型=4"
    If mrsBalance.EOF Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        Screen.MousePointer = 0
        MsgBox "一卡通读卡失败,请将IC卡放在读卡器中", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> NVL(mrsBalance!卡号) Then
        Screen.MousePointer = 0
        MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblTemp = Format(Val(NVL(mrsBalance!冲预交)), "0.00")
    If RoundEx(dblMoney, 6) <> Format(dblTemp, "0.00") Then
        Screen.MousePointer = 0
        MsgBox "一卡通结算必须全退,请检查!" & vbCrLf & vbCrLf & _
        "   结算金额" & Format(dblTemp, "0.00") & vbCrLf & _
        "   本次支付" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckOldOneCardIsValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function

Private Function CheckThreeSwapValied(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易验证
    '入参:objCard-三方卡
    '     dblMoney-刷卡金额,>=0表示收款;小于零表示退款
    '     bln退款-true,表示当前为退款检查;False表示当前为收款检查
    '出参:tyBrushCard-刷卡信息
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, cllSquareBalance As Collection
    Dim strXMLExpend As String, bln退现 As Boolean
    Dim dbl帐户余额 As Double, dbl未付金额 As Double
    Dim strExpand As String, strXmlIn As String
    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim blnCurInput As Boolean
    
    intMousePointer = Screen.MousePointer
    
    If dblMoney = 0 Then CheckThreeSwapValied = True: Exit Function
    
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then CheckThreeSwapValied = True: Exit Function
    
    On Error GoTo errHandle
    tyBrushCard.bln转帐 = False
    If dblMoney = 0 Then dblMoney = Val(txtReceive.Text): blnCurInput = True
    
    dbl未付金额 = RoundEx(mBalanceInfor.dbl未付合计 + dblMoney, 6)
     
    If dblMoney = 0 Then
        If dbl未付金额 = 0 Then
            CheckThreeSwapValied = True: Exit Function
        Else
            Screen.MousePointer = 0
            MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "刷卡") & "金额不能大于本次" & IIf(bln退款, "未退", "未付") & "金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Abs(dblMoney) <> Format(Abs(dbl未付金额), "0.00") And blnCurInput Then
        If mty_ModulePara.byt刷卡缺省金额操作 = 1 Then
            If MsgBox(IIf(bln退款, "退款", "刷卡") & "金额(" & Format(dblMoney, "0.00") & ")与本次" & IIf(bln退款, "未退", "未付") & "金额(" & Format(Abs(dbl未付金额), "0.00") & _
                ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf mty_ModulePara.byt刷卡缺省金额操作 = 2 Then
            MsgBox IIf(bln退款, "退款", "刷卡") & "金额(" & Format(dblMoney, "0.00") & ")与本次" & IIf(bln退款, "未退", "未付") & "金额(" & Format(Abs(dbl未付金额), "0.00") & _
                ")不同，不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Not bln退款 Then
        'zlBrushCard(frmMain As Object, _
           ByVal lngModule As Long, _
           ByVal rsClassMoney As ADODB.Recordset, _
           ByVal lngCardTypeID As Long, _
           ByVal bln消费卡 As Boolean, _
           ByVal strPatiName As String, ByVal strSex As String, _
           ByVal strOld As String, ByRef dbl金额 As Double, _
           Optional ByRef strCardNo As String, _
           Optional ByRef strPassWord As String, _
           Optional ByRef bln退费 As Boolean = False, _
           Optional ByRef blnShowPatiInfor As Boolean = False, _
           Optional ByRef bln退现 As Boolean = False, _
           Optional ByVal bln余额不足禁止 As Boolean = True, _
           Optional ByRef varSquareBalance As Variant) As Boolean
           '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            mPatiInfor.str姓名, mPatiInfor.str性别, mPatiInfor.str年龄, IIf(mPatiInfor.bln退款标志, -1, 1) * dblMoney, _
            tyBrushCard.str卡号, tyBrushCard.str密码, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
            '保存前,一些数据检查
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal strCardTypeID As Long, ByVal strCardNo As String, _
            ByVal dblMoney As Double, ByVal strNOs As String, _
            Optional ByVal strXMLExpend As String
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.接口序号, _
            objCard.消费卡, tyBrushCard.str卡号, dblMoney, "", strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
          '入参:frmMain-调用的主窗体
          '        lngModule-模块号
          '        strCardNo-卡号
          '        strExpand-预留，为空,以后扩展
          '出参:dblMoney-返回帐户余额
        If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
              tyBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡) = False Then Exit Function
        
        staThis.Panels(2).Text = Format(dbl帐户余额, "0.00")
        staThis.Panels(2).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
        tyBrushCard.dbl帐户余额 = RoundEx(dbl帐户余额, 2)
        If dbl帐户余额 <> 0 And dbl帐户余额 < dblMoney Then
            Screen.MousePointer = 0
            MsgBox objCard.结算方式 & "的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        CheckThreeSwapValied = True
        Exit Function
    End If
    
    '退款检查
    If mrsBalance Is Nothing Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    If mEditType = g_Ed_重新作废 Then
        mrsOldBalance.Filter = "类型=3 And 卡类别ID=" & objCard.接口序号
        If mrsOldBalance.EOF Then
            If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.结算方式 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        mrsBalance.Filter = "类型=3 And 卡类别ID=" & objCard.接口序号
        If mrsBalance.EOF Then
            If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.结算方式 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
   
    dblTemp = 0
    If mEditType = g_Ed_重新作废 Then
        With mrsOldBalance
            Do While Not .EOF
                dblTemp = dblTemp + Val(NVL(!冲预交))
                .MoveNext
            Loop
            mrsOldBalance.MoveFirst
            dblTemp = RoundEx(dblTemp, 5)
        End With
    Else
        With mrsBalance
            Do While Not .EOF
                dblTemp = dblTemp + Val(NVL(!冲预交))
                .MoveNext
            Loop
            mrsBalance.MoveFirst
            dblTemp = RoundEx(dblTemp, 5)
        End With
    End If
    
    If dblTemp = 0 Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.结算方式 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objCard.是否全退 Then
        If dblTemp <> dblMoney Then
            If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "注意:" & vbCrLf & objCard.名称 & "进行退款时，必须全退！" & vbCrLf & _
            "  剩余未退:" & Format(Abs(dblTemp), "0.00") & vbCrLf & _
            "  当前金额:" & Format(Abs(dblMoney), "0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        If dblMoney > dblTemp Then
            If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        End If
    End If
        
    
    'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户回退交易前的检查
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       lngCardTypeID-卡类别ID
        '       strCardNo-卡号
        '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       dblMoney-退款金额
        '       strSwapNo-交易流水号(退款时检查)
        '       strSwapMemo-交易说明(退款时传入)
        '       strXMLExpend    XML IN  可选参数:异常单据重新退费(1)
        '返回:退款合法,返回true,否则返回Flase
        
    strXMLExpend = ""
    If mEditType = g_Ed_重新作废 Then
        tyBrushCard.str卡号 = NVL(mrsOldBalance!卡号)
        tyBrushCard.str交易流水号 = NVL(mrsOldBalance!交易流水号)
        tyBrushCard.str交易说明 = NVL(mrsOldBalance!交易说明)
    Else
        tyBrushCard.str卡号 = NVL(mrsBalance!卡号)
        tyBrushCard.str交易流水号 = NVL(mrsBalance!交易流水号)
        tyBrushCard.str交易说明 = NVL(mrsBalance!交易说明)
    End If

    strBalanceIDs = "2|" & mBalanceInfor.lng结帐ID & IIf(mBalanceInfor.lng冲销ID = 0, "", "," & mBalanceInfor.lng冲销ID)
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, _
        strBalanceIDs, dblMoney, tyBrushCard.str交易流水号, tyBrushCard.str交易说明, strXMLExpend) = False Then Exit Function
    
    If objCard.是否退款验卡 Then
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
        Optional ByRef bln退现 As Boolean) As Boolean
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.接口序号, _
            objCard.消费卡, mPatiInfor.str姓名, mPatiInfor.str性别, _
            mPatiInfor.str年龄, IIf(mPatiInfor.bln退款标志, -1, 1) * dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
            True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    End If
    CheckThreeSwapValied = True
    Exit Function
    
GoTransferAccount:
    strXmlIn = "<IN><CZLX>1</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.接口序号, _
        objCard.消费卡, mPatiInfor.str姓名, mPatiInfor.str性别, _
        mPatiInfor.str年龄, IIf(bln退款, -1, 1) * dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
        True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    
    tyBrushCard.bln转帐 = True
    '调用转帐接口
    '    7.1.    zltransferAccountsCheck(转帐检查接口)
    'zlTransferAccountsCheck 转帐检查接口
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  HIS调用模块号
    'lngCardTypeID   Long    In  卡类别ID
    'strCardNo   String  In  卡号
    'dblMoney    Double  In  转帐金额(代扣时为负数)
    'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
    'strXMLExpend String In   XML串:
    '                            <IN>
    '                                <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务
    '                            </IN>
    '                    Out  XML串:
    '                            <OUT>
    '                               <ERRMSG>错误信息</ERRMSG >
    '                            </OUT>
    '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
    '说明:
    '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '构造XML串
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.接口序号, _
        tyBrushCard.str卡号, dblMoney, mBalanceInfor.lng结帐ID, strXMLExpend) = False Then
        Screen.MousePointer = 0
        Call zlShowThreeSwapErrInfor(0, strXMLExpend)
        Exit Function
    End If
    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
          tyBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡)
    If dbl帐户余额 <> 0 Then
        staThis.Panels(2).Text = objCard.结算方式 & "帐户余额:" & Format(dbl帐户余额, "0.00")
        staThis.Panels(2).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
    End If
    tyBrushCard.dbl帐户余额 = RoundEx(dbl帐户余额, 2)
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function
 


Private Function ExecuteThreeSwapPayInterface(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, objCard As Card, ByVal dblMoney As Double, _
    ByRef cllBillPro As Collection, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     tyBrushCard-当前刷卡信息
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str结算方式  As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    If dblMoney = 0 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
   '调用之前,先处理数据
    'Zl_病人结帐结算_Modify
    strSQL = "Zl_病人结帐结算_Modify("
    '  操作类型_In     Number,
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    strSQL = strSQL & "1,"
    '  病人id_In       病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  结帐id_In       病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算方式_In     Varchar2,
    str结算方式 = objCard.结算方式
    str结算方式 = str结算方式 & "|" & dblMoney
    str结算方式 = str结算方式 & "|" & IIf(tyBrushCard.str结算号码 = "", " ", tyBrushCard.str结算号码)
    str结算方式 = str结算方式 & "|" & IIf(tyBrushCard.str结算摘要 = "", " ", tyBrushCard.str结算摘要)
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  冲预交_In       病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  退支票额_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & objCard.接口序号 & ","
    '  卡号_In         病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & tyBrushCard.str卡号 & "',"
    '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & tyBrushCard.str交易流水号 & "',"
    '  交易说明_In     病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & tyBrushCard.str交易说明 & "',"
    '  缴款_In         病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl缴款 & ","
    '  找补_In         病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl找补 & ","
    '  误差金额_In     门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "NULL,"
    '  结帐类型_In     Number := 2,
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
    '  缺省结算方式_In 结算方式.名称%Type := Null,
    strSQL = strSQL & "NULL,"
    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "NULL,"
    '  完成结算_In Number:=0
    strSQL = strSQL & "0)"
    zlAddArray cllPro, strSQL
    
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
    '       dblMoney-结算金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    str结帐IDs = lng结帐ID
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, _
         str结帐IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    tyBrushCard.str交易流水号 = strSwapGlideNO
    tyBrushCard.str交易说明 = strSwapMemo
    
    If objCard.消费卡 = False Then
        Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    End If
    Call zlAddThreeSwapSQLToCollection(False, str结帐IDs, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    mBalanceInfor.blnSaveBill = True
    ExecuteThreeSwapPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strKinds As String
    Dim intIdkind As Integer
    Dim strIdkind As String
    If mEditType = g_Ed_单据查看 Then Exit Sub
        
    On Error GoTo errHandle
    
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    'strKinds = "姓|姓名|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|;住|住院号|0|0|0|0|0|;就|就诊卡|0|0|0|0|0|"
    
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKinds, txtPatient)
    Call GetRegInFor(g私有模块, Me.Name, "IDKIND", strIdkind)
    If Val(strIdkind) > 0 And Val(strIdkind) <= IDKind.ListCount Then IDKind.IDKind = Val(strIdkind)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKind.Cards.按缺省卡查找
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域初始设置
    '编制:刘兴洪
    '日期:2014-05-26 10:30:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, lngHeight As Long
    Dim strReg As String
    Dim panThis As Pane, panThis1 As Pane
    lngHeight = picPati.Height \ Screen.TwipsPerPixelY
    Set panThis = dkpMain.CreatePane(mConPans.Pan_PatiCon, 200, lngHeight, DockLeftOf, Nothing)
    panThis.Title = "病人条件"
    panThis.Handle = picPati.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = mConPans.Pan_PatiCon
    panThis.MaxTrackSize.Height = lngHeight
    panThis.MinTrackSize.Height = lngHeight
    
    Set panThis1 = dkpMain.CreatePane(mConPans.Pan_FeeList, 250, 580, DockBottomOf, panThis)
    panThis1.Title = "费目表"
    panThis1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis1.Handle = picFeeList.hWnd
    panThis1.Tag = mConPans.Pan_FeeList
    
    If mEditType = g_Ed_单据查看 Then
'        Set panThis = dkpMain.CreatePane(mConPans.Pan_Deposit, 250, 580, DockRightOf, panThis1)
'        panThis.Title = "预交情况"
'        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'        panThis.Handle = picDeposit.hWnd
'        panThis.Tag = mConPans.Pan_Deposit
        Set panThis = dkpMain.CreatePane(mConPans.Pan_Balance, 250, 580, DockRightOf, panThis1)
        panThis.Title = "结帐列表"
        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        panThis.Handle = picBalanceBack.hWnd
        panThis.Tag = mConPans.Pan_Balance
        panThis.MaxTrackSize.Width = 7500 \ Screen.TwipsPerPixelY
        panThis.MinTrackSize.Width = panThis.MaxTrackSize.Width
    Else
        Set panThis = dkpMain.CreatePane(mConPans.Pan_Balance, 250, 580, DockRightOf, panThis1)
        panThis.Title = "结帐列表"
        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        panThis.Handle = picBalanceBack.hWnd
        panThis.Tag = mConPans.Pan_Balance
        panThis.MaxTrackSize.Width = 6500 \ Screen.TwipsPerPixelY
        panThis.MinTrackSize.Width = panThis.MaxTrackSize.Width
    End If
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    dkpMain.Options.LockSplitters = True
    dkpMain.VisualTheme = ThemeDefault
    dkpMain.RecalcLayout
End Sub

Private Sub txtPatiBegin_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt天数.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt中途.Value = True, 1, 0)
        If Val(txt天数.Text) = 0 Then txt天数.Text = 1
    Else
        txt天数.Text = ""
    End If
End Sub

Private Sub txtPatiBegin_GotFocus()
    zlControl.TxtSelAll txtPatiBegin
    mstrPatiBegin = txtPatiBegin.Text
End Sub

Private Sub txtPatiBegin_Validate(Cancel As Boolean)
    If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
        MsgBox "请输入正确的住院开始日期!", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtPatiEnd_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt天数.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt中途.Value = True, 1, 0)
        If Val(txt天数.Text) = 0 Then txt天数.Text = 1
    Else
        txt天数.Text = ""
    End If
End Sub

Private Sub txtPatiEnd_GotFocus()
    zlControl.TxtSelAll txtPatiEnd
    mstrPatiEnd = txtPatiEnd.Text
End Sub

Private Sub txtPatiEnd_Validate(Cancel As Boolean)
    If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
        MsgBox "请输入正确的住院结束日期!", vbInformation, gstrSysName
        Cancel = True
   End If
End Sub
Private Function YBIdentifyCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消医保病人身份验证
    '返回:返回假时不退出界面或清除操作
    '编制:刘兴洪
    '日期:2015-01-12 16:08:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, varData As Variant
    On Error GoTo errHandle
        
    YBIdentifyCancel = True
    If mYBInFor.strYBPati <> "" Then
        varData = Split(mYBInFor.strYBPati, ";")
        If UBound(varData) >= 8 Then lng病人ID = Val(varData(8))
        If lng病人ID <> 0 Then YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng病人ID, mYBInFor.intInsure)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function RecalcFeeTotalDate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置费用的统计时间
    '返回:计算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 16:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str主页Ids As String, strStartDate As String, strEndDate As String
    Dim i As Long, lngMax As Long, lngMin As Long
    Dim varData As Variant, lng病人ID As Long
    
    
    If mEditType = g_Ed_门诊结帐 Then RecalcFeeTotalDate = True: Exit Function
    
    If mrsInfo Is Nothing Then RecalcFeeTotalDate = True: Exit Function
    If mrsInfo.State = 0 Then RecalcFeeTotalDate = True: Exit Function
    
    
    varData = Split(zlGetAllTims(cboPatiNums.GetNodesCheckedDatas(False)), ",")
    For i = 0 To UBound(varData)
        If lngMax = 0 Then lngMax = Val(varData(i))
        If lngMin = 0 Then lngMin = Val(varData(i))
        If lngMax < Val(varData(i)) Then
            lngMax = Val(varData(i))
        End If
        If lngMin > Val(varData(i)) Then
            lngMin = Val(varData(i))
        End If
    Next
    
    If lngMin = 0 And lngMax = 0 Then
        MsgBox "请先选择住院次数!", vbInformation, Me.Caption
        Exit Function
    End If
    
    lng病人ID = Val(NVL(mrsInfo!病人ID)): str主页Ids = IIf(lngMin = lngMax, lngMax, lngMin & "," & lngMax)
    If mobjBalanceAll.GetPatiFeeDateRang(lng病人ID, str主页Ids, strStartDate, strEndDate, gint费用时间 = 0) = False Then
        strEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        strStartDate = Format(CDate(strEndDate), "yyyy-mm-dd") & " 00:00:00"
    End If
    txtBegin.Text = Format(strStartDate, "yyyy-mm-dd")
    txtEnd.Text = Format(strEndDate, "yyyy-mm-dd")
    
    RecalcFeeTotalDate = True
End Function
Private Function CheckFactIsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否合法
    '出参:objSetFocus -出错时,光标定位到哪个对象
    '编制:刘兴洪
    '日期:2015-01-13 10:21:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '不打印,直接返回true,不检查票据号
    If mobjFactProperty.打印方式 = 0 Then CheckFactIsValied = True: Exit Function
    
    '先结自费费用时不打印发票票据
    If mty_ModulePara.blnNotPrintInvioce And mobjBalanceCon.blnCurBalanceOwnerFee Then CheckFactIsValied = True:  Exit Function
    
    If Not mobjFactProperty.严格控制 Then      '非严格控制
        If Len(txtInvoice.Text) <> mobjFactProperty.票号长度 And txtInvoice.Text <> "" Then
            MsgBox "票据号码长度应该为 " & mobjFactProperty.票号长度 & " 位！", vbInformation, gstrSysName
            Set objSetFocus = txtInvoice
            Exit Function
        End If
        CheckFactIsValied = True
        Exit Function
    End If
    
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
        Set objSetFocus = txtInvoice
        Exit Function
    End If
    If zlGetInvoiceGroupUseID(mlng领用ID, 1, txtInvoice.Text) = False Then Exit Function
    CheckFactIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function PrintBill(ByVal lng病人ID As Long, ByVal strNO As String, ByVal lng结帐ID As Long, _
    ByVal dtBalanceDate As Date, ByVal dbl缴款 As Double, ByVal dbl找补 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否打印票据
    '入参:strNO-结帐单号
    '     lng结帐ID-结帐ID
    '     dtBalanceDate-结帐日期
    '返回:打印票据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-13 10:08:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln打印退款收据 As Boolean, bytKind As Byte
    Dim bln打印费用明细 As Boolean, bln自费清单 As Boolean, blnPrintBillEmpty As Boolean
        
    On Error GoTo errHandle
    
    bln打印退款收据 = False
    If mty_ModulePara.int退款票据 <> 0 And InStr(1, mstrPrivs, ";病人结帐退款收据;") > 0 Then
        '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
        If mty_ModulePara.int退款票据 = 1 Then
           If MsgBox("你是否要打印“病人结帐退款收据”？" & vbCrLf & _
                   "   『是』：打印病人结帐退款收据" & vbCrLf & _
                   "   『否』：不打印病人结帐退款收据", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                bln打印退款收据 = True
            End If
        Else
            bln打印退款收据 = True
        End If
    End If
  
    bln打印费用明细 = False
     Select Case mty_ModulePara.bytFeePrintSet
     Case 1  '打印.
         If MsgBox("你是否要打印“病人结帐费用明细”？" & vbCrLf & _
                 "   『是』：打印病人结帐费用明细" & vbCrLf & _
                 "   『否』：不打印病人结帐费用明细", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                 bln打印费用明细 = True
         End If
     Case 0  '不打印
     Case 2  '打印.但不提示
         bln打印费用明细 = True
     End Select
     
    If mobjBalanceCon.blnCurBalanceOwnerFee Then   '自费清单打印控制
       bln自费清单 = False
       Select Case Val(zlDatabase.GetPara("自费费用打印方式", glngSys, mlngModul, "0"))
           Case 2  '打印.
               If MsgBox("你是否要打印“病人自费费用清单”？" & vbCrLf & _
                       "   『是』：打印病人自费费用清单" & vbCrLf & _
                       "   『否』：不打印病人自费费用清单", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                       bln自费清单 = True
               End If
           Case 0  '不打印
           Case 1  '打印.但不提示
               bln自费清单 = True
       End Select
    End If
        
    If bln打印退款收据 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me, "结帐ID=" & lng结帐ID, 2)
    End If
    
    '票据打印
    If mblnPrintInvoice Or (mYBInFor.intInsure <> 0 And MCPAR.医保接口打印票据) Then
        '问题:44332
RePrint:
        Dim strNotValiedNos As String
        mobjFactProperty.LastUseID = mlng领用ID
        Call UpateStartInvoice(mBalanceInfor.strNO, txtInvoice.Text)
        Call frmPrint.ReportPrint(1, strNO, lng结帐ID, mobjFactProperty, txtInvoice.Text, _
             dtBalanceDate, CCur(dbl缴款), CCur(dbl找补), , mobjFactProperty.打印格式, blnPrintBillEmpty, mYBInFor.intInsure <> 0 And MCPAR.医保接口打印票据)
        If mEditType = g_Ed_门诊结帐 Then
            bytKind = mty_ModulePara.bytInvoiceKindMZ
        Else
            bytKind = mty_ModulePara.bytInvoiceKindZY
        End If
        If mobjFactProperty.严格控制 And blnPrintBillEmpty = False And _
            ((bytKind = 0 And InStr(1, mstrPrivs, ";收据打印;") > 0) _
               Or (bytKind <> 0 And InStr(1, mstrPrivs, ";打印门诊收费票据;") > 0)) Then    'blnPrintBillEmpty:55052
            '60155
             If zlIsNotSucceedPrintBill(3, strNO, strNotValiedNos) = True Then
                     If MsgBox("结帐单据为[" & strNotValiedNos & "]的结帐票据打印未成功,是否重新打印结帐票据?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
             End If
        End If
    End If
    

    If bln打印费用明细 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me, "病人ID=" & lng病人ID, "结帐ID=" & lng结帐ID, 2)
    End If
    
    If bln自费清单 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_4", Me, "病人ID=" & lng病人ID, "结帐ID=" & lng结帐ID, 2)
    End If
    
    If mblnDepositBillPrint Then
        '打印预交票据
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mBalanceInfor.str预交No, "病人ID=" & mPatiInfor.lng病人ID, "收款时间=" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS"), 2)
    End If
    
    PrintBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function UpateStartInvoice(ByVal strNO As String, ByVal strInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改开始发票号
    '入参:strNO-结帐单号
    '编制:刘兴洪
    '日期:2015-01-14 10:21:52
    '说明:可能外面存在事务,所以不能使用错误错误中心(由父窗口捕获)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_票据起始号_Update('" & strNO & "','" & Trim(strInvoice) & "',3)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
End Function
 

 
Private Function CheckInputConsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入条件的效性检查
    '出参:objSetFocus-光标移动到指定的控件
    '返回:结帐数据有效，返回True,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:03:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnNotFondPati As Boolean
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then blnNotFondPati = True
    If Not blnNotFondPati Then blnNotFondPati = mrsInfo.State = 0
    
    If blnNotFondPati Then
        MsgBox "没有确定结帐病人,不能进行结帐操作！", vbExclamation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If

    If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
        MsgBox "请输入一个有效的开始时间！", vbInformation, gstrSysName
        Set objSetFocus = txtPatiBegin
        Exit Function
    End If
    If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
        MsgBox "请输入一个有效的结束时间！", vbInformation, gstrSysName
        Set objSetFocus = txtPatiEnd
        Exit Function
    End If
    
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        If txtPatiEnd < txtPatiBegin.Text Then
            MsgBox "结束时间不能小于开始时间！", vbInformation, gstrSysName
            Set objSetFocus = txtPatiBegin
            Exit Function
        End If
    End If
    If IsDate(txtPatiBegin.Text) And Not IsDate(txtPatiEnd.Text) Then
        MsgBox "请一并输入有效的结束时间！", vbInformation, gstrSysName
        Set objSetFocus = txtPatiBegin
        Exit Function
    End If
    If Not IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        MsgBox "请一并输入有效的开始时间！", vbInformation, gstrSysName
        Set objSetFocus = txtPatiBegin
        Exit Function
    End If
    If mrsFeeList Is Nothing Then
        MsgBox "该设置病人没有需要结帐的费用条件！", vbInformation, gstrSysName
        Set objSetFocus = cmdMore
        Exit Function
    End If
    If mrsFeeList.State <> 1 Then
        MsgBox "该设置下病人没有需要结帐的费用！", vbInformation, gstrSysName
        Set objSetFocus = cmdMore
        Exit Function
    End If
    If mrsFeeList.RecordCount = 0 Then
        MsgBox "该设置下病人没有需要结帐的费用！", vbInformation, gstrSysName
        Set objSetFocus = cmdMore
        Exit Function
    End If
        
    If zlCommFun.StrIsValid(txtBalance(Idx_结帐说明).Text, 50, txtBalance(Idx_结帐说明).hWnd, "结帐说明") = False Then
        Set objSetFocus = txtBalance(Idx_结帐说明)
        Exit Function
    End If
    
    If Val(txtBalance(Idx_本次未结).Text) < Val(txtBalance(Idx_本次结帐).Text) Then
        Call MsgBox("当前结帐金额大于了未结金额，不能进行结帐操作。", vbInformation, gstrSysName)
        Set objSetFocus = txtBalance(Idx_本次结帐)
        Exit Function
    End If

    If Val(txtBalance(Idx_本次未结).Text) <> 0 And Val(txtBalance(Idx_本次结帐).Text) = 0 Then
        Call MsgBox("未输入本次要结帐的金额，不能进行结帐操作。", vbInformation, gstrSysName)
        Set objSetFocus = txtBalance(Idx_本次结帐)
        Exit Function
    End If
    
    If Val(txtBalance(Idx_本次未结).Text) <= 0 Then
        If MsgBox("病人实际没有可结费用,要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set objSetFocus = txtPatient
            Exit Function
        End If
    End If
    
    '检查发票是否有效
    If CheckFactIsValied(objSetFocus) = False Then Exit Function
    If CheckBusinessRuleIsValied(objSetFocus) = False Then Exit Function     '业务规则检查
    
    CheckInputConsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSaveStrickDepositSQL(ByRef cllDeposit As Collection, ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存预交款的数据
    '出参:cllDeposit-相关的数据集
    '     objSetFocus-获取失败时,缺省光标定位到指定的控件中
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-19 15:12:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int预交类别 As Integer, lng病人ID As Long
    Dim dbl冲预交合计 As Double, dblMoney As Double, dbl预交余额 As Double
    Dim dbl冲预交 As Double, dbl预交余额合计 As Double
    Dim strTime As String
    Dim rsDeposit As ADODB.Recordset, i As Long
    
    
    On Error GoTo errHandle
    lng病人ID = mPatiInfor.lng病人ID
    strTime = ""
    If mty_ModulePara.bln仅用指定预交款 Then
        strTime = IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime)
    End If
    Set objSetFocus = txtBalance(Idx_冲预交)
    
    int预交类别 = 2
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then int预交类别 = 1
    If cllDeposit Is Nothing Then Set cllDeposit = New Collection
    dblMoney = RoundEx(Val(txtBalance(Idx_冲预交).Text), 2)
    With vsDeposit
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            '重读可用预交,并发操作判断
            Set rsDeposit = GetDeposit(lng病人ID, mblnDateMoved, strTime, , , int预交类别, mrs结算方式)
            For i = 1 To .Rows - 1
                dbl预交余额 = Val(.TextMatrix(i, .ColIndex("余额")))
                dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                If dbl冲预交 <> 0 Then
                    rsDeposit.Filter = "ID=" & CLng(.TextMatrix(i, .ColIndex("ID"))) & _
                        " And NO='" & .TextMatrix(i, .ColIndex("单据号")) & "' And 记录状态=" & .RowData(i) & " And 金额=" & dbl预交余额
                    If rsDeposit.RecordCount = 0 Then
                        If MsgBox("由于并发操作,病人预交款已发生变化,请重新提取病人结帐,是否重新提取预交数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                             Call LoadDepositList(lng病人ID, strTime)
                        End If
                        Screen.MousePointer = 0
                        Exit Function
                    End If

                    strSQL = "zl_结帐预交记录_Insert(" & CLng(.TextMatrix(i, .ColIndex("ID"))) & "," & _
                        "'" & .TextMatrix(i, .ColIndex("单据号")) & "'," & .RowData(i) & "," & _
                        dbl冲预交 & "," & mBalanceInfor.lng结帐ID & "," & lng病人ID & ")"
                    zlAddArray cllDeposit, strSQL
                   dbl冲预交合计 = RoundEx(dbl冲预交合计 + dbl冲预交, 6)
                End If
                dbl预交余额合计 = RoundEx(dbl预交余额合计 + dbl预交余额, 6)
            Next
            '结帐冲过的预交单据在预交款管理中被作废后,会出现负的预交余额单据
            If Val(dbl冲预交合计) > Val(dbl预交余额合计) And dbl冲预交合计 <> 0 Then
                Call MsgBox("可用预交余额不足冲款金额!", vbInformation, gstrSysName)
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
    End With
    
    dbl冲预交合计 = RoundEx(dbl冲预交合计, 6)
    If Val(dbl冲预交合计) = Val(dblMoney) Then
        GetSaveStrickDepositSQL = True: Exit Function
    End If
    
    If MsgBox("当前冲预交金额与冲预交明细不一致,是否重新按当前未付金额冲预交?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        '操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按结帐金额来冲预交(按时间先后来分摊）;3-全冲
        dblMoney = RoundEx(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl已付合计, 2)
        If dblMoney < 0 Then
            dblMoney = 0
            Call RecalcDepositMoney(0)
        Else
            Call RecalcDepositMoney(2, dblMoney)
        End If
        mblnNotChange = True
        txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
        mblnNotChange = False
    End If
    Screen.MousePointer = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckDepositValied(Optional blnCurBrushDeposit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的预交款是否合法
    '出参:blnCurBrushDeposit-当前是刷的预交款
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 15:15:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney  As Double, rsDeposit As ADODB.Recordset, i As Long, strSQL As String
    Dim lng病人ID As Long, strTime As String, int预交类别 As Integer
    Dim dbl预交余额 As Double, dbl冲预交 As Double, dbl预交余额合计 As Double, dbl冲预交合计 As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    dblMoney = Val(txtBalance(Idx_冲预交).Text)
    '无预交款且在保存时检查
    If dblMoney = 0 Then CheckDepositValied = True: Exit Function
    
    '刷过预交款的，需要重新刷卡
    If mBalanceInfor.bln预交刷卡 Then CheckDepositValied = True: Exit Function
    
    blnCurBrushDeposit = True
    
    If Not IsNumeric(txtBalance(Idx_冲预交).Text) And txtBalance(Idx_冲预交).Text <> "" Then
        Screen.MousePointer = 0:
        MsgBox "无效数值！", vbInformation, gstrSysName
        Exit Function
    ElseIf Val(txtBalance(Idx_冲预交).Text) < 0 Then
        dblTemp = 0
        For i = 1 To vsDeposit.Rows - 1
            dblTemp = dblTemp + vsDeposit.TextMatrix(i, vsDeposit.ColIndex("余额"))
        Next i
        If dblTemp >= 0 Then
            mblnNotChange = True
            MsgBox "预存款冲款金额不能为负！", vbInformation, gstrSysName
            mblnNotChange = False
            Screen.MousePointer = 0: Exit Function
        End If
    Else
'        If Val(txtBalance(Idx_冲预交).Text) > 0 And mBalanceInfor.dbl未付合计 < 0 Then
'        Screen.MousePointer = 0:
'        mblnNotChange = True
'        MsgBox "当前为退款,不能使用预存款！", vbInformation, gstrSysName
'        mblnNotChange = False
'        txtBalance(Idx_冲预交).Text = "0.00": Exit Function
    End If
    
    If Val(dblMoney) > Val(mPatiInfor.dbl实际余额) Then
        Screen.MousePointer = 0
        mblnNotChange = True
        MsgBox "冲预交金额大于了病人的预交余额,请检查!" & vbCrLf & _
               "当前冲预:" & Format(dblMoney, "0.00") & vbCrLf & _
               "当前余额:" & Format(mPatiInfor.dbl预交余额, "0.00"), vbInformation + vbOKOnly, gstrSysName
        mblnNotChange = False
        Exit Function
    End If
    
    lng病人ID = mPatiInfor.lng病人ID
    With vsDeposit
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            For i = 1 To .Rows - 1
                dbl预交余额 = Val(.TextMatrix(i, .ColIndex("余额")))
                dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                dbl冲预交合计 = RoundEx(dbl冲预交合计 + dbl冲预交, 5)
                dbl预交余额合计 = RoundEx(dbl预交余额合计 + dbl预交余额, 5)
            Next
            '结帐冲过的预交单据在预交款管理中被作废后,会出现负的预交余额单据
            If Val(dbl冲预交合计) > Val(dbl预交余额合计) And dbl冲预交合计 <> 0 Then
                Screen.MousePointer = 0
                Call MsgBox("可用预交余额不足冲款金额!", vbInformation, gstrSysName)
                Exit Function
            End If
        End If
    End With
    
    dbl冲预交合计 = RoundEx(dbl冲预交合计, 6)
    If Val(dbl冲预交合计) <> Val(dblMoney) Then
        Screen.MousePointer = 0
        If MsgBox("当前冲预交金额与冲预交明细不一致,是否重新按当前未付金额冲预交?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按结帐金额来冲预交(按时间先后来分摊）;3-全冲
            dblMoney = RoundEx(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl已付合计, 2)
            If dblMoney < 0 Then
                dblMoney = 0
                Call RecalcDepositMoney(0)
            Else
                Call RecalcDepositMoney(2, dblMoney)
            End If
            mblnNotChange = True
            txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
            mblnNotChange = False
        End If
        Exit Function
    End If
    
    '进行刷卡验证
    If gdbl预存款消费验卡 = 0 Then
        txtBalance(Idx_冲预交).BackColor = &HE0E0E0
        mBalanceInfor.bln预交刷卡 = True
        CheckDepositValied = True: Exit Function
    End If
    
    '住院的不用刷卡验证
    If Not (mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo) Then
        mBalanceInfor.bln预交刷卡 = True
        txtBalance(Idx_冲预交).BackColor = &HE0E0E0
        CheckDepositValied = True: Exit Function
    End If
    If Not zlDatabase.PatiIdentify(Me, glngSys, lng病人ID, dblMoney, , , , IIf(-1 * gdbl预存款消费验卡 >= dblMoney, False, True), , , , (gdbl预存款消费验卡 = 2)) Then
        txtBalance(Idx_冲预交).BackColor = vbWhite
        Exit Function
    End If
    txtBalance(Idx_冲预交).BackColor = &H8000000F
    mBalanceInfor.bln预交刷卡 = True
    CheckDepositValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckCurBalanceIsValied(ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal bln预交 As Boolean = False, _
    Optional ByRef objSetFocus As Object, _
    Optional objInCard As Card, Optional dblInMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前结帐是否有效
    '出参:tyBrushCard当前刷卡信息
    '     objSetFocus-光标移动对象
    '返回:有效返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 14:57:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng病人ID As Long, varData As Variant
    Dim dblMoney As Double, i As Long, blnFind As Boolean
    Dim cllDeposit As Collection, int性质 As Integer
    Dim dblCheck As Double
    
    Dim intCount As Integer '多种结算方式(排开医保)
    On Error GoTo errHandle
    
    If Not objInCard Is Nothing Then Set objCard = objInCard
    dblMoney = dblInMoney
    
    '输入条件的有效性检查
    If Not mBalanceInfor.blnSaveBill And mblnNotify = False Then
        If CheckInputConsValied(objSetFocus) = False Then Exit Function
    End If
    
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, .ColIndex("结算号码")), "'") > 0 Then
                MsgBox "结算号码含有非法字符单引号,不允许结帐", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("结算号码"))) > 30 Then
                 MsgBox "结算号码最多只能输入30个字符或15个汉字,不允许结帐", vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
            
            If InStr(.TextMatrix(i, .ColIndex("备注")), "'") > 0 Then
                MsgBox "摘要含有非法字符单引号,不允许结帐", vbInformation + vbOKOnly, gstrSysName
                Exit Function
           End If
        
           If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("备注"))) > 50 Then
                MsgBox "摘要最多只能输入50个字符或25个汉字,不允许结帐", vbInformation + vbOKOnly, gstrSysName
                Exit Function
           End If
        
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If bln预交 Then
                If int性质 = 1 Then blnFind = True: Exit For
            End If

            If InStr("34", int性质) > 0 And mbln连续结帐 Then
                MsgBox "连续结帐模式下,不允许使用:" & .TextMatrix(i, .ColIndex("结算方式")) & "进行结帐!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            int性质 = Val(.TextMatrix(i, .ColIndex("结算性质")))
            dblCheck = Val(.TextMatrix(i, .ColIndex("结算金额")))
            If InStr(",1,2,", "," & int性质 & ",") > 0 And dblCheck <> 0 Then intCount = intCount + 1
        Next
        
        If blnFind Then
            Screen.MousePointer = 0
            If bln预交 Then
                MsgBox "已经用预存款支付,只有删除预存款后才能支付!", vbOKOnly, gstrSysName
            Else
                MsgBox objCard.结算方式 & " 已经支付了,不能再用" & objCard.结算方式 & "进行支付", vbOKOnly + vbDefaultButton1, gstrSysName
            End If
            Exit Function
        End If
    End With

    '数据检查接口数(目前只同时支持两种接口(含医保算一种接口)
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    
    If objCard Is Nothing Then
        Set objCard = GetCard(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算方式")))
    End If
        
    '1.消费卡检查
    If CheckSquareBalanceValied(objCard, tyBrushCard, dblMoney) = False Then
        Set objSetFocus = txtReceive
        Exit Function
    End If
     
    '2.三方帐户检查
    If CheckThreeSwapValied(objCard, dblMoney, tyBrushCard, mPatiInfor.bln退款标志) = False Then
        Set objSetFocus = txtReceive
        Exit Function
    End If
    
    '3.一卡通(老版)检查
    If CheckOldOneCardIsValied(objCard, dblMoney, tyBrushCard) = False Then
        Set objSetFocus = txtReceive
        Exit Function
    End If
    
    '4.检查现金结算方式
    If CheckCashValied(objCard) = False Then
        Set objSetFocus = txtReceive
        Exit Function
    End If
    
    
    '5.检查支票结算方式是否合法
    If CheckChequeValied(objCard) = False Then
        Set objSetFocus = txtReceive
        Exit Function
    End If
    
    '6.检查其他结算方式
    If CheckOtherValied(objCard) = False Then
        Set objSetFocus = txtReceive
        Exit Function
    End If
    
    CheckCurBalanceIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckChequeValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查支票结算方式的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl当前未付 As Double
    Dim intMousePointer As Integer
    Dim objTempCard As Card
    Dim blnCheck As Boolean
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then CheckChequeValied = True: Exit Function
    
    If objCard.结算性质 <> 2 Or Not objCard.结算方式 Like "*支票*" Then CheckChequeValied = True: Exit Function
    
    
    dbl当前未付 = RoundEx(mBalanceInfor.dbl未付合计 - mBalanceInfor.dbl冲预交合计, 5)
    
    strTittle = IIf(dbl当前未付 < 0, "退款", "收款")
    dblMoney = Format(Val(txtReceive.Text), "0.00")
     
    If strTittle = "收款" Then
    
        If RoundEx(dblMoney, 6) = 0 And Not mbln连续结帐 Then
            Screen.MousePointer = 0
            MsgBox "未输入收款金额！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If dblMoney > RoundEx(dbl当前未付, 2) Then
            blnCheck = False
            If objTempCard Is Nothing Then
                blnCheck = True
            Else
                If objTempCard.接口序号 = 1 Then blnCheck = True
            End If
            
            
            If mstr退支票 = "" And blnCheck Then
                Screen.MousePointer = 0
                MsgBox "在结算方式中没有设置应付款的结算方式,不能进行退支票处理", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        CheckChequeValied = True
        Exit Function
    End If
    
    '退款
    If RoundEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "未输入退款金额！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckChequeValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckOtherValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查其他结算方式(支票等)的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl当前未付 As Double
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then CheckOtherValied = True: Exit Function
    
    If objCard.接口序号 > 0 Or objCard.结算方式 Like "*支票*" Or objCard.结算性质 = 1 Then CheckOtherValied = True: Exit Function
    
    dbl当前未付 = RoundEx(mBalanceInfor.dbl未付合计 - mBalanceInfor.dbl冲预交合计, 5)
    strTittle = IIf(dbl当前未付 < 0, "退款", "收款")
    dblMoney = Format(Val(txtReceive.Text), "0.00")
  
    If strTittle = "收款" Then
        If RoundEx(dblMoney, 6) = 0 And Not mbln连续结帐 And dbl当前未付 <> 0 Then
            Screen.MousePointer = 0
            MsgBox "未输入" & strTittle & "金额！", vbInformation, gstrSysName
            Exit Function
        End If
        If dblMoney > RoundEx(dbl当前未付, 2) Then
            Screen.MousePointer = 0
            MsgBox "注意:" & vbCrLf & "    输入的" & strTittle & "金额大于了未支付的金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckOtherValied = True
        Exit Function
    End If
    
    '退款
    If RoundEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "未输入" & strTittle & "金额！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If dblMoney > RoundEx(Abs(dbl当前未付), 2) Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "    输入的退款金额大于了未退金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    Screen.MousePointer = 0

    CheckOtherValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckCashValied(ByVal objCard As Card, Optional ByVal bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查现金结算方式的一些合法情检查
    '入参:objCard－当前支付卡
    '     bln退款
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, strTittle As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer

    
    On Error GoTo errHandle
    If objCard Is Nothing Then CheckCashValied = True: Exit Function
    If objCard.结算性质 <> 1 Then CheckCashValied = True: Exit Function
    
    dblMoney = Format(Val(txtReceive.Text), "0.00")
    If Not bln退款 Then
        '43153
        '缴款控制:0-不进行控制;1-存在收取现金时,必须输入缴款.
        If mty_ModulePara.byt缴款输入控制 = 0 Then CheckCashValied = True: Exit Function
        If mbln连续结帐 Then CheckCashValied = True: Exit Function
        
        '问题号:109307,结帐金额=0时也要进行缴款检查
        If txtReceive.Text = "" Then
            Screen.MousePointer = 0
            MsgBox "你还未输入缴款金额,不能继续", vbExclamation, gstrSysName
            If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
            Exit Function
        Else
            If Val(GetCashSum) > 0 And Val(txtReceive.Text) < Val(GetCashSum) Then
                MsgBox "输入的缴款金额不足,请补充缴款金额!", vbExclamation, gstrSysName
                If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
                Exit Function
            ElseIf Val(GetCashSum) < 0 And Val(txtReceive.Text) > Val(GetCashSum) Then
                MsgBox "输入的退款金额不足,请补充退款金额!", vbExclamation, gstrSysName
                If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
                Exit Function
            End If
        End If

        CheckCashValied = True
        Exit Function
    End If
    
    '退款处理
    If dblMoney < Abs(GetCashSum) And RoundEx(dblMoney, 6) <> 0 Then
        Screen.MousePointer = 0
        MsgBox "输入的退款金额不足！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCashValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    
    Call SaveErrLog
End Function
Private Function ExcutePatiOutHosptial() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行病人出院操作
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-13 10:46:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOut As Boolean, rsTmp As ADODB.Recordset
    Dim bln门诊留观病人 As Boolean
    Dim lng当前科室id As Long
    
    On Error GoTo errHandle
    If mrsInfo.State = 0 Then Exit Function
    bln门诊留观病人 = Val(NVL(mrsInfo!病人性质)) = 1
    If Not mty_ModulePara.blnAutoOut Then ExcutePatiOutHosptial = True: Exit Function
    If mEditType = g_Ed_门诊结帐 And Not bln门诊留观病人 Or mblnCurMzBalanceNo Or mobjBalanceCon.blnCurBalanceOwnerFee Then ExcutePatiOutHosptial = True: Exit Function
    If mYBInFor.bytMCMode = 1 Then ExcutePatiOutHosptial = True: Exit Function
    
    '出院病人且出院结帐或在院病人且是中途结帐的,直接返回
    If bln门诊留观病人 Then
        If Val(NVL(mrsInfo!在院)) <> 1 Then ExcutePatiOutHosptial = True: Exit Function
        lng当前科室id = Val(NVL(mrsInfo!出院科室ID))
    Else
        If Not (Not IsNull(mrsInfo!当前科室id) And opt出院.Value) Then ExcutePatiOutHosptial = True: Exit Function
        lng当前科室id = Val("" & mrsInfo!当前科室id)
    End If
    '自动出院(出院结帐)
    blnOut = True
    If mYBInFor.intInsure <> 0 And Not MCPAR.未结清出院 Then
        Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , IIf(bln门诊留观病人, 1, 2))
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!费用余额, 0) <> 0 Then blnOut = False
        End If
    End If

    If gTy_System_Para.TY_Balance.bln医生允许才能出院 And blnOut Then
        If Not check医生下达出院医嘱(mrsInfo!病人ID, mrsInfo!主页ID) Then blnOut = False
    End If
    If Not blnOut Then ExcutePatiOutHosptial = True: Exit Function  '不允许出院，直接返回true
    
    frmAutoOut.mlng病人ID = mrsInfo!病人ID
    frmAutoOut.mlng主页ID = mrsInfo!主页ID
    frmAutoOut.mlngDepID = lng当前科室id
    frmAutoOut.mint险类 = mYBInFor.intInsure
    frmAutoOut.mstr性别 = NVL(mrsInfo!性别)
    frmAutoOut.Show 1, Me
    ExcutePatiOutHosptial = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckNotExcuteItemValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查未执行项目是否合法
    '出参:objSetFocus-不合法时,返回光标缺省定位控件
    '返回:合法返回返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt检查未执行 = 0 Then CheckNotExcuteItemValied = True: Exit Function
    
    strInfo = ExistWaitExe(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0))
    If strInfo = "" Then CheckNotExcuteItemValied = True: Exit Function
        
    If gTy_System_Para.TY_Balance.byt检查未执行 = 1 Then
        If MsgBox("发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & _
            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set objSetFocus = txtPatient
            Exit Function
        End If
    Else
        MsgBox "发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许出院结帐.", vbInformation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If
    CheckNotExcuteItemValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckNotSendDrug() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查未发药品是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, strInfo As String
    
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt检查未发药 = 0 Then CheckNotSendDrug = True: Exit Function
    strInfo = ExistWaitDrug(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0), 1)
    If strInfo = "" Then CheckNotSendDrug = True: Exit Function
    If gTy_System_Para.TY_Balance.byt检查未发药 = 1 Then
        If MsgBox("发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Function
        End If
    Else
        MsgBox "发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "不允许出院结帐。", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Function
    End If
    CheckNotSendDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMZNotExcuteItemValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查未执行项目是否合法
    '出参:objSetFocus-不合法时,返回光标缺省定位控件
    '返回:合法返回返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt门诊检查未执行 = 0 Then CheckMZNotExcuteItemValied = True: Exit Function
    
    strInfo = ExistWaitExe(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0), 1)
    If strInfo = "" Then CheckMZNotExcuteItemValied = True: Exit Function
        
    If gTy_System_Para.TY_Balance.byt门诊检查未执行 = 1 Then
        If MsgBox("发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & _
            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set objSetFocus = txtPatient
            Exit Function
        End If
    Else
        MsgBox "发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许门诊结帐.", vbInformation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If
    CheckMZNotExcuteItemValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMZNotSendDrug() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查未发药品是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, strInfo As String
    
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt门诊检查未发药 = 0 Then CheckMZNotSendDrug = True: Exit Function
    strInfo = ExistWaitDrug(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0), 1, 1)
    If strInfo = "" Then CheckMZNotSendDrug = True: Exit Function
    If gTy_System_Para.TY_Balance.byt门诊检查未发药 = 1 Then
        If MsgBox("发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Function
        End If
    Else
        MsgBox "发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "不允许门诊结帐。", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Function
    End If
    CheckMZNotSendDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelAppleyFeeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费申请检查
    '返回:退费申请检查合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mty_ModulePara.blnAutoOut Then CheckDelAppleyFeeValied = True: Exit Function
    If IsNull(mrsInfo!当前科室id) Or opt出院.Value = False Or mYBInFor.bytMCMode = 1 Then CheckDelAppleyFeeValied = True: Exit Function
    
    If GetUnAuditReFee(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0)) Then
        If MsgBox("病人" & txtPatient.Text & "存在已申请退费但未审核的记录,确定要进行出院结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    CheckDelAppleyFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Private Function 病历检查() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病历有效性检查
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 19:03:35
    '说明:30036(bug)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str病历原因 As String

    On Error GoTo errHandle
    mBalanceInfor.str病历原因 = ""

    If Not mty_ModulePara.bln结帐检查病历接收 Or opt出院.Value = False Then 病历检查 = True: Exit Function
    
    
    If IsCheck病历已接收(Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID))) Then 病历检查 = True: Exit Function
    
    If MsgBox("发现病人" & mrsInfo!姓名 & "没有进行病历审核," & _
        vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Function
    End If
    
    str病历原因 = ""
    If frmInputBox.InputBox(Me, "病历未接原因", "请输入病历未接原因信息:", 100, 3, True, False, str病历原因) = False Then
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Function
    End If
    mBalanceInfor.str病历原因 = str病历原因
    病历检查 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSaveBalanceSQL(ByRef cllBalaceData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次结帐的结帐相关的Sql
    '出参:cllBalaceData-结帐数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-13 11:10:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strSQL As String, dblMoney As Double, dblTemp As Double
    Dim str费用IDs As String, str保险信息 As String, intMaxTime As Integer
    Dim lngTmp As String, str费用ID  As String, strTemp As String, strNow As String
    Dim str住院次数 As String, cllPartBalance As Collection, strArray() As String
    Dim dblAvail As Double, cllTemp As Collection, intCounter As Integer, intCount As Integer
    Dim i As Long, dblTotal As Double
    
    On Error GoTo errHandle
    Set cllBalaceData = New Collection
    Set cllPartBalance = New Collection
    
    If mBalanceInfor.blnSaveBill = True Then GetSaveBalanceSQL = True: Exit Function
    
    Set cllTemp = New Collection
    '当前结帐信息
    With mBalanceInfor
        .strNO = zlDatabase.GetNextNo(15)
        .lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
        .dtBalanceDate = zlDatabase.Currentdate
    End With
    
    intInsure = mYBInFor.intInsure
    If intInsure <> 0 Then str保险信息 = IIf(mYBInFor.intInsure = 0, " ", mYBInFor.intInsure) & "," & NVL(mrsInfo!密码, " ") & "," & NVL(mrsInfo!医保号, " ")
    intMaxTime = 0
    intMaxTime = GetMinMaxTime(1)
    '1.病人结帐记录
    '问题:25596
    ' Zl_病人结帐记录_Insert
    strSQL = "zl_病人结帐记录_Insert("
    '  Id_In           病人结帐记录.ID%Type,
    strSQL = strSQL & "" & mBalanceInfor.lng结帐ID & ","
    '  单据号_In       病人结帐记录.NO%Type,
    strSQL = strSQL & "'" & mBalanceInfor.strNO & "',"
    '  病人id_In       病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & Val(NVL(mrsInfo!病人ID)) & ","
    '  收费时间_In     病人结帐记录.收费时间%Type,
    strSQL = strSQL & "To_Date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  开始日期_In     病人结帐记录.开始日期%Type,
    strSQL = strSQL & "" & IIf(IsDate(txtPatiBegin.Text), "To_Date('" & txtPatiBegin.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  结束日期_In     病人结帐记录.结束日期%Type,
    strSQL = strSQL & "" & IIf(IsDate(txtPatiEnd.Text), "To_Date('" & txtPatiEnd.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  中途结帐_In     病人结帐记录.中途结帐%Type := 0,
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐, 0, IIf(opt中途.Value, 1, 0)) & ","
    '  多病人结帐_In   Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  最大结帐次数_In Number := 0,
    strSQL = strSQL & "" & ZVal(intMaxTime) & ","
    '  备注_In         病人结帐记录.备注%Type := Null
    strSQL = strSQL & "" & IIf(Trim(txtBalance(Idx_结帐说明).Text) = "", "NULL", "'" & Trim(txtBalance(Idx_结帐说明).Text) & "'") & ","
    '   来源_In         Number := 1,1-门诊;2-住院
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐, 1, 2) & ","
    '  原因_In         病人结帐记录.原因%Type := Null
    strSQL = strSQL & "" & IIf(Trim(mBalanceInfor.str病历原因) = "", "NULL", "'" & Trim(mBalanceInfor.str病历原因) & "'") & ","
    '    结帐类型_In     病人结帐记录.结帐类型%type:=2
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐, 1, 2) & ","
    '  结算状态_In     病人结帐记录.结算状态%type:=0
    '结算状态:NULL-正常的结帐数据;1-异常的结帐或作废数据;2-正常作废的异常记录
    strSQL = strSQL & "" & 1 & ","
    ' 住院次数_In     病人结帐记录.住院次数%Type := Null,
    str住院次数 = ""
    str住院次数 = mobjBalanceCon.strTime
    If str住院次数 = "" Then str住院次数 = mobjBalanceAll.strAllTime
    strSQL = strSQL & "" & IIf(str住院次数 = "", "NULL", "'" & str住院次数 & "'") & ","
    ' 结帐金额_In     病人结帐记录.结帐金额%Type := Null
    strSQL = strSQL & "" & mBalanceInfor.dbl当前结帐 & ","
    ' 票据号_In     病人结帐记录.实际票号%Type := Null
    strSQL = strSQL & IIf(mblnPrintInvoice, IIf(txtInvoice.Text = "", "Null)", "'" & txtInvoice.Text & "')"), "Null)")

    zlAddArray cllBalaceData, strSQL
    
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据")) <> "" _
                And Not (Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))) = 0 And Val(.Cell(flexcpData, i, .ColIndex("未结金额"))) <> 0) Then
                If Val(.TextMatrix(i, .ColIndex("ID"))) = 0 Then
                    '  Zl_结帐费用记录_Insert
                    strSQL = "Zl_结帐费用记录_Insert("
                    '  Id_In       住院费用记录.ID%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("ID"))) & ","
                    '  No_In       住院费用记录.NO%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("单据")) & "',"
                    '  记录性质_In 住院费用记录.记录性质%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("记录性质"))) & ","
                    '  记录状态_In 住院费用记录.记录状态%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("记录状态"))) & ","
                    '  执行状态_In 住院费用记录.执行状态%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("执行状态"))) & ","
                    '  序号_In     住院费用记录.序号%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("序号"))) & ","
                    '  结帐金额_In 住院费用记录.结帐金额%Type,
                    strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))) & ","
                    '  结帐id_In   住院费用记录.结帐id%Type
                    strSQL = strSQL & "" & mBalanceInfor.lng结帐ID & ")"
                    zlAddArray cllTemp, strSQL
                Else
                    If Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))) = Val(.Cell(flexcpData, i, .ColIndex("未结金额"))) Then
                        str费用IDs = str费用IDs & Val(.TextMatrix(i, .ColIndex("ID"))) & ","
                    Else
                        '  Zl_结帐费用记录_Insert
                        strSQL = "Zl_结帐费用记录_Insert("
                        '  Id_In       住院费用记录.ID%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("ID"))) & ","
                        '  No_In       住院费用记录.NO%Type,
                        strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("单据")) & "',"
                        '  记录性质_In 住院费用记录.记录性质%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("记录性质"))) & ","
                        '  记录状态_In 住院费用记录.记录状态%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("记录状态"))) & ","
                        '  执行状态_In 住院费用记录.执行状态%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("执行状态"))) & ","
                        '  序号_In     住院费用记录.序号%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("序号"))) & ","
                        '  结帐金额_In 住院费用记录.结帐金额%Type,
                        strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))) & ","
                        '  结帐id_In   住院费用记录.结帐id%Type
                        strSQL = strSQL & "" & mBalanceInfor.lng结帐ID & ")"
                        zlAddArray cllBalaceData, strSQL
                    End If
                End If
            End If
        Next i
    End With
    
    While str费用IDs <> ""
        If Len(str费用IDs) > 3998 Then
            lngTmp = InStrRev(Mid(str费用IDs, 1, 3998), ",")
            str费用ID = Mid(str费用IDs, 1, lngTmp - 1)
            str费用IDs = Mid(str费用IDs, lngTmp + 1)
        Else
            str费用ID = Mid(str费用IDs, 1, Len(str费用IDs) - 1)
            str费用IDs = ""
        End If
        strSQL = "zl_结帐费用记录_Batch('" & str费用ID & "'," & mrsInfo!病人ID & "," & mBalanceInfor.lng结帐ID & ")"
        zlAddArray cllBalaceData, strSQL
    Wend
    '负数记帐-->结帐作废-->对负数记帐销帐-->再次结帐时，会造成不能结帐的处理：原因是先12或13记录时，检查了结帐总额与实收是否一致，由于未处理性质为2或3的记录,造成统计的金额有问题,
    '现处理方式:现需要先处理费用未结的，然后再处理12或13的记录
 
    For i = 1 To cllTemp.Count
        zlAddArray cllBalaceData, cllTemp(i)
    Next
    GetSaveBalanceSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function



Private Function CheckBusinessRuleIsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查业务规则是否合法
    '出参:objSetFocus-不合法时,光标缺省定位到哪个控件
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-12 18:12:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intState As Integer, strTime As String, i As Long, strTmp As String
       
    
    On Error GoTo errHandle
    '门诊结帐,直接返回true
    If mYBInFor.bytMCMode <> 1 And mEditType <> g_Ed_门诊结帐 Then
        intState = GetPatientState
        If mYBInFor.intInsure <> 0 And opt出院.Value Then
            If MCPAR.出院结算必须出院 And intState <> 0 Then
                '问题号:115055,焦博,2017/10/16,检查数据合法性时会报错
                If IsNull(mrsInfo!当前科室id) Then
                    MsgBox "病人在结帐期间被撤销出院,医保病人出院结帐前必须先出院！", vbInformation, gstrSysName
                Else
                    MsgBox "医保病人出院结帐前必须先出院！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
          
        '是否在院
        If gTy_System_Para.TY_Balance.bln在院不准结帐 And opt出院.Value And (intState = 1 Or intState = 2) Then '  ' 30572:预出院也是在院.
            MsgBox "当前病人在院，不允许出院结帐。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '检查是否还有代收费用未退还病人
        If opt出院.Value = True Then
            If PatiHaveStorage(mrsInfo!病人ID) Then Exit Function
        End If
                      
        'bytAuditing:0-不检查,1-检查并提示,2-检查并禁止
        '问题:37369:中途结帐不检查
        With gTy_System_Para.TY_Balance
            If .bytAuditing <> 0 And opt出院.Value Then
                If HaveNOAuditing(mrsInfo!病人ID, mobjBalanceCon.strTime) Then
                    If .bytAuditing = 1 Then
                        '在读取病人信息时,已经检查了
                    ElseIf .bytAuditing = 2 Then
                         Call MsgBox("该病人还存在未审核的记帐费用,禁止结帐!", vbInformation + vbOKOnly, gstrSysName)
                         Exit Function
                    End If
                End If
            End If
        End With
          
        '需要再次检查,以防结帐期间已审核的病人被取消审核
        If (InStr(mstrPrivs, ";未审核病人中途结帐;") = 0 And opt中途.Value _
            Or InStr(mstrPrivs, ";未审核病人出院结帐;") = 0 And opt出院.Value) _
                And mEditType = g_Ed_住院结帐 Then
            strTime = IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime)
            If strTime <> "" Then
                For i = 0 To UBound(Split(strTime, ","))
                    strTmp = Split(strTime, ",")(i)
                    If Val(strTmp) <> 0 Then
                        If Not Chk病人审核(mrsInfo!病人ID, Val(strTmp)) Then
                            MsgBox "待结帐费用中包含病人第" & strTmp & "次住院未审核的费用记录。" & vbCrLf & _
                                "你不能对未审核的费用进行结帐！", vbInformation, gstrSysName
                            If cmdMore.Visible And cmdMore.Enabled Then cmdMore.SetFocus
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    If mEditType = g_Ed_门诊结帐 Then
        If CheckMZNotExcuteItemValied(objSetFocus) = False Then Exit Function   '检查未执行项目是否合法
        If CheckMZNotSendDrug = False Then Exit Function '检查未发药品
    Else
        If opt出院.Value Then
            If CheckNotExcuteItemValied(objSetFocus) = False Then Exit Function   '检查未执行项目是否合法
            If CheckNotSendDrug = False Then Exit Function '检查未发药品
        End If
    End If
    
    If Not CheckDelAppleyFeeValied Then Exit Function '检查退费申请的合法性
    If 病历检查 = False Then Exit Function
    CheckBusinessRuleIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetPatientState() As Integer
'功能:获取病人状态
'返回:0-出院,1-在院,2-预出院,-1-访问数据库出错
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    GetPatientState = -1
    On Error GoTo errH
    strSQL = "Select a.当前科室id, a.主页id As 最后主页id, b.主页id, b.状态" & vbNewLine & _
            "From 病人信息 A, 病案主页 B" & vbNewLine & _
            "Where a.病人id = b.病人id And Nvl(b.主页id, 0) = (Select Max(Column_Value) From Table(f_str2list([2]))) And b.病人id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(mrsInfo!病人ID), IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime))
    
    If rsTmp.RecordCount > 0 Then
        If Val(NVL(rsTmp!最后主页ID)) > Val(NVL(rsTmp!主页ID)) Then
            GetPatientState = 0
        Else
            If Not IsNull(rsTmp!当前科室id) Then
                If Val("" & rsTmp!状态) = 3 Then
                    GetPatientState = 2
                Else
                    GetPatientState = 1
                End If
            Else
                GetPatientState = 0
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub WriteZYInforToCard(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional blnDelete As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将住院信息写入卡中
    '入参:blnDelete-是否退费
    '编制:刘兴洪
    '日期:2015-01-13 11:04:01
    '问题:56615
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strExpend As String
    Dim objCard As Card

    On Error GoTo errHandle
        
    '未确定刷卡类别,直接退出
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then
        If InStr(1, mstrPrivs, ";门诊信息写卡;") = 0 Then Exit Sub
    Else
        If InStr(1, mstrPrivs, ";住院信息写卡;") = 0 Then Exit Sub
    End If
    If lng病人ID = 0 Then Exit Sub
    
    If mlngCardTypeID = 0 Then
        If blnDelete Then GoTo goDelete:
        Exit Sub
    End If
    
    If IDKind.GetCurCard.接口序号 = mlngCardTypeID Then
        Set objCard = IDKind.GetCurCard
    Else
        Set objCard = IDKind.GetIDKindCard(mlngCardTypeID, CardTypeID)
    End If
    
    If objCard Is Nothing Then Exit Sub
    If objCard.是否写卡 = False Or objCard.接口序号 <= 0 Then Exit Sub '不准写卡的,不调用接口
    lngCardTypeID = objCard.接口序号
goDelete:
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then
        Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng病人ID, lng结帐ID, strExpend)
    Else
        Call gobjSquare.objSquareCard.zlZyInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng病人ID, lng结帐ID, strExpend)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function ExistWaitDrug(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal lng检查离院带药 As Long = 0, Optional ByVal int门诊标志 As Integer) As String
'功能：检查病人在药房是否还有未发药的药品或卫材
'返回：药房和发料部门名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(1,[1],[2],-1,[3],[4]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitDrug", lng病人ID, lng主页ID, lng检查离院带药, int门诊标志)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = NVL(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function IsCheck病历已接收(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病历是否已经接收
    '入参:
    '出参:
    '返回:已接收返回True,否则返回False
    '编制:刘兴洪
    '日期:2010-05-24 16:39:47
    '说明;
    '     问题:30036
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "select nvl(信息值,0) as 病历接收 from 病案主页从表 where 病人id=[1] and 主页id=[2] and 信息名='病历接收'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
            IsCheck病历已接收 = Val(NVL(rsTemp!病历接收)) = 1
    Else
            IsCheck病历已接收 = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function PatiHaveStorage(ByVal lng病人ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    strSQL = "Select A.结算方式,Sum(A.金额) as 金额" & _
        " From 病人预交记录 A,结算方式 B" & _
        " Where A.记录性质=1 And A.结算方式=B.名称 And B.性质=5 And A.病人ID=[1]" & _
        " Group by A.结算方式 Having Sum(A.金额)<>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID)
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            strMsg = strMsg & vbCrLf & rsTmp!结算方式 & "：" & Format(rsTmp!金额, "0.00")
            rsTmp.MoveNext
        Loop
    End If
    If strMsg <> "" Then
        If mty_ModulePara.byt结帐检查代收款项 = 1 Then
            If MsgBox("还有以下代收费用没有退还病人：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "要继续结帐吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                PatiHaveStorage = True
            Else
                PatiHaveStorage = False
            End If
        Else
            MsgBox "还有以下代收费用没有退还病人：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "请先将费用退还给病人再结帐。", vbInformation, gstrSysName
            PatiHaveStorage = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveBalaceCharge(ByVal bln预交 As Boolean, ByRef tyBrushCard As TY_BrushCard, _
    Optional ByRef blnChargeEnd As Boolean, _
    Optional ByRef objSetFocus As Object, _
    Optional ByVal objInCard As Card, _
    Optional ByVal lngRow As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结算数据
    '入参:bln预交-当前是缴预交款
    '出参:blnChargeEnd-收费完成操作(true,完成收费;False-还未完成)
    '     objSetFocus-结算失败时,缺省定位光标位置
    '返回:保存结算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 10:35:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str消费卡结算 As String, str收费结算 As String, strSQL As String
    Dim strCardNo As String, strErrMsg As String
    Dim blnHaveMoney As Boolean, blnFind As Boolean, blnTrans As Boolean
    Dim dbl剩余金额 As Double, dblTemp As Double, dbl未付金额 As Double
    Dim dblMoney As Double, dbl退支票额 As Double, bln存预交 As Boolean
    Dim i As Long, j As Long, varData As Variant, cllDeposit As Collection
    Dim cllUpdate As Collection, cllThreeSwap As Collection, cllPro As Collection
    Dim objCard As Card, lng病人ID As Long
    Dim intSign As Integer, rsTmp As ADODB.Recordset
    Dim bytSign As Byte, str住院次数 As String
    Dim intMousePointer As Integer
    Dim strArray() As String, k As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    '检查当前条件是否有效
    blnChargeEnd = False
    If objInCard Is Nothing Then
        Set objCard = IDKindPaymentsType.GetCurCard
    Else
        Set objCard = objInCard
    End If
    
    lng病人ID = mPatiInfor.lng病人ID
    
    With mBalanceInfor
        .dbl缴款 = 0: .dbl找补 = 0
        .dbl现金 = 0
    End With
    
    If Not bln预交 Then
'        bln存预交 = IDKind找补.GetCurCard.接口序号 <> 1
    End If
    
    If bln预交 Then
        dblMoney = Val(txtBalance(Idx_冲预交).Text)
        If dblMoney <> mBalanceInfor.dbl冲预交合计 Then Exit Function
        dbl剩余金额 = RoundEx(mBalanceInfor.dbl未付合计 - mBalanceInfor.dbl冲预交合计, 6)
    Else
        If lngRow = 0 Then
            dblMoney = RoundEx(Val(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算金额"))), 6)
            mBalanceInfor.dbl缴款 = Val(txtReceive.Text)
            dblTemp = dbl未付金额: dbl剩余金额 = 0
            dblMoney = GetCentMoney(dblTemp)
            mBalanceInfor.dbl现金 = dblMoney
            dbl剩余金额 = 0
        Else
            dblMoney = RoundEx(Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("结算金额"))), 6)
            dbl剩余金额 = Val(mBalanceInfor.dbl未付合计)
        End If
    End If
    
    Call Show误差金额(bln预交)
    
    '误差不能大于1.5块钱
    If Abs(mBalanceInfor.dbl误差额) > 1.5 Then
        Screen.MousePointer = 0
        Call MsgBox("误差过大,请检查是否正确!", vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    
    If dbl剩余金额 > 0 Then blnHaveMoney = True
  
 
    Set cllPro = New Collection: Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    If GetSaveBalanceSQL(cllPro) = False Then Exit Function
    
    If Not bln预交 And Not objCard Is Nothing Then
        '执行一卡通(老版)接口
        If ExecuteOldOneCardPayInterface(lng病人ID, mBalanceInfor.lng结帐ID, objCard, dblMoney, tyBrushCard, cllPro) = False Then Exit Function
        '执持三方帐户交易接口
        If tyBrushCard.bln转帐 Then
            If ExecuteThreeSwapTransferPay(objCard, dblMoney, cllPro, tyBrushCard) = False Then Exit Function
        Else
            If ExecuteThreeSwapPayInterface(lng病人ID, mBalanceInfor.lng结帐ID, objCard, dblMoney, cllPro, tyBrushCard) = False Then Exit Function
        End If
    End If
    
    If dbl剩余金额 = 0 And bln预交 = False Then
        '完成结帐
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("组合信息")) <> "" And Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                    If Val(.TextMatrix(i, .ColIndex("是否转账"))) = 0 Then
                        If Val(.Cell(flexcpData, i, .ColIndex("组合信息"))) = 1 Then
                            If ExecuteThreeSwapDelSingle(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("卡类别ID")), CardTypeID), _
                                                     RoundEx(-1 * .TextMatrix(i, .ColIndex("结算金额")), 2), .Cell(flexcpData, i, .ColIndex("卡号")), _
                                                    .TextMatrix(i, .ColIndex("交易说明")), .TextMatrix(i, .ColIndex("交易流水号")), _
                                                     Val(.TextMatrix(i, .ColIndex("组合信息"))), cllPro) = False Then
                                '接口失败

                                For k = 1 To vsDeposit.Rows - 1
                                    If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("预交ID"))) = Val(.TextMatrix(i, .ColIndex("组合信息"))) Then
                                        vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbRed
                                    End If
                                Next k
                                Exit Function
                            Else
                                '接口成功


                                For k = 1 To vsDeposit.Rows - 1
                                    If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("预交ID"))) = Val(.TextMatrix(i, .ColIndex("组合信息"))) Then
                                        vsDeposit.TextMatrix(k, vsDeposit.ColIndex("编辑状态")) = 1
                                        vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbBlack
                                    End If
                                Next k
                            End If
                        Else
                            strArray = Split(.TextMatrix(i, .ColIndex("组合信息")), "|")
                            If ExecuteThreeSwapDelBatch(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("卡类别ID")), CardTypeID), _
                                                         RoundEx(-1 * Val(.TextMatrix(i, .ColIndex("结算金额"))), 2), .TextMatrix(i, .ColIndex("组合信息")), _
                                                        cllPro) = False Then
                                '接口失败
                                For j = 0 To UBound(strArray)
                                    For k = 1 To vsDeposit.Rows - 1
                                        If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("预交ID"))) = Val(Split(strArray(j), ",")(4)) Then
                                            vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbRed
                                        End If
                                    Next k
                                Next j
                                Exit Function
                            Else
                                '接口成功
                                For j = 0 To UBound(strArray)
                                    For k = 1 To vsDeposit.Rows - 1
                                        If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("预交ID"))) = Val(Split(strArray(j), ",")(4)) Then
                                            vsDeposit.TextMatrix(k, vsDeposit.ColIndex("编辑状态")) = 1
                                            vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbBlack
                                        End If
                                    Next k
                                Next j
                            End If
                        End If
                        mBalanceInfor.blnSaveBill = True
                        .TextMatrix(i, .ColIndex("组合信息")) = ""
                        .TextMatrix(i, .ColIndex("结算状态")) = 1
                        .TextMatrix(i, .ColIndex("编辑状态")) = 0
                    Else
                        If CheckThreeSwapValied(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("卡类别ID")), CardTypeID), Val(.TextMatrix(i, .ColIndex("结算金额"))), tyBrushCard, True) = False Then Exit Function
                        If ExecuteThreeSwapTransferPay(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("卡类别ID")), CardTypeID), Val(.TextMatrix(i, .ColIndex("结算金额"))), cllPro, tyBrushCard) = False Then Exit Function
                        mBalanceInfor.blnSaveBill = True
                        .TextMatrix(i, .ColIndex("组合信息")) = ""
                        .TextMatrix(i, .ColIndex("结算状态")) = 1
                        .TextMatrix(i, .ColIndex("编辑状态")) = 0
                    End If
                End If
            Next i
        End With
        '处理冲预交
        If GetSaveStrickDepositSQL(cllDeposit, objSetFocus) = False Then Exit Function
        For i = 1 To cllDeposit.Count
            cllPro.Add cllDeposit(i)
        Next
        
        If ExcuteBalanceEnd(dbl退支票额, cllPro) = False Then Exit Function
        
        If opt出院.Value = True And mEditType = g_Ed_住院结帐 Then
            '出院结帐,检查是否结清
            Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 2)
            If Not rsTmp Is Nothing Then
                '结清,调整自动记帐标志
                If NVL(rsTmp!费用余额, 0) = 0 Then
                    str住院次数 = ""
                    str住院次数 = mobjBalanceCon.strTime
                    If str住院次数 = "" Then str住院次数 = mobjBalanceAll.strAllTime
                    If str住院次数 <> "" Then
                        strSQL = "zl_病人自动记帐_Stop(" & mrsInfo!病人ID & ",'" & str住院次数 & "')"
                        zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    End If
                End If
            End If
        End If
        
        '打印票据
        Call PrintBill(mPatiInfor.lng病人ID, mBalanceInfor.strNO, mBalanceInfor.lng结帐ID, mBalanceInfor.dtBalanceDate, mBalanceInfor.dbl缴款, mBalanceInfor.dbl找补)
        '81697:李南春,2015/6/8,评价器
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.InPatiCashierAfter(mPatiInfor.lng病人ID, mBalanceInfor.lng结帐ID)
            Err.Clear
        End If
        
        If Not mbln连续结帐 Then
            Call ExcutePatiOutHosptial '病人出院
        End If
        '住院信息写卡:56615
        Call WriteZYInforToCard(mPatiInfor.lng病人ID, mBalanceInfor.lng结帐ID)

        zlDatabase.SetPara "默认出院结帐", IIf(opt出院.Value = True, "1", "0"), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        
        blnChargeEnd = True
        If mEditType = g_Ed_重新结帐 Then Unload Me: Exit Function
        SaveBalaceCharge = True
        Exit Function
    End If
NextBalance:
    Err = 0: On Error GoTo errHandle:
GoEnd:
    Set objSetFocus = txtReceive
    txtReceive.Text = ""
    Call LoadCurOwnerPayInfor
    Call LedDisplayBank
    
    SaveBalaceCharge = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then
            Screen.MousePointer = intMousePointer
            Resume
        End If
    End If
End Function
 

Public Function Get消费卡结算方式() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡结算信息
    '编制:刘兴洪
    '日期:2015-01-30 15:29:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str消费卡结算 As String, i As Long
    Dim objCard As Card
    Dim lngCardTypeID As Long
    
    On Error GoTo errHandle
 
    str消费卡结算 = ""  '卡类别ID|卡号|消费卡ID|消费金额||....
    With vsBlance
       For i = 1 To .Rows - 1
           '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
           '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
           '结算状态:是否已结算:1-已结算;0-未结算
           If Val(.TextMatrix(i, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                str消费卡结算 = str消费卡结算 & "||" & Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                str消费卡结算 = str消费卡结算 & "|" & Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
                str消费卡结算 = str消费卡结算 & "|" & Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                str消费卡结算 = str消费卡结算 & "|" & RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
           End If
       Next
    End With
'
'    If Not mcllCurSquareBalance Is Nothing Then
'        For i = 1 To mcllCurSquareBalance.Count
'            'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
'            str消费卡结算 = str消费卡结算 & "||" & Val(mcllCurSquareBalance(i)(0))
'            str消费卡结算 = str消费卡结算 & "|" & Trim(mcllCurSquareBalance(i)(3))
'            str消费卡结算 = str消费卡结算 & "|" & Val(mcllCurSquareBalance(i)(1))
'            str消费卡结算 = str消费卡结算 & "|" & RoundEx(Val(mcllCurSquareBalance(i)(2)), 6)
'        Next
'    End If
    If str消费卡结算 <> "" Then str消费卡结算 = Mid(str消费卡结算, 3)
    Get消费卡结算方式 = str消费卡结算
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function Get普通结算方式() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费结算数据
    '返回:收费用结算方式,格式如下:
    '       结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
    '编制:刘兴洪
    '日期:2015-01-14 16:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, i As Long, int性质 As Integer
    Dim strBalance As String, dblMoney As Double, varData As Variant
    Dim objCard As Card, objTempCard As Card
    Dim bln存预交 As Boolean
    '结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
    '收费完成
    strBalance = ""
    With vsBlance
        For i = .Rows - 1 To 1 Step -1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If str结算方式 <> "" And int性质 = 0 Then
                strBalance = strBalance & "||" & str结算方式
                strBalance = strBalance & "|" & Val(.TextMatrix(i, .ColIndex("结算金额")))
                strBalance = strBalance & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("结算号码"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("结算号码"))))
                strBalance = strBalance & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("备注"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("备注"))))
            End If
        Next
        Set objCard = IDKindPaymentsType.GetCurCard
'        Set objTempCard = IDKind找补.GetCurCard
        
        bln存预交 = Not objTempCard Is Nothing
        If bln存预交 Then
            bln存预交 = objTempCard.接口序号 > 1
        End If
        
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    Get普通结算方式 = strBalance
    
End Function
 
Private Function ExcuteBalanceEnd(ByVal dbl退支票 As Double, _
    ByVal cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐结束操作
    '入参:dbl退支票-当前退支票金额
    '     cllPro-结帐数据
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 16:06:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllPro As Collection, strSQL As String
    Dim lng结帐ID As Long, str普通结算 As String, str消费卡结算 As String
    Dim dbl预交 As Double, int预交类别 As Integer
    Dim lng病人ID As Long, lng主页ID As Long
    Dim dblMoney As Double
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    str普通结算 = Get普通结算方式
    str消费卡结算 = Get消费卡结算方式
    
    lng病人ID = mPatiInfor.lng病人ID
    lng结帐ID = mBalanceInfor.lng结帐ID
    lng主页ID = mPatiInfor.lng主页ID
    
    On Error GoTo errHandle
    
    If str消费卡结算 <> "" Then
        '调用之前,先处理数据
        'Zl_病人结帐结算_Modify
        strSQL = "Zl_病人结帐结算_Modify("
        '  操作类型_In     Number,
        '  --操作类型_In:
        '--   3-消费卡结算:
        '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
        strSQL = strSQL & "3,"
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "'" & str消费卡结算 & "',"
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "NULL,"
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  结帐类型_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0
        strSQL = strSQL & "" & IIf(str普通结算 = "", 1, "0") & " )"
        zlAddArray cllPro, strSQL
    End If
    
    If str普通结算 <> "" Or str消费卡结算 = "" Then
         '调用之前,先处理数据
         'Zl_病人结帐结算_Modify
        strSQL = "Zl_病人结帐结算_Modify("
        '  操作类型_In     Number,
        '  --操作类型_In:
        '--   0-普通收费方式:
        '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        strSQL = strSQL & "0,"
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "'" & str普通结算 & "',"
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "" & dbl退支票 & ","
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & IIf(InStr(mstrForceNote, "强制退现") + 4 = Len(mstrForceNote), "", mstrForceNote) & "',"
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl缴款 & ","
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "" & IIf(txtCaculated.ForeColor = vbRed, txtCaculated.Text, 0) & ","
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl误差额 & ","
        '  结帐类型_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
         strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        
        '  完成结算_In Number:=0
        strSQL = strSQL & "1)"
        zlAddArray cllPro, strSQL
    End If
    If GetSaveAddDepositSQL(lng病人ID, lng主页ID, mBalanceInfor.lng结帐ID, cllPro) = False Then Exit Function
    
    '异常记录时间处理
    If mEditType = g_Ed_重新结帐 Then
        strSQL = "Zl_病人结帐异常_Update("
        strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        strSQL = strSQL & "" & lng结帐ID & ")"
        zlAddArray cllPro, strSQL
    End If
    
    Err = 0: On Error GoTo ErrTrans:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    If mbln连续结帐 Then
        mPatiInfor.dbl未付累计 = RoundEx(mPatiInfor.dbl未付累计 + Val(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算金额"))), 6)
        mPatiInfor.bln连续结帐 = mbln连续结帐
        Set mPatiInfor.objCard = IDKindPaymentsType.GetCurCard
    Else
        mPatiInfor.dbl未付累计 = 0
        mPatiInfor.bln连续结帐 = False
        Set mPatiInfor.objCard = Nothing
    End If
    
    ExcuteBalanceEnd = True
    Exit Function
ErrTrans:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetSaveAddDepositSQL(ByVal lng病人ID As Long, lng主页ID As Long, _
     ByVal lng结帐ID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存预交的SQL
    '入参:lng结帐ID-存入预交对应的结帐ID
    '出参:cllPro-将保存预交的SQL增加到该集合中
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-30 13:46:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, int预交类别 As Integer, str结算方式 As String
    Dim dblMoney As Double, strSQL As String
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then GetSaveAddDepositSQL = True: Exit Function
    If objCard.接口序号 <= 1 Then GetSaveAddDepositSQL = True: Exit Function
    If IDKindPaymentsType.GetCurCard Is Nothing Then Exit Function
        
    str结算方式 = IDKindPaymentsType.GetCurCard.结算方式
    
    int预交类别 = objCard.接口序号 - 1    '1-门诊预交;2-住院预交
    
    '存为预交款
    dblMoney = RoundEx(Val(IIf(lblCaculated.Caption = "找补", txtCaculated.Text, 0)), 6)
    If dblMoney < 0 Then Exit Function
    
    mBalanceInfor.lng预交ID = zlDatabase.GetNextId("病人预交记录")
    mBalanceInfor.str预交No = zlDatabase.GetNextNo(11)
    
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & mBalanceInfor.lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & mBalanceInfor.str预交No & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "'" & mstrDepositInvioce & "',"
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  主页id_In     病人预交记录.主页id%Type,:42329
    If int预交类别 = 2 Then
       strSQL = strSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    Else
       strSQL = strSQL & "NULL,"
    End If
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "" & UserInfo.部门ID & ","
    
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSQL = strSQL & "Null,"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSQL = strSQL & "Null,"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSQL = strSQL & "Null,"
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'结帐存预交',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(mlng预交领用ID = 0, "NULL", mlng预交领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & int预交类别 & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "Null,"
    '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & "Null,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "Null,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "Null,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "Null,"
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strSQL = strSQL & "Null,"
    '  收款时间_In   病人预交记录.收款时间%Type := Null
    strSQL = strSQL & "to_date('" & mBalanceInfor.dtBalanceDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '   操作类型_In Integer:=0 :0-正常缴预交;1-存为划价单
    strSQL = strSQL & "0,"
    '  结帐id_In     病人预交记录.结帐id%Type >0时,表示某次结帐时,同步产生的预交记录
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算性质_In     病人预交记录.结算性质%Type >0时,结帐产生的预交款,性质为2
    strSQL = strSQL & "" & 12 & ")"
    zlAddArray cllPro, strSQL
    GetSaveAddDepositSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function zlCheckMulitInterfaceNumValied(Optional bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是正同时存在两种以上接口(不含两种)
    '返回:不含两种以上接口的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int性质 As Integer, str结算方式 As String
    Dim varData As Variant, strErrMsg As String
    Dim objCard As Card
    Dim intMousePointer As Integer
    On Error GoTo errHandle
    strErrMsg = ""
    intMousePointer = Screen.MousePointer
    Set objCard = IDKindPaymentsType.GetCurCard
    
    If objCard Is Nothing Then zlCheckMulitInterfaceNumValied = True: Exit Function
        
    If bln预交 Or objCard.接口序号 <= 0 Then zlCheckMulitInterfaceNumValied = True:        Exit Function
    
   '医保算一个接口
   If mYBInFor.intInsure <> 0 And mBalanceInfor.blnSaveBill Then intCount = intCount + 1: strErrMsg = strErrMsg & "医保结算:" & Format(mBalanceInfor.dbl医保支付合计, gstrDec)
   With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
 
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If InStr("34", int性质) > 0 Then
                intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str结算方式 & ":" & .TextMatrix(i, .ColIndex("结算金额"))
            End If
        Next
    End With
    If intCount > 2 Then
        Screen.MousePointer = 0
        Call MsgBox("注意:" & vbCrLf & "   本系统目前只支持两种以下接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function

Private Sub Show误差金额(ByVal bln预交 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示误差金额
    '入参:bln预交-预交额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 11:33:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl退支付额 As Double
    Dim dbl剩余金额 As Double, dblTemp As Double, dbl未付款 As Double
    Dim intSign As Integer, objCard As Card
    Dim i As Long, lngError As Long
    
    If mEditType = g_Ed_单据查看 Then Exit Sub
    
    dblMoney = Val(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算金额")))
    With mBalanceInfor
        .dbl误差额 = 0
        dbl未付款 = RoundEx(dblMoney + (.dbl当前结帐 - .dbl已付合计 - .dbl冲预交合计), 6)
    End With
    
    dbl退支付额 = 0: dbl剩余金额 = RoundEx(dbl未付款 - dblMoney, 6)
    
    If bln预交 Then
        '输入预交时
        mBalanceInfor.dbl误差额 = RoundEx(dbl未付款 - RoundEx(dbl未付款, 2), 6): GoTo Show误差:
        Exit Sub
    End If
    Set objCard = GetCard(vsBlance.TextMatrix(1, vsBlance.ColIndex("结算方式")))
    If Not objCard Is Nothing Then
        If objCard.结算性质 = 1 Then
            dblTemp = dbl未付款: dbl剩余金额 = 0
            dblMoney = GetCentMoney(dblTemp)
            mBalanceInfor.dbl误差额 = RoundEx(dbl未付款 - dblMoney, 6)
            GoTo Show误差:
        End If
    End If
    mBalanceInfor.dbl误差额 = RoundEx(dbl未付款 - RoundEx(dbl未付款, 2), 6): GoTo Show误差:
    
    If mYBInFor.intInsure <> 0 And mBalanceInfor.blnSaveBill = False Then mBalanceInfor.dbl误差额 = 0
Show误差:
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("类型"))) = 9 Then
                lngError = i
                Exit For
            End If
        Next i
        
        If mBalanceInfor.dbl误差额 = 0 Then
            If lngError <> 0 Then
                Call DeletePayInfor(lngError, True)
            End If
            Exit Sub
        End If
        
        If lngError <> 0 Then
            .TextMatrix(lngError, .ColIndex("结算金额")) = FormatEx(mBalanceInfor.dbl误差额, 6, , , 2)
        Else
            If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "" Then
                .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "误差费"
                .TextMatrix(.Rows - 1, .ColIndex("编辑状态")) = 0
                .TextMatrix(.Rows - 1, .ColIndex("类型")) = 9
                .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = FormatEx(mBalanceInfor.dbl误差额, 6, , , 2)
                .Rows = .Rows + 1
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "误差费"
                .TextMatrix(.Rows - 1, .ColIndex("编辑状态")) = 0
                .TextMatrix(.Rows - 1, .ColIndex("类型")) = 9
                .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = FormatEx(mBalanceInfor.dbl误差额, 6, , , 2)
            End If
        End If
    End With
End Sub
Private Function CheckSquareBalanceValied(ByVal objCard As Card, ByRef tyBrushCard As TY_BrushCard, _
                                        Optional ByVal dblInMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡结算交易检查
    '入参:objCard-三方卡
    '出参:dblMoney-当前刷卡金额
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl帐户余额 As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln退现 As Boolean, dbl未付金额 As Double
    Dim intMousePointer As Integer, strXmlIn As String
    Dim lng消费卡ID As Long, str卡号 As String, str密码 As String
    Dim str限制类别 As String, byt是否密文   As Byte
    Dim cllBushSquare As Collection, i As Long
    
    
    intMousePointer = Screen.MousePointer
    If objCard Is Nothing Then CheckSquareBalanceValied = True: Exit Function

    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then CheckSquareBalanceValied = True: Exit Function
    
    On Error GoTo errHandle
    
    tyBrushCard = strBrushCard
    If dblInMoney <> 0 Then
        dblMoney = dblInMoney
    Else
        dblMoney = Val(txtReceive.Text)
    End If
 
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "收款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
 
    '先检查对应的接口
    If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Then
        If zlGetClassMoney(0, rsMoney) = False Then Exit Function
    Else
        If zlGetClassMoney(mBalanceInfor.lng结帐ID, rsMoney) = False Then Exit Function
    End If
    
     '构建消费卡的刷卡信息
     Set cllSquareBalance = New Collection
     Set mcllCurSquareBalance = New Collection
     With vsBlance
        For i = 1 To .Rows - 1
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
            '结算状态:是否已结算:1-已结算;0-未结算
            lng消费卡ID = Val(.TextMatrix(i, .ColIndex("消费卡ID")))
            
            If Val(.TextMatrix(i, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = objCard.接口序号 _
                And Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 And lng消费卡ID <> 0 Then
              
                dblTemp = RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
                str卡号 = Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
                str密码 = Trim(.Cell(flexcpData, i, .ColIndex("消费卡ID")))  '密码
                str限制类别 = Trim(.Cell(flexcpData, i, .ColIndex("卡类别ID")))  '限制类别
                byt是否密文 = Val(.TextMatrix(i, .ColIndex("是否密文")))
                
                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                cllSquareBalance.Add Array(objCard.接口序号, lng消费卡ID, dblTemp, str卡号, str密码, str限制类别, byt是否密文)
            End If
        Next
     End With
     For i = 1 To cllSquareBalance.Count
        mcllCurSquareBalance.Add cllSquareBalance(i)
     Next
     
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "", _
        Optional ByVal byt业务场合 As Byte = 1, _
        Optional ByVal str费用来源 As String, _
        Optional ByVal lng病人ID As Long) As Boolean
    'varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
    strXmlIn = "<IN><CZLX>0</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, _
            objCard.接口序号, objCard.消费卡, _
            "" & mPatiInfor.str姓名, "" & mPatiInfor.str性别, "" & mPatiInfor.str年龄, dblMoney, _
            tyBrushCard.str卡号, tyBrushCard.str密码, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn, _
            GetFeeFromType(), mPatiInfor.lng病人ID) = False Then Exit Function
       
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.接口序号, _
        objCard.消费卡, tyBrushCard.str卡号, dblMoney, "", strXMLExpend) = False Then Exit Function
    '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    '    ByVal strCardTypeID As Long, _
    '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    'If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
          tyBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡) = False Then Exit Function
    '已经更改了结算金额
    
      
    Set mcllCurSquareBalance = cllSquareBalance
    
    Call AddSquareBalance(objCard)
'    If RoundEx(dblMoney, 6) <> Val(txtReceive.Text) Then
'        txtReceive.Text = Format(dblMoney, "0.00")
'    End If
    CheckSquareBalanceValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSquareDelValied(ByVal objCard As Card, _
     ByRef tyBrushCard As TY_BrushCard, _
     Optional ByVal lng消费卡ID As Long, _
     Optional dblDelMoney As Double _
     ) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡退费检查
    '入参:objCard-三方卡
    '     dblDelMoney-退款金额
    '出参:tyBrushCard-返回刷卡对象
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-23 11:07:58
    '说明:同步验证了接口和刷卡接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl帐户余额 As Double
    Dim cllSquareBalance As Collection
    Dim strExpand As String, bln退现 As Boolean
    Dim dblTotal As Double, dblBrushMoney As Double
    Dim cllBalance As Collection, strXmlIn As String
    Dim varData As Variant, varTemp As Variant, i As Long, j As Integer
    Dim rsBalance As ADODB.Recordset
    On Error GoTo errHandle
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then CheckSquareDelValied = True: Exit Function
     
    If zlGetClassMoney(mBalanceInfor.lng结帐ID, rsMoney) = False Then Exit Function
    If dblDelMoney = 0 Then
        If Val(txtReceive.Text) = 0 Then
            MsgBox "未输入退费金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        End If
    End If
     
    '退款检查
    If Not mrsOldBalance Is Nothing Then
        Set rsBalance = mrsOldBalance '原记录集
    Else
        Set rsBalance = mrsBalance
    End If
    
    If rsBalance Is Nothing Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsBalance.State <> 1 Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    If lng消费卡ID <> 0 Then
        rsBalance.Filter = "类型=5 And 结算卡序号=" & objCard.接口序号 & " And 消费卡ID=" & lng消费卡ID
    Else
        rsBalance.Filter = "类型=5 And 结算卡序号=" & objCard.接口序号
    End If
    
    If rsBalance.EOF Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblDelMoney <> 0 Then
        dblMoney = dblDelMoney
    Else
        dblMoney = Val(txtReceive.Text)
    End If
    
    dblTotal = 0
    Set cllSquareBalance = New Collection
    Set cllBalance = New Collection
    Set mcllCurSquareBalance = New Collection
    dblTemp = dblMoney
    
    With rsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTotal = dblTotal + Val(NVL(!冲预交))
            
            'dblBrushMoney = GetSquareBrushMoney(objCard.接口序号, Val(Nvl(!消费卡ID)), Nvl(!卡号))
            'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
            cllSquareBalance.Add Array(objCard.接口序号, Val(NVL(!消费卡ID)), _
             0, NVL(!卡号), "", "", 0, Val(NVL(!冲预交)))
            
            If dblTemp > Val(NVL(!冲预交)) And dblTemp <> 0 Then
                cllBalance.Add Array(objCard.接口序号, Val(NVL(!消费卡ID)), _
                Format(Val(NVL(!冲预交)), "0.00"), NVL(!卡号), "", "", 0)
                dblTemp = dblTemp - Val(NVL(!冲预交))
            ElseIf dblTemp <> 0 Then
                cllBalance.Add Array(objCard.接口序号, Val(NVL(!消费卡ID)), _
                Format(dblTemp, "0.00"), NVL(!卡号), "", "", 0)
                dblTemp = 0
            End If
            .MoveNext
        Loop
    End With
    
    If RoundEx(dblTotal, 6) < RoundEx(dblMoney, 6) Then
        MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.结算方式 & "的未退金额,请检查!" & vbCrLf & _
               "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If RoundEx(dblTotal, 6) <> RoundEx(dblMoney, 6) Then
        If objCard.是否全退 Then
            MsgBox "注意:" & vbCrLf & "   " & objCard.结算方式 & "必须全退,请检查!" & vbCrLf & _
                   "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If gbln消费卡退费验卡 Then
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
        Optional ByRef bln退现 As Boolean) As Boolean
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, objCard.接口序号, _
            objCard.消费卡, mPatiInfor.str姓名, mPatiInfor.str性别, _
            mPatiInfor.str年龄, dblMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
        Set cllBalance = cllSquareBalance
    End If
    For i = 1 To cllBalance.Count
        varData = cllBalance(i)
        dblTemp = Val(varData(2)) + dblTemp
        mcllCurSquareBalance.Add varData
    Next
    
    If dblDelMoney = 0 Then
        txtReceive.Text = Format(dblTemp, "0.00")
    End If
    CheckSquareDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSquareBrushMoney(ByVal lngCardTypeID As Long, ByVal lng消费卡ID As Long, ByVal strCardNo As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡已刷卡金额
    '入参:lngCardTypeId-消费卡接口编号
    '     lng消费卡ID-消费卡ID
    '     strCardNo-卡号
    '出参:
    '返回:返回刷卡金额
    '编制:刘兴洪
    '日期:2014-08-12 11:51:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, varTemp As Variant
    Dim dblMoney As Double, lngRow As Long
    Dim lng卡类别ID As Long, lng消费卡ID1 As Long, strBalance As String
    dblMoney = 0
    With vsBlance
        For lngRow = 1 To .Rows - 1
            lng卡类别ID = Val(.TextMatrix(lngRow, .ColIndex("卡类别ID")))
            lng消费卡ID1 = Val(.TextMatrix(lngRow, .ColIndex("消费卡ID")))
            strBalance = .TextMatrix(lngRow, .ColIndex("结算方式"))
            If Val(.TextMatrix(lngRow, .ColIndex("类型"))) = 5 And strBalance <> "" Then
                If lngCardTypeID = lng卡类别ID And (lng消费卡ID1 = lng消费卡ID Or lng消费卡ID = 0) Then
                    dblMoney = RoundEx(dblMoney + Val(.TextMatrix(lngRow, .ColIndex("结算金额"))), 2)
                End If
            End If
        Next
    End With
    GetSquareBrushMoney = dblMoney
End Function
Private Function zlGetClassMoney(ByRef lng结帐ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, dblMoney As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    '初始化数据结构
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "金额", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng结帐ID <> 0 Then
        strSQL = "" & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 门诊费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 " & _
        "   Union ALL " & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 住院费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 "
        strSQL = "Select 收费类别,Sum(金额) as 金额 From (" & strSQL & ")  Group by  收费类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "收费类别='" & NVL(!收费类别, "无") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!收费类别 = NVL(!收费类别, "无")
                rsMoney!金额 = Val(NVL(rsMoney!金额)) + Val(NVL(!金额))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        zlGetClassMoney = True
        Exit Function
    End If
  
    With mrsFeeList
        dblMoney = mBalanceInfor.dbl当前结帐
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTemp = Val(NVL(!未结金额))
            If RoundEx(dblMoney - dblTemp, gbytDec) <= 0 Then
                dblTemp = dblMoney
            End If
            If dblTemp <> 0 And dblMoney <> 0 Then
                rsMoney.Find "收费类别='" & NVL(!收费类别, "无") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!收费类别 = NVL(!收费类别, "无")
                rsMoney!金额 = Val(NVL(rsMoney!金额)) + dblTemp
                rsMoney.Update
            End If
            dblMoney = dblMoney - dblTemp
            .MoveNext
        Loop
    End With
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LedDisplayBank(Optional ByVal blnLedAsked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示保价信息
    '编制:刘兴洪
    '日期:2011-12-15 13:40:46
    '问题:52117
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, i As Long
    Dim str医保 As String, str三方交易 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String, varData As Variant
    If Not gblnLED Then Exit Sub
    
    
    With vsBlance
        For i = 1 To .Rows - 1
            '医保交易
            If .TextMatrix(i, .ColIndex("结算方式")) <> "" Then
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 1 '医保
                    str医保 = str医保 & "||" & .TextMatrix(i, .ColIndex("结算方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00")
                Case 2 '三方接口交易
                    str三方交易 = str三方交易 & "||" & .TextMatrix(i, .ColIndex("结算方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00")
                Case 3   ' 一卡通交易
                    str老一卡通 = str老一卡通 & "||" & .TextMatrix(i, .ColIndex("结算方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00")
                Case Else
                    str普通结算 = str普通结算 & "||" & .TextMatrix(i, .ColIndex("结算方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00")
                End Select
            End If
        Next
    End With
     
    str结算方式 = ""
    If str医保 <> "" Then str结算方式 = str结算方式 & "||医保结算:||帐户余额:" & Format(mYBInFor.cur个帐余额, "0.00") & str医保
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
        If str结算方式 > "" Then str结算方式 = Mid(str结算方式, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str结算方式
    End Select
    If blnLedAsked = False Then
        If Format(mBalanceInfor.dbl预结算总额, gstrDec) <> Format(mBalanceInfor.dbl医保支付合计, gstrDec) Then
            '虚结算不一致时,需要再次提醒
            zl9LedVoice.Speak "#21 " & Format(Val(mBalanceInfor.dbl未付合计), "0.00")
        End If
    End If
End Sub

Private Sub SetOperatonCommandCaption()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置操作控件的Caption
    '编制:刘兴洪
    '日期:2015-01-21 16:11:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim EditType As gBalanceBill
    
    EditType = mEditType
    If chkCancel.Value = 1 Then EditType = g_Ed_结帐作废
    
    Select Case EditType
    Case g_Ed_重新作废
        cmdOK.Caption = "确定(&O)"
        cmdCancel.Caption = "取消(&C)"
        lblBalance(3).Caption = "退 预 交"
    Case g_Ed_取消结帐
        cmdOK.Caption = "作废(&O)"
        cmdCancel.Caption = "取消(&C)"
        lblBalance(3).Caption = "冲 预 交"
    Case g_Ed_结帐作废
        cmdOK.Caption = "确定(&O)"
        cmdCancel.Caption = "取消(&C)"
        lblBalance(3).Caption = "退 预 交"
    Case Else
        cmdOK.Caption = "完成结算(&O)"
        cmdCancel.Caption = "取消结算(&C)"
        lblBalance(3).Caption = "冲 预 交"
    End Select
    Call picBalanceBack_Resize
End Sub
Private Function GetLocalePayCard(ByVal lng卡类别ID As Long, _
    ByVal bln消费卡 As Boolean, Optional ByRef intKindIdex As Integer) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的卡对象
    '出参:intKindIdex-IDkind的索引
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 15:58:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, i As Long
    
    On Error GoTo errHandle
    intKindIdex = -1
    For i = 1 To IDKindPaymentsType.ListCount
        Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
          If objCard Is Nothing Then Exit Function
        If lng卡类别ID = objCard.接口序号 And objCard.消费卡 = bln消费卡 Then
            intKindIdex = i
            Set GetLocalePayCard = objCard: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetLocaleOldOneCard(ByVal str结算方式 As String, _
     Optional ByRef intKindIdex As Integer) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的老一卡通的卡对象
    '入参:str结算方式-老版一卡通的结算方
    '出参:intKindIdex-IDkind的索引
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 15:58:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, i As Long
    On Error GoTo errHandle
    intKindIdex = -1
    For i = 1 To IDKindPaymentsType.ListCount
        Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
        If objCard.结算方式 = str结算方式 Then
            intKindIdex = i
            Set GetLocaleOldOneCard = objCard: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CancelIsValied(ByVal objCard As Card, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废结帐检查
    '入参:objCard-当前卡对象
    '出参:tyBrushCard-当前刷卡信息
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 15:28:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim dbl剩余金额 As Double, bln退款 As Boolean
    Dim dblMoney As Double, strBalance As String, i As Long
    Dim dblCurMoney As Double, objTemp As Card
    Dim strSquares As String, cllSquare As Collection 'array(ID,名称)
    Dim lngCardTypeID As Long
    On Error GoTo errHandle
      
    If mYBInFor.intInsure > 0 Then
        If Not MCPAR.出院病人结算作废 And mYBInFor.bytMCMode <> 1 Then
            If Not isYBPati(mPatiInfor.lng病人ID, True) Then
                MsgBox "该参保病人已经出院，不能取消该结帐单！", vbInformation, gstrSysName: Exit Function
            End If
        End If
        If gclsInsure.CheckInsureValid(mYBInFor.intInsure) = False Then Exit Function
    End If
    
    With mBalanceInfor
        dbl剩余金额 = RoundEx(.dbl未付合计 - .dbl冲预交合计, 5)
        bln退款 = dbl剩余金额 > 0
    End With
    
    dblCurMoney = Val(txtReceive.Text)
    Set cllSquare = New Collection
    With vsBlance
        For i = 1 To .Rows - 1
            strBalance = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            If strBalance <> "" Then
               If dblCurMoney <> 0 And objCard.结算方式 = strBalance Then
                    MsgBox "在退款列表中已经存在『" & strBalance & "』的退款方式," & vbCrLf & _
                           "不能再使用该退款方式!", vbInformation + vbOKOnly, gstrSysName
                    If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
                    Exit Function
               End If
               
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6) * IIf(mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1, -1, 1)
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 0 '普通结算
                Case 1 '预交款
                Case 2 '医保
                Case 3 '一卡通
                    '医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                    If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                        Set objTemp = GetCard(Val(.TextMatrix(i, .ColIndex("卡类别ID"))))
                        If objTemp Is Nothing Then
                            MsgBox "当前站点不支持" & strBalance & "方式支付!", vbInformation + vbOKOnly, gstrSysName
                            Exit Function
                        End If
                        If Val(.TextMatrix(i, .ColIndex("结算金额"))) > 0 And (mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1) Then bln退款 = True
                        If CheckThreeSwapValied(objTemp, dblMoney, tyBrushCard, bln退款) = False Then Exit Function
                    End If
                Case 4 '一卡通(老版本)
                    Set objTemp = GetLocaleOldOneCard(strBalance)
                    If objTemp Is Nothing Then
                        MsgBox "当前站点不支持" & strBalance & "方式支付!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, .ColIndex("结算金额"))) > 0 And (mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Or chkCancel.Value = 1) Then bln退款 = True
                    If CheckOldOneCardIsValied(objTemp, dblMoney, tyBrushCard, bln退款) = False Then Exit Function
                Case 5 '消费卡
                    lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                    If InStr(strSquares & ",", "," & lngCardTypeID & ",") = 0 Then
                        strSquares = strSquares & "," & lngCardTypeID
                        cllSquare.Add Array(lngCardTypeID, strBalance)
                    End If
                Case Else
                End Select
            End If
        Next
    End With
    For i = 1 To cllSquare.Count
        '医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
        Set objTemp = GetLocalePayCard(Val(cllSquare(i)(0)), True)
        If objTemp Is Nothing Then
            MsgBox "当前站点不支持" & cllSquare(i)(1) & "方式支付!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        dblMoney = GetSquareBrushMoney(Val(cllSquare(i)(0)), 0, "")
        If CheckSquareDelValied(objTemp, tyBrushCard, 0, dblMoney) = False Then Exit Function
        Call AddSquareBalance(objTemp)
    Next
    '----------------------------------------------------------------
    '当前刷卡检查
    
    '当前已经完全冲销完成,直接返回
    If dblCurMoney = 0 And dbl剩余金额 = 0 Then CancelIsValied = True: Exit Function
    
    '现金检查
    If CheckCashValied(objCard, bln退款) = False Then Exit Function
    If objCard.结算性质 = 1 Then CancelIsValied = True: Exit Function
    
    
    If dblCurMoney = 0 Then
        MsgBox "当前" & IIf(bln退款, "退款", "收款") & "金额未输入!", vbInformation + vbOKOnly, gstrSysName
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
        Exit Function
    End If
    
    '支票检查
    If CheckChequeValied(objCard) = False Then Exit Function
    
    '消费卡检查
    If bln退款 Then
        If CheckSquareDelValied(objCard, tyBrushCard, 0, dblCurMoney) = False Then Exit Function
    Else
        If CheckSquareBalanceValied(objCard, tyBrushCard) = False Then Exit Function
    End If
            
    '非三方刷卡和消费卡,直接返回true
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then CancelIsValied = True: Exit Function
    
    '三方交易检查
    If CheckThreeSwapValied(objCard, dblCurMoney, tyBrushCard, bln退款) = False Then Exit Function
    '老版一卡通检查
    If CheckOldOneCardIsValied(objCard, dblCurMoney, tyBrushCard, bln退款) = False Then Exit Function
    CancelIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function GetCancelBalance(ByVal bytFun As Byte, ByRef strBalances As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结帐作废的普通结算方式
    '入参:bytFun-0-普通;1-医保;2-消费卡
    '出参:
    '    bytfunc=0:strBalances的格式:结算方式|结算金额|结算号码||...
    '    bytfunc=1:strBalances的格式:结算方式|结算金额||...
    '    bytfunc=2:strBalances的格式:卡类别ID|卡号|消费卡ID|消费金额||.
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 16:20:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPTBalance As String, i As Long, dblMoney As Double
    Dim strYbBalance As String, strBalance As String, varData As Variant
    Dim strXFBalance As String
    
    On Error GoTo errHandle
    With vsBlance
        '收集退款方式及金额
        strPTBalance = "": strYbBalance = "": strXFBalance = ""
        For i = 1 To .Rows - 1
            dblMoney = -1 * RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
            strBalance = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            If strBalance <> "" Then
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 0 '普通结算
                    '结算方式|结算金额|结算号码|结算摘要||..
                    strPTBalance = strPTBalance & "||" & strBalance
                    strPTBalance = strPTBalance & "|" & dblMoney
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("结算号码")) = "", " ", .TextMatrix(i, .ColIndex("结算号码")))
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("备注")) = "", " ", .TextMatrix(i, .ColIndex("备注")))
                Case 1 '预交款
                Case 2 '医保
                    '结算方式|结算金额||...
                    strYbBalance = strYbBalance & "||" & .TextMatrix(i, .ColIndex("结算方式")) & "|" & dblMoney
                Case 3 '一卡通
                Case 4 '一卡通(老版本)
                Case 5 '消费卡
                    '卡类别ID|卡号|消费卡ID|消费金额||.
                    strXFBalance = strXFBalance & "||" & Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                    strXFBalance = strXFBalance & "|" & Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
                    strXFBalance = strXFBalance & "|" & Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                    strXFBalance = strXFBalance & "|" & dblMoney
                Case Else
                End Select
            End If
        Next
    End With
    If strPTBalance <> "" Then strPTBalance = Mid(strPTBalance, 3)
    If strYbBalance <> "" Then strYbBalance = Mid(strYbBalance, 3)
    If strXFBalance <> "" Then strXFBalance = Mid(strXFBalance, 3)
    
    If bytFun = 0 Then
        strBalances = strPTBalance
    ElseIf bytFun = 1 Then
        strBalances = strYbBalance
    Else
       strBalances = strXFBalance
    End If
    GetCancelBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function ExecuteBalaceCancel(objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行结帐取消操作
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-21 16:25:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, i As Long
    Dim strYbBalance As String '结算方式|金额||...
    Dim strSQL As String, strCardNo As String
    Dim lng冲销ID As Long
    Dim dbl剩余金额 As Double, bln退款 As Boolean
    Dim dblCurMoney As Double, dblMoney As Double
    Dim tyBrushCardInfor As TY_BrushCard
    Dim dblTemp As Double
    Dim objBackCard As Card
    
    If objCard Is Nothing Then Exit Function
    
    On Error GoTo errHandle

    If Not mEditType = g_Ed_重新作废 And mblnNotify = False Then
        If MsgBox("确实要将单据[" & mBalanceInfor.strNO & "]进行取消结帐吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        mblnPrintInvoice = False
        Select Case mobjRedProperty.打印方式
        Case 0  '不打印
        Case 1
            mblnPrintInvoice = True '自动打印
        Case 2  '提示打印
            If MsgBox("是否打印结帐作废票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
        End Select
        mblnNotify = True
    End If
    If CheckDepositFactValied = False Then Exit Function
    If CancelIsValied(objCard, tyBrushCardInfor) = False Then Exit Function
    
    
    With mBalanceInfor
        dbl剩余金额 = RoundEx(.dbl未付合计 - .dbl冲预交合计, 5)
        bln退款 = dbl剩余金额 > 0
    End With
    
    dblCurMoney = IIf(bln退款, 1, -1) * Val(txtReceive.Text)
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill("", mBalanceInfor.lng结帐ID) = False Then Exit Function
    End If
'
    Set cllPro = New Collection
    
    If mBalanceInfor.blnSaveBill = False Then
         lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
         With mBalanceInfor
             .lng冲销ID = lng冲销ID
             .dtBalanceDate = zlDatabase.Currentdate
        End With
        
         '先退结算记录及费用
         strSQL = "Zl_病人结帐记录_Cancel("
         '  No_In         病人结帐记录.No%Type,
         strSQL = strSQL & "'" & mBalanceInfor.strNO & "',"
         '  冲销id_In     病人结帐记录.Id%Type,
         strSQL = strSQL & "" & lng冲销ID & ","
         '  操作员编号_In 病人结帐记录.操作员编号%Type,
         strSQL = strSQL & "'" & UserInfo.编号 & "',"
         '  操作员姓名_In 病人结帐记录.操作员姓名%Type
         strSQL = strSQL & "'" & UserInfo.姓名 & "',"
         '  作废时间_In   病人结帐记录.收费时间%Type := Null
         strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
         strSQL = strSQL & ")"
         zlAddArray cllPro, strSQL
         '执行医保退费操作
         If ExecuteInsureDel(cllPro) = False Then Exit Function
    End If
    
    '执行三方帐户和老一卡通交易退费
    If ExcuteBalanceListThreeDelSwap(cllPro) = False Then Exit Function
    
    '执行当前操作

    If dblCurMoney <> 0 Then
       If dblCurMoney > 0 Then
            '当前退款
            '1.执行当前老一卡通操作
            If ExecuteOneCardDelInterface(objCard, dblCurMoney, cllPro) = False Then Exit Function
            
            '2.执行当前三方帐户操作
            If tyBrushCardInfor.bln转帐 Then
                If ExecuteThreeSwapTransferAccount(objCard, dblCurMoney, cllPro, tyBrushCardInfor, False) = False Then Exit Function
            Else
                If mEditType = g_Ed_重新作废 Then
                    If ExecuteThreeSwapDelInterface(objCard, dblCurMoney, cllPro, True) = False Then Exit Function
                Else
                    If ExecuteThreeSwapDelInterface(objCard, dblCurMoney, cllPro) = False Then Exit Function
                End If
            End If
       Else
            '当前收款
            '1.执行当前老一卡通操作
            If ExecuteOldOneCardPayInterface(mPatiInfor.lng病人ID, mBalanceInfor.lng冲销ID, objCard, -1 * dblCurMoney, tyBrushCardInfor, cllPro) = False Then Exit Function
            '2.执行当前三方帐户操作
            If ExecuteThreeSwapPayInterface(mPatiInfor.lng病人ID, mBalanceInfor.lng冲销ID, objCard, -1 * dblCurMoney, cllPro, tyBrushCardInfor) = False Then Exit Function
       End If
    End If
    
    If objCard.结算性质 = 1 Then
        dblTemp = dbl剩余金额: dbl剩余金额 = 0
        mBalanceInfor.dbl缴款 = RoundEx(IIf(bln退款, -1, 1) * dblCurMoney, 5)
        mBalanceInfor.dbl找补 = Val(txtCaculated.Text)
        dblMoney = GetCentMoney(dblTemp)
        mBalanceInfor.dbl现金 = dblMoney
    Else
        dblTemp = dblCurMoney
        If Not objBackCard Is Nothing And dblCurMoney = 0 Then
            If objBackCard.接口序号 <> 1 And lblCaculated.Caption = "找补" Then
               dblTemp = RoundEx(Val(txtCaculated.Text), 6)
            End If
        End If
        dblMoney = GetCentMoney(dblTemp)
        dbl剩余金额 = RoundEx(dbl剩余金额 - dblCurMoney - mBalanceInfor.dbl误差额, 5)
    End If
    
    Call Show误差金额(False)
    '完成退费操作
    If dbl剩余金额 = 0 Then
        If ExecuteOverBalanceCancel(objCard, cllPro, dblMoney) = False Then Exit Function
        mblnNotify = False
        
        strSQL = "Zl_病人自动记帐_Restore('" & mBalanceInfor.strNO & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxInErase(gcnOracle, mBalanceInfor.lng结帐ID)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        
        If mblnPrintInvoice Then
            '红票打印
            Call frmPrint.ReportPrint(3, mBalanceInfor.strNO, mBalanceInfor.lng冲销ID, mobjRedProperty, _
                mstrInvoice, mBalanceInfor.dtBalanceDate, , , mPatiInfor.lng病人ID, _
                mobjRedProperty.打印格式, , mYBInFor.intInsure <> 0 And MCPAR.医保接口打印票据)
        End If
        
        If mYBInFor.intInsure <> 0 Then
            If MCPAR.结帐作废后打印回单 And InStr(1, mstrPrivs, ";病人退费回单;") > 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "结帐ID=" & mBalanceInfor.lng冲销ID, 2)
            End If
        ElseIf InStr(1, mstrPrivs, ";病人退费回单;") > 0 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "结帐ID=" & mBalanceInfor.lng冲销ID, 2)
        End If
        If mblnDepositBillPrint Then
            '打印预交票据
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mBalanceInfor.str预交No, "病人ID=" & mPatiInfor.lng病人ID, "收款时间=" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS"), 2)
        End If
        Call WriteZYInforToCard(mPatiInfor.lng病人ID, mBalanceInfor.lng冲销ID, True)
        If mintPreEditType >= 0 Then mEditType = mintPreEditType
        If mEditType = g_Ed_结帐作废 Or mEditType = g_Ed_重新作废 Then
            mBalanceInfor.blnSaveBill = False
            Unload Me: ExecuteBalaceCancel = True: Exit Function
        End If
        
        mblnNotChange = True
        chkCancel.Value = 0
        Call NewBill
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        
        mblnNotChange = False
        ExecuteBalaceCancel = True
        Exit Function
    End If

    '加入退费信息
    With vsBlance
        If objCard.消费卡 Then
            Call AddSquareBalance(objCard)
        Else
            If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("结算方式"))) = "") Then
                .Rows = .Rows + 1
                .RowPosition(.Rows - 1) = 1
            End If
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            strCardNo = tyBrushCardInfor.str卡号
            .TextMatrix(1, .ColIndex("是否密文")) = 0
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If objCard.结算性质 = 7 And objCard.接口序号 < 0 Then
                .TextMatrix(1, .ColIndex("类型")) = 4
                .TextMatrix(1, .ColIndex("编辑状态")) = 0   '0-禁止删除;1-允许编辑金额;2-允许删除
                .TextMatrix(1, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
            ElseIf objCard.接口序号 > 0 Then
                .TextMatrix(1, .ColIndex("类型")) = 3
                .TextMatrix(1, .ColIndex("卡类别ID")) = objCard.接口序号
                .TextMatrix(1, .ColIndex("卡类别名称")) = objCard.名称
                .TextMatrix(1, .ColIndex("编辑状态")) = 0   '0-禁止删除;1-允许编辑金额;2-允许删除
                .TextMatrix(1, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                .TextMatrix(1, .ColIndex("是否密文")) = IIf(objCard.卡号密文规则 <> "", 1, 0)
            Else
                .TextMatrix(1, .ColIndex("类型")) = 0
                .TextMatrix(1, .ColIndex("编辑状态")) = 2   '0-禁止删除;1-允许编辑金额;2-允许删除
                .TextMatrix(1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
            End If
            .TextMatrix(1, .ColIndex("结算方式")) = objCard.结算方式
            .TextMatrix(1, .ColIndex("结算性质")) = objCard.结算性质
            .TextMatrix(1, .ColIndex("卡类别ID")) = objCard.接口序号
            .TextMatrix(1, .ColIndex("消费卡ID")) = 0

            .TextMatrix(1, .ColIndex("结算金额")) = Format(dblMoney, "0.00")
            .Cell(flexcpData, 1, .ColIndex("结算金额")) = Format(dblMoney, "0.00")
            .TextMatrix(1, .ColIndex("结算号码")) = ""
            .TextMatrix(1, .ColIndex("备注")) = ""

            If objCard.接口序号 > 0 Then
                .TextMatrix(1, .ColIndex("卡号")) = IIf(objCard.卡号密文规则 <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = tyBrushCardInfor.str卡号
                .TextMatrix(1, .ColIndex("交易流水号")) = tyBrushCardInfor.str交易流水号
                .TextMatrix(1, .ColIndex("交易说明")) = tyBrushCardInfor.str交易说明
                .TextMatrix(1, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(1, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(1, .ColIndex("卡类别名称")) = objCard.名称
            End If
            mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 + dblMoney, 6)
            mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 - dblMoney, 6)
        End If
        For i = 1 To IDKindPaymentsType.ListCount
            '缺省定位在现金上
            Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
            If objCard.结算性质 = 1 Then IDKindPaymentsType.IDKind = i: Exit For
        Next
    End With
    
    txtReceive.Text = ""
    If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
    Call LedDisplayBank

    ExecuteBalaceCancel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ExecuteOverBalanceCancel(ByVal objCard As Card, _
    ByRef cllDelBalancePro As Collection, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行完成退费操作
    '入参:objCard-当前支付类别
    '     cllDelBalancePro-执行的退费单据
    '     dblMoney-当前退款金额
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-23 09:31:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String
    Dim strPTBalance As String
    Dim strSquareBalance As String, bln退预交 As Boolean, i As Long
    On Error GoTo errHandle
    Set cllPro = New Collection
    For i = 1 To cllDelBalancePro.Count
        cllPro.Add cllDelBalancePro(i)
    Next
    
    If GetCancelBalance(0, strPTBalance) = False Then Exit Function
    If GetCancelBalance(2, strSquareBalance) = False Then Exit Function
    
    If objCard.接口序号 <= 0 And InStr(",1,2,", "," & objCard.结算性质 & ",") > 0 Then
        strPTBalance = strPTBalance & IIf(strPTBalance = "", "", "||")
        strPTBalance = strPTBalance & objCard.结算方式
        strPTBalance = strPTBalance & "|" & -1 * dblMoney
        strPTBalance = strPTBalance & "|" & vsBlance.TextMatrix(1, vsBlance.ColIndex("结算号码"))
        strPTBalance = strPTBalance & "|" & vsBlance.TextMatrix(1, vsBlance.ColIndex("备注"))
    ElseIf objCard.接口序号 > 0 And objCard.消费卡 And dblMoney <> 0 Then
        For i = 1 To mcllCurSquareBalance.Count
            '卡类别ID|卡号|消费卡ID|消费金额||.
            'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
            strSquareBalance = strSquareBalance & IIf(strSquareBalance = "", "", "||")
            strSquareBalance = strSquareBalance & Val(mcllCurSquareBalance(i)(0))
            strSquareBalance = strSquareBalance & "|" & Trim(mcllCurSquareBalance(i)(3))
            strSquareBalance = strSquareBalance & "|" & Val(mcllCurSquareBalance(i)(1))
            strSquareBalance = strSquareBalance & "|" & -1 * Val(mcllCurSquareBalance(i)(2))
        Next
    End If
    If strSquareBalance <> "" Then
        'Zl_病人结帐作废_Modify
        strSQL = "Zl_病人结帐作废_Modify("
        '  操作类型_In   Number,
        '--   1-普通退费方式:
        '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '--   2.三方卡退费结算:
        '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '--   4-消费卡结算:
        '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        strSQL = strSQL & "" & 4 & ","
        '  病人id_In     病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & strSquareBalance & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & IIf(InStr(mstrForceNote, "强制退现") + 4 = Len(mstrForceNote), "", mstrForceNote) & "',"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "NULL,"
        '  找补_In       病人预交记录.找补%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In   病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  预交金额_In   病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '冲预交病人ids_In Varchar2 := Null,
        ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
        strSQL = strSQL & "NULL,"
        '  完成作废_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End If
    
    bln退预交 = Val(txtBalance(Idx_冲预交).Text) <> 0
    'Zl_病人结帐作废_Modify
    strSQL = "Zl_病人结帐作废_Modify("
    '  操作类型_In   Number,
    '--   1-普通退费方式:
    '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '--   2.三方卡退费结算:
    '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '--   4-消费卡结算:
    '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    strSQL = strSQL & "" & 1 & ","
    '  病人id_In     病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & strPTBalance & "',"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl缴款 & ","
    '  找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl找补 & ","
    '  误差金额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & -1 * mBalanceInfor.dbl误差额 & ","
    '  预交金额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & -1 * mBalanceInfor.dbl冲预交合计 & ","
    '操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '冲预交病人ids_In Varchar2 := Null,
    ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
    strSQL = strSQL & "NULL,"
    '  完成作废_In Number:=0
    strSQL = strSQL & "2)"
    zlAddArray cllPro, strSQL
    
    If GetSaveAddDepositSQL(mPatiInfor.lng病人ID, mPatiInfor.lng主页ID, mBalanceInfor.lng冲销ID, cllPro) = False Then Exit Function
    
    If mEditType = g_Ed_重新作废 Then
        strSQL = "Zl_病人结帐异常_Update("
        strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ")"
        zlAddArray cllPro, strSQL
    End If
    
    Err = 0: On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ExecuteOverBalanceCancel = True
    Exit Function
ErrRoll:
     gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function


Private Function ExecuteInsureDel(ByRef cllDelBalancePro As Collection, Optional bln异常作废 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保退费操作
    '入参:cllDelBalancePro-执行的退费单据
    '     bln异常作废-是否异常作废
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 16:39:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, strYbBalance As String
    Dim blnTransMC As Boolean, blnTrans As Boolean, i As Long
    Dim strAdvance  As String
    Dim blnReload As Boolean
    
    If mYBInFor.intInsure = 0 Then ExecuteInsureDel = True: Exit Function
    
    '获取医保结算方式
    strYbBalance = ""
    If bln异常作废 = False Then
        If GetCancelBalance(1, strYbBalance) = False Then Exit Function
    End If
    
    On Error GoTo errHandle
    
    Set cllPro = New Collection
    For i = 1 To cllDelBalancePro.Count
        cllPro.Add cllDelBalancePro(i)
    Next
    If mYBInFor.bytMCMode = 1 Then
        If MCPAR.门诊病人结算作废 = False Then  '不支持门诊结算作废,则直接返回
            If strYbBalance = "" Then ExecuteInsureDel = True: Exit Function
            MsgBox "由于该医保不支持门诊结帐作废,因此,不能执行退费操作!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        If Not MCPAR.出院病人结算作废 Then
            If Not isYBPati(mPatiInfor.lng病人ID, True) Then
                MsgBox "该参保病人已经出院，不能取消该结帐单！", vbInformation, gstrSysName: Exit Function
            End If
        End If
        If MCPAR.住院结算作废 = False Then  '不支持门诊结算作废,则直接返回
            If strYbBalance = "" Then ExecuteInsureDel = True: Exit Function
            MsgBox "由于该医保不支持住院结帐作废,因此,不能执行退费操作!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
     
    If strYbBalance <> "" Then
         'Zl_病人结帐作废_Modify
         strSQL = "Zl_病人结帐作废_Modify("
         '  操作类型_In   Number,
         '--   1-普通退费方式:
         '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
         '--   2.三方卡退费结算:
         '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
         '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
         '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
         '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
         '--   4-消费卡结算:
         '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
         
        strSQL = strSQL & "" & 3 & ","
        '  病人id_In     病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & strYbBalance & "',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,0,1)"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   病人预交记录.冲预交%Type := Null,
        '  预交金额_In   病人预交记录.冲预交%Type := Null,
        '操作员编号_In    病人预交记录.操作员编号%Type := Null,
        '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        '收款时间_In      病人预交记录.操作员姓名%Type := Null,
        '冲预交病人ids_In Varchar2 := Null,
        ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
        '  完成作废_In Number:=0
        zlAddArray cllPro, strSQL
    End If
    
    '执行医保退费
    Err = 0: On Error GoTo ErrRoll:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    blnTransMC = False
    
    If mYBInFor.bytMCMode = 1 Then
        strAdvance = mBalanceInfor.lng冲销ID & "|0"
        If Not gclsInsure.ClinicDelSwap(mBalanceInfor.lng结帐ID, , mYBInFor.intInsure, strAdvance) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        blnTransMC = True
        If zlInsureCheck(strYbBalance, strAdvance) Then
            'Zl_病人结帐作废_Modify
             strSQL = "Zl_病人结帐作废_Modify("
             '  操作类型_In   Number,
             '--   1-普通退费方式:
             '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
             '--   2.三方卡退费结算:
             '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
             '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
             '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
             '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
             '--   4-消费卡结算:
             '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
             
            strSQL = strSQL & "" & 3 & ","
            '  病人id_In     病人结帐记录.病人id%Type,
            strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "'" & strAdvance & "')"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            '  卡号_In       病人预交记录.卡号%Type := Null,
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            '  缴款_In       病人预交记录.缴款%Type := Null,
            '  找补_In       病人预交记录.找补%Type := Null,
            '  误差金额_In   病人预交记录.冲预交%Type := Null,
            '  预交金额_In   病人预交记录.冲预交%Type := Null,
            '操作员编号_In    病人预交记录.操作员编号%Type := Null,
            '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            '收款时间_In      病人预交记录.操作员姓名%Type := Null,
            '冲预交病人ids_In Varchar2 := Null,
            ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
            '  完成作废_In Number:=0
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            blnReload = True
        Else
            If strYbBalance <> "" Then
                'Zl_病人结帐作废_Modify
                 strSQL = "Zl_病人结帐作废_Modify("
                 '  操作类型_In   Number,
                 '--   1-普通退费方式:
                 '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
                 '--   2.三方卡退费结算:
                 '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
                 '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
                 '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
                 '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
                 '--   4-消费卡结算:
                 '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
                 
                strSQL = strSQL & "" & 3 & ","
                '  病人id_In     病人结帐记录.病人id%Type,
                strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
                '  冲销id_In     病人预交记录.结帐id%Type,
                strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
                '  结算方式_In   Varchar2,
                strSQL = strSQL & "'" & strYbBalance & "')"
                '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                '  卡号_In       病人预交记录.卡号%Type := Null,
                '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                '  交易说明_In   病人预交记录.交易说明%Type := Null,
                '  缴款_In       病人预交记录.缴款%Type := Null,
                '  找补_In       病人预交记录.找补%Type := Null,
                '  误差金额_In   病人预交记录.冲预交%Type := Null,
                '  预交金额_In   病人预交记录.冲预交%Type := Null,
                '操作员编号_In    病人预交记录.操作员编号%Type := Null,
                '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
                '收款时间_In      病人预交记录.操作员姓名%Type := Null,
                '冲预交病人ids_In Varchar2 := Null,
                ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
                '  完成作废_In Number:=0
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                blnReload = True
            End If
        End If
    Else
        strAdvance = ""
        If Not gclsInsure.SettleDelSwap(mBalanceInfor.lng结帐ID, mYBInFor.intInsure, strAdvance) Then
            gcnOracle.RollbackTrans:  Exit Function
        End If
        blnTransMC = True
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If blnReload Then
        i = 1
        With vsBlance
            Do While i <= .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("类型"))) = 2 Then
                    Call DeletePayInfor(i, True)
                Else
                    i = i + 1
                End If
            Loop
        End With
        Call LoadBalancePayData(mPatiInfor.lng病人ID, mBalanceInfor.lng冲销ID, , False, True, -1)
        Call LoadCurOwnerPayInfor
        MsgBox "医保退款情况已发生变化,请根据新的退款金额重新处理作废！", vbInformation, gstrSysName
        mBalanceInfor.blnSaveBill = True
        Exit Function
    End If
    
    Set cllDelBalancePro = New Collection   '清空保存作废结帐单据数据
    mBalanceInfor.blnSaveBill = True
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mYBInFor.bytMCMode = 1, 交易Enum.Busi_ClinicDelSwap, 交易Enum.Busi_SettleDelSwap), True, mYBInFor.intInsure)
    ExecuteInsureDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Exit Function
ErrRoll:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    
End Function

Private Function zlInsureCheck(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")

    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsureCheck = blnMedicareCheck
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ExcuteBalanceListThreeDelSwap(ByRef cllDelBalancePro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行结算列表中的三方交易退费
    '入参:cllDelBalancePro-执行的退费单据
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 17:20:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strBalance As String, objTemp As Card, i As Long
    Dim lngTypeCardTypeID As Long
    Dim strName As String
    
    On Error GoTo errHandle
  
    With vsBlance
        '收集退款方式及金额
        For i = 1 To .Rows - 1
            dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
            strBalance = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            If strBalance <> "" Then
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 0 '普通结算
                Case 1 '预交款
                Case 2 '医保
                Case 3 '一卡通
                    lngTypeCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                    Set objTemp = GetCard(CStr(lngTypeCardTypeID))
                    If objTemp Is Nothing Then
                        strName = Trim(.TextMatrix(i, .ColIndex("卡类别名称")))
                        MsgBox "本站点不支持使用『" & IIf(strName = "", strBalance, strName) & "』方式进行退款!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                        If dblMoney <> 0 Then
                            '执行退费
                            If Not ExecuteThreeSwapDelInterface(objTemp, dblMoney, cllDelBalancePro) Then Exit Function
                           .TextMatrix(i, .ColIndex("结算状态")) = 1
                        Else
                            
                        End If
                    End If
                Case 4 '一卡通(老版本)
                    If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                        Set objTemp = GetLocaleOldOneCard(strBalance)
                        If objTemp Is Nothing Then
                            strName = Trim(.TextMatrix(i, .ColIndex("卡类别名称")))
                            MsgBox "本站点不支持使用『" & IIf(strName = "", strBalance, strName) & "』方式进行退款!", vbInformation + vbOKOnly, gstrSysName
                            Exit Function
                        End If
                        If dblMoney >= 0 Then
                            '执行退费用
                            If Not ExecuteOneCardDelInterface(objTemp, dblMoney, cllDelBalancePro) Then Exit Function
                            .TextMatrix(i, .ColIndex("结算状态")) = 1
                        Else
                            
                        End If
                    End If
                Case 5 '消费卡
                Case Else
                End Select
            End If
        Next
    End With
    ExcuteBalanceListThreeDelSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteThreeSwapDelSingle(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByVal str卡号 As String, ByVal str交易说明 As String, _
    ByVal str交易流水号 As String, ByVal lng预交ID As Long, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(单笔三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strValue As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, dblZFJE As Double
    Dim strCardNo As String, str结算方式   As String
    Dim strOutXML As String, strInXML As String, strExpend As String
    Dim objXml As New clsXML, strArray() As String, lngRow As Long
    Dim strExpendAfterXml As String, strBalanceIDs As String
    Dim j As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapDelSingle = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
    
    With mrsBalance
        str结算方式 = objCard.结算方式
        
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & -1 * dblMoney
        str结算方式 = str结算方式 & "|" & ""
        str结算方式 = str结算方式 & "|" & ""
        
        '调用之前,先处理数据
        'Zl_病人结帐结算_Modify
        strSQL = "Zl_病人结帐结算_Modify("
        '  操作类型_In     Number,
        '  --操作类型_In:
        '  --   0-普通收费方式:
        '  --   1.三方卡结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退支票额_In:传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        strSQL = strSQL & "1,"
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & objCard.接口序号 & ","
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  结帐类型_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    strBalanceIDs = "1|" & lng预交ID
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.接口序号, False, str卡号, strBalanceIDs, _
         dblMoney, str交易流水号, str交易说明, strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
    
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, objCard.接口序号, False, str卡号, strBalanceIDs, _
         dblMoney, str交易流水号, str交易说明, strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
         
    strSQL = "Zl_三方退款信息_Insert("
    strSQL = strSQL & mBalanceInfor.lng结帐ID & ","
    strSQL = strSQL & lng预交ID & ","
    strSQL = strSQL & dblMoney & ",'"
    strSQL = strSQL & str卡号 & "','"
    strSQL = strSQL & str交易流水号 & "','"
    strSQL = strSQL & str交易说明 & "',"
    strSQL = strSQL & 0 & ")"
    zlAddArray cllThreeSwap, strSQL
    
    Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng结帐ID, objCard.接口序号, objCard.消费卡, "", strExpend, cllThreeSwap, lng预交ID)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapDelSingle = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteOneCardDelInterface(ByVal objCard As Card, _
    ByVal dblDelMoney As Double, _
    ByRef cllBillPro As Collection, Optional ByVal bln异常作废 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通退费接口(老版)
    '入参:cllBillPro-保存单据的SQL
    '     bln异常作废-异常作废调用(true,为异常作废调用,False-正常调用)
    '编制:刘兴洪
    '日期:2014-07-10 10:36:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String '医院编码
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String
    Dim str结算方式 As String
    Dim cllPro As Collection, blnTrans As Boolean
    
    '非一卡通支付,直接返回
    If objCard.结算性质 <> 7 Then ExecuteOneCardDelInterface = True: Exit Function

     mOldOneCard.rsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mOldOneCard.rsOneCard.EOF Then
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        ExecuteOneCardDelInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    On Error GoTo errHandle
    If mrsBalance Is Nothing Then Exit Function
    If mrsBalance.State <> 1 Then Exit Function
    mrsBalance.Filter = "类型=4"
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        .MoveFirst
        Do While Not .EOF
            dblMoney = dblMoney + Val(NVL(mrsBalance!冲预交))
            .MoveNext
        Loop
        .MoveFirst
    End With
    If RoundEx(dblMoney, 6) = 0 Then Exit Function
    
    If dblDelMoney <> dblMoney Then
        MsgBox objCard.结算方式 & " 必须全退!" & vbCrLf & "原结算金额:" & Format(dblMoney, "0.00") & vbCrLf & " 现退款金额:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    '一卡通(旧):只能使用一种
    With mrsBalance
        strCardNo = NVL(!卡号)
        str结算方式 = NVL(!结算方式)
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & -1 * dblMoney
        str结算方式 = str结算方式 & "|" & IIf(Trim(NVL(!结算号码)) = "", " ", Trim(NVL(!结算号码)))
        str结算方式 = str结算方式 & "| "
         
         
        'Zl_病人结帐作废_Modify
        strSQL = "Zl_病人结帐作废_Modify("
        '  操作类型_In   Number,
        '--   1-普通退费方式:
        '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '--   2.三方卡退费结算:
        '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '--   4-消费卡结算:
        '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        strSQL = strSQL & "" & 2 & ","
        '  病人id_In     病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & NVL(!交易流水号) & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & NVL(!交易说明) & "')"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   病人预交记录.冲预交%Type := Null,
        '  预交金额_In   病人预交记录.冲预交%Type := Null,
        '  完成作废_In Number:=0
       If Not bln异常作废 Then zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Err = 0: On Error GoTo ErrRoll:
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "一卡通退费交易调用失败,不能继续退费操作！", vbExclamation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    ExecuteOneCardDelInterface = True
    mBalanceInfor.blnSaveBill = True

    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapDelBatch(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByVal strInput As String, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(多笔三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strValue As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, dblZFJE As Double
    Dim strCardNo As String, str结算方式   As String
    Dim strOutXML As String, strInXML As String, strExpend As String
    Dim objXml As New clsXML, strArray() As String, lngRow As Long
    Dim strExpendAfterXml As String, strBalanceIDs As String
    Dim j As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapDelBatch = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
    
    strArray = Split(mstrBalanceLimit, "|")
    
    For i = 0 To UBound(strArray)
        If Val(Split(strArray(i), ",")(0)) = objCard.接口序号 Then
            If dblMoney > Abs(Val(Split(strArray(i), ",")(1))) Then
                MsgBox objCard.结算方式 & " 的退款金额超过了最大退款金额!" & vbCrLf & "最大退款金额:" & Format(Val(Split(strArray(i), ",")(1)), "0.00") & vbCrLf & " 现退款金额:" & Format(dblMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next i
    
    With mrsBalance
        str结算方式 = objCard.结算方式
        
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & -1 * dblMoney
        str结算方式 = str结算方式 & "|" & ""
        str结算方式 = str结算方式 & "|" & ""
        
        '调用之前,先处理数据
        'Zl_病人结帐结算_Modify
        strSQL = "Zl_病人结帐结算_Modify("
        '  操作类型_In     Number,
        '  --操作类型_In:
        '  --   0-普通收费方式:
        '  --   1.三方卡结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退支票额_In:传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        strSQL = strSQL & "1,"
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & objCard.接口序号 & ","
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  结帐类型_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    objXml.ClearXmlText
    Call objXml.AppendNode("JSLIST")
    strArray = Split(strInput, "|")
    For i = 0 To UBound(strArray)
        Call objXml.AppendNode("JS")
            Call objXml.appendData("KH", Split(strArray(i), ",")(0))
            Call objXml.appendData("JYLSH", TruncStringEx(Split(strArray(i), ",")(1), True))
            Call objXml.appendData("JYSM", TruncStringEx(Split(strArray(i), ",")(2), True))
            Call objXml.appendData("ZFJE", RoundEx(-1 * Val(Split(strArray(i), ",")(3)), 2))
            Call objXml.appendData("JSLX", 1)
            Call objXml.appendData("ID", Split(strArray(i), ",")(4))
        Call objXml.AppendNode("JS", True)
        
        strSQL = "Zl_三方退款信息_Insert("
        strSQL = strSQL & mBalanceInfor.lng结帐ID & ","
        strSQL = strSQL & Val(Split(strArray(i), ",")(4)) & ","
        strSQL = strSQL & -1 * Val(Split(strArray(i), ",")(3)) & ",'"
        strSQL = strSQL & TruncStringEx(Split(strArray(i), ",")(0), True) & "','"
        strSQL = strSQL & TruncStringEx(Split(strArray(i), ",")(1), True) & "','"
        strSQL = strSQL & TruncStringEx(Split(strArray(i), ",")(2), True) & "')"
        zlAddArray cllThreeSwap, strSQL
        strBalanceIDs = strBalanceIDs & "," & Val(Split(strArray(i), ",")(4))
    Next i
    Call objXml.AppendNode("JSLIST", True)

    strInXML = objXml.XmlText
    strExpend = objXml.XmlText
    If strBalanceIDs <> "" Then strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
    
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.接口序号, objCard.消费卡, "", strBalanceIDs, _
         dblMoney, "", "", strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
    
    If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, objCard.接口序号, objCard.消费卡, strInXML, _
         mBalanceInfor.lng结帐ID, strOutXML, strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
    
    
    If strOutXML <> "" Then
        If zlXML_Init = False Then Exit Function
        If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
        Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
        For i = 0 To lngRow - 1
            Call zlXML_GetNodeValue("ID", i, strValue)
            strSQL = "Zl_三方退款信息_Insert("
            strSQL = strSQL & mBalanceInfor.lng结帐ID & ","
            strSQL = strSQL & Val(strValue) & ","
            For j = 0 To UBound(strArray)
                If Val(Split(strArray(i), ",")(4)) = Val(strValue) Then
                    dblZFJE = -1 * Val(Split(strArray(i), ",")(3))
                    Exit For
                End If
            Next j
            strSQL = strSQL & dblZFJE & ",'"
            Call zlXML_GetNodeValue("KH", i, strValue)
            strSQL = strSQL & strValue & "','"
            Call zlXML_GetNodeValue("TKLSH", i, strValue)
            strSQL = strSQL & strValue & "','"
            Call zlXML_GetNodeValue("TKSM", i, strValue)
            strSQL = strSQL & strValue & "',"
            strSQL = strSQL & 1 & ")"
            zlAddArray cllThreeSwap, strSQL
        Next i
    End If
    
    If strExpend <> "" Then
        If zlXML_LoadXMLToDOMDocument(strExpend, False) = False Then Exit Function
        Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
        For i = 0 To lngRow - 1
            Call zlXML_GetNodeValue("XMMC", i, strValue)
            strExpendAfterXml = strExpendAfterXml & "||" & strValue
            Call zlXML_GetNodeValue("XMNR", i, strValue)
            strExpendAfterXml = strExpendAfterXml & "|" & strValue
        Next i
    End If
    If strExpendAfterXml <> "" Then strExpendAfterXml = Mid(strExpendAfterXml, 3)
    Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng结帐ID, objCard.接口序号, objCard.消费卡, "", strExpendAfterXml, cllThreeSwap)
    
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    ExecuteThreeSwapDelBatch = True
    
    
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapDelInterface(ByVal objCard As Card, _
    ByVal dblDelMoney As Double, ByRef cllBillPro As Collection, _
    Optional ByVal bln异常作废 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     bln异常作废-异常作废时调用:true-异常作废;false-正常作废操作
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, dblMoney As Double, str结算方式  As String
    
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapDelInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    
    If bln异常作废 = True Then
        If Not mrsOldBalance Is Nothing Then
            If mrsOldBalance.State <> 1 Then Exit Function
            
            mrsOldBalance.Filter = "类型=3 And 卡类别ID=" & objCard.接口序号
            If mrsOldBalance.RecordCount = 0 Then Exit Function
        
            With mrsOldBalance
                .MoveFirst
                Do While Not .EOF
                    dblMoney = dblMoney + Val(NVL(mrsOldBalance!冲预交))
                    .MoveNext
                Loop
                .MoveFirst
            End With
            
            If RoundEx(dblMoney, 6) = 0 Then Exit Function
            If dblDelMoney > dblMoney Then
                MsgBox objCard.结算方式 & " 的退款金额超过了原始结算金额!" & vbCrLf & "原结算金额:" & Format(dblMoney, "0.00") & vbCrLf & " 现退款金额:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            
            With mrsOldBalance
                strCardNo = NVL(!卡号)
                strSwapNO = NVL(!交易流水号)
                strSwapMemo = NVL(!交易说明)
                str结算方式 = NVL(!结算方式)
                
                '结算方式|结算金额|结算号码|结算摘要||..
                str结算方式 = str结算方式 & "|" & -1 * dblDelMoney
                str结算方式 = str结算方式 & "|" & IIf(Trim(NVL(!结算号码)) = "", " ", Trim(NVL(!结算号码)))
                str结算方式 = str结算方式 & "|" & IIf(Trim(NVL(!摘要)) = "", " ", Trim(NVL(!摘要)))
                'Zl_病人结帐作废_Modify
                strSQL = "Zl_病人结帐作废_Modify("
                '  操作类型_In   Number,
                '--   1-普通退费方式:
                '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
                '--   2.三方卡退费结算:
                '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
                '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
                '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
                '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
                '--   4-消费卡结算:
                '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
                strSQL = strSQL & "" & 2 & ","
                '  病人id_In     病人结帐记录.病人id%Type,
                strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
                '  冲销id_In     病人预交记录.结帐id%Type,
                strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
                '  结算方式_In   Varchar2,
                strSQL = strSQL & "'" & str结算方式 & "',"
                '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                strSQL = strSQL & "" & objCard.接口序号 & ","
                '  卡号_In       病人预交记录.卡号%Type := Null,
                strSQL = strSQL & "'" & strCardNo & "',"
                '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                strSQL = strSQL & "'" & NVL(!交易流水号) & "',"
                '  交易说明_In   病人预交记录.交易说明%Type := Null,
                strSQL = strSQL & "'" & NVL(!交易说明) & "')"
                '  缴款_In       病人预交记录.缴款%Type := Null,
                '  找补_In       病人预交记录.找补%Type := Null,
                '  误差金额_In   病人预交记录.冲预交%Type := Null,
                '  预交金额_In   病人预交记录.冲预交%Type := Null,
                '  完成作废_In Number:=0
                zlAddArray cllPro, strSQL
            End With
        End If
    Else
        If Not mrsBalance Is Nothing Then
            If mrsBalance.State <> 1 Then Exit Function
            
            mrsBalance.Filter = "类型=3 And 卡类别ID=" & objCard.接口序号
            If mrsBalance.RecordCount = 0 Then Exit Function
        
            With mrsBalance
                .MoveFirst
                Do While Not .EOF
                    dblMoney = dblMoney + Val(NVL(mrsBalance!冲预交))
                    .MoveNext
                Loop
                .MoveFirst
            End With
            
            If RoundEx(dblMoney, 6) = 0 Then Exit Function
            If dblDelMoney > dblMoney Then
                MsgBox objCard.结算方式 & " 的退款金额超过了原始结算金额!" & vbCrLf & "原结算金额:" & Format(dblMoney, "0.00") & vbCrLf & " 现退款金额:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            
            With mrsBalance
                strCardNo = NVL(!卡号)
                strSwapNO = NVL(!交易流水号)
                strSwapMemo = NVL(!交易说明)
                str结算方式 = NVL(!结算方式)
                
                '结算方式|结算金额|结算号码|结算摘要||..
                str结算方式 = str结算方式 & "|" & -1 * dblDelMoney
                str结算方式 = str结算方式 & "|" & IIf(Trim(NVL(!结算号码)) = "", " ", Trim(NVL(!结算号码)))
                str结算方式 = str结算方式 & "|" & IIf(Trim(NVL(!摘要)) = "", " ", Trim(NVL(!摘要)))
                'Zl_病人结帐作废_Modify
                strSQL = "Zl_病人结帐作废_Modify("
                '  操作类型_In   Number,
                '--   1-普通退费方式:
                '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
                '--   2.三方卡退费结算:
                '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
                '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
                '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
                '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
                '--   4-消费卡结算:
                '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
                strSQL = strSQL & "" & 2 & ","
                '  病人id_In     病人结帐记录.病人id%Type,
                strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
                '  冲销id_In     病人预交记录.结帐id%Type,
                strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
                '  结算方式_In   Varchar2,
                strSQL = strSQL & "'" & str结算方式 & "',"
                '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                strSQL = strSQL & "" & objCard.接口序号 & ","
                '  卡号_In       病人预交记录.卡号%Type := Null,
                strSQL = strSQL & "'" & strCardNo & "',"
                '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                strSQL = strSQL & "'" & NVL(!交易流水号) & "',"
                '  交易说明_In   病人预交记录.交易说明%Type := Null,
                strSQL = strSQL & "'" & NVL(!交易说明) & "')"
                '  缴款_In       病人预交记录.缴款%Type := Null,
                '  找补_In       病人预交记录.找补%Type := Null,
                '  误差金额_In   病人预交记录.冲预交%Type := Null,
                '  预交金额_In   病人预交记录.冲预交%Type := Null,
                '  完成作废_In Number:=0
                zlAddArray cllPro, strSQL
            End With
        End If
    End If
    
    On Error GoTo ErrRoll:
    
    str结帐IDs = mBalanceInfor.lng冲销ID & IIf(mBalanceInfor.lng结帐ID <> 0, "," & mBalanceInfor.lng结帐ID, "")
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, objCard.接口序号, objCard.消费卡, strCardNo, _
        "2|" & mBalanceInfor.lng结帐ID, dblDelMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
    'Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, strCardNO, strSwapNO, strSwapMemo, cllUpdate, 2)
    Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng冲销ID, objCard.接口序号, objCard.消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
    
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    mBalanceInfor.blnSaveBill = True
    ExecuteThreeSwapDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapTransferPay(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef cllBillPro As Collection, _
    ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通转帐支付(三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     tyBrushCard-转帐刷卡信息
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, str结算方式   As String
    Dim strXMLExpend As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapTransferPay = True: Exit Function
    If Not objCard.是否转帐及代扣 Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
  
    With mrsBalance
        strCardNo = tyBrushCard.str卡号
        strSwapNO = tyBrushCard.str交易流水号
        strSwapMemo = tyBrushCard.str交易说明
        str结算方式 = objCard.结算方式
        
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & dblMoney
        str结算方式 = str结算方式 & "|" & " "
        str结算方式 = str结算方式 & "|" & "转帐结算"
        
        '调用之前,先处理数据
        'Zl_病人结帐结算_Modify
        strSQL = "Zl_病人结帐结算_Modify("
        '  操作类型_In     Number,
        '  --操作类型_In:
        '  --   0-普通收费方式:
        '  --   1.三方卡结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退支票额_In:传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        strSQL = strSQL & "1,"
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & objCard.接口序号 & ","
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & strSwapNO & "',"
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & strSwapMemo & "',"
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl缴款 & ","
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl找补 & ","
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  结帐类型_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    With mBalanceInfor
        dblMoney = RoundEx(IIf(RoundEx(.dbl未付合计 - .dbl冲预交合计, 6) < 0, -1, 1) * dblMoney, 5)
    End With
    'zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long,
    '     ByVal lngCardTypeID As Long, _
    '    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double,
    '    Optional ByRef strSwapGlideNO As String, _
    '    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String,
    '    Optional ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方帐户转帐
    '入参:
    '   frmMain-调用的主窗体
    '   lngModule-HIS调用模块号
    '   lngCardTypeID-卡类别ID
    '   strCardNo-卡号
    '   strBalanceID-结算ID
    '   dblMoney-转帐金额
    '    strSwapExtendInfor-退费业务时，传入本次退费的冲销ID:
    '                        格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                        收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '   strXMLExpend-XML串:
    '       <IN>
    '             <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务
    '       </IN>
    '出参:
    '   strSwapGlideNO-交易流水号
    '   strSwapMemo -交易说明
    '   strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '   strXMLExpend-XML串:
    '        <OUT>
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    strSwapExtendInfor = "2|" & mBalanceInfor.lng结帐ID
    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.接口序号, _
        strCardNo, mBalanceInfor.lng结帐ID, Abs(dblMoney), strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
        gcnOracle.RollbackTrans: Call zlShowThreeSwapErrInfor(1, strXMLExpend): Exit Function
    End If
    
    Call zlAddUpdateSwapSQL(False, mBalanceInfor.lng结帐ID, objCard.接口序号, False, tyBrushCard.str卡号, strSwapNO, strSwapMemo, cllUpdate, 2)
'    strSQL = "Zl_三方退款信息_Insert(" & mBalanceInfor.lng结帐ID & "," & objCard.接口序号 & "," & dblMoney & ",'" & strCardNo & "'," & "'" & strSwapNO & "'," & "'" & strSwapMemo & "',0)"
'    zlAddArray cllUpdate, strSQL
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    mBalanceInfor.blnSaveBill = True
    If strSwapExtendInfor <> "2|" & mBalanceInfor.lng结帐ID Then
        Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng结帐ID, objCard.接口序号, objCard.消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapTransferPay = True
    
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapTransferAccount(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef cllBillPro As Collection, _
    ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal bln异常作废 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通转帐支付(三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     tyBrushCard-转帐刷卡信息
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, str结算方式   As String
    Dim strXMLExpend As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapTransferAccount = True: Exit Function
    If Not objCard.是否转帐及代扣 Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
  
    With mrsBalance
        strCardNo = tyBrushCard.str卡号
        strSwapNO = tyBrushCard.str交易流水号
        strSwapMemo = tyBrushCard.str交易说明
        str结算方式 = objCard.结算方式
        
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & -1 * dblMoney
        str结算方式 = str结算方式 & "|" & " "
        str结算方式 = str结算方式 & "|" & "转帐结算"
        
        'Zl_病人结帐作废_Modify
        strSQL = "Zl_病人结帐作废_Modify("
        '  操作类型_In   Number,
        '--   1-普通退费方式:
        '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '--   2.三方卡退费结算:
        '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '--   4-消费卡结算:
        '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        strSQL = strSQL & "" & 2 & ","
        '  病人id_In     病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & objCard.接口序号 & ","
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & strSwapNO & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & strSwapMemo & "')"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   病人预交记录.冲预交%Type := Null,
        '  预交金额_In   病人预交记录.冲预交%Type := Null,
        '  完成作废_In Number:=0
        If bln异常作废 = False Then zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    str结帐IDs = mBalanceInfor.lng冲销ID & IIf(mBalanceInfor.lng结帐ID <> 0, "," & mBalanceInfor.lng结帐ID, "")
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    'zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long,
    '     ByVal lngCardTypeID As Long, _
    '    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double,
    '    Optional ByRef strSwapGlideNO As String, _
    '    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String,
    '    Optional ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方帐户转帐
    '入参:
    '   frmMain-调用的主窗体
    '   lngModule-HIS调用模块号
    '   lngCardTypeID-卡类别ID
    '   strCardNo-卡号
    '   strBalanceID-结算ID
    '   dblMoney-转帐金额
    '    strSwapExtendInfor-退费业务时，传入本次退费的冲销ID:
    '                        格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                        收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '   strXMLExpend-XML串:
    '       <IN>
    '             <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务
    '       </IN>
    '出参:
    '   strSwapGlideNO-交易流水号
    '   strSwapMemo -交易说明
    '   strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '   strXMLExpend-XML串:
    '        <OUT>
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    strSwapExtendInfor = "2|" & mBalanceInfor.lng冲销ID
    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.接口序号, _
        strCardNo, mBalanceInfor.lng结帐ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
        gcnOracle.RollbackTrans: Call zlShowThreeSwapErrInfor(1, strXMLExpend): Exit Function
    End If
    
    Call zlAddUpdateSwapSQL(False, mBalanceInfor.lng冲销ID, objCard.接口序号, False, tyBrushCard.str卡号, strSwapNO, strSwapMemo, cllUpdate, 2)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    mBalanceInfor.blnSaveBill = True
    If strSwapExtendInfor <> "2|" & mBalanceInfor.lng冲销ID Then
        Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng冲销ID, objCard.接口序号, objCard.消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapTransferAccount = True
    
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub AddSquareBalance(ByVal objCard As Card)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加消费卡结算方式到结算方式列表
    '编制:刘兴洪
    '日期:2015-01-23 15:09:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As New Collection
    Dim j As Integer, dblMoney As Double, strCardNo As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    With vsBlance
      '先清除原始的消费卡部分,再重新退费
        Call ClearSquareBalance(objCard.接口序号)
        Set cllBalance = mcllCurSquareBalance
        For j = 1 To cllBalance.Count
            If objCard.接口序号 = Val(cllBalance(j)(0)) Then
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                If .Rows = 1 Then .Rows = .Rows + 1
                
                If Trim(.TextMatrix(.Rows - 1, .ColIndex("结算方式"))) <> "" Then
                    .Rows = .Rows + 1
                End If
          
                '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                dblMoney = cllBalance(j)(2)
            
                .TextMatrix(.Rows - 1, .ColIndex("类型")) = 5
                .TextMatrix(.Rows - 1, .ColIndex("是否密文")) = Val(cllBalance(j)(6))
                .TextMatrix(.Rows - 1, .ColIndex("结算性质")) = objCard.结算性质
                If zlSquareIsDelCash(objCard.接口序号) Then
                    .TextMatrix(.Rows - 1, .ColIndex("编辑状态")) = 2 '0-禁止删除;1-允许编辑金额;2-允许删除
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("编辑状态")) = 0 '0-禁止删除;1-允许编辑金额;2-允许删除
                End If
                
                .TextMatrix(.Rows - 1, .ColIndex("结算状态")) = 0 '是否已结算:1-已结算;0-未结算
                .TextMatrix(.Rows - 1, .ColIndex("卡类别ID")) = objCard.接口序号
                .TextMatrix(.Rows - 1, .ColIndex("消费卡ID")) = Val(cllBalance(j)(1))
                .Cell(flexcpData, .Rows - 1, .ColIndex("消费卡ID")) = cllBalance(j)(4) '密码
                .Cell(flexcpData, .Rows - 1, .ColIndex("卡类别ID")) = cllBalance(j)(5) '限制类别
                
                .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = objCard.结算方式
                 strCardNo = Trim(cllBalance(j)(3))
                .TextMatrix(.Rows - 1, .ColIndex("卡号")) = IIf(objCard.卡号密文规则 <> "" And objCard.卡号密文规则 <> "0", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, .Rows - 1, .ColIndex("卡号")) = strCardNo
                .TextMatrix(.Rows - 1, .ColIndex("结算金额")) = Format(dblMoney, "0.00")
                .Cell(flexcpData, .Rows - 1, .ColIndex("结算金额")) = Format(dblMoney, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("结算号码")) = ""
                .TextMatrix(.Rows - 1, .ColIndex("备注")) = ""
                .TextMatrix(.Rows - 1, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("卡类别名称")) = objCard.名称
                
                mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 + dblMoney, 6)
                mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 - dblMoney, 6)
                
            End If
        Next
    End With
End Sub

Private Sub ClearSquareBalance(ByVal lngCardTypeID As Long, _
    Optional ByVal lng消费卡ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除消费卡结算
    '编制:刘兴洪
    '日期:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = lngCardTypeID _
                And (lng消费卡ID = 0 Or (lng消费卡ID <> 0 And Val(.TextMatrix(j, .ColIndex("消费卡ID"))) = lng消费卡ID)) Then
                dblMoney = Val(.TextMatrix(j, .ColIndex("结算金额")))
                
                mBalanceInfor.dbl已付合计 = RoundEx(mBalanceInfor.dbl已付合计 - dblMoney, 6)
                mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl未付合计 + dblMoney, 6)
                If .Rows >= 2 Then
                    .RemoveItem j
                Else
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                   .RowData(1) = ""
                   j = 2
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub

Private Sub vsDeposit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsDeposit.EditCell
    vsDeposit.EditSelStart = 0
    vsDeposit.EditSelLength = 100
End Sub

Private Sub vsDeposit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsDeposit.ColIndex("冲预交") Then
        If Val(vsDeposit.EditText) = Val(vsDeposit.TextMatrix(Row, Col)) Then mblnNoTrigger = True
    End If
End Sub

Private Sub vsDetailList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress vsDeposit, KeyAscii, m金额式
End Sub

Private Sub vsDetailList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, dblMoney As Double
    With vsDetailList
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDec)
        .Cell(flexcpData, Row, Col) = Val(.TextMatrix(Row, Col))
        For i = 1 To .Rows - 1
            dblMoney = dblMoney + Val(.Cell(flexcpData, i, .ColIndex("结帐金额")))
        Next i
    End With
    mblnNotChange = True
    txtBalance(Idx_本次结帐).Text = Format(dblMoney, gstrDec)
    mblnNotChange = False
    mBalanceInfor.dbl当前结帐 = dblMoney
    mBalanceInfor.dbl未付合计 = RoundEx(mBalanceInfor.dbl当前结帐 - mBalanceInfor.dbl已付合计, 5)
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor(True)
    If vsDetailList.Row + 1 <= vsDetailList.Rows - 1 Then
        vsDetailList.Select vsDetailList.Row + 1, vsDetailList.ColIndex("结帐金额")
    End If
    mbln已报价 = False
End Sub

Private Sub vsDetailList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mrsInfo Is Nothing Then Cancel = True: Exit Sub
    If mrsInfo.State = 0 Then Cancel = True: Exit Sub
    If mrsInfo.RecordCount = 0 Then Cancel = True: Exit Sub
    If mYBInFor.intInsure <> 0 Then Cancel = True: Exit Sub
    
    If InStr(mstrPrivs, ";结帐设置;") = 0 Then Cancel = True: Exit Sub
     
    With vsDetailList
        If Col <> .ColIndex("结帐金额") Then
            Cancel = True
        Else
            If .Cell(flexcpBackColor, Row, .ColIndex("结帐金额")) = .Cell(flexcpBackColor, Row, .ColIndex("日期")) _
                Or .TextMatrix(Row, .ColIndex("单据")) = "" Then
                Cancel = True
            End If
            '负数金额不允许修改
            If Val(.Cell(flexcpData, Row, .ColIndex("未结金额"))) < 0 Then Cancel = True
        End If
    End With
End Sub

Private Sub vsDetailList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsDetailList.EditCell
    vsDetailList.EditSelStart = 0
    vsDetailList.EditSelLength = 100
End Sub

Private Sub vsDetailList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDetailList
        If IsNumeric(.EditText) = False And .EditText <> "" Then Cancel = True: Exit Sub
        If Val(.Cell(flexcpData, Row, .ColIndex("未结金额"))) < 0 Then
            If Val(.EditText) > 0 Then Cancel = True: Exit Sub
            If Val(.EditText) < Val(.Cell(flexcpData, Row, .ColIndex("未结金额"))) Then
                .EditText = Val(.Cell(flexcpData, Row, .ColIndex("未结金额")))
            End If
        Else
            If Val(.EditText) < 0 Then Cancel = True: Exit Sub
            If Val(.EditText) > Val(.Cell(flexcpData, Row, .ColIndex("未结金额"))) Then
                .EditText = Val(.Cell(flexcpData, Row, .ColIndex("未结金额")))
            End If
        End If
    End With
End Sub

Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsFeeList, Me.Name, "费用列表"
End Sub

Private Sub vsFeeList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsFeeList, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
     zl_vsGrid_Para_Save mlngModul, vsFeeList, Me.Name, "费用列表"
End Sub

Private Sub vsFeeList_GotFocus()
    zl_VsGridGotFocus vsFeeList, &HFFC0C0
End Sub
Private Sub vsFeeList_LostFocus()
   zl_VsGridLOSTFOCUS vsFeeList
End Sub

Private Sub vsDetailList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDetailList, Me.Name, "明细列表"
End Sub

Private Sub vsDetailList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    vsDetailList.Cell(flexcpBackColor, OldRow, 0, OldRow, 3) = vbWhite
    vsDetailList.Cell(flexcpBackColor, NewRow, 0, NewRow, 3) = 16772055
    vsDetailList.Select NewRow, 4
End Sub

Private Sub vsDetailList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
     zl_vsGrid_Para_Save mlngModul, vsDetailList, Me.Name, "明细列表"
End Sub

Private Sub vsDetailList_GotFocus()
    vsDetailList.Cell(flexcpBackColor, vsDetailList.Row, 0, vsDetailList.Row, 3) = 16772055
End Sub
Private Sub vsDetailList_LostFocus()
    vsDetailList.Cell(flexcpBackColor, vsDetailList.Row, 0, vsDetailList.Row, 3) = GRD_LOSTFOCUS_COLORSEL
End Sub

Private Sub vsDeposit_GotFocus()
    zl_VsGridGotFocus vsDeposit, &HFFC0C0
End Sub
Private Sub vsDeposit_LostFocus()
   zl_VsGridLOSTFOCUS vsDeposit
End Sub
Private Sub vsBlance_GotFocus()
    If vsBlance.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vsBlance, &HFFEBD7
    
End Sub
Private Sub vsBlance_LostFocus()

    If mEditType = g_Ed_单据查看 Then Exit Sub
    If vsBlance.Row = 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vsBlance
     OS.OpenIme False
End Sub
Private Function GetOldBalanceMoney(ByVal int类型 As Integer, _
    ByVal objCard As Card) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据类型，确定原结算方式的金额
    '入参:int类型-类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '返回:返回原结算金额
    '编制:刘兴洪
    '日期:2015-01-30 17:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, rsBalance As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not mrsOldBalance Is Nothing Then
        Set rsBalance = mrsOldBalance
    Else
        Set rsBalance = mrsBalance
    End If
    If rsBalance Is Nothing Then Exit Function
    If rsBalance.State <> 1 Then Exit Function
     
    If objCard.接口序号 > 0 Then
        If objCard.消费卡 = False Then '一卡通
            rsBalance.Filter = "类型=" & int类型 & " And 卡类别ID=" & objCard.接口序号
        Else '消费卡
            rsBalance.Filter = "类型=" & int类型 & " And 结算卡序号=" & objCard.接口序号
        End If
    Else
        rsBalance.Filter = "类型=" & int类型
    End If
    
    If rsBalance.EOF Then
        If objCard.是否转帐及代扣 Then
           GetOldBalanceMoney = RoundEx(Val(mBalanceInfor.dbl未付合计), 6)
        End If
        rsBalance.Filter = 0: Exit Function
    End If
    
    rsBalance.MoveFirst
    Do While Not rsBalance.EOF
        dblMoney = dblMoney + Val(NVL(rsBalance!冲预交))
        rsBalance.MoveNext
    Loop
    GetOldBalanceMoney = dblMoney
    rsBalance.Filter = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function PatiErrBillPay(ByVal lng病人ID As Long, Optional ByVal strCheckNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人,对异常单据进行重新结帐或作废处理
    '入参:lng病人ID-指定的病人ID
    '返回:存在异常单据,并成功读取异常单据返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-02-03 11:30:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNO As String, lng结帐ID As Long
    Dim str操作员姓名 As String, strTittle As String
    Dim blnDel As Boolean, blnErrCancel As Boolean
    Dim strDelTime As String
    
    If mEditType = g_Ed_单据查看 Then Exit Function
'    If mEditType = g_Ed_门诊结帐 Or mEditType <> g_Ed_住院结帐 Then Exit Function
    
    On Error GoTo errHandle
    If strCheckNO = "" Then
        strSQL = " " & _
        "    Select  a.No, a.ID, a.操作员姓名, decode(记录状态,2,2,1) As 异常类型,A.收费时间 " & _
        "    From 病人结帐记录 A" & _
        "    Where nvl(结算状态,0) = 1 and 病人ID=[1]   And Rownum < 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    Else
        strSQL = " " & _
        "    Select  a.No, a.ID, a.操作员姓名, decode(记录状态,2,2,1) As 异常类型,A.收费时间 " & _
        "    From 病人结帐记录 A" & _
        "    Where nvl(结算状态,0) = 1 and NO=[1]   And Rownum < 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCheckNO)
    End If
    If rsTemp.EOF Then
        If strCheckNO <> "" Then PatiErrBillPay = True
        Exit Function
    End If
    
    strNO = NVL(rsTemp!NO): lng结帐ID = Val(NVL(rsTemp!ID))
    blnDel = Val(NVL(rsTemp!异常类型)) = 2
    strTittle = IIf(Not blnDel, "结帐", "重退")
    strDelTime = Format(rsTemp!收费时间, "yyyy-mm-dd HH:MM:SS")
    str操作员姓名 = NVL(rsTemp!操作员姓名)
    
    If str操作员姓名 <> UserInfo.姓名 Then
        '100703
         If MsgBox("注意:" & vbCrLf & _
                            "       该病人存在异常的" & strTittle & "单据" & IIf(str操作员姓名 <> UserInfo.姓名, ",该单据是操作员[" & str操作员姓名 & "]收取的,你无法处理," & vbCrLf, "") & " ,是否不对异常单据进行处理,继续进行结帐操作" & "?" & vbCrLf & vbCrLf & _
                            "『是』代表不对异常单据进行处理,继续进行结帐操作. " & vbCrLf & _
                            "『否』代表中止结帐操作.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            PatiErrBillPay = False
            Exit Function
         Else
            PatiErrBillPay = True
            Exit Function
         End If
    End If
    
    If MsgBox("注意:" & vbCrLf & _
                        "       该病人存在异常的" & strTittle & "单据" & IIf(str操作员姓名 <> UserInfo.姓名, ",该单据是操作员[" & str操作员姓名 & "]收取的," & vbCrLf, "") & " ,是否重新对该单据进行" & strTittle & "?" & vbCrLf & vbCrLf & _
                        "『是』代表重新对异常单据 " & strTittle & vbCrLf & _
                        "『否』代表不对异常单据进行处理,继续进行结帐操作.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Function
    End If

    If strCheckNO <> "" Then
        PatiErrBillPay = True
        Exit Function
    End If
    
    mintPreEditType = mEditType
    If blnDel Then
        Call frmPatiBalanceSplit.ShowMe(Me, g_Ed_重新作废, mstrPrivs, , , strNO, True)
    Else
        mEditType = IIf(blnDel, g_Ed_重新作废, g_Ed_重新结帐)
        mblnViewCancel = blnDel
        Call SetFeeListColumnShow
        Call SetPatiConsControlVisible
        Call SetOperatonCommandCaption
        
        If ReadBalance(strNO) Then PatiErrBillPay = True: Exit Function
    End If
    mEditType = mintPreEditType
    Call LoadBalanceBill
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DeleteBalance(Optional blnDelBalance As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐作废处理(异常作废)
    '返回:作废成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-02-03 16:36:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card, tyBrushCard As TY_BrushCard
    Dim lng冲销ID As Long, lngCount As Long, dblMoney As Double
    Dim i As Long, strBalance As String, strSQL As String
    Dim cllPro As Collection
    Dim strYbBalance As String
    Dim rsTmp As ADODB.Recordset
    
    
    On Error GoTo errHandle
    If mYBInFor.intInsure > 0 Then
        If Not MCPAR.出院病人结算作废 And mYBInFor.bytMCMode <> 1 Then
            If Not isYBPati(mPatiInfor.lng病人ID, True) Then
                MsgBox "该参保病人已经出院，不能取消该结帐单！", vbInformation, gstrSysName: Exit Function
            End If
            If MCPAR.住院结算作废 = False Then
                MsgBox "该医保不支持门诊结帐作废，不能取消该结帐单！", vbInformation, gstrSysName: Exit Function
            End If
        ElseIf mYBInFor.bytMCMode = 1 And Not MCPAR.门诊病人结算作废 Then
                MsgBox "该医保不支持门诊结帐作废，不能取消该结帐单！", vbInformation, gstrSysName: Exit Function
        End If
        If gclsInsure.CheckInsureValid(mYBInFor.intInsure) = False Then Exit Function
    End If
    
    Set objTemp = Nothing
    With vsBlance
        For i = 1 To .Rows - 1
            strBalance = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            
            If strBalance <> "" Then
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 0 '普通结算
                Case 1 '预交款
                Case 2 '医保
                    strYbBalance = strYbBalance & "," & strBalance
                    
                Case 3 '一卡通
                    Set objTemp = GetCard(strBalance)  'GetLocalePayCard(Val(.TextMatrix(i, .ColIndex("卡类别ID"))), False)
                    If objTemp Is Nothing Then
                        MsgBox "当前站点不支持" & strBalance & "方式进行退费处理,不允许作废!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                     dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
                    If CheckThreeSwapValied(objTemp, dblMoney, tyBrushCard, True) = False Then Exit Function
                    lngCount = lngCount + 1
                Case 4 '一卡通(老版本)
                    dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
                    Set objTemp = GetLocaleOldOneCard(strBalance)
                    If objTemp Is Nothing Then
                        MsgBox "当前站点不支持" & strBalance & "进行退费处理,不允许作废!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    If CheckOldOneCardIsValied(objTemp, dblMoney, tyBrushCard, True) = False Then Exit Function
                    lngCount = lngCount + 1
                Case 5 '消费卡
                Case Else
                End Select
            End If
        Next
    End With
    If Not mrsBalance Is Nothing Then
        mrsBalance.Filter = 0
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            If Val(NVL(mrsBalance!类型)) = 2 And InStr(strYbBalance & ",", "," & mrsBalance!结算方式 & ",") = 0 Then
                MsgBox "该医保不支持“" & mrsBalance!结算方式 & "”原样退回处理,不允许作废!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            mrsBalance.MoveNext
        Loop
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    End If
    
    strSQL = "Select 1 From 三方退款信息 Where 结帐ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBalanceInfor.lng结帐ID)
    If Not rsTmp.EOF Then
        MsgBox IIf(blnDelBalance, "结帐", "异常") & "作废暂不支持包含多笔退款接口的交易,请重新收费结帐!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If lngCount + IIf(mYBInFor.intInsure > 0, 1, 0) > 1 Then
        MsgBox IIf(blnDelBalance, "结帐", "异常") & "作废暂不支持两种接口以上的交易,你必须完成结帐!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
    With mBalanceInfor
        .lng冲销ID = lng冲销ID
        .dtBalanceDate = zlDatabase.Currentdate
    End With
    Set cllPro = New Collection
     '先退结算记录及费用
     strSQL = "Zl_病人结帐记录_Cancel("
     '  No_In         病人结帐记录.No%Type,
     strSQL = strSQL & "'" & mBalanceInfor.strNO & "',"
     '  冲销id_In     病人结帐记录.Id%Type,
     strSQL = strSQL & "" & lng冲销ID & ","
     '  操作员编号_In 病人结帐记录.操作员编号%Type,
     strSQL = strSQL & "'" & UserInfo.编号 & "',"
     '  操作员姓名_In 病人结帐记录.操作员姓名%Type
     strSQL = strSQL & "'" & UserInfo.姓名 & "')"
     zlAddArray cllPro, strSQL
     
     
    'Zl_病人结帐作废_Modify
    strSQL = "Zl_病人结帐作废_Modify("
    '  操作类型_In   Number,
    strSQL = strSQL & "" & 0 & ","
    '  病人id_In     病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & mPatiInfor.lng病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL,"
    '  找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  预交金额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '冲预交病人ids_In Varchar2 := Null,
    ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
    strSQL = strSQL & "NULL,"
    '  完成作废_In Number:=0
    strSQL = strSQL & "1)"
    zlAddArray cllPro, strSQL
    '执行医保退费操作
    If ExecuteInsureDel(cllPro, True) = False Then Exit Function
    
    If Not objTemp Is Nothing Then
        If ExecuteThreeSwapDelInterface(objTemp, dblMoney, cllPro, True) = False Then Exit Function
        If ExecuteOneCardDelInterface(objTemp, dblMoney, cllPro, True) = False Then Exit Function
    End If
    
    strSQL = "Zl_病人结帐异常_Update("
    strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    strSQL = strSQL & "" & lng冲销ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    DeleteBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SkipSetFocus(ByVal bytCurOper As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:光标移动位置
    '入参:bytCurOper-当前正处于的操作(0-查找病人;1-当前在结帐说明,2-当前输入结帐金额)
    '
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-11 17:39:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case bytCurOper
    Case 0 '查找病人
        If Not (vsBlance.Enabled And vsBlance.Visible) Then zlCommFun.PressKey vbKeyTab: Exit Sub
        '定位在结算方式框
        With vsBlance
            If .Row <= 0 And .Rows > 1 Then .Row = 1
            If .Col <= 0 And .Cols >= .ColIndex("结算方式") Then .Col = .ColIndex("结算方式")
            .ShowCell .Row, .Col
            .SetFocus
        End With
        Exit Sub
    Case 1 '结帐说明
        If cmdYBBalance.Enabled And cmdYBBalance.Visible Then cmdYBBalance.SetFocus: Exit Sub
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus
        Exit Sub
    Case 2 '当前输入结帐金额
        If txtReceive.Enabled And txtReceive.Visible Then txtReceive.SetFocus: Exit Sub
        If vsBlance.Enabled And vsBlance.Visible Then vsBlance.SetFocus
        Exit Sub
       Exit Sub
    Case Else
    End Select
End Sub
Private Function CheckPatiFromZyNumIsYB(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef intInsure As Integer, Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择指定住院次是否为医保病人
    '入参:
    '出参:intInsure-返回医保序号
    '     strInsureName-医保名称
    '返回:是医保返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-13 09:53:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    intInsure = 0
    If Not mobjBalanceAll.rsAllTime Is Nothing Then
        With mobjBalanceAll.rsAllTime
            If .State = 1 Then
                .Filter = "主页ID=" & lng主页ID
                If Not .EOF Then
                    intInsure = Val(NVL(!险类))
                    strInsureName = Trim(NVL(!保险名称))
                    CheckPatiFromZyNumIsYB = intInsure <> 0
                    Exit Function
                End If
            End If
        End With
    End If
    
    strSQL = "Select Nvl(a.险类,0) As 险类,b.名称 From 病案主页 A,保险类别  b Where a.险类=b.序号(+) and A.病人ID = [1] And A.主页ID =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(mrsInfo!病人ID)), lng主页ID)
    If rsTemp.EOF Then Exit Function
    
    intInsure = Val(NVL(rsTemp!险类))
    strInsureName = Trim(NVL(rsTemp!保险名称))
    CheckPatiFromZyNumIsYB = intInsure <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function LoadDataPatiNumsToComBox(ByVal lng病人ID As Long, ByVal str主页Ids As String, ByRef blnAllSel As Boolean, _
    ByRef rsTimeAll As ADODB.Recordset, ByRef intInsure As Integer, Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载住院次数，给下拉列表框
    '入参: str主页IDs-所有住院次数,用逗号分隔
    '出参:blnAllSel-当前是否选择了所有住院次数
    '     intInsure-返回第一个选择的医保序号
    '     strInsureName-返回第一个选择的医保名称
    '返回:加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-13 11:23:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, int主页ID As Long, strTag As String
    Dim i As Long, intInsure1 As Integer, strInsureName1 As String
    
    On Error GoTo errHandle
    
    cboPatiNums.Clear
    If mEditType <> g_Ed_住院结帐 Then
        cboPatiNums.AddItem "R", "所有门诊", True, True, True, , "0"
        varTemp = Split(str主页Ids, ",")
        blnAllSel = True
        For i = 0 To UBound(varTemp)
            If Val(varTemp(i)) = 0 Then
                cboPatiNums.AddItem Val(varTemp(i)), "普通门诊", False, True
            Else
                cboPatiNums.AddItem Val(varTemp(i)), "第" & Val(varTemp(i)) & "次留观", False, True
            End If
        Next
        Call cboPatiNums.Refresh
        Set rsTimeAll = Nothing
        LoadDataPatiNumsToComBox = True
        Exit Function
    End If
    
    cboPatiNums.AddItem "R", "所有住院", True, True, True, , "0"
    '获取当前未结住院次所涉及的医保数据集
    Call mobjBalanceAll.zlGetTimeRecordFromTimeString(lng病人ID, str主页Ids, rsTimeAll)

    '加载住院次数文本框
    Dim blnSelect As Boolean
    With rsTimeAll
        intInsure = 0
        If .RecordCount <> 0 Then
            .MoveFirst:  intInsure = Val(NVL(!险类)): strInsureName = NVL(!保险名称)
        End If
        
        i = 1: blnAllSel = True
        Do While Not .EOF
            '自费的，先缺省全选,最后一次住院为医保的，则先结医保的
            
            blnSelect = mobjBalanceAll.strAllOwnerFeeType <> "" Or (intInsure <> 0 And i = 1) Or intInsure = 0
            If Not blnSelect And intInsure <> 0 And MCPAR.允许结多次住院费用 Then blnSelect = True
            
            If blnAllSel And Not blnSelect Then blnAllSel = False
            
            int主页ID = Val(NVL(!主页ID)): intInsure1 = Val(NVL(!险类)): strInsureName1 = NVL(!保险名称)
            strTag = int主页ID & "|" & Val(NVL(!险类)) & "|" & NVL(!保险名称)
            
            cboPatiNums.AddItem int主页ID, "第" & int主页ID & "次住院" & IIf(Val(NVL(!险类)) <> 0, "(医保)", ""), False, blnSelect, , , strTag
            i = i + 1
            .MoveNext
        Loop
     End With
     Call cboPatiNums.Refresh
    LoadDataPatiNumsToComBox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeFromType() As String
    '获取收费单据来源类型
    '返回：费用来源，多个用逗号分隔
    Dim i As Long
    Dim str费用来源 As String, byt费用来源 As Byte
    
    On Error GoTo errHandle
    If mEditType = g_Ed_门诊结帐 Or mblnCurMzBalanceNo Then '门诊
        str费用来源 = ""
    Else '住院
        GetFeeFromType = "2": Exit Function
    End If
    
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据")) <> "" Then
                If Not (Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))) = 0 And Val(.Cell(flexcpData, i, .ColIndex("未结金额"))) <> 0) Then
                    byt费用来源 = Decode(Val(.Cell(flexcpData, i, .ColIndex("序号"))), 4, 3, 2, 2, 1)
                    If InStr(str费用来源, byt费用来源) = 0 Then
                        str费用来源 = str费用来源 & "," & byt费用来源
                    End If
                End If
            End If
        Next
    End With
    If Left(str费用来源, 1) = "," Then str费用来源 = Mid(str费用来源, 2)
    GetFeeFromType = str费用来源
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DepositMonyVerfy(Optional blnSaveCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:预交金额合法性佼对
    '入参:blnSaveCheck-true:点击完成时，没效对的检查;False-文本框校对检查(valied事件调用)
    '出参:
    '返回:校对成功rue,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-28 11:31:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, blnNoRecal As Boolean
    
    On Error GoTo errHandle
    
    If chkDeposit.Visible Then DepositMonyVerfy = True: Exit Function
    
    dblMoney = RoundEx(Val(txtBalance(Idx_冲预交).Text), 6)
    
    If mblnNotChange = False Then
        If Val(dblMoney) > Val(mPatiInfor.dbl实际余额) Then
            MsgBox "当前输入的冲预交大于预交余额,不能继续!" & vbCrLf & "实际余额:" & Format(mPatiInfor.dbl实际余额, "0.00") & vbCrLf & "冲预交:" & Format(Val(txtBalance(Idx_冲预交).Text), "0.00")
            Exit Function
        End If
    End If
    
    blnNoRecal = dblMoney = mBalanceInfor.dbl冲预交合计 And dblMoney <> 0
    
    If blnNoRecal = False Then
        '金额相等，就不用再重新计算
        If GetDepositTotal = dblMoney Then mBalanceInfor.dbl冲预交合计 = dblMoney
    End If
    
    '操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按结帐金额来冲预交(按时间先后来分摊）;3-全冲
    If dblMoney <> mBalanceInfor.dbl冲预交合计 And mBalanceInfor.bln预交刷卡 = False Then
        If dblMoney = 0 Then
            Call RecalcDepositMoney(0)
        Else
            Call RecalcDepositMoney(2, dblMoney)
        End If
        
        mblnNotChange = True
        txtBalance(Idx_冲预交).Text = Format(mBalanceInfor.dbl冲预交合计, "0.00")
        mblnNotChange = False
    End If
    If mblnNotChange Then DepositMonyVerfy = True: Exit Function
    
    If Not mBalanceInfor.bln预交刷卡 Then
        If CheckDepositValied(True) = False Then Exit Function
    End If
    
    If Not blnNoRecal Then
        Call LoadIntendBalance
    End If
'    If blnSaveCheck And dblMoney = mBalanceInfor.dbl冲预交合计 Then
'        '金额未发生变化，不再计算结算信息相关
'        DepositMonyVerfy = True: Exit Function
'    End If
    Call LoadCurOwnerPayInfor(True)
    DepositMonyVerfy = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetDepositTotal(Optional ByVal bln余额 As Boolean = False) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取冲预交总额或余额总额
    '入参:bln余额-获取余额总额
    '出参:
    '返回:返回冲预交总额或余额总额
    '编制:刘兴洪
    '日期:2017-12-28 11:31:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Integer, i As Long
    Dim dblTemp As Double
    With vsDeposit
        dblTemp = 0
        For i = 1 To .Rows - 1
            intCol = IIf(bln余额, .ColIndex("余额"), .ColIndex("冲预交"))
            If intCol >= 0 Then
              dblTemp = dblTemp + Val(.TextMatrix(i, intCol))
            End If
        Next i
        dblTemp = RoundEx(dblTemp, 5)
    End With
End Function

