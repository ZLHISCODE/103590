VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanEdit 
   Caption         =   "出诊安排设置"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15120
   Icon            =   "frmClinicPlanEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picVerify 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   11460
      ScaleHeight     =   375
      ScaleWidth      =   1755
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   900
      Width           =   1755
      Begin VB.CheckBox chkAotuVerify 
         Caption         =   "保存后立即审核"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   90
         TabIndex        =   8
         Top             =   90
         Width           =   1605
      End
   End
   Begin VB.PictureBox picFeeItem 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7800
      ScaleHeight     =   375
      ScaleWidth      =   2955
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   990
      Width           =   2955
      Begin VB.ComboBox cboFeeItem 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   30
         Width           =   2055
      End
      Begin VB.Label lblFeeItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费项目"
         Height          =   180
         Left            =   90
         TabIndex        =   53
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.PictureBox picSignal 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4290
      ScaleHeight     =   375
      ScaleWidth      =   2505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   990
      Width           =   2505
      Begin VB.TextBox txtSignal 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   1
         ToolTipText     =   "当前允许输入号码，医生姓名或简码，科室名称或简码进行查找。"
         Top             =   30
         Width           =   2025
      End
      Begin VB.Label lblSignal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号码"
         Height          =   180
         Left            =   30
         TabIndex        =   52
         Top             =   90
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList imglist16 
      Left            =   1260
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":6852
            Key             =   "plan_deleted"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":6DEC
            Key             =   "plan_nosave"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":7386
            Key             =   "plan_nothing"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":7920
            Key             =   "plan_saved"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSouceList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   30
      ScaleHeight     =   2400
      ScaleWidth      =   3195
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5910
      Width           =   3195
      Begin zl9RegEvent.ShowSourceInfor SourceInfor 
         Height          =   1635
         Left            =   570
         TabIndex        =   28
         Top             =   390
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2884
         BackColor       =   16773091
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
      Begin VB.Image imgSignalSource 
         Height          =   240
         Left            =   45
         Picture         =   "frmClinicPlanEdit.frx":7EBA
         Top             =   60
         Width           =   240
      End
      Begin VB.Label lblSourceTittle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号源信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   51
         Top             =   90
         Width           =   780
      End
      Begin VB.Shape shpSourceLine 
         BackColor       =   &H00ECEDC2&
         BorderColor     =   &H80000003&
         Height          =   915
         Left            =   30
         Top             =   390
         Width           =   360
      End
   End
   Begin VB.PictureBox picSourceAndPlan 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   3000
      ScaleHeight     =   3285
      ScaleWidth      =   5595
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7050
      Visible         =   0   'False
      Width           =   5595
      Begin VB.PictureBox picPlan 
         BackColor       =   &H00FFEFE3&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   2010
         ScaleHeight     =   2775
         ScaleWidth      =   3135
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   150
         Width           =   3135
         Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
            Height          =   2085
            Left            =   180
            TabIndex        =   26
            Top             =   600
            Width           =   2805
            _cx             =   4948
            _cy             =   3678
            Appearance      =   2
            BorderStyle     =   0
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
            BackColor       =   16773091
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16773091
            BackColorAlternate=   16773091
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16773091
            FocusRect       =   0
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClinicPlanEdit.frx":8444
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
            FrozenCols      =   2
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblPlanInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "安排预览"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   270
            TabIndex        =   49
            Top             =   90
            Width           =   780
         End
         Begin VB.Image imgPlanInfo 
            Height          =   240
            Left            =   45
            Picture         =   "frmClinicPlanEdit.frx":84D7
            Top             =   60
            Width           =   240
         End
         Begin VB.Shape shpPlanLine 
            BackColor       =   &H00ECEDC2&
            BorderColor     =   &H80000003&
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
      Begin XtremeSuiteControls.TabControl tbPageSourceAndPlan 
         Height          =   705
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   1605
         _Version        =   589884
         _ExtentX        =   2831
         _ExtentY        =   1244
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picValidTime 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5310
      ScaleHeight     =   375
      ScaleWidth      =   5205
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   5205
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   675
         TabIndex        =   3
         Top             =   30
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483641
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   221446147
         CurrentDate     =   38091
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3060
         TabIndex        =   4
         Top             =   30
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483641
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   221446147
         CurrentDate     =   38091
      End
      Begin VB.Label lblValidTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "有效期"
         Height          =   180
         Left            =   90
         TabIndex        =   30
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblTimeRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   180
         Left            =   2820
         TabIndex        =   50
         Top             =   90
         Width           =   180
      End
   End
   Begin MSComctlLib.ImageList img11 
      Left            =   3810
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":8A61
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":8F6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":9475
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanEdit.frx":997F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDateList 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   150
      ScaleHeight     =   3060
      ScaleWidth      =   3045
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   420
      Width           =   3045
      Begin zl9RegEvent.CalendarSel cldsCalenbarSel 
         Height          =   1845
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3254
         BackColor       =   16773091
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowStyle       =   2
      End
      Begin VB.Shape shpItemSel 
         BorderColor     =   &H80000003&
         Height          =   2625
         Left            =   0
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.PictureBox picWorkTimeList 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   90
      ScaleHeight     =   2400
      ScaleWidth      =   3195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3450
      Width           =   3195
      Begin MSComctlLib.ListView lvwWorkTime 
         Height          =   1035
         Left            =   240
         TabIndex        =   12
         Top             =   330
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1826
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16773091
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "时间段"
            Object.Width           =   9596
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "开始时间"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "终止时间"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblCalendbarTittle 
         BackStyle       =   0  'Transparent
         Caption         =   "上班时段"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   31
         Top             =   90
         Width           =   810
      End
      Begin VB.Image imgWork 
         Height          =   240
         Left            =   60
         Picture         =   "frmClinicPlanEdit.frx":9E89
         Top             =   75
         Width           =   240
      End
      Begin VB.Shape shpWorkLine 
         BackColor       =   &H00FFEFE3&
         BorderColor     =   &H80000003&
         Height          =   2295
         Left            =   30
         Top             =   30
         Width           =   3150
      End
   End
   Begin VB.PictureBox picDetailedList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7500
      Left            =   3900
      ScaleHeight     =   7500
      ScaleWidth      =   13935
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1950
      Width           =   13935
      Begin zl9RegEvent.ClinicPlanDetailPages CPDPages 
         Height          =   2505
         Left            =   2040
         TabIndex        =   16
         Top             =   2220
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   4419
         BackColor       =   -2147483628
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
      Begin VB.PictureBox picApply 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   2460
         ScaleHeight     =   825
         ScaleMode       =   0  'User
         ScaleWidth      =   8580
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   8580
         Begin VB.PictureBox picApplyRule 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   900
            ScaleHeight     =   345
            ScaleWidth      =   7905
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   7905
            Begin VB.Frame fraLoopSkip 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   3975
               TabIndex        =   39
               Top             =   30
               Width           =   3735
               Begin VB.ComboBox cboDays 
                  Height          =   300
                  Left            =   1110
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   0
                  Width           =   1260
               End
               Begin VB.TextBox txtSkip 
                  Height          =   285
                  Left            =   2820
                  Locked          =   -1  'True
                  TabIndex        =   20
                  Text            =   "7"
                  Top             =   15
                  Width           =   330
               End
               Begin MSComCtl2.UpDown updSkip 
                  Height          =   285
                  Left            =   3150
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   503
                  _Version        =   393216
                  Value           =   1
                  BuddyControl    =   "txtSkip"
                  BuddyDispid     =   196650
                  OrigLeft        =   3225
                  OrigTop         =   15
                  OrigRight       =   3480
                  OrigBottom      =   300
                  Max             =   30
                  Min             =   1
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.Label lblLoopDate 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "开始轮循日期"
                  Height          =   180
                  Left            =   0
                  TabIndex        =   41
                  Top             =   60
                  Width           =   1080
               End
               Begin VB.Label lblLoopSkipDays 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "间隔       天"
                  Height          =   180
                  Left            =   2445
                  TabIndex        =   42
                  Top             =   60
                  Width           =   1170
               End
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "轮循"
               Height          =   240
               Index           =   4
               Left            =   3150
               TabIndex        =   33
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "星期"
               Height          =   240
               Index           =   3
               Left            =   2325
               TabIndex        =   38
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "双日"
               Height          =   240
               Index           =   2
               Left            =   1530
               TabIndex        =   37
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "单日"
               Height          =   240
               Index           =   1
               Left            =   735
               TabIndex        =   36
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "当前"
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   75
               Value           =   -1  'True
               Width           =   705
            End
         End
         Begin VB.PictureBox picApplyWeek 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   930
            ScaleHeight     =   255
            ScaleWidth      =   7020
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   390
            Width           =   7020
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周日"
               Height          =   180
               Index           =   6
               Left            =   5685
               TabIndex        =   40
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周六"
               Height          =   180
               Index           =   5
               Left            =   4735
               TabIndex        =   48
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周五"
               Height          =   180
               Index           =   4
               Left            =   3788
               TabIndex        =   47
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周四"
               Height          =   180
               Index           =   3
               Left            =   2841
               TabIndex        =   46
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周三"
               Height          =   180
               Index           =   2
               Left            =   1894
               TabIndex        =   45
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周二"
               Height          =   180
               Index           =   1
               Left            =   947
               TabIndex        =   44
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "周一"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   43
               Top             =   30
               Width           =   690
            End
         End
         Begin VB.Label lblApply 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应用于(&Y)"
            Height          =   180
            Left            =   15
            TabIndex        =   34
            Top             =   90
            Width           =   810
         End
      End
      Begin zl9RegEvent.CustomButton btnLeft 
         Height          =   315
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   75
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
      End
      Begin zl9RegEvent.CustomButton btnRight 
         Height          =   315
         Left            =   390
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   75
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
      End
      Begin VB.Shape shpDetailedList 
         BackColor       =   &H00FFEFE3&
         BorderColor     =   &H80000003&
         Height          =   1305
         Left            =   0
         Top             =   0
         Width           =   2280
      End
      Begin VB.Label lblTittle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "星期二"
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
         Left            =   840
         TabIndex        =   32
         Top             =   75
         Width           =   990
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   10575
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicPlanEdit.frx":A413
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21590
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmClinicPlanEdit.frx":ACA7
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmClinicPlanEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private Enum m_Fun
    F_FixedRule = 0
    F_MonthPlan = 1
    F_WeekPlan = 2
    F_Templet = 3
    F_MonthTemplet = 4
End Enum
Private mbytPlanType As m_Fun '0-固定排班,1-按月排班,2-按周排班,3-模板,4-按天安排出诊的月模板
Private mbytFun As G_Enum_Fun

Public Enum Pancel_Index
    Pan_日历 = 1001
    Pan_时间段 = 1002
    Pan_号源 = 1003
    Pan_详情 = 1004
End Enum
Private mblnFirst As Boolean
Private mblnTimeChanged As Boolean, mblnFeeItemChanged As Boolean

Private mobj出诊安排 As 出诊安排, mlng出诊ID As Long
Private mlng号源Id As Long, mlng安排ID As Long
Private mstr缺省日期 As String
Private mlngSavedRecords As Long
Private mblnNotClick As Boolean

Private mstrCurDay As String
Private mdtToday As Date '当前服务器日期
Private mrsVisitedRecord As ADODB.Recordset '该号源已出诊记录
Private mobj停诊记录集 As 停诊记录集 '当前号源的出诊记录集
Private mrsVisitedRecordByDate As ADODB.Recordset '当前号源某个日期的出诊记录
Private Type FixedPlanDateRange
    dtStart As Date
    dtEnd As Date
End Type
Private mFixedPlanDateRange As FixedPlanDateRange
Private mstrPrivs As String
Private mcllFixedPlan As Collection  '临时安排，记录已有预约挂号记录的出诊记录，Array(出诊日期,限制项目,上班时段,开始时间,终止时间)

Private mblnCheckedByDay As Boolean '是否已按天检查
Private mblnValiedCanSave As Boolean

Private Enum mPgIndex 'TabPage索引
    Pg_号源信息 = 0
    Pg_安排预览 = 1
End Enum

Private Enum mMenuID
    M_Signal = 1
    M_ValidTime = 2
    M_FeeItem = 3
    M_Verify = 4
End Enum

Public Function ShowMe(frmParent As Object, ByVal bytPlanType As Byte, ByVal bytFun As G_Enum_Fun, ByVal lng出诊ID As Long, _
    Optional ByVal lng号源Id As Long, Optional ByVal lng安排ID As Long, _
    Optional ByVal str缺省日期 As String, Optional ByVal strPrivs As String) As Boolean
    '功能：程序入口
    '入参：
    '   bytPlanType 0-固定排班,1-按月排班,2-按周排班,3-模板,4-按天安排出诊的月模板
    '   bytFun '0-查看,1-新增,2-修改,3-删除,4-临时安排(固定出诊表),5-调整合作单位预约挂号,6-新增号源,7-发布后临时出诊
    '   lng号源ID/lng安排ID 6-新增号源时可不传入
    '   str缺省日期 缺省选中日期
    mbytPlanType = bytPlanType: mbytFun = bytFun: mlng出诊ID = lng出诊ID
    mlng号源Id = lng号源Id: mlng安排ID = lng安排ID
    mstr缺省日期 = str缺省日期
    mstrPrivs = strPrivs
    mlngSavedRecords = 0
    mstrCurDay = ""
    Set mobj出诊安排 = New 出诊安排

    On Error Resume Next
    Me.Show 1, frmParent
    ShowMe = mlngSavedRecords > 0
End Function

Private Function CheckDepend(ByVal bytMode As Byte, ByVal strCurDate As String, _
    Optional ByVal dtStartTime As Date, Optional ByVal dtEndTime As Date, _
    Optional ByVal blnShowErr As Boolean = True) As Boolean
    '功能:检查上班时间内是否可出诊
    '参数:
    '   bytMode - 0,检查今日是否可出诊;1,检查时间范围dtStart~dtEnd是否可出诊
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj停诊记录 As 停诊记录
    Dim obj出诊记录 As 出诊记录, blnAllInvalid As Boolean

    On Error GoTo ErrHandler
    If IsDate(strCurDate) = False Then CheckDepend = True: Exit Function
    If mbytPlanType = F_MonthTemplet Then CheckDepend = True: Exit Function '按天安排出诊的月模板退出
    If Not (mbytFun = Fun_TempPlanRecord _
        Or mbytFun = Fun_AddSignalSourcePlan And mlng号源Id <> 0 _
        Or mbytFun = Fun_UpdatePlan) Then CheckDepend = True: Exit Function
    
    If mbytFun = Fun_UpdatePlan Then
        '如果时段全部停诊或已使用，则不允许调整
        If bytMode = 0 Then
            If mobj出诊安排(1).Count > 0 Then
                blnAllInvalid = True
                For Each obj出诊记录 In mobj出诊安排(1)
                    If obj出诊记录.是否固定 = False Then blnAllInvalid = False
                Next
                If blnAllInvalid Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 没有可调整的安排（这些安排已失效或已停诊或已用于预约挂号），不允许调整！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, Format(strCurDate, "yyyy-mm-dd")) Then
            If blnShowErr Then
                If mbytFun = Fun_TempPlanRecord Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 是在其它出诊表中设置的出诊安排，不能在当前出诊表中进行临时出诊安排！", vbInformation, gstrSysName
                Else
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 已在其它出诊表中设置了出诊安排，不能重复安排！", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
        
        If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, Format(strCurDate, "yyyy-mm-dd"), False, mlng安排ID) Then
            If blnShowErr Then
                If mbytFun = Fun_TempPlanRecord Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 是在当前出诊表的其它安排中设置的出诊安排，不能在当前安排中进行临时出诊安排！", vbInformation, gstrSysName
                Else
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 已在当前出诊表的其它安排中设置的出诊安排，不能重复安排！", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
        
        If mobj出诊安排.排班方式 = 0 Then
            '在月/周安排中的固定安排不能临时出诊
            strSQL = "Select b.排班方式" & vbNewLine & _
                    " From 临床出诊安排 A, 临床出诊表 B" & vbNewLine & _
                    " Where a.出诊id = b.Id And a.号源id = [2] And Nvl(b.排班方式, 0) In (1, 2) And [3] Between a.开始时间 And a.终止时间" & vbNewLine & _
                    "       And Exists(Select 1" & vbNewLine & _
                    "           From 临床出诊安排 M, 临床出诊表 N" & vbNewLine & _
                    "           Where m.出诊id = n.Id And Nvl(n.排班方式, 0) = 0 And m.号源id = a.号源id And n.Id = [1]) And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlng出诊ID, mlng号源Id, CDate(strCurDate))
            If Not rsTemp.EOF Then
                If blnShowErr Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 已在" & IIf(Val(Nvl(rsTemp!排班方式)) = 1, "月", "周") & "出诊表中，不能在当前安排中进行临时出诊安排！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            
            strSQL = "Select c.排班方式" & vbNewLine & _
                    " From 临床出诊安排 A, 临床出诊表 B, 临床出诊号源 C" & vbNewLine & _
                    " Where a.出诊id = b.Id And Nvl(b.排班方式, 0) In (1, 2) And a.号源id = c.Id And Nvl(c.排班方式, 0) <> 0" & vbNewLine & _
                    "       And c.Id = [1] And a.开始时间 < [2] And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlng号源Id, CDate(strCurDate))
            If Not rsTemp.EOF Then
                If blnShowErr Then
                    MsgBox "当前号源已调整为了按" & IIf(Val(Nvl(rsTemp!排班方式)) = 1, "月", "周") & "排班，不能在当前安排中进行临时出诊安排！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
    End If

    If bytMode = 0 Then
        dtStartTime = CDate(Format(strCurDate, "yyyy-mm-dd"))
        dtEndTime = CDate(Format(strCurDate, "yyyy-mm-dd") & " 23:59:59")
    End If
    dtEndTime = GetWorkTrueDate(dtStartTime, dtEndTime)

    '不能对历史的安排进行操作
    If bytMode = 0 Then
        If DateDiff("d", strCurDate, Format(zlDatabase.Currentdate, "yyyy-mm-dd")) > 0 Then
            If blnShowErr Then MsgBox "不能对历史的出诊日期进行" & _
                IIf(mbytFun = Fun_UpdatePlan, "调整安排！", IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊安排！"), vbInformation, gstrSysName
            Exit Function
        End If
        '当前日期是否在有效时间范围内
        If mobj出诊安排.排班方式 = 0 Then
            If Not (DateDiff("d", dtEndTime, mobj出诊安排.开始时间) <= 0 And DateDiff("d", dtStartTime, mobj出诊安排.终止时间) >= 0) Then
                If blnShowErr Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 不在该号源的可预约时间范围内不能进行" & _
                        IIf(mbytFun = Fun_UpdatePlan, "调整安排！", IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊安排！"), vbInformation, gstrSysName
                End If
                Exit Function
            End If
            If Not (DateDiff("d", dtEndTime, mFixedPlanDateRange.dtStart) <= 0 And DateDiff("d", dtStartTime, mFixedPlanDateRange.dtEnd) >= 0) Then
                If blnShowErr Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " 不在该号源出诊安排的有效范围内不能进行" & _
                        IIf(mbytFun = Fun_UpdatePlan, "调整安排！", IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊安排！"), vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
    Else
        '当前上班时段是否在有效时间范围内
        If mobj出诊安排.排班方式 = 0 Then
            If Not (DateDiff("d", dtEndTime, mFixedPlanDateRange.dtStart) <= 0 And DateDiff("d", dtStartTime, mFixedPlanDateRange.dtEnd) >= 0) Then
                If blnShowErr Then MsgBox "当前上班时段不在该出诊安排的有效范围内不能进行" & IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊安排！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mbytFun = Fun_UpdatePlan Then CheckDepend = True: Exit Function
    If mblnCheckedByDay Then CheckDepend = True: Exit Function
    If mobj停诊记录集 Is Nothing Then
        Set mobj停诊记录集 = GetStopVisitObjects(GetStopVisit(mlng号源Id, mobj出诊安排.开始时间, mobj出诊安排.终止时间, False))
    End If
    If mobj停诊记录集 Is Nothing Then CheckDepend = True: Exit Function
    If mobj停诊记录集.Count = 0 Then CheckDepend = True: Exit Function
    For Each obj停诊记录 In mobj停诊记录集
        '整体都停诊，则不允许设置，否则在选择上班时段时再判断
        If DateDiff("s", obj停诊记录.开始时间, dtStartTime) >= 0 _
            And DateDiff("s", obj停诊记录.终止时间, dtEndTime) <= 0 Then
            Select Case obj停诊记录.类型
            Case 1   '停诊安排
                If bytMode = 0 Then
                    mblnCheckedByDay = True
                    If blnShowErr Then
                        If MsgBox("注意：" & vbCrLf & _
                                  "    当前号源的出诊医生在今日已停诊，你确定要安排" & IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊吗？", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    If blnShowErr Then
                        If MsgBox("注意：" & vbCrLf & _
                                  "    当前号源的出诊医生在该上班时段的时间范围内已停诊，你确定要安排" & IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊吗？", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            Case 2   '法定节假日
                If bytMode = 0 Then
                    mblnCheckedByDay = True
                    If blnShowErr Then
                        If MsgBox("注意：" & vbCrLf & _
                                  "    当前号源今日因为法定节假日(" & obj停诊记录.停诊原因 & ")已停诊，你确定要安排" & IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊吗？", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    If blnShowErr Then
                        If MsgBox("注意：" & vbCrLf & _
                                  "    当前上班时段在法定节假日(" & obj停诊记录.停诊原因 & ")停诊时间范围内，你确定要安排" & IIf(mbytFun = Fun_TempPlanRecord, "临时", "") & "出诊吗？", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            End Select
        End If
    Next
    CheckDepend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom

    Err = 0: On Error GoTo Errhand:
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
    End With
    cbsThis.EnableCustomization False

    '菜单定义
    cbsThis.DeleteAll

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched

    With cbrToolBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_Signal, "号码"): cbrCustom.Handle = picSignal.Hwnd
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_ValidTime, "有效期"): cbrCustom.Handle = picValidTime.Hwnd
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_FeeItem, "收费项目"): cbrCustom.Handle = picFeeItem.Hwnd
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_Verify, "保存后立即审核"): cbrCustom.Handle = picVerify.Hwnd
        
        If mbytFun = Fun_Delete Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "确定"): cbrControl.flags = xtpFlagRightAlign
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "取消    "): cbrControl.flags = xtpFlagRightAlign
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.flags = xtpFlagRightAlign
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出    "): cbrControl.flags = xtpFlagRightAlign
        End If
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlLabel And cbrControl.Type <> xtpControlCustom Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next

    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
    End With

    zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitPanel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Docking控件
    '编制:刘兴洪
    '日期:2016-01-08 14:34:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    Dim strReg As String
    Dim panThis As Pane, panLeft As Pane

    On Error GoTo Errhand
    dkpMain.SetCommandBars cbsThis
    dkpMain.VisualTheme = ThemeOffice2003 '设置显示风格
    sngWidth = 200
    sngHeight = 200
    
    Set panLeft = dkpMain.CreatePane(Pancel_Index.Pan_日历, sngWidth, sngHeight, DockTopOf, Nothing)
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Title = "": panLeft.Tag = Pancel_Index.Pan_日历
    panLeft.Handle = picDateList.Hwnd
    '固定大小
    If cldsCalenbarSel.ShowStyle = Show_Plan_Week Then
        panLeft.MinTrackSize.Height = 120
        panLeft.MaxTrackSize.Height = 120
    ElseIf cldsCalenbarSel.ShowStyle = Show_Plan_Rule Then
        panLeft.MinTrackSize.Height = 165
        panLeft.MaxTrackSize.Height = 165
    Else
        panLeft.MinTrackSize.Height = 200
        panLeft.MaxTrackSize.Height = 200
    End If
    panLeft.MinTrackSize.Width = 200
    panLeft.MaxTrackSize.Width = 200

    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_详情, sngWidth, 300, DockRightOf, panLeft)
    panThis.Title = ""
    panThis.Tag = Pancel_Index.Pan_详情
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picDetailedList.Hwnd
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_时间段, sngWidth, 250, DockBottomOf, panLeft)
    panThis.Title = "上班时间"
    panThis.Tag = Pancel_Index.Pan_时间段
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picWorkTimeList.Hwnd

    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_号源, sngWidth, 300, DockBottomOf, panThis)
    panThis.Title = "当前号源信息"
    panThis.Tag = Pancel_Index.Pan_号源
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    If mbytPlanType = F_FixedRule Then
        Call InitPage
        picSourceAndPlan.Visible = True
        Set picSouceList.Container = picSourceAndPlan
        panThis.Handle = picSourceAndPlan.Hwnd
    Else
        panThis.Handle = picSouceList.Hwnd
    End If
    '固定最大高度
    If cldsCalenbarSel.ShowStyle = Show_Plan_Week Then
        panThis.MaxTrackSize.Height = 240
    Else
        panThis.MaxTrackSize.Height = 150
    End If

    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Call picDateList_Resize
    'Set dkpMain.PaintManager.CaptionFont = use.Font

    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
    InitPanel = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub RemovePlan(obj出诊安排 As 出诊安排, ByVal strItem As String)
    '从未保存出诊安排和当前安排中删除某一个安排
    Dim strKey As String

    strKey = GetPlanKey(strItem)
    If obj出诊安排.未保存出诊安排.Exits(strKey) Then obj出诊安排.未保存出诊安排.Remove strKey
    If obj出诊安排.Exits(strKey) Then obj出诊安排.Remove strKey
    Call ChangeCurPlan(obj出诊安排, strItem)
End Sub

Private Sub cboDays_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboFeeItem_Click()
    Dim obj号源 As 出诊号源
    
    Err = 0: On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If cboFeeItem.ListIndex = -1 Then Exit Sub
    If Val(cboFeeItem.Tag) = cboFeeItem.ItemData(cboFeeItem.ListIndex) Then Exit Sub
    
    mblnFeeItemChanged = True
    Set obj号源 = mobj出诊安排.出诊号源.Clone
    obj号源.项目ID = cboFeeItem.ItemData(cboFeeItem.ListIndex)
    obj号源.项目名称 = cboFeeItem.Text
    Call SourceInfor.LoadData(obj号源)
    Set obj号源 = Nothing
    Exit Sub
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboFeeItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim obj出诊安排 As 出诊安排, obj出诊记录集 As 出诊记录集
    Dim cllPro As New Collection, ObjItem As 出诊记录集
    Dim strKey As String, lngCount As Long
    Dim blnHavePlan As Boolean

    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_Save '保存数据
        If IsValied() = False Then Control.Enabled = True: Exit Sub
        Set obj出诊安排 = Get出诊安排

        If Not obj出诊安排.已保存出诊安排 Is Nothing Then
            For Each ObjItem In obj出诊安排.已保存出诊安排
                If ObjItem.是否删除 Then lngCount = lngCount + 1
                blnHavePlan = True
            Next
        End If
        If Not obj出诊安排.未保存出诊安排 Is Nothing Then
            For Each ObjItem In obj出诊安排.未保存出诊安排
                lngCount = lngCount + 1
            Next
        End If

        If lngCount = 0 Then
            If blnHavePlan And (mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel _
                Or mblnTimeChanged Or mblnFeeItemChanged _
                Or chkAotuVerify.Visible And chkAotuVerify.Value = vbChecked) Then
                '自动审核
                If chkAotuVerify.Visible And chkAotuVerify.Value = vbChecked Then
                    If TempPlanVerifyOrCancel(obj出诊安排.安排ID, obj出诊安排.出诊号源.ID, 1) Then
                        mlngSavedRecords = mlngSavedRecords + 1
                        Unload Me: Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Else
                MsgBox "当前没有需要保存的有效安排！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If

        Control.Enabled = False
        If SaveData(obj出诊安排) Then
            mlngSavedRecords = mlngSavedRecords + 1
            mlng安排ID = obj出诊安排.安排ID

            '临时出诊和修改预约挂号控制时退出
            If mbytFun = Fun_TempPlanRecord _
                Or mbytFun = Fun_UpdateUnit _
                Or mbytFun = Fun_TempPlanVerify _
                Or mbytFun = Fun_TempPlanCancel _
                Or mbytFun = Fun_UpdatePlan Then
                Unload Me: Exit Sub
            End If
            
            mblnTimeChanged = False: mblnFeeItemChanged = False
            mblnCheckedByDay = False
            stbThis.Panels(2).Text = "保存成功！"
            
            '自动审核
            If (chkAotuVerify.Visible And chkAotuVerify.Value = vbChecked) Then
                If TempPlanVerifyOrCancel(obj出诊安排.安排ID, obj出诊安排.出诊号源.ID, 1) Then
                    Unload Me: Exit Sub
                End If
            End If

            '重新加载数据
            Set mobj出诊安排 = New 出诊安排
            mstr缺省日期 = mstrCurDay '调整缺省日期
            If InitData(mobj出诊安排, mlng出诊ID, mlng号源Id, mlng安排ID) Then
                Call CheckTempFixedPlan(mobj出诊安排, dtpBegin.Value, dtpEnd.Value)
                Call LoadData
            End If
        End If
        Control.Enabled = True
    Case conMenu_File_Exit '退出
        If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
            lngCount = 0
            Set obj出诊安排 = Get出诊安排
            '检查是否有已修改安排
            If Not obj出诊安排.已保存出诊安排 Is Nothing Then
                For Each ObjItem In obj出诊安排.已保存出诊安排
                    If ObjItem.是否删除 Then lngCount = lngCount + 1
                Next
            End If
            If Not obj出诊安排.未保存出诊安排 Is Nothing Then
                For Each ObjItem In obj出诊安排.未保存出诊安排
                    lngCount = lngCount + 1
                Next
            End If

            If lngCount > 0 Or mblnTimeChanged Or mblnFeeItemChanged Then
                If MsgBox("部分安排可能已被修改，是否不保存直接退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        Unload Me
    End Select
    Exit Sub
ErrHandler:
    Control.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Err = 0: On Error Resume Next
    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    Select Case Control.ID
    Case mMenuID.M_Signal
        Control.Visible = mbytFun = Fun_AddSignalSourcePlan
        Control.Enabled = Control.Visible
    Case mMenuID.M_ValidTime
        Control.Visible = mbytPlanType = F_FixedRule
        Control.Enabled = Control.Visible And (mbytFun = Fun_TempPlan _
                                            Or mbytFun = Fun_AddSignalSourcePlan _
                                            Or mbytFun = Fun_Add _
                                            Or mbytFun = Fun_Update)
        dtpBegin.Enabled = Control.Enabled: dtpEnd.Enabled = Control.Enabled
    Case mMenuID.M_FeeItem
        Control.Visible = mbytPlanType = F_FixedRule _
            And mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "所有科室")
        Control.Enabled = Control.Visible
    Case mMenuID.M_Verify
        Control.Visible = mbytPlanType = F_FixedRule _
            And (mbytFun = Fun_TempPlan Or mbytFun = Fun_AddSignalSourcePlan) _
            And zlStr.IsHavePrivs(mstrPrivs, "审核临时固定安排") And mobj出诊安排.发布时间 <> ""
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Save '保存数据
        Control.Visible = mbytFun <> Fun_View
        Control.Enabled = Control.Visible
    Case conMenu_File_Exit

    End Select
End Sub

Private Sub chkAotuVerify_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkWeek_Click(index As Integer)
    Dim i As Integer
    
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    
    mblnNotClick = True
    If Val(chkWeek(index).Tag) = 1 Then
        chkWeek(index).Value = vbChecked
    ElseIf Not mcllFixedPlan Is Nothing Then
        'Array(出诊日期,限制项目,上班时段,开始时间,终止时间)
        For i = 1 To mcllFixedPlan.Count
            If mcllFixedPlan(i)(1) = chkWeek(index).Caption Then
                chkWeek(index).Value = vbUnchecked
                Exit For
            End If
        Next
    End If
    mblnNotClick = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkWeek_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cldsCalenbarSel_SelectedChangeBefore(ByVal OldDate As String, NewDate As String, Cancel As Boolean)
    Dim strApply As String
    Dim i As Integer

    '临时出诊安排不能切换日期
    If mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan Then Cancel = True: Exit Sub
    If Format(OldDate, "yyyy-mm-dd") = Format(NewDate, "yyyy-mm-dd") Then Cancel = True: Exit Sub

    If IsDate(NewDate) Then
        '只能选择开始时间与终止时间范围内的
        If mobj出诊安排.排班方式 = 0 And cldsCalenbarSel.ShowStyle = Show_Plan_Day Then
            If DateDiff("d", NewDate, mFixedPlanDateRange.dtStart) > 0 Or DateDiff("d", NewDate, mFixedPlanDateRange.dtEnd) < 0 Then
                Cancel = True: Exit Sub
            End If
        Else
            If DateDiff("d", NewDate, mobj出诊安排.开始时间) > 0 Or DateDiff("d", NewDate, mobj出诊安排.终止时间) < 0 Then
                Cancel = True: Exit Sub
            End If
        End If

        '检查上班时间内是否可出诊
        If mobj出诊安排(1).Count > 0 Then
            mblnCheckedByDay = False
            If CheckDepend(0, OldDate) = False Then
                For i = 1 To lvwWorkTime.ListItems.Count
                    If lvwWorkTime.ListItems(i).Checked Then
                        lvwWorkTime.ListItems(i).Checked = False
                        lvwWorkTime_ItemCheck lvwWorkTime.ListItems(i)
                    End If
                Next
                Cancel = True: Exit Sub
            End If
        End If
    End If

    If mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan _
        Or (mbytPlanType = F_Templet And mobj出诊安排.排班规则 = 1) _
        Or mbytPlanType = F_MonthTemplet Then
        If CheckExistRecord(0, Replace(GetApplyToStr(), ",", "|"), mobj出诊安排) Then
            If MsgBox("注意：" & vbCrLf & _
                      "      部分被应用的日期当前已存在出诊安排，应用后这部分安排将会被覆盖！是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        End If
    End If

    '检查当前安排信息
    If CPDPages.IsValied() = False Then
        Cancel = True: Exit Sub
    End If

    '获取当前安排信息
    Call Get当前出诊安排(mobj出诊安排)
    If (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) _
        And mobj出诊安排(1).是否修改 And IsDate(mstrCurDay) _
        And mbytPlanType <> F_MonthTemplet Then
        If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, CDate(mstrCurDay)) Then
            MsgBox Format(mstrCurDay, "yyyy-mm-dd") & " 已在其它出诊表中设置了出诊安排，不能重复安排！", vbInformation, gstrSysName
            mobj出诊安排(1).RemoveAll
            mobj出诊安排(1).是否修改 = False
            Call LoadDetailData
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub cldsCalenbarSel_SelectedChanged(ByVal OldDate As String, NewDate As String)
    mstrCurDay = NewDate
    Call SetButtonEnabled
    Call LoadDetailData
End Sub

Private Function CheckTempFixedPlan(obj出诊安排 As 出诊安排, _
    ByVal dtStartTime As Date, ByVal dtEndTime As Date, _
    Optional ByVal blnTimeRangeChanged As Boolean, _
    Optional ByRef blnUnloadForm As Boolean, Optional ByVal blnSaveBeforeValid As Boolean) As Boolean
    '检查固定安排在指定时间范围内能否添加临时安排
    '入参：
    '   obj出诊安排 出诊安排对象
    '   dtStartTime、dtEndTime 临时安排的有效期
    '   blnTimeRangeChanged 是否时间范围发生了变化
    '说明：
    '   检查规则如下：
    '   1.当前已调整为月/周排班且已按月/周制定了安排，则不允许再新增临时安排
    '   2.检查是否有按月/周排班的，若有，则不允许新增
    '   3.无预约挂号数据，则按正常新增
    '   4.有预约挂号数据，则
    '     在同一周几(如周一)有预约挂号数据的上班时段之间时间范围有交叉，则提示并禁止，如有挂号数据的上班时段为"1日(周一)-上午"、"8日(周一)-白天"；
    '     否则，把有预约挂号数据的上班时段自动加到新增的临时安排中，且不允许修改
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim rsFixedRecord As ADODB.Recordset, rsNewFixedRecord As ADODB.Recordset
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim strPriorDay As String, strPriorItem As String
    Dim strPrior上班时段 As String, dtPriorStart As Date, dtPriorEnd As Date
    Dim strMsgInfo As String, strErrorInfo As String
    Dim i As Long, j As Long, strFixedFilter As String, strKey As String
    Dim blnFindItem As Boolean, blnChangedItem As Boolean
    Dim bln提示 As Boolean, blnTemp As Boolean
    
    On Error GoTo ErrHandler
    Set mcllFixedPlan = New Collection 'Array(出诊日期,限制项目,上班时段,开始时间,终止时间)
    blnUnloadForm = False
    If Not (mbytFun = Fun_TempPlan _
        Or mbytFun = Fun_TempPlanVerify _
        Or mbytFun = Fun_TempPlanCancel) Then CheckTempFixedPlan = True: Exit Function
    
    If dtStartTime < mdtToday Then dtStartTime = mdtToday
    '取消审核时的检查
    If mbytFun = Fun_TempPlanCancel Then
        '1.并发检查，是否被他人取消审核
        '2.一旦安排被使用就不能再取消审核了
        strSQL = "Select 1 From 临床出诊安排 A Where a.Id = [1] And a.审核时间 Is Not Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "并发检查", mlng安排ID)
        If rsTemp.EOF Then
            MsgBox "当前安排已被他人取消审核或删除，不能再取消审核！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊记录 A, 病人挂号记录 B" & vbNewLine & _
                " Where a.Id = b.出诊记录id And a.安排id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前安排是否有预约挂号数据", mlng安排ID)
        If Not rsTemp.EOF Then
            MsgBox "当前安排已存在预约挂号数据，不能取消审核！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C" & vbNewLine & _
                " Where a.号源id = b.号源id And a.出诊id = c.Id And c.排班方式 = 0 And a.Id <> b.Id And b.Id = [1] And a.登记时间 > b.登记时间" & vbNewLine & _
                "       And a.审核时间 Is Not Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查之后是否还有已审核安排", mlng安排ID)
        If Not rsTemp.EOF Then
            MsgBox "该号源在当前安排之后还存在已审核的安排，你不能取消审核当前安排！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        CheckTempFixedPlan = True: Exit Function
    ElseIf mbytFun = Fun_TempPlanVerify Then
        '1.并发检查，是否被他人审核
        strSQL = "Select 1 From 临床出诊安排 A Where a.Id = [1] And a.审核时间 Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "并发检查", mlng安排ID)
        If rsTemp.EOF Then
            MsgBox "当前安排已被他人审核或删除，不能再审核！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        '2.没有有效安排的不能审核
        strSQL = "Select 1 From 临床出诊限制 A Where a.安排Id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有有效安排", mlng安排ID)
        If rsTemp.EOF Then
            MsgBox "当前安排中无任何有效安排，不能审核！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C" & vbNewLine & _
                " Where a.号源id = b.号源id And a.出诊id = c.Id And c.排班方式 = 0 And a.Id <> b.Id And b.Id = [1] And a.登记时间 < b.登记时间" & vbNewLine & _
                "       And a.审核时间 Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查之前是否还有未审核安排", mlng安排ID)
        If Not rsTemp.EOF Then
            MsgBox "该号源在当前安排之前还存在未审核的安排，你不能审核当前安排！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
    End If
    
    strSQL = "Select 1 From 临床出诊号源 A " & vbNewLine & _
            " Where Nvl(a.排班方式,0)<>0 and a.Id = [1] And Rownum < 2" & vbNewLine & _
            "       And Exists(Select 1" & vbNewLine & _
            "                  From 临床出诊记录 M,临床出诊安排 N,临床出诊表 P" & vbNewLine & _
            "                  Where m.安排ID=n.ID And n.出诊ID=p.ID And n.号源ID=a.ID And Nvl(p.排班方式,0)=Nvl(a.排班方式,0) And Rownum<2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有安排", mlng号源Id)
    If Not rsTemp.EOF Then
        MsgBox "当前号源调整为了其它非固定排班方式后已制定了安排，不能再在该出诊表中制定临时安排！", vbInformation + vbOKOnly, gstrSysName
        blnUnloadForm = True: Exit Function
    End If
    
    strSQL = "Select 1 From 临床出诊表 A Where a.Id = [1] And a.发布时间 Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否发布", mlng出诊ID)
    If Not rsTemp.EOF Then
        strSQL = "Select 1 From 临床出诊安排 A, 临床出诊限制 B Where a.Id = b.安排id And a.出诊id = [1] And a.号源ID = [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有安排", mlng出诊ID, mlng号源Id)
        If rsTemp.EOF Then
            MsgBox "该号源在当前未发布的出诊表中还未制定任何安排，不能制定临时安排！", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
    End If
    
    '检查是否有按月/周排班的
    strSQL = "Select b.排班方式, a.开始时间, a.终止时间" & vbNewLine & _
            " From 临床出诊安排 A, 临床出诊表 B" & vbNewLine & _
            " Where a.出诊id = b.Id And b.排班方式 In (1, 2) And a.号源id = [1] And a.开始时间 < [3] And a.终止时间 > [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有按月/周排班的", obj出诊安排.出诊号源.ID, dtStartTime, dtEndTime)
    If Not rsTemp.EOF Then
        dtCurStart = Nvl(rsTemp!开始时间): dtCurEnd = Nvl(rsTemp!终止时间)
        If dtCurStart < dtStartTime Then dtCurStart = dtStartTime
        If dtCurEnd > dtEndTime Then dtCurEnd = dtEndTime
        
        MsgBox "当前号源在时间范围(" & Format(dtCurStart, "yyyy-mm-dd hh:mm:ss") & _
            "-" & Format(dtCurEnd, "yyyy-mm-dd hh:mm:ss") & ")内已按" & _
            IIf(Val(rsTemp!排班方式) = 1, "月", "周") & "进行了排班，不能再在这段时间范围内制定临时安排！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '1.检查是否有预约挂号数据
    Set rsFixedRecord = Get预约挂号记录(obj出诊安排.出诊号源.ID, _
        CDate(Format(dtStartTime, "yyyy-mm-dd")), CDate(Format(dtEndTime, "yyyy-mm-dd")))
    If rsFixedRecord.EOF Then
        Call SetPlanFixed(obj出诊安排, False, "", "", blnChangedItem) '将所有出诊记录都标记为可以修改的
        If blnChangedItem Then Call LoadData
        CheckTempFixedPlan = True: Exit Function
    End If
    
    '1.在同一周几(如周一)有预约挂号数据的上班时段之间时间范围有交叉，则提示并禁止
    Do While Not rsFixedRecord.EOF
        '由于是按限制项目和开始时间(hh:mm:ss)排序的，
        '所以每一个限制项目只需要检查相邻上班时段的时间范围是否有交叉即可
        dtCurStart = Format(mdtToday, "yyyy-mm-dd ") & Format(Nvl(rsFixedRecord!开始时间), "hh:mm:ss")
        dtCurEnd = Format(mdtToday, "yyyy-mm-dd ") & Format(Nvl(rsFixedRecord!终止时间), "hh:mm:ss")
        dtCurEnd = GetWorkTrueDate(dtCurStart, dtCurEnd)
        
        If strPriorItem = Nvl(rsFixedRecord!限制项目) Then
            If Not (DateDiff("n", dtCurStart, dtPriorEnd) <= 0 Or DateDiff("n", dtCurEnd, dtPriorStart) >= 0) Then
                '上班时段时间范围有交叉，组织提示信息
                strErrorInfo = strErrorInfo & vbCrLf & _
                    strPriorDay & "(" & strPriorItem & ")上班时段【" & strPrior上班时段 & "(" & Format(dtPriorStart, "hh:mm") & "-" & Format(dtPriorEnd, "hh:mm") & ")】" & _
                    "与" & Format(Nvl(rsFixedRecord!出诊日期), "yyyy-mm-dd ") & "(" & Nvl(rsFixedRecord!限制项目) & ")上班时段【" & Nvl(rsFixedRecord!上班时段) & "(" & Format(dtCurStart, "hh:mm") & "-" & Format(dtCurEnd, "hh:mm") & ")】"
            End If
        End If
        
        strPriorItem = Nvl(rsFixedRecord!限制项目): strPriorDay = Format(Nvl(rsFixedRecord!出诊日期), "yyyy-mm-dd ")
        strPrior上班时段 = Nvl(rsFixedRecord!上班时段)
        dtPriorStart = dtCurStart: dtPriorEnd = dtCurEnd
        
        '上班时段不能修改固定加入，组织提示信息
        strMsgInfo = strMsgInfo & vbCrLf & _
            strPriorDay & "(" & strPriorItem & ") 上班时段【" & strPrior上班时段 & "(" & Format(dtPriorStart, "hh:mm") & "-" & Format(dtPriorEnd, "hh:mm") & ")】"
        
        'Array(出诊日期,限制项目,上班时段,开始时间,终止时间)
        strKey = "K" & strPriorItem & "_" & strPrior上班时段
        If CollExitsValue(mcllFixedPlan, strKey) = False Then
            mcllFixedPlan.Add Array(strPriorItem, strPriorItem, strPrior上班时段, dtPriorStart, dtPriorEnd), strKey
        End If
        rsFixedRecord.MoveNext
    Loop
    
    If strErrorInfo <> "" Then
        MsgBox "在当前选择的有效时间范围内当前号源在同一个出诊项目(星期)下存在交叉的上班时段，" & _
               "且在这些上班时段时间范围内都存在预约挂号记录，不能再在该有效时间范围内制定临时" & _
               "安排，你可以修改有效时间范围然后继续！" & vbCrLf & _
               strErrorInfo, vbInformation + vbOKOnly, gstrSysName
        If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
        Exit Function
    End If
    
    If blnTimeRangeChanged = False And obj出诊安排.安排ID <> 0 And IsDate(obj出诊安排.登记时间) Then
        '2.检查自新增以后是否有新的不在已有安排中的预约挂号记录
        strSQL = "Select Decode(To_Char(出诊日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) As 限制项目," & vbNewLine & _
                "        ID As 记录id, 出诊日期, 上班时段, 开始时间, 终止时间, 是否独占" & vbNewLine & _
                " From (Select a.Id, a.出诊日期, a.上班时段, a.开始时间, a.终止时间, 是否独占," & vbNewLine & _
                "               Row_Number() Over(Partition By To_Char(a.出诊日期, 'D'), a.上班时段 Order By a.出诊日期) As 行号" & vbNewLine & _
                "        From 临床出诊记录 A, 病人挂号记录 B" & vbNewLine & _
                "        Where a.Id = b.出诊记录id And a.上班时段 Is Not Null And a.号源id = [1] And a.出诊日期 Between [2] And [3]" & vbNewLine & _
                "              And b.登记时间 > [4])" & vbNewLine & _
                " Where 行号 < 2"
        Set rsNewFixedRecord = zlDatabase.OpenSQLRecord(strSQL, "获取有预约挂号数据的出诊记录", obj出诊安排.出诊号源.ID, _
            CDate(Format(dtStartTime, "yyyy-mm-dd")), CDate(Format(dtEndTime, "yyyy-mm-dd")), CDate(obj出诊安排.登记时间))
        
        strMsgInfo = ""
        Do While Not rsNewFixedRecord.EOF
            '上班时段不能修改固定加入，组织提示信息
            strMsgInfo = strMsgInfo & vbCrLf & _
                Format(Nvl(rsNewFixedRecord!出诊日期), "yyyy-mm-dd ") & "(" & Nvl(rsNewFixedRecord!限制项目) & _
                    ") 上班时段【" & Nvl(rsNewFixedRecord!上班时段) & "(" & Format(Nvl(rsNewFixedRecord!开始时间), "hh:mm") & "-" & Format(Nvl(rsNewFixedRecord!终止时间), "hh:mm") & ")】"
            rsNewFixedRecord.MoveNext
        Loop
        
        If strMsgInfo <> "" Then
            MsgBox "在当前选择的有效时间范围内，当前号源自上次新增或修改到现在这段时间范围内有不在该安排中的上班时段产生了新" & _
                   "的预约挂号记录(这些上班时段必须包含在新的安排中)，" & _
                   IIf(mbytFun <> Fun_TempPlanVerify, "需要重新调整安排！", "必须重新调整安排否则不能审核！") & vbCrLf & _
                    strMsgInfo, vbInformation + vbOKOnly, gstrSysName
            If mbytFun = Fun_TempPlanVerify Then blnUnloadForm = True: Exit Function
        End If
    Else
        bln提示 = True
    End If
    
    blnChangedItem = False
    If rsFixedRecord.RecordCount > 0 Then
        Call GetFixedPlan(obj出诊安排, rsFixedRecord, blnTemp)
        blnChangedItem = blnChangedItem Or blnTemp
    End If
    
    'Array(出诊日期,限制项目,上班时段,开始时间,终止时间)
    '清除不在固定不能修改的安排集合内的安排
    Call SetPlanFixed(obj出诊安排, False, "", "", blnTemp, mcllFixedPlan)
    blnChangedItem = blnChangedItem Or blnTemp
    
    If blnChangedItem Then
        If bln提示 Then
            MsgBox "在当前选择的有效时间范围内，当前号源在如下日期存在有预约挂号记录的上班时段，" & _
                    "这些上班时段都将不能修改，且必须加入到新的临时安排中！" & vbCrLf & _
                    strMsgInfo, vbInformation + vbOKOnly, gstrSysName
        End If
        Call LoadData
        If blnSaveBeforeValid And strMsgInfo <> "" Then Exit Function
    End If
    CheckTempFixedPlan = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetPlanFixed(obj出诊安排 As 出诊安排, ByVal blnFixed As Boolean, _
    Optional ByVal str限制项目 As String, Optional ByVal str上班时段 As String, _
    Optional ByRef blnChangedItem As Boolean, Optional ByVal cllFixedPlan As Collection) As Boolean
    '设置安排的固定标记，主要用于固定安排临时安排
    '入参：
    '   blnFixed 是否固定
    '   str限制项目 如果为空，则是所有的出诊日期
    '   str上班时段 如果为空，则是所有的上班时段
    '返回：
    '   str限制项目不为空且str上班时段不为空时返回是否找到
    '   blnChangedItem 是否有改变的项目
    '说明：
    '   cllFixedPlan元素大于零时，用于清除不在集合中的固定标识
    Dim blnFindItem As Boolean, blnTemp As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    blnChangedItem = False
    blnFindItem = SetPlanFixedSub(obj出诊安排, blnFixed, str限制项目, str上班时段, blnTemp, cllFixedPlan)
    blnChangedItem = blnTemp
    blnFindItem = blnFindItem Or SetPlanFixedSub(obj出诊安排.已保存出诊安排, blnFixed, str限制项目, str上班时段, blnTemp, cllFixedPlan)
    blnChangedItem = blnChangedItem Or blnTemp
    blnFindItem = blnFindItem Or SetPlanFixedSub(obj出诊安排.未保存出诊安排, blnFixed, str限制项目, str上班时段, blnTemp, cllFixedPlan)
    blnChangedItem = blnChangedItem Or blnTemp
    
    SetPlanFixed = blnFindItem
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetPlanFixedSub(obj出诊安排 As 出诊安排, ByVal blnFixed As Boolean, _
    Optional ByVal str限制项目 As String, Optional ByVal str上班时段 As String, _
    Optional ByRef blnChangedItem As Boolean, Optional ByVal cllFixedPlan As Collection) As Boolean
    '设置安排的固定标记，主要用于固定安排临时安排
    '入参：
    '   blnFixed 是否固定
    '   str限制项目 如果为空，则是所有的出诊日期
    '   str上班时段 如果为空，则是所有的上班时段
    '返回：
    '   str限制项目不为空且str上班时段不为空时返回是否找到
    '   blnChangedItem 是否有改变的项目
    '说明：
    '   cllFixedPlan元素大于零时，用于清除不在集合中的固定标识
    Dim blnFind As Boolean, blnFindItem As Boolean
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    
    blnChangedItem = False
    If obj出诊安排 Is Nothing Then Exit Function
    If cllFixedPlan Is Nothing Then Set cllFixedPlan = New Collection
    If cllFixedPlan.Count = 0 Then
        For Each obj出诊记录集 In obj出诊安排
            If blnFind And str限制项目 <> "" Then Exit For
            If obj出诊记录集.出诊日期 = str限制项目 Or str限制项目 = "" Then
                For Each obj出诊记录 In obj出诊记录集
                    If obj出诊记录.时间段 = str上班时段 Or str上班时段 = "" Then
                        If obj出诊记录.是否固定 <> blnFixed Then
                            obj出诊记录.是否固定 = blnFixed
                            blnChangedItem = True
                        End If
                        blnFind = True: blnFindItem = True
                        If str上班时段 <> "" Then Exit For
                    End If
                Next
            End If
        Next
    Else
        For Each obj出诊记录集 In obj出诊安排
            For Each obj出诊记录 In obj出诊记录集
                If CollExitsValue(cllFixedPlan, "K" & obj出诊记录集.出诊日期 & "_" & obj出诊记录.时间段) = False Then
                    If obj出诊记录.是否固定 <> blnFixed Then
                        obj出诊记录.是否固定 = blnFixed
                        blnChangedItem = True
                    End If
                End If
            Next
        Next
    End If
    SetPlanFixedSub = blnFindItem
End Function

Private Sub GetFixedPlan(obj出诊安排 As 出诊安排, rsRecord As ADODB.Recordset, _
    Optional ByRef blnNewAdd As Boolean)
    '将固定不能修改的安排调整为记录集
    '出参:
    '   blnNewAdd - 是否有新增/修改的
    '说明：
    '   在数据加载时将数据全部定为已修改的原因是，
    '   制定临时安排时无法判断是否有新的不能修改的上班时段(有新的预约挂号记录的上班时段)
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim obj合作单位 As 合作单位控制, obj合作单位Tmp As 合作单位控制
    Dim strItem As String, rsUnitReg As ADODB.Recordset
    Dim strKey As String
    
    '获取不变的出诊安排对象
    Err = 0: On Error GoTo ErrHandler
    blnNewAdd = False
    If obj出诊安排.已保存出诊安排 Is Nothing Then Set obj出诊安排.已保存出诊安排 = New 出诊安排
    If obj出诊安排.未保存出诊安排 Is Nothing Then Set obj出诊安排.未保存出诊安排 = New 出诊安排
    If rsRecord.RecordCount = 0 Then Exit Sub
    
    rsRecord.MoveFirst
    Set obj出诊记录集 = New 出诊记录集
    obj出诊记录集.出诊日期 = Nvl(rsRecord!限制项目)
    
    strKey = "K" & obj出诊记录集.出诊日期
    If obj出诊安排.Exits(strKey) Then
        Set obj出诊记录集 = obj出诊安排(strKey).Clone
    ElseIf obj出诊安排.未保存出诊安排.Exits(strKey) Then
        Set obj出诊记录集 = obj出诊安排.未保存出诊安排(strKey).Clone
    ElseIf obj出诊安排.已保存出诊安排.Exits(strKey) Then
        Set obj出诊记录集 = obj出诊安排.已保存出诊安排(strKey).Clone
    End If
    Do While Not rsRecord.EOF
        '转换成对象
        If strItem <> "" And strItem <> Nvl(rsRecord!限制项目) Then
            If obj出诊记录集.Count > 0 Then
                '加入到安排中
                strKey = "K" & obj出诊记录集.出诊日期
                If obj出诊安排.已保存出诊安排.Exits(strKey) Then obj出诊安排.已保存出诊安排(strKey).是否删除 = True '标记为删除
                If obj出诊安排.未保存出诊安排.Exits(strKey) Then obj出诊安排.未保存出诊安排.Remove strKey '存在先移除
                obj出诊安排.未保存出诊安排.AddItem obj出诊记录集, strKey
                
                If obj出诊安排.Exits(strKey) Then
                    obj出诊安排.Remove strKey '存在先移除
                    obj出诊安排.AddItem obj出诊记录集.Clone, strKey
                End If
            End If
            Set obj出诊记录集 = New 出诊记录集
            obj出诊记录集.出诊日期 = Nvl(rsRecord!限制项目)
            
            strKey = "K" & obj出诊记录集.出诊日期
            If obj出诊安排.Exits(strKey) Then
                Set obj出诊记录集 = obj出诊安排(strKey).Clone
            ElseIf obj出诊安排.未保存出诊安排.Exits(strKey) Then
                Set obj出诊记录集 = obj出诊安排.未保存出诊安排(strKey).Clone
            ElseIf obj出诊安排.已保存出诊安排.Exits(strKey) Then
                Set obj出诊记录集 = obj出诊安排.已保存出诊安排(strKey).Clone
            End If
        End If
        
        '1.出诊记录
        Set obj出诊记录 = GetVisitTimesObject(GetVisitTime(Val(rsRecord!记录ID), True))
        '上班时段
        If obj出诊安排.所有上班时段.Exits("K" & obj出诊记录.时间段) Then
            Set obj出诊记录.上班时段 = obj出诊安排.所有上班时段("K" & obj出诊记录.时间段).Clone
        Else
            '出诊记录时，上班时段可能已被删除
            Set obj出诊记录.上班时段 = New 上班时段
            With obj出诊记录.上班时段
                .开始时间 = obj出诊记录.开始时间
                .结束时间 = obj出诊记录.终止时间
            End With
        End If

        '2.分诊诊室
        Set obj出诊记录.安排门诊诊室集 = GetVisitRoomsObjects(GetVisitRooms(Val(rsRecord!记录ID), True))
        obj出诊记录.安排门诊诊室集.分诊方式 = obj出诊记录.分诊方式
        obj出诊记录.安排门诊诊室集.医生姓名 = obj出诊安排.出诊号源.医生姓名

        '3.号序信息
        Set obj出诊记录.号序信息集 = GetTimeIntervalObjects(GetTimeInterval(Val(rsRecord!记录ID), True))
        With obj出诊记录.号序信息集
            .出诊频次 = obj出诊安排.出诊号源.出诊频次
            .是否分时段 = obj出诊记录.是否分时段
            .是否序号控制 = obj出诊记录.是否序号控制
            .限号数 = obj出诊记录.限号数
            .限约数 = obj出诊记录.限约数
            .预约控制 = obj出诊记录.预约控制
        End With

        '4.合作单位预约控制
        Set obj出诊记录.合作单位控制集 = New 合作单位控制集
        obj出诊记录.合作单位控制集.是否独占 = Val(Nvl(rsRecord!是否独占)) = 1
        For Each obj合作单位 In obj出诊安排.所有合作单位
            Set rsUnitReg = GetUnitReg(Val(rsRecord!记录ID), obj合作单位.合作单位名称, obj合作单位.类型, True)
            If Not rsUnitReg.EOF Then
                Set obj合作单位Tmp = New 合作单位控制
                obj合作单位Tmp.合作单位名称 = obj合作单位.合作单位名称
                obj合作单位Tmp.类型 = obj合作单位.类型
                obj合作单位Tmp.预约控制方式 = Val(rsUnitReg!控制方式)
                Set obj合作单位Tmp.号序信息集 = GetTimeIntervalObjects(rsUnitReg)
                
                obj出诊记录.合作单位控制集.AddItem obj合作单位Tmp, "K" & obj合作单位Tmp.合作单位名称
            End If
        Next
        obj出诊记录.是否固定 = True '在进行临时出诊时，不能删除
        
        strKey = "K" & obj出诊记录.时间段
        If obj出诊记录集.Exits(strKey) Then
            If obj出诊记录集(strKey).是否固定 = False Then
                blnNewAdd = True
                obj出诊记录.是否修改 = True
            End If
            obj出诊记录集.Remove strKey
        Else
            blnNewAdd = True
            obj出诊记录.是否修改 = True
        End If
        obj出诊记录集.AddItem obj出诊记录, "K" & obj出诊记录.时间段
        strItem = Nvl(rsRecord!限制项目)
        rsRecord.MoveNext
    Loop
    
    '遍历完后，检查最后一个是否还有有效安排
    If obj出诊记录集.Count > 0 Then
        '加入到安排中
        strKey = "K" & obj出诊记录集.出诊日期
        If obj出诊安排.已保存出诊安排.Exits(strKey) Then obj出诊安排.已保存出诊安排(strKey).是否删除 = True '标记为删除
        If obj出诊安排.未保存出诊安排.Exits(strKey) Then obj出诊安排.未保存出诊安排.Remove strKey '存在先移除
        obj出诊安排.未保存出诊安排.AddItem obj出诊记录集, strKey
        
        If obj出诊安排.Exits(strKey) Then
            obj出诊安排.Remove strKey '存在先移除
            obj出诊安排.AddItem obj出诊记录集.Clone, strKey
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CPDPages_DataIsChanged(index As Integer)
    '显示当前安排
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim obj出诊安排 As 出诊安排
    
    Err = 0: On Error GoTo ErrHandler
    stbThis.Panels(2).Text = ""
    If mbytPlanType <> F_FixedRule Then Exit Sub

    Set obj出诊记录集 = CPDPages.Get出诊记录集()
    If obj出诊记录集 Is Nothing Then Exit Sub
    If obj出诊记录集.Count = 0 Then Exit Sub
    
    Set obj出诊安排 = New 出诊安排
    obj出诊安排.AddItem obj出诊记录集, GetPlanKey(obj出诊记录集.出诊日期)
    
    LoadPlanToGrid obj出诊安排, 0
    Set obj出诊安排 = Nothing
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpBegin_LostFocus()
    If mblnFirst Then Exit Sub
    If IsDate(dtpBegin.Tag) Then
        If DateDiff("s", dtpBegin.Tag, dtpBegin.Value) = 0 Then Exit Sub
    End If
    mblnTimeChanged = True: stbThis.Panels(2).Text = ""
    mblnValiedCanSave = CheckTempFixedPlan(mobj出诊安排, dtpBegin.Value, dtpEnd.Value, True)
    dtpBegin.Tag = dtpBegin.Value
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_LostFocus()
    If mblnFirst Then Exit Sub
    If IsDate(dtpEnd.Tag) Then
        If DateDiff("s", dtpEnd.Tag, dtpEnd.Value) = 0 Then Exit Sub
    End If
    mblnTimeChanged = True: stbThis.Panels(2).Text = ""
    mblnValiedCanSave = CheckTempFixedPlan(mobj出诊安排, dtpBegin.Value, dtpEnd.Value, True)
    dtpEnd.Tag = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    Dim blnUnloadForm As Boolean
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    On Error GoTo ErrHandler
    If mbytFun = Fun_TempPlan _
        Or mbytFun = Fun_Update And mobj出诊安排.是否临时安排 _
        Or mbytFun = Fun_TempPlanVerify _
        Or mbytFun = Fun_TempPlanCancel Then
        
        If CheckTempFixedPlan(mobj出诊安排, dtpBegin.Value, dtpEnd.Value, False, blnUnloadForm) = False Then
            If blnUnloadForm Then Unload Me: Exit Sub
        End If
    End If
    
    If (mbytFun = Fun_Add Or mbytFun = Fun_Update) _
        And (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) And IsDate(mstr缺省日期) Then
        If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, CDate(mstr缺省日期)) Then
            MsgBox "注意，" & Format(mstr缺省日期, "yyyy-mm-dd") & " 已在其它出诊表中设置了出诊安排！", vbInformation, gstrSysName
        End If
    End If

    If CheckDepend(0, mstr缺省日期) = False Then Unload Me: Exit Sub

    If txtSignal.Visible And txtSignal.Enabled Then txtSignal.SetFocus

    lvwWorkTime.View = lvwList
    lvwWorkTime.View = lvwReport
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    cldsCalenbarSel.KeyShift = Shift        '是否按下Ctrl键
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    cldsCalenbarSel.KeyShift = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    mblnFirst = True
    
    mblnTimeChanged = False: mblnFeeItemChanged = False
    mblnCheckedByDay = False
    If mlng出诊ID = 0 Then
        MsgBox "出诊表信息未找到，请刷新后重试！", vbInformation + vbOKOnly, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Select Case mbytFun
    Case Fun_View   '查看
        Me.Caption = "查看安排"
    Case Fun_Add   '新增
        Me.Caption = "新增安排"
    Case Fun_Update, Fun_UpdatePlan '编辑,调整已发布后的安排
        Me.Caption = "调整安排"
    Case Fun_Delete   '删除
        Me.Caption = "删除安排"
    Case Fun_TempPlan   '临时安排(固定出诊表)
        Me.Caption = "临时安排"
    Case Fun_UpdateUnit   '调整合作单位预约挂号
        Me.Caption = "调整预约挂号控制"
    Case Fun_AddSignalSourcePlan   '新增号源
        Me.Caption = "新增号源安排"
    Case Fun_TempPlanRecord  '发布后临时出诊
        Me.Caption = "临时出诊"
    Case Fun_TempPlanVerify   '临时安排(固定出诊表)审核
        Me.Caption = "审核临时安排"
    Case Fun_TempPlanCancel   '临时安排(固定出诊表)取消审核
        Me.Caption = "取消审核临时安排"
    End Select
    
    Select Case mbytPlanType
    Case F_Templet
        cldsCalenbarSel.ShowStyle = Show_Plan_Rule
    Case F_FixedRule
        cldsCalenbarSel.ShowStyle = Show_Plan_Week
    Case F_MonthPlan, F_WeekPlan, F_MonthTemplet
        cldsCalenbarSel.ShowStyle = Show_Plan_Day
    End Select
    
    If zlDefCommandBars() = False Then Unload Me: Exit Sub
    If InitPanel() = False Then Unload Me: Exit Sub
    If InitData(mobj出诊安排, mlng出诊ID, mlng号源Id, mlng安排ID) = False Then Unload Me: Exit Sub

    Screen.MousePointer = vbHourglass
    If LoadData() = False Then Screen.MousePointer = vbDefault: Unload Me: Exit Sub
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetLastValidPlanFeeItem(ByVal lng号源Id As Long) As ADODB.Recordset
    '获取最后一次有效安排的收费项目
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select b.id, b.名称" & vbNewLine & _
            " From (Select 项目id From 临床出诊安排" & vbNewLine & _
            "       Where 号源id = [1] And 审核时间 Is Not Null" & vbNewLine & _
            "       Order By 登记时间 Desc) A, 收费项目目录 B" & vbNewLine & _
            " Where a.项目id = b.Id And Rownum < 2"
    Set GetLastValidPlanFeeItem = zlDatabase.OpenSQLRecord(strSQL, "获取最后一次有效安排的收费项目", lng号源Id)
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData(ByRef obj出诊安排 As 出诊安排, ByVal lng出诊ID As Long, _
    ByVal lng号源Id As Long, Optional ByVal lng安排ID As Long) As Boolean
    '功能：构造出诊安排对象
    Dim rsSignalSource As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim obj所有上班时段 As 上班时段集, obj所有分诊诊室 As 分诊诊室集, obj所有合作单位 As 合作单位控制集
    Dim obj合作单位控制 As 合作单位控制, obj合作单位 As 合作单位控制
    Dim blnRecord As Boolean, strTemp As String
    Dim obj号序信息集 As 号序信息集

    '门诊诊室、号序信息、合作单位号序信息均根据安排ID一次性从数据库中读取出来
    Dim rs出诊项目 As ADODB.Recordset, rs出诊记录 As ADODB.Recordset
    Dim rs门诊诊室 As ADODB.Recordset, rs号序信息 As ADODB.Recordset, rs合作单位号序信息 As ADODB.Recordset
    Dim obj出诊记录 As 出诊记录, obj出诊记录集 As 出诊记录集
    Dim strSQL As String, rsLastValidPlanFeeItem As ADODB.Recordset
    Dim dtNow As Date
    
    Err = 0: On Error GoTo ErrHandler
    '加载收费项目
    If mbytPlanType = F_FixedRule _
            And mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "所有科室") Then
        strSQL = "Select ID,名称 From 收费项目目录 " & _
                " Where 类别='1' And (撤档时间 >=To_date('3000-01-01','yyyy-mm-dd') Or 撤档时间 Is Null)" & _
                " And (站点='" & gstrNodeNo & "' Or 站点 is Null) " & _
                " Order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        cboFeeItem.Clear
        Do While Not rsTemp.EOF
            cboFeeItem.AddItem rsTemp!名称
            cboFeeItem.ItemData(cboFeeItem.NewIndex) = Val(Nvl(rsTemp!ID))
            rsTemp.MoveNext
        Loop
    End If
    
    '出诊安排信息
    Set obj出诊安排 = GetVisitPlanObjects(GetVisitPlan(lng安排ID, lng出诊ID))
    If obj出诊安排.出诊ID = 0 Then
        MsgBox "出诊表信息未找到，请刷新后重试！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '号源信息
    Set rsSignalSource = GetSignalSource("", IIf(mbytFun = Fun_AddSignalSourcePlan And lng号源Id = 0, -1, lng号源Id))
    If rsSignalSource.RecordCount = 0 Then
        If mbytFun <> Fun_AddSignalSourcePlan Then
            MsgBox "临床出诊号源信息未找到，请刷新后重试！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    dtNow = zlDatabase.Currentdate
    mdtToday = CDate(Format(dtNow, "yyyy-mm-dd"))
    blnRecord = Not (mbytPlanType = F_Templet Or mbytPlanType = F_FixedRule Or mbytPlanType = F_MonthTemplet)
    If mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel Then
        mstr缺省日期 = "周一" '审核和取消审核缺省为周一
    End If
    
    If lng安排ID = 0 Then
        '设置缺省时间范围
        Select Case mbytPlanType
        Case F_FixedRule
            obj出诊安排.开始时间 = Format(DateAdd("d", IIf(mbytFun = Fun_TempPlan, Get预约天数(lng出诊ID, lng号源Id) + 1, 1), mdtToday), "yyyy-MM-dd hh:mm:ss")
            If mbytFun = Fun_TempPlan Then
                '缺省一个星期
                obj出诊安排.终止时间 = Format(DateAdd("d", 7, obj出诊安排.开始时间), "yyyy-MM-dd 23:59:59")
            Else
                obj出诊安排.终止时间 = "3000-01-01"
            End If
        Case F_Templet
            
        Case F_MonthTemplet
            obj出诊安排.开始时间 = "1900-01-01"
            obj出诊安排.终止时间 = "1900-01-31"
        Case Else
            Dim varDateRange As Variant
            varDateRange = GetDateRange(obj出诊安排.年份, obj出诊安排.月份, IIf(mbytPlanType = F_WeekPlan, obj出诊安排.周数, 0))
            obj出诊安排.开始时间 = Format(varDateRange(0), "yyyy-mm-dd hh:mm:ss")
            obj出诊安排.终止时间 = Format(varDateRange(1), "yyyy-mm-dd hh:mm:ss")
        End Select
    Else
        '修改固定安排出诊记录的开始日期为当前日期，防止性能问题
        If obj出诊安排.排班方式 = 0 And blnRecord Then
            With mFixedPlanDateRange
                .dtStart = IIf(DateDiff("d", obj出诊安排.开始时间, mdtToday) > 0, mdtToday, obj出诊安排.开始时间)
                .dtEnd = mobj出诊安排.终止时间
                If .dtEnd > CDate(Format(mdtToday + Get预约天数(lng出诊ID, mlng号源Id), "yyyy-mm-dd 23:59:59")) Then
                    .dtEnd = CDate(Format(mdtToday + Get预约天数(lng出诊ID, mlng号源Id), "yyyy-mm-dd 23:59:59"))
                End If
            End With
            
            obj出诊安排.开始时间 = Format(mdtToday, "yyyy-MM-dd")
            obj出诊安排.终止时间 = Format(mdtToday + Get预约天数(lng出诊ID, mlng号源Id), "yyyy-mm-dd 23:59:59")
        End If
        If (mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan) And IsDate(mstr缺省日期) Then
            obj出诊安排.开始时间 = mstr缺省日期
            obj出诊安排.终止时间 = Format(mstr缺省日期, "yyyy-mm-dd 23:59:59")
        End If
        If mbytPlanType = F_MonthTemplet Then
            obj出诊安排.开始时间 = "1900-01-01"
            obj出诊安排.终止时间 = "1900-01-31"
        End If
    End If
    
    '号源信息,科室、医生、项目取安排中的
    Set obj出诊安排.出诊号源 = GetSignalSourceObject(rsSignalSource)
    If obj出诊安排.安排ID <> 0 Then
        With obj出诊安排.出诊号源
            .项目ID = obj出诊安排.项目ID
            .项目名称 = obj出诊安排.项目名称
            .医生ID = obj出诊安排.医生ID
            .医生姓名 = obj出诊安排.医生姓名
            .医生职称 = obj出诊安排.医生职称
        End With
    Else
        If mbytPlanType = F_FixedRule And mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "所有科室") Then
            '取最后一次有效安排的收费项目
            Set rsLastValidPlanFeeItem = GetLastValidPlanFeeItem(mlng号源Id)
            If rsLastValidPlanFeeItem.RecordCount > 0 Then
                obj出诊安排.出诊号源.项目ID = Val(Nvl(rsLastValidPlanFeeItem!ID))
                obj出诊安排.出诊号源.项目名称 = Nvl(rsLastValidPlanFeeItem!名称)
            End If
        End If
    End If
    
    mblnNotClick = True
    zlControl.CboLocate cboFeeItem, obj出诊安排.出诊号源.项目ID, True
    If cboFeeItem.ListIndex = -1 Then
        cboFeeItem.AddItem obj出诊安排.出诊号源.项目名称
        cboFeeItem.ItemData(cboFeeItem.NewIndex) = obj出诊安排.出诊号源.项目ID
        cboFeeItem.ListIndex = cboFeeItem.NewIndex
    End If
    mblnNotClick = False

    '基础数据
    Set obj所有分诊诊室 = GetVisitRoomsObjects(GetDoctorRooms(obj出诊安排.出诊号源.科室ID))
    Set obj所有上班时段 = GetWorkTimesObjects(GetWorkTimes(obj出诊安排.出诊号源.站点, obj出诊安排.出诊号源.号类))
    Set obj所有合作单位 = GetUnitsObjects(GetUnitAll())

    If Not (mbytFun = Fun_View Or mbytFun = Fun_UpdateUnit _
        Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel) Then
        If mbytFun <> Fun_UpdatePlan Then
            '调整安排时不能新增上班时段
            Set obj出诊安排.所有上班时段 = obj所有上班时段
        End If
        Set obj出诊安排.所有分诊诊室 = obj所有分诊诊室
    End If
    Set obj出诊安排.所有合作单位 = obj所有合作单位

    Set obj出诊安排.号源安排 = GetClinicRecordFromSignalSource(obj出诊安排.出诊号源.ID)
    If blnRecord Then
        '获取停诊安排
        Set obj出诊安排.停诊记录 = GetStopVisitObjects(GetStopVisit(obj出诊安排.出诊号源.ID, obj出诊安排.开始时间, obj出诊安排.终止时间))
        If mbytFun = Fun_UpdatePlan Then
            Set mrsVisitedRecordByDate = GetVisitedRecordByDate(mlng安排ID, mstr缺省日期)
        Else
            Set mrsVisitedRecord = GetVisitedRecord(obj出诊安排.出诊号源.ID, obj出诊安排.开始时间, obj出诊安排.终止时间)
        End If
    End If

    obj出诊安排.临时出诊 = (mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan)
    obj出诊安排.更新合作单位 = mbytFun = Fun_UpdateUnit
    '将出诊安排信息复制到未保存和已保存集合中
    Set obj出诊安排.未保存出诊安排 = obj出诊安排.Clone
    Set obj出诊安排.已保存出诊安排 = obj出诊安排.Clone
    
    '从数据库读取数据
    If blnRecord Then
        If mbytPlanType = F_FixedRule Then
            '出诊日期
            strSQL = "Select To_Char(b.出诊日期,'yyyy-mm-dd') As 出诊日期" & vbNewLine & _
                    " From 临床出诊安排 A, 临床出诊记录 B" & vbNewLine & _
                    " Where a.Id = b.安排id And a.出诊Id = [1] And a.号源ID = [2]" & vbNewLine & _
                    "       And b.上班时段 Is Not Null And b.出诊日期 Between [3] And [4]" & vbNewLine & _
                    " Group By To_Char(b.出诊日期,'yyyy-mm-dd')"
            Set rs出诊项目 = zlDatabase.OpenSQLRecord(strSQL, "获取出诊项目", mlng出诊ID, mlng号源Id, _
                CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '出诊时间段信息
            strSQL = "Select b.ID As 记录ID,To_Char(b.出诊日期,'yyyy-mm-dd') As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, b.开始时间, b.终止时间," & vbNewLine & _
                    "        b.限号数, b.已挂数, b.限约数, b.已约数, b.分诊方式, b.预约控制, b.是否临时出诊," & vbNewLine & _
                    "        b.替诊医生姓名, b.科室ID, b.项目ID, c.名称 As 项目名称, b.医生ID, b.医生姓名, b.是否独占," & vbNewLine & _
                    "        b.停诊开始时间, b.停诊终止时间, b.停诊原因" & vbNewLine & _
                    " From 临床出诊安排 A, 临床出诊记录 B, 收费项目目录 C" & vbNewLine & _
                    " Where a.Id = b.安排id And b.项目ID = c.Id And a.出诊Id = [1] And a.号源ID = [2]" & vbNewLine & _
                    "       And b.上班时段 Is Not Null And b.出诊日期 Between [3] And [4]"
            Set rs出诊记录 = zlDatabase.OpenSQLRecord(strSQL, "获取出诊时段", mlng出诊ID, mlng号源Id, _
                CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '门诊诊室
            strSQL = "Select c.ID As 记录ID,a.诊室ID, b.名称" & vbNewLine & _
                    " From 临床出诊安排 D,临床出诊诊室记录 A, 临床出诊记录 C, 门诊诊室 B" & vbNewLine & _
                    " Where a.记录ID = c.ID And a.诊室id = b.Id And d.出诊Id = [1] And d.号源ID = [2]" & vbNewLine & _
                    "       And c.出诊日期 Between [3] And [4]"
            Set rs门诊诊室 = zlDatabase.OpenSQLRecord(strSQL, "获取临床出诊诊室", mlng出诊ID, mlng号源Id, _
                CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '号序信息
            '"       And (a.开始时间 <> a. 终止时间 Or a.开始时间 Is Null And a. 终止时间 Is Null)" & vbNewLine & _'开始时间与终止时间相等的是加号的序号
            strSQL = "Select b.ID As 记录ID,a.序号, a.开始时间, a. 终止时间, a.数量, a.是否预约, a.是否停诊" & vbNewLine & _
                    " From 临床出诊安排 D,临床出诊序号控制 A,临床出诊记录 B" & vbNewLine & _
                    " Where a.记录ID = b.ID And d.出诊Id = [1] And d.号源ID = [2]" & vbNewLine & _
                    "       And b.出诊日期 Between [3] And [4] " & vbNewLine & _
                    "       And (a.开始时间 <> a. 终止时间 Or a.开始时间 Is Null And a. 终止时间 Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(b.是否分时段,0)=1 And Nvl(b.是否序号控制,0)=0)" & vbNewLine & _
                    "               Or Nvl(b.是否分时段,0)=1 And Nvl(b.是否序号控制,0)=0 And a.预约顺序号 IS NULL)"
            Set rs号序信息 = zlDatabase.OpenSQLRecord(strSQL, "获取号序信息", mlng出诊ID, mlng号源Id, _
                CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '预约挂号控制
            strSQL = "Select c.ID As 记录ID,a.名称,a.类型,a.性质,a.控制方式, a.序号, b.开始时间, b.终止时间, a.数量, b.是否预约, b.是否停诊" & vbNewLine & _
                    " From 临床出诊安排 D,临床出诊挂号控制记录 A, 临床出诊序号控制 B,临床出诊记录 C" & vbNewLine & _
                    " Where a.记录id = b.记录id(+) And a.序号 = b.序号(+)  And a.记录ID=c.ID" & vbNewLine & _
                    "       And d.出诊Id = [1] And d.号源ID = [2] And c.出诊日期 Between [3] And [4]" & vbNewLine & _
                    "       And (b.开始时间 <> b. 终止时间 Or b.开始时间 Is Null And b. 终止时间 Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(c.是否分时段,0)=1 And Nvl(c.是否序号控制,0)=0) " & vbNewLine & _
                    "               Or Nvl(c.是否分时段,0)=1 And Nvl(c.是否序号控制,0)=0 And b.预约顺序号 IS NULL)"
            Set rs合作单位号序信息 = zlDatabase.OpenSQLRecord(strSQL, "获取号序信息", mlng出诊ID, mlng号源Id, _
                CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))
        ElseIf mlng安排ID <> 0 Then
            '出诊日期
            strSQL = "Select To_Char(b.出诊日期,'yyyy-mm-dd') As 出诊日期" & _
                    " From 临床出诊记录 B" & vbNewLine & _
                    " Where b.安排id = [1] And b.上班时段 Is Not Null And b.出诊日期 Between [2] And [3]" & vbNewLine & _
                    " Group By To_Char(b.出诊日期,'yyyy-mm-dd')"
            Set rs出诊项目 = zlDatabase.OpenSQLRecord(strSQL, "获取出诊项目", lng安排ID, _
                CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '出诊时间段信息
            strSQL = "Select b.ID As 记录ID,To_Char(b.出诊日期,'yyyy-mm-dd') As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, b.开始时间, b.终止时间," & vbNewLine & _
                    "        b.限号数, b.已挂数, b.限约数, b.已约数, b.分诊方式, b.预约控制, b.是否临时出诊," & vbNewLine & _
                    "        b.替诊医生姓名, b.科室ID, b.项目ID, c.名称 As 项目名称, b.医生ID, b.医生姓名, b.是否独占," & vbNewLine & _
                    "        b.停诊开始时间, b.停诊终止时间, b.停诊原因" & vbNewLine & _
                    " From 临床出诊记录 B, 收费项目目录 C" & vbNewLine & _
                    " Where b.项目ID = c.Id And b.安排Id = [1] And b.上班时段 Is Not Null And b.出诊日期 Between [2] And [3]"
            Set rs出诊记录 = zlDatabase.OpenSQLRecord(strSQL, "获取出诊时段", lng安排ID, CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '门诊诊室
            strSQL = "Select c.ID As 记录ID,a.诊室ID, b.名称" & vbNewLine & _
                    " From 临床出诊诊室记录 A, 临床出诊记录 C, 门诊诊室 B" & vbNewLine & _
                    " Where a.记录ID = c.ID And a.诊室id = b.Id And c.安排ID = [1] And c.出诊日期 Between [2] And [3]"
            Set rs门诊诊室 = zlDatabase.OpenSQLRecord(strSQL, "获取临床出诊诊室", lng安排ID, CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '号序信息
            '"       And (a.开始时间 <> a. 终止时间 Or a.开始时间 Is Null And a. 终止时间 Is Null)" & vbNewLine & _'开始时间与终止时间相等的是加号的序号
            strSQL = "Select b.ID As 记录ID,a.序号, a.开始时间, a. 终止时间, a.数量, a.是否预约, a.是否停诊" & vbNewLine & _
                    " From 临床出诊序号控制 A,临床出诊记录 B" & vbNewLine & _
                    " Where a.记录ID = b.ID And b.安排ID=[1] And b.出诊日期 Between [2] And [3] " & vbNewLine & _
                    "       And (a.开始时间 <> a. 终止时间 Or a.开始时间 Is Null And a. 终止时间 Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(b.是否分时段,0)=1 And Nvl(b.是否序号控制,0)=0) Or Nvl(b.是否分时段,0)=1 And Nvl(b.是否序号控制,0)=0 And a.预约顺序号 IS NULL)"
            Set rs号序信息 = zlDatabase.OpenSQLRecord(strSQL, "获取号序信息", lng安排ID, CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))

            '预约挂号控制
            strSQL = "Select c.ID As 记录ID,a.名称,a.类型,a.性质,a.控制方式, a.序号, b.开始时间, b.终止时间, a.数量, b.是否预约, b.是否停诊" & vbNewLine & _
                    " From 临床出诊挂号控制记录 A, 临床出诊序号控制 B,临床出诊记录 C" & vbNewLine & _
                    " Where a.记录id = b.记录id(+) And a.序号 = b.序号(+)  And a.记录ID=c.ID" & vbNewLine & _
                    "       And c.安排id = [1] And c.出诊日期 Between [2] And [3]" & vbNewLine & _
                    "       And (b.开始时间 <> b. 终止时间 Or b.开始时间 Is Null And b. 终止时间 Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(c.是否分时段,0)=1 And Nvl(c.是否序号控制,0)=0) Or Nvl(c.是否分时段,0)=1 And Nvl(c.是否序号控制,0)=0 And b.预约顺序号 IS NULL)"
            Set rs合作单位号序信息 = zlDatabase.OpenSQLRecord(strSQL, "获取号序信息", lng安排ID, CDate(obj出诊安排.开始时间), CDate(obj出诊安排.终止时间))
        End If
    ElseIf mlng安排ID <> 0 Then
        '限制项目
        strSQL = "Select " & IIf(mbytPlanType = F_MonthTemplet, "'1900-01-' || Replace(b.限制项目, '日', '')", "b.限制项目") & " As 出诊日期" & _
                " From 临床出诊限制 B" & vbNewLine & _
                " Where b.安排id = [1] And b.上班时段 Is Not Null" & vbNewLine & _
                " Group By b.限制项目"
        Set rs出诊项目 = zlDatabase.OpenSQLRecord(strSQL, "获取出诊项目", lng安排ID)

        '时间段信息
        strSQL = "Select b.ID as 记录ID," & IIf(mbytPlanType = F_MonthTemplet, "'1900-01-' || Replace(b.限制项目, '日', '')", "b.限制项目") & " As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, NULL as 开始时间, NULL as 终止时间," & vbNewLine & _
                "        b.限号数, 0 as 已挂数, b.限约数, 0 as 已约数, b.分诊方式, b.预约控制, 0 As 是否临时出诊," & vbNewLine & _
                "        '' As 替诊医生姓名,0 As 科室ID, 0 as 项目ID, '' As 项目名称, 0 as 医生ID, '' As 医生姓名, b.是否独占," & vbNewLine & _
                "        NULL as 停诊开始时间, NULL as 停诊终止时间, NULL as 停诊原因" & vbNewLine & _
                " From 临床出诊限制 B" & vbNewLine & _
                " Where b.安排id = [1] And b.上班时段 Is Not Null"
        Set rs出诊记录 = zlDatabase.OpenSQLRecord(strSQL, "获取出诊时段", lng安排ID)

        '门诊诊室
        strSQL = "Select c.ID As 记录ID,a.诊室ID, b.名称" & vbNewLine & _
                " From 临床出诊诊室 A, 临床出诊限制 C, 门诊诊室 B" & vbNewLine & _
                " Where a.限制ID=c.ID And a.诊室id = b.Id And c.安排id = [1]"
        Set rs门诊诊室 = zlDatabase.OpenSQLRecord(strSQL, "获取临床出诊诊室", lng安排ID)

        '号序信息
        strSQL = "Select b.ID As 记录ID,a.序号, a.开始时间, a. 终止时间, a.限制数量  As 数量, a.是否预约, 0 As 是否停诊" & vbNewLine & _
                " From 临床出诊时段 A,临床出诊限制 B" & vbNewLine & _
                " Where a.限制ID=b.ID And b.安排ID = [1]"
        Set rs号序信息 = zlDatabase.OpenSQLRecord(strSQL, "获取号序信息", lng安排ID)

        '预约挂号控制
        strSQL = "Select c.ID As 记录ID,a.名称,a.类型,a.性质,a.控制方式, a.序号, b.开始时间, b.终止时间, a.数量, b.是否预约, 0 As 是否停诊" & vbNewLine & _
                " From 临床出诊挂号控制 A, 临床出诊时段 B,临床出诊限制 C" & vbNewLine & _
                " Where a.限制ID = b.限制ID(+) And a.序号 = b.序号(+) And a.限制ID=c.ID " & vbNewLine & _
                "       And c.安排ID = [1]"
        Set rs合作单位号序信息 = zlDatabase.OpenSQLRecord(strSQL, "获取合作单位挂号控制", lng安排ID)
    End If

    '转换成对象
    If Not rs出诊项目 Is Nothing Then
        Do While Not rs出诊项目.EOF
            Set obj出诊记录集 = New 出诊记录集
            rs出诊记录.Filter = "出诊日期='" & Nvl(rs出诊项目!出诊日期) & "'"
            Do While Not rs出诊记录.EOF
                Set obj出诊记录 = GetVisitTimesObject(rs出诊记录)

                '上班时段
                If obj所有上班时段.Exits("K" & obj出诊记录.时间段) Then
                    Set obj出诊记录.上班时段 = obj所有上班时段("K" & obj出诊记录.时间段).Clone
                Else
                    '出诊记录时，上班时段可能已被删除
                    Set obj出诊记录.上班时段 = New 上班时段
                    With obj出诊记录.上班时段
                        .开始时间 = obj出诊记录.开始时间
                        .结束时间 = obj出诊记录.终止时间
                    End With
                End If

                '分诊诊室
                rs门诊诊室.Filter = "记录ID=" & obj出诊记录.记录ID
                Set obj出诊记录.安排门诊诊室集 = GetVisitRoomsObjects(rs门诊诊室)
                obj出诊记录.安排门诊诊室集.分诊方式 = obj出诊记录.分诊方式
                obj出诊记录.安排门诊诊室集.医生姓名 = obj出诊安排.出诊号源.医生姓名

                '号序信息
                rs号序信息.Filter = "记录ID=" & obj出诊记录.记录ID
                Set obj出诊记录.号序信息集 = GetTimeIntervalObjects(rs号序信息)
                With obj出诊记录.号序信息集
                    .出诊频次 = obj出诊安排.出诊号源.出诊频次
                    .是否分时段 = obj出诊记录.是否分时段
                    .是否序号控制 = obj出诊记录.是否序号控制
                    .限号数 = obj出诊记录.限号数
                    .限约数 = obj出诊记录.限约数
                    .预约控制 = obj出诊记录.预约控制
                End With

                '合作单位预约控制
                Set obj出诊记录.合作单位控制集 = New 合作单位控制集
                obj出诊记录.合作单位控制集.是否独占 = Val(Nvl(rs出诊记录!是否独占))
                Set obj号序信息集 = Nothing
                strTemp = ""

                rs合作单位号序信息.Filter = "记录ID=" & obj出诊记录.记录ID
                rs合作单位号序信息.Sort = "类型,性质,名称,序号"
                Do While Not rs合作单位号序信息.EOF
                    If strTemp <> Nvl(rs合作单位号序信息!类型) & "-" & Nvl(rs合作单位号序信息!性质) & "-" & Nvl(rs合作单位号序信息!名称) Then
                        If Not obj号序信息集 Is Nothing Then
                            Set obj合作单位控制.号序信息集 = obj号序信息集
                            obj出诊记录.合作单位控制集.AddItem obj合作单位控制, "K" & obj合作单位控制.合作单位名称
                        End If
                        Set obj合作单位控制 = New 合作单位控制
                        obj合作单位控制.合作单位名称 = Nvl(rs合作单位号序信息!名称)
                        obj合作单位控制.类型 = Val(Nvl(rs合作单位号序信息!类型))
                        obj合作单位控制.预约控制方式 = Val(Nvl(rs合作单位号序信息!控制方式))
                        Set obj号序信息集 = New 号序信息集

                        strTemp = Nvl(rs合作单位号序信息!类型) & "-" & Nvl(rs合作单位号序信息!性质) & "-" & Nvl(rs合作单位号序信息!名称)
                    End If

                    obj号序信息集.AddItem GetTimeIntervalObject(rs合作单位号序信息)
                    rs合作单位号序信息.MoveNext
                Loop
                If Not obj号序信息集 Is Nothing Then
                    Set obj合作单位控制.号序信息集 = obj号序信息集
                    obj出诊记录.合作单位控制集.AddItem obj合作单位控制, "K" & obj合作单位控制.合作单位名称
                End If
                
                obj出诊记录.是否固定 = False
                If mbytFun = Fun_TempPlanRecord Then '在进行临时出诊时，不能删除
                    obj出诊记录.是否固定 = True
                ElseIf mbytFun = Fun_UpdatePlan Then '已停诊和已被用于挂号预约的不能调整
                    If CheckPlanIsStopOrUsed(obj出诊记录.记录ID) Then
                        obj出诊记录.是否固定 = True
                    ElseIf CDate(obj出诊记录.终止时间) < dtNow Then '终止时间已小于当前时间的不能调整
                        obj出诊记录.是否固定 = True
                    End If
                End If

                obj出诊记录集.AddItem obj出诊记录, "K" & obj出诊记录.时间段
                rs出诊记录.MoveNext
            Loop

            obj出诊记录集.出诊日期 = Format(Nvl(rs出诊项目!出诊日期), "yyyy-mm-dd")
            obj出诊记录集.是否删除 = False

            obj出诊安排.已保存出诊安排.AddItem obj出诊记录集, "K" & obj出诊记录集.出诊日期
            If obj出诊记录集.出诊日期 = Format(mstr缺省日期, "yyyy-mm-dd") Then
                obj出诊安排.AddItem obj出诊记录集.Clone, "K" & obj出诊记录集.出诊日期
            End If

            rs出诊项目.MoveNext
        Loop
    End If
    
    '缺省选择日期
    If Not (mbytFun = Fun_AddSignalSourcePlan Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel) Then
        obj出诊安排.缺省出诊日期 = mstr缺省日期
        If obj出诊安排.Count = 0 And mbytPlanType <> F_MonthTemplet Then
            '模板其他规则时，缺省到第一个安排项目
            If obj出诊安排.已保存出诊安排.Count > 0 And Not (obj出诊安排.排班规则 = 0 Or obj出诊安排.排班规则 = 1) Then
                obj出诊安排.AddItem obj出诊安排.已保存出诊安排(1).Clone, GetPlanKey(obj出诊安排.已保存出诊安排(1).出诊日期)
                obj出诊安排.缺省出诊日期 = obj出诊安排.已保存出诊安排(1).出诊日期
            End If
        End If
    End If

    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function LoadData() As Boolean
    'blnReLoadTimeRange 是否重新加载时间范围
    Dim obj上班时段 As 上班时段
    Dim objListItem As ListItem
    
    On Error GoTo Errhand

    '有效时间范围
    If mbytPlanType = F_FixedRule And mblnFirst Then
        dtpBegin.Value = Format(mobj出诊安排.开始时间, "yyyy-MM-dd hh:mm:ss")
        dtpEnd.Value = Format(mobj出诊安排.终止时间, "yyyy-MM-dd hh:mm:ss")
        
        dtpBegin.Tag = dtpBegin.Value: dtpEnd.Tag = dtpEnd.Value
    End If
    
    '出诊日期
    cldsCalenbarSel.LoadData mobj出诊安排

    '出诊时间
    lvwWorkTime.ListItems.Clear
    If Not mobj出诊安排.所有上班时段 Is Nothing Then
        For Each obj上班时段 In mobj出诊安排.所有上班时段
            Set objListItem = lvwWorkTime.ListItems.Add(, "K" & obj上班时段.时间段, obj上班时段.时间段 & _
                "(" & Format(obj上班时段.开始时间, "hh:mm") & "-" & Format(obj上班时段.结束时间, "hh:mm") & ")")
            objListItem.SubItems(1) = obj上班时段.开始时间
            objListItem.SubItems(2) = obj上班时段.结束时间
            objListItem.Tag = obj上班时段.时间段
            '用颜色区分是否号源中设置的时间段
            If mobj出诊安排.号源安排.Exits("K" & obj上班时段.时间段) Then
                objListItem.ForeColor = vbBlue
            End If
        Next
    End If
    If Not lvwWorkTime.SelectedItem Is Nothing Then
        lvwWorkTime.SelectedItem.Selected = False '去掉选中项背景
    End If

    '号源信息
    SourceInfor.LoadData mobj出诊安排.出诊号源

    '出诊安排
    Call LoadDetailData

    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadDetailData() As Boolean
    Dim objTemp As 出诊记录集, i As Integer

    On Error GoTo Errhand

    Screen.MousePointer = vbHourglass
    '当前项目
    If mobj出诊安排.Count = 0 Then
        Set objTemp = New 出诊记录集
    Else
        Set objTemp = mobj出诊安排(1).Clone
    End If

    mstrCurDay = objTemp.出诊日期
    Call SetTitleText

    '时间段
    CheckWorkTime objTemp
    
    '安排预览
    Call ShowPlan(mobj出诊安排)
    
    '安排
    CPDPages.LoadData objTemp, mobj出诊安排.所有分诊诊室, mobj出诊安排.所有合作单位, True

    '恢复应用
    If mbytPlanType = F_MonthPlan Or mbytPlanType = F_MonthTemplet Then
        optRule(0).Value = True
        mblnNotClick = True
        For i = chkWeek.LBound To chkWeek.UBound
            chkWeek(i).Value = vbUnchecked
        Next
        mblnNotClick = False
    End If

    SetEnabled Not (mbytFun = Fun_View _
            Or mbytFun = Fun_UpdateUnit Or mlng号源Id = 0 _
            Or mbytFun = Fun_TempPlanVerify _
            Or mbytFun = Fun_TempPlanCancel)
    CPDPages.EditMode = IIf(mbytFun = Fun_View Or mlng号源Id = 0 Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel, _
        ED_RegistPlan_View, IIf(mbytFun = Fun_UpdateUnit, ED_RegistPlan_UpdateUnit, ED_RegistPlan_Edit))

    If IsDate(mstrCurDay) And (mbytFun = Fun_AddSignalSourcePlan _
        Or mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan) Then
        If DateDiff("d", mstrCurDay, mdtToday) > 0 Then
            SetEnabled False
            CPDPages.EditMode = ED_RegistPlan_View
        End If
    End If

    Screen.MousePointer = vbDefault
    LoadDetailData = True
    Exit Function
Errhand:
    Screen.MousePointer = vbDefault
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SavePlanData(ByVal lng安排ID As Long, ByVal obj出诊安排 As 出诊安排, _
    ByVal obj出诊号源 As 出诊号源, cllPro As Collection, ByVal blnRecord As Boolean, _
    ByVal dtCurdate As Date, ByVal obj已保存出诊安排 As 出诊安排) As Boolean
    '功能：保存安排详细数据
    Dim strSQL As String
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim strTemp As String, lngTemp As Long, lng记录ID As Long
    Dim str诊室 As String, obj诊室 As 分诊诊室, byt分诊方式 As Byte
    Dim str号序 As String, obj号序 As 号序信息, cll号序 As Collection
    Dim obj合作单位 As 合作单位控制, bln是否独占 As Boolean
    Dim i As Long, blnPublished As Boolean
        
    Dim obj出诊记录集1 As 出诊记录集, obj出诊记录1 As 出诊记录
    Dim blnFind As Boolean, blnFindPlan As Boolean
    Dim lngSavedCount As Long, lngNewCount As Long

    Err = 0: On Error GoTo ErrHandler
    If obj出诊安排 Is Nothing Then SavePlanData = True: Exit Function
    If obj出诊安排.Count = 0 Then SavePlanData = True: Exit Function

    If cllPro Is Nothing Then Set cllPro = New Collection
    
    If Not obj已保存出诊安排 Is Nothing Then
        '已保存安排如果全部被删除，且没有新的安排，则删除临床出诊安排
        For Each obj出诊记录集 In obj已保存出诊安排
            lngSavedCount = lngSavedCount + obj出诊记录集.Count
        Next
        For Each obj出诊记录集 In obj出诊安排
            lngNewCount = lngNewCount + obj出诊记录集.Count
        Next
        
        '清除在已保存中的出诊记录，不在修改后中的出诊记录
        '注意：如果安排在已保存中，而不在未保存中，则表示未进行过查看，肯定没有修改，不用删除
        For Each obj出诊记录集 In obj已保存出诊安排
            blnFindPlan = False
            If obj出诊记录集.是否删除 = False Then ' 全部删除的已在前面处理
                For Each obj出诊记录 In obj出诊记录集
                    blnFind = False
                    For Each obj出诊记录集1 In obj出诊安排
                        If blnFind Then Exit For
                        If obj出诊记录集1.出诊日期 = obj出诊记录集.出诊日期 Then
                            blnFindPlan = True '未保存中未找到表示没有修改
                            For Each obj出诊记录1 In obj出诊记录集1
                                '用记录ID判断
                                'If obj出诊记录1.时间段 = obj出诊记录.时间段 Then
                                If obj出诊记录1.记录ID = obj出诊记录.记录ID Then
                                    blnFind = True: Exit For
                                End If
                            Next
                        End If
                    Next
                    
                    If blnFindPlan And blnFind = False Then
                        lngSavedCount = lngSavedCount - 1
                        'Zl_临床出诊上班时段_Delete(
                        strSQL = "Zl_临床出诊上班时段_Delete("
                        '安排id_In       临床出诊限制.安排id%Type,
                        strSQL = strSQL & "" & lng安排ID & ","
                        '项目_In         临床出诊限制.限制项目%Type,
                        strSQL = strSQL & "'" & IIf(mbytPlanType = F_MonthTemplet And blnRecord = False, _
                            FormatApplyToStr(obj出诊记录集.出诊日期), obj出诊记录集.出诊日期) & "',"
                        '出诊记录_In     Number := 0,
                        strSQL = strSQL & "" & IIf(blnRecord, 1, 0) & ","
                        '上班时段_In     临床出诊限制.上班时段%Type,
                        strSQL = strSQL & "'" & obj出诊记录.时间段 & "',"
                        '删除出诊安排_In Number:=0
                        strSQL = strSQL & "" & IIf(lngSavedCount = 0 And lngNewCount = 0, 1, 0) & ")"
                        cllPro.Add strSQL
                    End If
                Next
            End If
        Next
    End If
    
    '出诊记录
    If blnRecord Then
        For Each obj出诊记录集 In obj出诊安排
            '保存出诊记录
            For Each obj出诊记录 In obj出诊记录集
                '固定的未修改，不保存
                If obj出诊记录.是否固定 = False Then
                    lng记录ID = obj出诊记录.记录ID
                    If lng记录ID = 0 Then lng记录ID = zlDatabase.GetNextId("临床出诊记录")
                    bln是否独占 = obj出诊记录.合作单位控制集.是否独占
                    obj出诊记录.开始时间 = Format(obj出诊记录集.出诊日期, "yyyy-mm-dd ") & Format(obj出诊记录.上班时段.开始时间, "hh:mm:ss")
                    obj出诊记录.终止时间 = GetWorkTrueDate(obj出诊记录.开始时间, obj出诊记录.上班时段.结束时间)

                    '门诊诊室
                    byt分诊方式 = obj出诊记录.安排门诊诊室集.分诊方式
                    str诊室 = ""
                    For Each obj诊室 In obj出诊记录.安排门诊诊室集
                        '诊室_In:诊室1,诊室2,...
                        str诊室 = str诊室 & "," & obj诊室.诊室ID
                    Next
                    If str诊室 <> "" Then str诊室 = Mid(str诊室, 2)

                    '出诊时段
                    Set cll号序 = New Collection: str号序 = ""
                    For Each obj号序 In obj出诊记录.号序信息集
                        strTemp = obj号序.序号 & "," & _
                            GetWorkTrueDate(obj出诊记录.开始时间, ZDate(obj号序.开始时间, obj出诊记录.开始时间, False), , False) & "," & _
                            GetWorkTrueDate(obj出诊记录.开始时间, ZDate(obj号序.终止时间, obj出诊记录.终止时间, False)) & "," & _
                            obj号序.数量 & "," & IIf(obj号序.是否预约, 1, 0)
                        If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                            '时段_In:序号,开始时间,终止时间,限制数量,预约标志|...
                            str号序 = Mid(str号序, 2)
                            cll号序.Add str号序
                            str号序 = ""
                        End If
                        str号序 = str号序 & "|" & strTemp
                    Next
                    If str号序 <> "" Then
                        str号序 = Mid(str号序, 2)
                        cll号序.Add str号序
                    End If
                    For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                        'Zl_临床出诊记录_Insert(
                        strSQL = "Zl_临床出诊记录_Insert("
                        'Id_In           临床出诊记录.Id%Type,
                        strSQL = strSQL & "" & lng记录ID & ","
                        '安排id_In       临床出诊限制.安排id%Type,
                        strSQL = strSQL & "" & lng安排ID & ","
                        '号源id_In       临床出诊记录.号源id%Type,
                        strSQL = strSQL & "" & obj出诊号源.ID & ","
                        '出诊日期_In     临床出诊记录.出诊日期%Type,
                        strSQL = strSQL & "To_Date('" & obj出诊记录集.出诊日期 & "','yyyy-mm-dd'),"
                        '上班时段_In     临床出诊记录.上班时段%Type,
                        strSQL = strSQL & "'" & obj出诊记录.时间段 & "',"
                        '开始时间_In     临床出诊记录.开始时间%Type,
                        strSQL = strSQL & "" & ZDate(obj出诊记录.开始时间) & ","
                        '终止时间_In     临床出诊记录.终止时间%Type,
                        strSQL = strSQL & "" & ZDate(obj出诊记录.终止时间) & ","
                        '缺省预约时间_In 临床出诊记录.缺省预约时间%Type,
                        strSQL = strSQL & "" & ZDate(GetWorkTrueDate(obj出诊记录.开始时间, obj出诊记录.上班时段.缺省预约时间)) & ","
                        '提前挂号时间_In 临床出诊记录.提前挂号时间%Type,
                        strSQL = strSQL & "" & ZDate(GetWorkTrueDate(obj出诊记录.开始时间, obj出诊记录.上班时段.提前挂号时间, False)) & ","
                        '限号数_In       临床出诊记录.限号数%Type,
                        strSQL = strSQL & "" & ZVal(obj出诊记录.限号数) & ","
                        '限约数_In       临床出诊记录.限约数%Type,
                        strSQL = strSQL & "" & ZVal(obj出诊记录.限约数) & ","
                        '是否序号控制_In 临床出诊记录.是否序号控制%Type,
                        strSQL = strSQL & "" & IIf(obj出诊记录.是否序号控制, 1, 0) & ","
                        '是否分时段_In   临床出诊记录.是否分时段%Type,
                        strSQL = strSQL & "" & IIf(obj出诊记录.是否分时段, 1, 0) & ","
                        '预约控制_In     临床出诊记录.预约控制%Type,
                        strSQL = strSQL & "" & obj出诊记录.预约控制 & ","
                        '是否独占_In     临床出诊记录.是否独占%Type,
                        strSQL = strSQL & "" & IIf(bln是否独占, 1, 0) & ","
                        '项目id_In       临床出诊记录.项目id%Type,
                        strSQL = strSQL & "" & ZVal(obj出诊号源.项目ID) & ","
                        '科室id_In       临床出诊记录.科室id%Type,
                        strSQL = strSQL & "" & obj出诊号源.科室ID & ","
                        '医生id_In       临床出诊记录.医生id%Type,
                        strSQL = strSQL & "" & ZVal(obj出诊号源.医生ID) & ","
                        '医生姓名_In     临床出诊记录.医生姓名%Type,
                        strTemp = obj出诊号源.医生姓名
                        strSQL = strSQL & "" & IIf(strTemp = "", "NULL", "'" & strTemp & "'") & ","
                        '分诊方式_In     临床出诊记录.分诊方式%Type,
                        strSQL = strSQL & "" & byt分诊方式 & ","
                        '是否临时出诊_In 临床出诊记录.是否临时出诊%Type,
                        strSQL = strSQL & "" & IIf(mbytFun = Fun_UpdatePlan, "NULL", IIf(mbytFun = Fun_TempPlanRecord, 1, 0)) & ","
                        '登记人_In       临床出诊记录.登记人%Type,
                        strSQL = strSQL & IIf(mbytFun = Fun_UpdatePlan, "NULL", "'" & UserInfo.姓名 & "'") & ","
                        '登记时间_In     临床出诊记录.登记时间%Type,
                        strSQL = strSQL & IIf(mbytFun = Fun_UpdatePlan, "NULL", ZDate(dtCurdate)) & ","
                        '是否发布_In     临床出诊记录.是否发布%Type,
                        '1.临时出诊
                        '2.月安排或周安排增加号源
                        blnPublished = mbytFun = Fun_TempPlanRecord _
                                    Or mbytFun = Fun_AddSignalSourcePlan And (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan)
                        strSQL = strSQL & IIf(mbytFun = Fun_UpdatePlan, "NULL", IIf(blnPublished, 1, 0)) & ","
                        '诊室_In         Varchar2 := Null,
                        strSQL = strSQL & "'" & str诊室 & "',"
                        '时段_In         Varchar2 := Null,
                        str号序 = ""
                        If cll号序.Count > 0 Then str号序 = cll号序(i)
                        strSQL = strSQL & "'" & str号序 & "',"
                        '删除序号_In Number:=0
                        strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                        cllPro.Add strSQL
                    Next
                    '出诊挂号控制
                    For Each obj合作单位 In obj出诊记录.合作单位控制集
                        '预约控制:0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
                        '类型:1-三方机构;2-预约方式
                        Set cll号序 = New Collection: str号序 = ""
                        For Each obj号序 In obj合作单位.号序信息集
                            strTemp = obj号序.序号 & "," & obj号序.数量
                            If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                                '安排控制_in:序号1,数量|序号2,数量|...
                                str号序 = Mid(str号序, 2)
                                cll号序.Add str号序
                                str号序 = ""
                            End If
                            str号序 = str号序 & "|" & strTemp
                        Next
                        If str号序 <> "" Then
                            str号序 = Mid(str号序, 2)
                            cll号序.Add str号序
                        End If
                        For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                            'Zl_临床出诊挂号控制记录_Insert(
                            strSQL = "Zl_临床出诊挂号控制记录_Insert("
                            '记录id_In   临床出诊挂号控制记录.记录id%Type,
                            strSQL = strSQL & "" & lng记录ID & ","
                            '类型_In     临床出诊挂号控制记录.类型%Type,
                            strSQL = strSQL & "" & obj合作单位.类型 & ","
                            '性质_In     临床出诊挂号控制记录.性质%Type,
                            strSQL = strSQL & "" & 1 & ","
                            '名称_In     临床出诊挂号控制记录.名称%Type,
                            strSQL = strSQL & "'" & obj合作单位.合作单位名称 & "',"
                            '控制方式_In 临床出诊挂号控制记录.控制方式%Type,
                            strSQL = strSQL & "" & obj合作单位.预约控制方式 & ","
                            '是否独占_In 临床出诊记录.是否独占%Type,
                            strSQL = strSQL & "" & IIf(obj出诊记录.合作单位控制集.是否独占, 1, 0) & ","
                            '安排控制_In Varchar2,
                            str号序 = ""
                            If cll号序.Count > 0 Then str号序 = cll号序(i)
                            strSQL = strSQL & "'" & str号序 & "',"
                            '删除_In Number:=0
                            strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                            cllPro.Add strSQL
                        Next
                    Next
                End If
            Next
        Next
    Else  '模板/固定规则
        For Each obj出诊记录集 In obj出诊安排
            '保存限制项目
            For Each obj出诊记录 In obj出诊记录集
                lng记录ID = obj出诊记录.记录ID
                If lng记录ID = 0 Then lng记录ID = zlDatabase.GetNextId("临床出诊限制")
                bln是否独占 = obj出诊记录.合作单位控制集.是否独占
                '门诊诊室
                byt分诊方式 = obj出诊记录.安排门诊诊室集.分诊方式
                str诊室 = ""
                For Each obj诊室 In obj出诊记录.安排门诊诊室集
                    '诊室_In:诊室1,诊室2,...
                    str诊室 = str诊室 & "," & obj诊室.诊室ID
                Next
                If str诊室 <> "" Then str诊室 = Mid(str诊室, 2)
    
                '出诊时段
                Set cll号序 = New Collection: str号序 = ""
                For Each obj号序 In obj出诊记录.号序信息集
                    strTemp = obj号序.序号 & ","
                    strTemp = strTemp & GetWorkTrueDate(obj出诊记录.开始时间, ZDate(obj号序.开始时间, obj出诊记录.开始时间, False), , False) & ","
                    strTemp = strTemp & GetWorkTrueDate(obj出诊记录.开始时间, ZDate(obj号序.终止时间, obj出诊记录.终止时间, False)) & ","
                    strTemp = strTemp & obj号序.数量 & "," & IIf(obj号序.是否预约, 1, 0)
    
                    If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                        '时段_In:序号,开始时间,终止时间,限制数量,预约标志|...
                        str号序 = Mid(str号序, 2)
                        cll号序.Add str号序
                        str号序 = ""
                    End If
                    str号序 = str号序 & "|" & strTemp
                Next
                If str号序 <> "" Then
                    str号序 = Mid(str号序, 2)
                    cll号序.Add str号序
                End If
                For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                    'Zl_临床出诊限制_Insert(
                    strSQL = "Zl_临床出诊限制_Insert("
                    'Id_In           临床出诊限制.Id%Type,
                    strSQL = strSQL & "" & lng记录ID & ","
                    '安排id_In       临床出诊限制.安排id%Type,
                    strSQL = strSQL & "" & lng安排ID & ","
                    '限制项目_In     临床出诊限制.限制项目%Type,
                    strSQL = strSQL & "'" & IIf(mbytPlanType = F_MonthTemplet, FormatApplyToStr(obj出诊记录集.出诊日期), obj出诊记录集.出诊日期) & "',"
                    '上班时段_In     临床出诊限制.上班时段%Type,
                    strSQL = strSQL & "'" & obj出诊记录.时间段 & "',"
                    '限号数_In       临床出诊限制.限号数%Type,
                    strSQL = strSQL & "" & ZVal(obj出诊记录.限号数) & ","
                    '限约数_In       临床出诊限制.限约数%Type,
                    strSQL = strSQL & "" & ZVal(obj出诊记录.限约数) & ","
                    '是否分时段_In   临床出诊限制.是否分时段%Type,
                    strSQL = strSQL & "" & IIf(obj出诊记录.是否分时段, 1, 0) & ","
                    '是否序号控制_In 临床出诊限制.是否序号控制%Type,
                    strSQL = strSQL & "" & IIf(obj出诊记录.是否序号控制, 1, 0) & ","
                    '预约控制_In     临床出诊限制.预约控制%Type,
                    strSQL = strSQL & "" & obj出诊记录.预约控制 & ","
                    '是否独占_In     临床出诊限制.是否独占%Type,
                    strSQL = strSQL & "" & IIf(bln是否独占, 1, 0) & ","
                    '分诊方式_In     临床出诊限制.分诊方式%Type := Null,
                    strSQL = strSQL & "" & byt分诊方式 & ","
                    '诊室_In         Varchar2 := Null,
                    strSQL = strSQL & "'" & str诊室 & "',"
                    '时段_In         Varchar2 := Null,
                    str号序 = ""
                    If cll号序.Count > 0 Then str号序 = cll号序(i)
                    strSQL = strSQL & "'" & str号序 & "',"
                    '删除序号_In Number:=0
                    strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                    cllPro.Add strSQL
                Next
                '出诊挂号控制
                For Each obj合作单位 In obj出诊记录.合作单位控制集
                    '预约控制:0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
                    '类型:1-三方机构;2-预约方式
                    Set cll号序 = New Collection: str号序 = ""
                    For Each obj号序 In obj合作单位.号序信息集
                        strTemp = obj号序.序号 & "," & obj号序.数量
                        If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                            '安排控制_in:序号1,数量|序号2,数量|...
                            str号序 = Mid(str号序, 2)
                            cll号序.Add str号序
                            str号序 = ""
                        End If
                        str号序 = str号序 & "|" & strTemp
                    Next
                    If str号序 <> "" Then
                        str号序 = Mid(str号序, 2)
                        cll号序.Add str号序
                    End If
                    For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                        'Zl_临床出诊挂号控制_Insert(
                        strSQL = "Zl_临床出诊挂号控制_Insert("
                        '限制id_In   临床出诊挂号控制.限制id%Type,
                        strSQL = strSQL & "" & lng记录ID & ","
                        '类型_In     临床出诊挂号控制.类型%Type,
                        strSQL = strSQL & "" & obj合作单位.类型 & ","
                        '性质_In     临床出诊挂号控制.性质%Type,
                        strSQL = strSQL & "" & 1 & ","
                        '名称_In     临床出诊挂号控制.名称%Type,
                        strSQL = strSQL & "'" & obj合作单位.合作单位名称 & "',"
                        '控制方式_In 临床出诊挂号控制.控制方式%Type,
                        strSQL = strSQL & "" & obj合作单位.预约控制方式 & ","
                        '是否独占_In 临床出诊限制.是否独占%Type,
                        strSQL = strSQL & "" & IIf(obj出诊记录.合作单位控制集.是否独占, 1, 0) & ","
                        '安排控制_In Varchar2,
                        str号序 = ""
                        If cll号序.Count > 0 Then str号序 = cll号序(i)
                        strSQL = strSQL & "'" & str号序 & "',"
                        '删除_In Number:=0
                        strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                        cllPro.Add strSQL
                    Next
                Next
            Next
        Next
    End If

    SavePlanData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    If mbytFun = Fun_AddSignalSourcePlan And mlng号源Id <> 0 And mlngSavedRecords = 0 Then
        If mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan Then
            '重新生成出诊记录，在开始时删除了未被使用的固定安排的出诊记录的
            strSQL = "Zl1_Auto_Buildingregisterplan(Null)"
            zlDatabase.ExecuteProcedure strSQL, "恢复出诊记录"
        End If
    End If
    
    Set mobj出诊安排 = Nothing
    Set mrsVisitedRecord = Nothing
    Set mobj停诊记录集 = Nothing
    Set mrsVisitedRecordByDate = Nothing
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveUpdateUnit(ByRef obj出诊安排 As 出诊安排) As Boolean
    '功能：保存合作单位
    Dim strSQL As String, cllPro As Collection, strTemp As String
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim str号序 As String, obj号序 As 号序信息, cll号序 As Collection
    Dim obj合作单位 As 合作单位控制, i As Integer
    Dim blnTrans As Boolean

    Err = 0: On Error GoTo ErrHandler
    Set cllPro = New Collection
    For Each obj出诊记录集 In obj出诊安排
        Select Case mbytPlanType
        Case F_Templet, F_FixedRule, F_MonthTemplet
            For Each obj出诊记录 In obj出诊记录集
                If obj出诊记录.合作单位控制集.是否修改 Then
                    '出诊挂号控制
                    For Each obj合作单位 In obj出诊记录.合作单位控制集
                        '预约控制:0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
                        '类型:1-三方机构;2-预约方式
                        If Not (obj出诊记录.预约控制 = 1 _
                            Or (obj出诊记录.预约控制 = 2 And obj合作单位.类型 = 1)) Then

                            Set cll号序 = New Collection: str号序 = ""
                            For Each obj号序 In obj合作单位.号序信息集
                                strTemp = obj号序.序号 & "," & obj号序.数量
                                If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                                    '安排控制_in:序号1,数量|序号2,数量|...
                                    str号序 = Mid(str号序, 2)
                                    cll号序.Add str号序
                                    str号序 = ""
                                End If
                                str号序 = str号序 & "|" & strTemp
                            Next
                            If str号序 <> "" Then
                                str号序 = Mid(str号序, 2)
                                cll号序.Add str号序
                            End If
                            For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                                'Zl_临床出诊挂号控制_Insert(
                                strSQL = "Zl_临床出诊挂号控制_Insert("
                                '限制id_In   临床出诊挂号控制.限制id%Type,
                                strSQL = strSQL & "" & obj出诊记录.记录ID & ","
                                '类型_In     临床出诊挂号控制.类型%Type,
                                strSQL = strSQL & "" & obj合作单位.类型 & ","
                                '性质_In     临床出诊挂号控制.性质%Type,
                                strSQL = strSQL & "" & 1 & ","
                                '名称_In     临床出诊挂号控制.名称%Type,
                                strSQL = strSQL & "'" & obj合作单位.合作单位名称 & "',"
                                '控制方式_In 临床出诊挂号控制.控制方式%Type,
                                strSQL = strSQL & "" & obj合作单位.预约控制方式 & ","
                                '是否独占_In 临床出诊记录.是否独占%Type,
                                strSQL = strSQL & "" & IIf(obj出诊记录.合作单位控制集.是否独占, 1, 0) & ","
                                '安排控制_In Varchar2,
                                str号序 = ""
                                If cll号序.Count > 0 Then str号序 = cll号序(i)
                                strSQL = strSQL & "'" & str号序 & "',"
                                '删除_In Number:=0
                                strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                                cllPro.Add strSQL
                            Next
                        End If
                    Next
                End If
            Next
        Case Else
            For Each obj出诊记录 In obj出诊记录集
                If obj出诊记录.合作单位控制集.是否修改 Then
                    For Each obj合作单位 In obj出诊记录.合作单位控制集
                        '预约控制:0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
                        '类型:1-三方机构;2-预约方式
                        If Not (obj出诊记录.预约控制 = 1 _
                            Or (obj出诊记录.预约控制 = 2 And obj合作单位.类型 = 1)) Then

                            Set cll号序 = New Collection: str号序 = ""
                            For Each obj号序 In obj合作单位.号序信息集
                                strTemp = obj号序.序号 & "," & obj号序.数量
                                If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                                    '安排控制_in:序号1,数量|序号2,数量|...
                                    str号序 = Mid(str号序, 2)
                                    cll号序.Add str号序
                                    str号序 = ""
                                End If
                                str号序 = str号序 & "|" & strTemp
                            Next
                            If str号序 <> "" Then
                                str号序 = Mid(str号序, 2)
                                cll号序.Add str号序
                            End If
                            For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                                'Zl_临床出诊挂号控制记录_Insert(
                                strSQL = "Zl_临床出诊挂号控制记录_Insert("
                                '记录id_In   临床出诊挂号控制记录.记录id%Type,
                                strSQL = strSQL & "" & obj出诊记录.记录ID & ","
                                '类型_In     临床出诊挂号控制记录.类型%Type,
                                strSQL = strSQL & "" & obj合作单位.类型 & ","
                                '性质_In     临床出诊挂号控制记录.性质%Type,
                                strSQL = strSQL & "" & 1 & ","
                                '名称_In     临床出诊挂号控制记录.名称%Type,
                                strSQL = strSQL & "'" & obj合作单位.合作单位名称 & "',"
                                '控制方式_In 临床出诊挂号控制记录.控制方式%Type,
                                strSQL = strSQL & "" & obj合作单位.预约控制方式 & ","
                                '是否独占_In 临床出诊记录.是否独占%Type,
                                strSQL = strSQL & "" & IIf(obj出诊记录.合作单位控制集.是否独占, 1, 0) & ","
                                '安排控制_In Varchar2,
                                str号序 = ""
                                If cll号序.Count > 0 Then str号序 = cll号序(i)
                                strSQL = strSQL & "'" & str号序 & "',"
                                '删除_In Number:=0
                                strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                                cllPro.Add strSQL
                            Next
                        End If
                    Next
                End If
            Next
        End Select
    Next

    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    SaveUpdateUnit = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TempPlanVerifyOrCancel(ByVal lng安排ID As Long, ByVal lng号源Id As Long, _
    ByVal blnVerify As Boolean) As Boolean
    '功能：审核或取消审核临时安排
    Dim strSQL As String

    Err = 0: On Error GoTo ErrHandler
    If lng安排ID = 0 Then Exit Function
    
    If blnVerify Then
        'Zl_临床出诊临时安排_Verify(
        strSQL = "Zl_临床出诊临时安排_Verify("
        '安排id_In In 临床出诊安排.Id%Type,
        strSQL = strSQL & "" & lng安排ID & ","
        '审核人_in in 临床出诊安排.审核人%type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '审核时间_in in 临床出诊安排.审核时间%type
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    Else
        'Zl_临床出诊临时安排_Cancel(
        strSQL = "Zl_临床出诊临时安排_Cancel("
        '安排id_In In 临床出诊安排.Id%Type
        strSQL = strSQL & "" & lng安排ID & ")"
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    TempPlanVerifyOrCancel = True
    
    '重新生成出诊记录
    strSQL = "Zl1_Auto_Buildingregisterplan(Null," & lng号源Id & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData(obj出诊安排 As 出诊安排) As Boolean
    '保存数据
    Dim cllPro As Collection, blnTrans As Boolean
    Dim dtCurdate As Date, strTemp As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnRecord As Boolean, blnDeletePlan As Boolean
    Dim obj出诊安排Temp As 出诊安排, obj出诊记录集 As 出诊记录集
    Dim ObjItem As 出诊记录集, blnUpdatePlan As Boolean
    Dim blnExistPlan As Boolean

    On Error GoTo ErrHandler
    If mbytFun = Fun_UpdateUnit Then
        '更新预约挂号控制
        SaveData = SaveUpdateUnit(obj出诊安排.未保存出诊安排)
        Exit Function
    ElseIf mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel Then
        SaveData = TempPlanVerifyOrCancel(obj出诊安排.安排ID, obj出诊安排.出诊号源.ID, mbytFun = Fun_TempPlanVerify)
        Exit Function
    End If

    Set cllPro = New Collection
    If obj出诊安排.安排ID = 0 Then
        '获取安排ID
        obj出诊安排.安排ID = zlDatabase.GetNextId("临床出诊安排")
    End If
    dtCurdate = zlDatabase.Currentdate
    blnRecord = mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan

    '如果已保存中被删除，未保存中又有，则此时要重新修改安排；
    blnUpdatePlan = False
    If mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan Then
        '发布后临时出诊或调整安排不能调整原临床出诊安排记录，因为已发布
    Else
        For Each obj出诊记录集 In obj出诊安排.已保存出诊安排
            If obj出诊记录集.是否删除 Then
                For Each ObjItem In obj出诊安排.未保存出诊安排
                    If obj出诊记录集.出诊日期 = ObjItem.出诊日期 Then
                        blnUpdatePlan = True: Exit For
                    End If
                Next
            End If
        Next
    End If
    
    '1.调整模板后要更新安排
    '2.改变固定安排时间范围后要调整安排的开始时间和终止时间
    '3.改变收费项目后要调整安排的项目ID
    If mbytPlanType = F_Templet Or mblnTimeChanged Or mblnFeeItemChanged Then
        blnUpdatePlan = True
    End If

    '模板规则变化后删除原有出诊项目信息
    '模板排班规则为2(单日),3(双日),4(月内轮循),5(无限轮循)时删除重新添加
    If obj出诊安排.已保存出诊安排.Count > 0 And mbytPlanType = F_Templet _
        And (obj出诊安排.已保存出诊安排.排班规则 <> obj出诊安排.排班规则 _
            Or InStr(",2,3,4,5,", obj出诊安排.排班规则) > 0) Then
        'Zl_临床出诊安排_Batchdelete
        strSQL = "Zl_临床出诊安排_Batchdelete("
        '出诊id_In 临床出诊表.Id%Type,
        strSQL = strSQL & "" & obj出诊安排.出诊ID & ","
        '人员id_In 人员表.Id%Type := 0,--不等于0则删除人员所在科室的所有号源安排
        strSQL = strSQL & "" & "NULL" & ","
        '站点_In   部门表.站点%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '号源id_In 临床出诊安排.号源id%Type := 0--不等于0则删除该号源的所有安排
        strSQL = strSQL & "" & obj出诊安排.出诊号源.ID & ")"
        cllPro.Add strSQL
    End If

    '删除未设置时段的安排
    blnDeletePlan = mbytFun = Fun_Update Or mbytFun = Fun_TempPlan
    If DeletePlan(obj出诊安排.已保存出诊安排, cllPro, blnRecord, blnExistPlan, blnDeletePlan) = False Then Exit Function
    
    If blnUpdatePlan Or obj出诊安排.未保存出诊安排.Count > 0 Then
        If (blnUpdatePlan Or blnExistPlan = False And obj出诊安排.未保存出诊安排.Count > 0) _
            And Not (mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan) Then
            'Zl_临床出诊安排_Insert(
            strSQL = "Zl_临床出诊安排_Insert("
            'Id_In           临床出诊安排.Id%Type,
            strSQL = strSQL & "" & obj出诊安排.安排ID & ","
            '出诊id_In       临床出诊安排.出诊id%Type,
            strSQL = strSQL & "" & obj出诊安排.出诊ID & ","
            '号源id_In       临床出诊安排.号源id%Type,
            strSQL = strSQL & "" & obj出诊安排.出诊号源.ID & ","
            '项目id_In       临床出诊安排.项目id%Type,
            If obj出诊安排.项目ID = 0 Then
                strTemp = obj出诊安排.出诊号源.项目ID
            Else
                strTemp = obj出诊安排.项目ID
            End If
            strSQL = strSQL & "" & ZVal(strTemp) & ","
            '医生id_In       临床出诊安排.医生id%Type,
            If obj出诊安排.医生ID = 0 Then
                strTemp = obj出诊安排.出诊号源.医生ID
            Else
                strTemp = obj出诊安排.医生ID
            End If
            strSQL = strSQL & "" & ZVal(strTemp) & ","
            '医生姓名_In     临床出诊安排.医生姓名%Type,
            If obj出诊安排.医生姓名 = "" Then
                strTemp = obj出诊安排.出诊号源.医生姓名
            Else
                strTemp = obj出诊安排.医生姓名
            End If
            strSQL = strSQL & "" & IIf(strTemp = "", "NULL", "'" & strTemp & "'") & ","
            '排班规则_In     临床出诊安排.排班规则%Type,
            If mbytPlanType = F_Templet Then
                strTemp = obj出诊安排.排班规则
            Else
                strTemp = IIf(mbytPlanType = F_MonthTemplet, 6, "NULL")
            End If
            strSQL = strSQL & "" & strTemp & ","
            '是否周六出诊_In 临床出诊安排.是否周六出诊%Type,
            If mbytPlanType = F_Templet And InStr("2,3,4,5", obj出诊安排.排班规则) > 0 Then
                strTemp = IIf(obj出诊安排.周六不出诊, "0", "1")
            Else
                strTemp = "NULL"
            End If
            strSQL = strSQL & "" & strTemp & ","
            '是否周日出诊_In 临床出诊安排.是否周日出诊%Type,
            If mbytPlanType = F_Templet And InStr("2,3,4,5", obj出诊安排.排班规则) > 0 Then
                strTemp = IIf(obj出诊安排.周日不出诊, "0", "1")
            Else
                strTemp = "NULL"
            End If
            strSQL = strSQL & "" & strTemp & ","
            '开始时间_In     临床出诊安排.开始时间%Type,
            strSQL = strSQL & "" & IIf(mbytPlanType = F_MonthTemplet, "NULL", ZDate(obj出诊安排.开始时间)) & ","
            '终止时间_In     临床出诊安排.终止时间%Type,
            strSQL = strSQL & "" & IIf(mbytPlanType = F_MonthTemplet, "NULL", ZDate(obj出诊安排.终止时间)) & ","
            '操作员姓名_In   临床出诊安排.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '登记时间_In     临床出诊安排.登记时间%Type
            strSQL = strSQL & "" & ZDate(dtCurdate) & ","
            '是否审核_In     number
            If mbytFun = Fun_AddSignalSourcePlan _
                And (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            strSQL = strSQL & "" & strTemp & ","
            '是否临时安排_In 临床出诊安排.是否临时安排%Type := 0
            strSQL = strSQL & "" & IIf(mbytFun = Fun_TempPlan, 1, 0) & ")"
            cllPro.Add strSQL
        End If
        
        '保存未保存的出诊安排
        If SavePlanData(obj出诊安排.安排ID, obj出诊安排.未保存出诊安排, obj出诊安排.出诊号源, _
            cllPro, blnRecord, dtCurdate, obj出诊安排.已保存出诊安排) = False Then Exit Function
    Else
        obj出诊安排.安排ID = 0 '标记已无安排了
    End If

    If cllPro.Count = 0 Then
        MsgBox "当前没有需要保存的有效安排！", vbInformation, gstrSysName
        Exit Function
    End If

    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    
    SaveData = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DeletePlan(ByVal obj出诊安排 As 出诊安排, _
    ByRef cllPro As Collection, _
    ByVal blnRecord As Boolean, _
    ByRef blnExistPlan As Boolean, _
    ByVal blnDeletePlan As Boolean) As Boolean
    '功能：删除出诊安排
    '出参：
    '   blnExistPlan 删除后是否还存在已保存安排
    Dim strSQL As String
    Dim ObjItem As 出诊记录集

    Err = 0: On Error GoTo ErrHandler
    blnExistPlan = False
    If obj出诊安排 Is Nothing Then DeletePlan = True: Exit Function
    If obj出诊安排.Count = 0 Then DeletePlan = True: Exit Function
    
    If cllPro Is Nothing Then Set cllPro = New Collection
    For Each ObjItem In obj出诊安排
        If ObjItem.是否删除 Then
            'Zl_临床出诊上班时段_Delete(
            strSQL = "Zl_临床出诊上班时段_Delete("
            '安排id_In   临床出诊限制.安排id%Type,
            strSQL = strSQL & "" & obj出诊安排.安排ID & ","
            'Id_In       临床出诊限制.限制项目%Type := 0,
            strSQL = strSQL & "'" & IIf(mbytPlanType = F_MonthTemplet, FormatApplyToStr(ObjItem.出诊日期), ObjItem.出诊日期) & "',"
            '出诊记录_In Number:=0,
            strSQL = strSQL & "" & IIf(blnRecord, 1, 0) & ","
            '上班时段_In     临床出诊限制.上班时段%Type := Null,
            strSQL = strSQL & "" & "NULL" & ","
            '删除出诊安排_In Number:=0
            strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
            cllPro.Add strSQL
        Else
            blnExistPlan = True
        End If
    Next

    DeletePlan = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwWorkTime_GotFocus()
    stbThis.Panels(2).Text = "说明：字体颜色为蓝色的上班时段表示已在号源中设置了缺省安排。"
End Sub

Private Sub lvwWorkTime_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Item.Selected = False '去掉选中项背景
    stbThis.Panels(2).Text = "说明：字体颜色为蓝色的上班时段表示已在号源中设置了缺省安排。"
End Sub

Private Sub lvwWorkTime_LostFocus()
    stbThis.Panels(2).Text = ""
End Sub

Private Sub optRule_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub picPlan_Resize()
    Err = 0: On Error Resume Next
    With picPlan
        shpPlanLine.Top = .ScaleTop
        shpPlanLine.Width = .ScaleWidth
        shpPlanLine.Left = .ScaleLeft
        shpPlanLine.Height = .ScaleHeight

        vsfPlan.Left = lblSourceTittle.Left
        vsfPlan.Top = lblPlanInfo.Top + lblPlanInfo.Height + 50
        vsfPlan.Width = .ScaleWidth - vsfPlan.Left - 50
        vsfPlan.Height = .ScaleHeight - vsfPlan.Top - 50
    End With
End Sub

Private Sub picSourceAndPlan_Resize()
    Err = 0: On Error Resume Next
    With picSourceAndPlan
        tbPageSourceAndPlan.Left = 0
        tbPageSourceAndPlan.Top = 0
        tbPageSourceAndPlan.Width = .ScaleWidth - tbPageSourceAndPlan.Left
        tbPageSourceAndPlan.Height = .ScaleHeight - tbPageSourceAndPlan.Top
    End With
End Sub

Private Sub txtSignal_GotFocus()
    stbThis.Panels(2).Text = "当前允许输入号码，医生姓名或简码，科室名称或简码进行查找。"
End Sub

Private Sub txtSignal_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim strSQL As String, strWhere As String
    Dim blnCancel As Boolean, rsTemp As ADODB.Recordset
    Dim vRect As RECT, strKey As String

    If KeyAscii <> vbKeyReturn Then Exit Sub
    strText = Trim(txtSignal.Text)
    If Trim(txtSignal.Text) = "" Then Exit Sub

    Err = 0: On Error GoTo ErrHandler
    '根据号码、科室名称、医生姓名进行模糊过滤
    If strText <> "" Then
        strKey = gstrLike & UCase(strText) & "%"
        If IsNumeric(strText) Then   '输入的是全数字
            strWhere = " And a.号码 Like [2]"
        ElseIf zlCommFun.IsCharAlpha(strText) Then  '输入的是全字母
            strWhere = " And (Upper(c.简码) Like [2] Or Upper(d.简码) Like [2])"
        ElseIf zlCommFun.IsCharChinese(strText) Then '是否含有汉字,'含有汉字,肯定是找名称
            strWhere = " And (c.名称 Like [2] Or a.医生姓名 Like [2])"
        Else
            strWhere = " And (a.号码 Like [2] Or c.名称 Like [2] Or Upper(c.简码) Like [2] Or a.医生姓名 Like [2] Or Upper(d.简码) Like [2])"
        End If

        strSQL = "Select a.Id, a.号类, a.号码, b.名称 As 项目, c.名称 As 科室, a.医生姓名" & vbNewLine & _
                " From 临床出诊号源 A, 收费项目目录 B, 部门表 C, 人员表 D" & vbNewLine & _
                " Where a.项目id = b.Id And a.科室id = c.Id And a.医生id = d.Id(+)" & vbNewLine & _
                "       And a.排班方式 In (Select 排班方式 From 临床出诊表 Where ID = [1])" & vbNewLine & _
                "       And Not Exists (Select 1 From 临床出诊安排 Where 号源id = a.Id And 出诊id = [1])" & vbNewLine & _
                "       And Not Exists (Select 1 From 临床出诊安排 P,临床出诊表 Q,临床出诊表 H" & vbNewLine & _
                "                       Where p.出诊ID = q.ID And q.排班方式 = h.排班方式 And q.出诊表名 = h.出诊表名" & vbNewLine & _
                "                             And p.号源id = a.Id And h.id = [1])" & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And Nvl(Nvl(c.站点,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
                strWhere
        vRect = zlControl.GetControlRect(txtSignal.Hwnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "出诊号源", False, "", "请选择出诊号源", False, False, True, _
            vRect.Left, vRect.Top, txtSignal.Height, blnCancel, True, False, mlng出诊ID, strKey, UserInfo.ID, _
            gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
        If blnCancel Then zlControl.TxtSelAll txtSignal: Exit Sub
        If rsTemp Is Nothing Then
            MsgBox "没有找到需要新增的有效出诊号源，请检查！", vbInformation, gstrSysName
            zlControl.TxtSelAll txtSignal
            Exit Sub
        End If
        If rsTemp.EOF Then
            MsgBox "没有找到需要新增的有效出诊号源，请检查！", vbInformation, gstrSysName
            zlControl.TxtSelAll txtSignal
            Exit Sub
        End If
        mlng号源Id = Nvl(rsTemp!ID)
        txtSignal.Text = Nvl(rsTemp!号码)
        zlControl.TxtSelAll txtSignal

        If CheckSignalSource(mbytPlanType = F_FixedRule, mlng号源Id, mlng出诊ID, _
            mobj出诊安排.开始时间, mobj出诊安排.终止时间) = False Then
            mlng号源Id = 0
            Exit Sub
        End If

        If InitData(mobj出诊安排, mlng出诊ID, mlng号源Id, mlng安排ID) = False Then Exit Sub
        Call LoadData
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    '设置控件可用状态
    lvwWorkTime.Enabled = blnEnabled
    If lvwWorkTime.Enabled Then lvwWorkTime.BackColor = lvwWorkTime.BackColor
    picApply.Enabled = blnEnabled
    cldsCalenbarSel.Enabled = blnEnabled
End Sub


Private Function CheckWorkTime(ByVal obj出诊记录集 As 出诊记录集) As Boolean
    '选择时间段
    Dim objListItem As ListItem, i As Integer
    Dim ObjItem As 出诊记录, blnFind As Boolean
    Dim strTmp As String, blnClearAllWorkTime As Boolean
    Dim strSetedTime As String

    On Error GoTo Errhand
    Call LockWindowUpdate(lvwWorkTime.Hwnd) '防止闪烁
    '清除所有上班时段
    If mobj出诊安排.所有上班时段 Is Nothing Then
        blnClearAllWorkTime = True
    ElseIf mobj出诊安排.所有上班时段.Count = 0 Then
        blnClearAllWorkTime = True
    End If
    
    If blnClearAllWorkTime Then
        lvwWorkTime.ListItems.Clear
    Else
        '取消所有的选择
        For i = 1 To lvwWorkTime.ListItems.Count
            lvwWorkTime.ListItems(i).Checked = False
        Next
    End If
    If obj出诊记录集 Is Nothing Then Exit Function
    lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width - 600
    For Each ObjItem In obj出诊记录集
        blnFind = False
        If ObjItem.时间段 <> "" Then
            For i = 1 To lvwWorkTime.ListItems.Count
                If InStr(";" & strSetedTime & ";", ";" & lvwWorkTime.ListItems(i).Tag & ";") = 0 Then
                    '清除停诊状态
                    strTmp = lvwWorkTime.ListItems(i).Text
                    lvwWorkTime.ListItems(i).Text = Left(strTmp, IIf(InStr(strTmp, ")(") = 0, Len(strTmp), InStr(strTmp, ")(")))
                End If
                
                If ObjItem.时间段 = lvwWorkTime.ListItems(i).Tag Then
                    lvwWorkTime.ListItems(i).Checked = True
                    '显示停诊状态
                    lvwWorkTime.ListItems(i).Text = lvwWorkTime.ListItems(i).Text & _
                        Decode(CheckRecordStopVisit(ObjItem), 1, "(部分停诊)", 2, "(已停诊)", "")
                    '103078,调整首列宽度
                    If InStr(lvwWorkTime.ListItems(i).Text, "(部分停诊)") > 0 Then
                        lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width + 200
                    ElseIf InStr(lvwWorkTime.ListItems(i).Text, "(已停诊)") > 0 Then
                        lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width
                    End If
                    blnFind = True
                    strSetedTime = strSetedTime & ";" & lvwWorkTime.ListItems(i).Tag
                    Exit For
                End If
            Next
            If blnFind = False Then
                Set objListItem = lvwWorkTime.ListItems.Add(, "K" & ObjItem.时间段, ObjItem.时间段 & _
                    "(" & Format(ObjItem.开始时间, "hh:mm") & "-" & Format(ObjItem.终止时间, "hh:mm") & ")")
                '显示停诊状态
                objListItem.Text = objListItem.Text & Decode(CheckRecordStopVisit(ObjItem), 1, "(部分停诊)", 2, "(已停诊)", "")
                objListItem.SubItems(1) = ObjItem.开始时间
                objListItem.SubItems(2) = ObjItem.终止时间
                objListItem.Tag = ObjItem.时间段
                objListItem.Checked = True
                '103078,调整首列宽度
                If InStr(objListItem.Text, "(部分停诊)") > 0 Then
                    lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width + 200
                ElseIf InStr(objListItem.Text, "(已停诊)") > 0 Then
                    lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width
                End If
                strSetedTime = strSetedTime & ";" & objListItem.Tag
                '用颜色区分是否号源中设置的时间段
                If mobj出诊安排.号源安排.Exits("K" & ObjItem.时间段) Then
                    objListItem.ForeColor = vbBlue
                End If
            End If
        End If
    Next
    If Not lvwWorkTime.SelectedItem Is Nothing Then
        lvwWorkTime.SelectedItem.Selected = False '去掉选中项背景
    End If
    lvwWorkTime.View = lvwList
    lvwWorkTime.View = lvwReport
    Call LockWindowUpdate(0)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckRecordStopVisit(ByVal obj出诊记录 As 出诊记录) As Integer
    '判断出诊记录的停诊情况
    '入参：
    '   obj出诊记录 - 出诊记录
    '出参：
    '   停诊类型：0-没有停诊，1-部分时段停诊，2-全部时段停诊
    Err = 0: On Error GoTo Errhand
    If obj出诊记录 Is Nothing Then Exit Function
    If obj出诊记录.时间段 = "" _
        Or obj出诊记录.停诊开始时间 = "" Or obj出诊记录.停诊终止时间 = "" Then Exit Function

    If DateDiff("s", obj出诊记录.开始时间, obj出诊记录.停诊开始时间) = 0 _
        And DateDiff("s", obj出诊记录.终止时间, obj出诊记录.停诊终止时间) = 0 Then
        CheckRecordStopVisit = 2
    Else
        CheckRecordStopVisit = 1
    End If
    Exit Function
Errhand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Function CheckPlanIsStopOrUsed(ByVal lng记录ID As Long) As Boolean
    '检查出诊记录是否已停诊或已被使用
    Err = 0: On Error GoTo Errhand
    If lng记录ID = 0 Then Exit Function
    If mrsVisitedRecordByDate Is Nothing Then Exit Function
    
    mrsVisitedRecordByDate.Filter = "ID=" & lng记录ID
    If mrsVisitedRecordByDate.EOF Then Exit Function
    
    If Val(Nvl(mrsVisitedRecordByDate!已使用)) = 1 _
        Or Val(Nvl(mrsVisitedRecordByDate!已停诊)) = 1 Then
        CheckPlanIsStopOrUsed = True: Exit Function
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub btnLeft_Click()
    Dim intIndex As Integer
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String

    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then Exit Sub

    strOldDate = mstrCurDay
    If IsDate(strOldDate) Then
        strNewDate = Format(DateAdd("d", -1, strOldDate), "yyyy-mm-dd")
    Else
        If mobj出诊安排.排班方式 = 3 And mobj出诊安排.排班规则 <> 1 Then
            If mobj出诊安排.排班规则 = 6 Then
                If Val(strOldDate) - 1 > 0 Then strNewDate = Val(strOldDate) - 1 & "日"
            End If
        Else '星期
            intIndex = GetWeekIndex(strOldDate)
            If intIndex - 1 >= 0 And intIndex - 1 <= 6 Then
                strNewDate = GetWeekName(intIndex - 1)
            End If
        End If
    End If
    Call cldsCalenbarSel_SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then Exit Sub

    Call ChangeCurPlan(mobj出诊安排, strNewDate)
    Call cldsCalenbarSel_SelectedChanged(strOldDate, strNewDate)

    cldsCalenbarSel.LoadData mobj出诊安排
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub btnRight_Click()
    Dim intIndex As Integer
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String

    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then Exit Sub

    strOldDate = mstrCurDay
    If IsDate(strOldDate) Then
        strNewDate = Format(DateAdd("d", 1, strOldDate), "yyyy-mm-dd")
    Else
        If mobj出诊安排.排班方式 = 3 And mobj出诊安排.排班规则 <> 1 Then
            If mobj出诊安排.排班规则 = 6 Then
                If Val(strOldDate) + 1 <= 31 Then strNewDate = Val(strOldDate) + 1 & "日"
            End If
        Else '星期
            intIndex = GetWeekIndex(strOldDate)
            If intIndex + 1 >= 0 And intIndex + 1 <= 6 Then
                strNewDate = GetWeekName(intIndex + 1)
            End If
        End If
    End If
    Call cldsCalenbarSel_SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then Exit Sub

    Call ChangeCurPlan(mobj出诊安排, strNewDate)
    Call cldsCalenbarSel_SelectedChanged(strOldDate, strNewDate)

    cldsCalenbarSel.LoadData mobj出诊安排
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnabled()
    '设置上一个、下一个按钮可用状态
    Dim intIndex As Integer

    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then
        btnLeft.Enabled = False
        btnRight.Enabled = False
    ElseIf mobj出诊安排.临时出诊 Then
        btnLeft.Enabled = False
        btnRight.Enabled = False
    Else
        If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj出诊安排.排班规则 <> 1 Then
            '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
            btnLeft.Enabled = False
            btnRight.Enabled = False
            '多选时不可切换，单选时可切换
            If mobj出诊安排.排班规则 = 6 Then
                If mobj出诊安排.Count = 1 Then
                    btnLeft.Enabled = Val(mstrCurDay) > 1
                    btnRight.Enabled = Val(mstrCurDay) < 31
                End If
            End If
        Else
            btnLeft.Enabled = True
            btnRight.Enabled = True
            If IsDate(mstrCurDay) Then '日期
                If DateDiff("d", mobj出诊安排.开始时间, mstrCurDay) <= 0 Then
                    btnLeft.Enabled = False
                End If
                If DateDiff("d", mobj出诊安排.终止时间, mstrCurDay) >= 0 Then
                    btnRight.Enabled = False
                End If
            Else '星期
                intIndex = GetWeekIndex(mstrCurDay)
                If intIndex <= 0 Then
                    btnLeft.Enabled = False
                End If
                If intIndex >= 6 Or intIndex < 0 Then
                    btnRight.Enabled = False
                End If
            End If
        End If
    End If

    Set btnLeft.Picture = img11.ListImages(IIf(btnLeft.Enabled, 1, 3)).Picture
    Set btnRight.Picture = img11.ListImages(IIf(btnRight.Enabled, 2, 4)).Picture
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Get当前出诊安排(obj出诊安排 As 出诊安排, Optional ByVal blnApply As Boolean = True)
    '获取当前出诊安排
    Dim obj出诊记录集 As 出诊记录集
    Dim ObjItem As 出诊记录集, objRecord As 出诊记录
    Dim strApplyTo As String, varApplyTo As Variant
    Dim i As Integer, blnAdd As Boolean
    Dim obj应用出诊记录集 As 出诊记录集
    Dim blnFindCur As Boolean '是否应用于当前选择项目

    On Error GoTo Errhand
    Set obj出诊记录集 = CPDPages.Get出诊记录集
    If obj出诊记录集 Is Nothing Then Exit Sub

    '可能有多个出诊安排（规则-制定日期），将当前出诊安排应用于所有出诊安排
    For Each ObjItem In obj出诊安排
        '移除后再新增
        ObjItem.RemoveAll
        ObjItem.是否修改 = ObjItem.是否修改 Or obj出诊记录集.是否修改
        For Each objRecord In obj出诊记录集
            ObjItem.AddItem objRecord.Clone, "K" & objRecord.时间段
        Next
    Next

    If blnApply = False Then Exit Sub
    '应用于其它日期
    blnFindCur = False
    varApplyTo = Split(GetApplyToStr, ",")
    For i = 0 To UBound(varApplyTo)
        blnAdd = True
'        If IsDate(varApplyTo(i)) Then
'            '小于当前日期的不允许调整
'            If DateDiff("d", varApplyTo(i), mdtToday) > 0 Then blnAdd = False
'        End If
        If blnAdd Then
            If IsDate(varApplyTo(i)) Then
                If GetPlanKey(mstrCurDay) = GetPlanKey(varApplyTo(i)) Then blnFindCur = True
            Else
                blnFindCur = True
            End If
            Set obj应用出诊记录集 = New 出诊记录集
            With obj应用出诊记录集
                .出诊日期 = IIf(IsDate(varApplyTo(i)), Format(varApplyTo(i), "yyyy-mm-dd"), varApplyTo(i))
                .是否修改 = True '新增的肯定是修改
                For Each objRecord In obj出诊记录集
                    If GetPlanKey(mstrCurDay) <> GetPlanKey(varApplyTo(i)) Then
                        '被应用于的其它日期需要产生新的记录ID
                        objRecord.记录ID = 0
                        objRecord.是否固定 = False
                        objRecord.是否修改 = True
                    End If
                    .AddItem objRecord.Clone, "K" & objRecord.时间段
                Next
            End With
            If obj出诊安排.Exits(GetPlanKey(varApplyTo(i))) = False Then
                obj出诊安排.AddItem obj应用出诊记录集, GetPlanKey(varApplyTo(i))
            End If
        End If
    Next

    '当前选择日期不在应用于中，则移除
    If UBound(varApplyTo) > -1 And blnFindCur = False Then
        If obj出诊安排.Exits(GetPlanKey(mstrCurDay)) Then
            obj出诊安排.Remove GetPlanKey(mstrCurDay)
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pancel_Index.Pan_日历
        Item.Handle = picDateList.Hwnd
    Case Pancel_Index.Pan_时间段
        Item.Handle = picWorkTimeList.Hwnd
    Case Pancel_Index.Pan_号源
        If mbytPlanType = F_FixedRule Then
            Item.Handle = picSourceAndPlan.Hwnd
        Else
            Item.Handle = picSouceList.Hwnd
        End If
    Case Pancel_Index.Pan_详情
        Item.Handle = picDetailedList.Hwnd
    End Select
End Sub

Private Sub lvwWorkTime_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, obj出诊安排 As 出诊安排
    Dim blnChecked As Boolean, ObjItem As 出诊安排, objTemp As 出诊安排
    Dim lngFindIndex As Long, objItemTemp As 出诊记录
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim dtStart As Date, dtEnd As Date
    Dim dtStopStart As Date, dtStopEnd As Date
    Dim blnNotCheck As Boolean
    Dim byt缺省分诊方式 As Byte, obj缺省分诊诊室 As 分诊诊室集

    On Error GoTo Errhand
    blnChecked = Item.Checked
    Item.Checked = Not blnChecked
    If mobj出诊安排.Count = 0 Then
        MsgBox IIf(cldsCalenbarSel.ShowStyle = Show_Plan_Rule, "出诊规则未选择！", "出诊日期未选择！"), vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If blnChecked Then
        '限制上班时间段，不能有交叉，注意排除停诊的时间
        dtStart = Item.SubItems(1): dtEnd = Item.SubItems(2)
        If IsDate(mstrCurDay) Then
            dtStart = Format(mstrCurDay, "yyyy-mm-dd ") & Format(dtStart, "hh:mm:ss")
            dtEnd = Format(mstrCurDay, "yyyy-mm-dd ") & Format(dtEnd, "hh:mm:ss")
        End If
        dtEnd = GetWorkTrueDate(dtStart, dtEnd)
        For i = 1 To lvwWorkTime.ListItems.Count
            blnNotCheck = False
            If lvwWorkTime.ListItems(i).Checked Then
                dtCurStart = lvwWorkTime.ListItems(i).SubItems(1): dtCurEnd = lvwWorkTime.ListItems(i).SubItems(2)
                If IsDate(mstrCurDay) Then
                    dtCurStart = Format(mstrCurDay, "yyyy-mm-dd ") & Format(dtCurStart, "hh:mm:ss")
                    dtCurEnd = Format(mstrCurDay, "yyyy-mm-dd ") & Format(dtCurEnd, "hh:mm:ss")
                End If
                dtCurEnd = GetWorkTrueDate(dtCurStart, dtCurEnd)

                '排除停诊的时间范围
                If mobj出诊安排.已保存出诊安排.Exits(GetPlanKey(mstrCurDay)) Then
                    If mobj出诊安排.已保存出诊安排(GetPlanKey(mstrCurDay)).Exits("K" & lvwWorkTime.ListItems(i).Tag) Then
                        With mobj出诊安排.已保存出诊安排(GetPlanKey(mstrCurDay))("K" & lvwWorkTime.ListItems(i).Tag)
                            If .停诊开始时间 <> "" Then
                                dtStopStart = Format(dtCurStart, "yyyy-mm-dd ") & Format(.停诊开始时间, "hh:mm:ss")
                                dtStopEnd = Format(dtCurStart, "yyyy-mm-dd ") & Format(.停诊终止时间, "hh:mm:ss")
                                dtStopEnd = GetWorkTrueDate(dtStopStart, dtStopEnd)
                                If DateDiff("n", dtCurStart, dtStopStart) <= 0 And DateDiff("n", dtCurEnd, dtStopEnd) >= 0 Then
                                    '全部停诊不用检查
                                    blnNotCheck = True
                                Else
                                    '部分停诊，取出未停诊的时间范围
                                    If DateDiff("n", dtCurStart, dtStopStart) = 0 Then '1.停诊前部分[dtStopEnd,dtCurEnd]
                                        dtCurStart = dtStopEnd
                                    ElseIf DateDiff("n", dtCurEnd, dtStopEnd) = 0 Then '2.停诊后部分[dtCurStart,dtStopStart]
                                        dtCurEnd = dtStopStart
                                    Else '3.停诊中间部分
                                        '前,先检查[dtCurStart,dtStopStart]
                                        If CheckTimeBucketIsCross(dtCurStart, dtStopStart, dtStart, dtEnd) Then
                                            MsgBox "当前上班时段的时间范围与已选择的有效上班时段【" & lvwWorkTime.ListItems(i).Tag & "】的时间范围有交叉，不能同时选择！", vbInformation + vbOKOnly, gstrSysName
                                            Exit Sub
                                        End If
                                        '后[dtStopEnd,dtCurEnd]
                                        dtCurStart = dtStopEnd
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
                
                
                If blnNotCheck = False And CheckTimeBucketIsCross(dtCurStart, dtCurEnd, dtStart, dtEnd) Then
                    MsgBox "当前上班时段的时间范围与已选择的有效上班时段【" & lvwWorkTime.ListItems(i).Tag & "】的时间范围有交叉，不能同时选择！", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        Next

        '当前日期+选择时间段开始时间必须大于当前时间
        If IsDate(mstrCurDay) And (mbytFun = Fun_AddSignalSourcePlan Or mbytFun = Fun_TempPlanRecord _
            Or mbytFun = Fun_UpdatePlan) Then
            If DateDiff("s", dtEnd, zlDatabase.Currentdate) > 0 Then
                MsgBox "当前上班时段的终止时间小于了当前时间，不能安排出诊！", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If

            '检查上班时间内是否可出诊
            If CheckDepend(1, mstrCurDay, dtStart, dtEnd) = False Then
                Exit Sub
            End If
        End If
    Else
        '临时出诊时已有时段不能删除
        If mobj出诊安排.临时出诊 Or mbytFun = Fun_TempPlan Then
            If mobj出诊安排(1).Exits("K" & Item.Tag) Then
                If mobj出诊安排(1)("K" & Item.Tag).是否固定 Then
                    If mbytFun = Fun_TempPlan Then
                        MsgBox "当前上班时段在临时安排中必须包含，不能取消！", vbInformation, gstrSysName
                    ElseIf mbytFun = Fun_TempPlanRecord Then
                        If mobj出诊安排(1)("K" & Item.Tag).是否临时出诊 Then
                            MsgBox "临时出诊时只能新增上班时段安排，若需要取消该上班时段安排，请使用调整安排功能操作！", vbInformation, gstrSysName
                        Else
                            MsgBox "临时出诊时只能新增上班时段安排，若需要取消该上班时段安排，请使用停诊功能操作！", vbInformation, gstrSysName
                        End If
                    ElseIf mbytFun = Fun_UpdatePlan Then
                        MsgBox "当前上班时段安排已停诊或已存在预约挂号记录，不允许调整！", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    If mbytFun = Fun_UpdatePlan And mobj出诊安排(1)("K" & Item.Tag).是否临时出诊 = False Then
                        MsgBox "调整安排时不能取消非临时出诊的上班时段安排，若需要取消该上班时段安排，请使用停诊功能操作！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If

        '删除所有时段表示删除安排
        If mobj出诊安排.已保存出诊安排.Exits(GetPlanKey(mobj出诊安排(1).出诊日期)) _
            Or mobj出诊安排.未保存出诊安排.Exits(GetPlanKey(mobj出诊安排(1).出诊日期)) Then
            If IsClearAll(Item.index) Then
                If MsgBox("你确定要删除 " & IIf(mbytPlanType = F_MonthTemplet, FormatApplyToStr(mobj出诊安排(1).出诊日期), mobj出诊安排(1).出诊日期) & " 的出诊安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    Item.Checked = blnChecked

    Get当前出诊安排 mobj出诊安排, False

    If Item.Checked Then
        '取已保存中的
        If mobj出诊安排.已保存出诊安排.Exits("K" & Format(mobj出诊安排(1).出诊日期, "yyyy-mm-dd")) Then
            If mobj出诊安排.已保存出诊安排("K" & Format(mobj出诊安排(1).出诊日期, "yyyy-mm-dd")).Exits("K" & Item.Tag) Then
                Set obj出诊记录 = mobj出诊安排.已保存出诊安排("K" & Format(mobj出诊安排(1).出诊日期, "yyyy-mm-dd"))("K" & Item.Tag).Clone
            End If
        End If

        '取号源设置了的
        If obj出诊记录 Is Nothing And mobj出诊安排.号源安排.Exits("K" & Item.Tag) Then
            Set obj出诊记录 = mobj出诊安排.号源安排("K" & Item.Tag).Clone
            obj出诊记录.记录ID = 0 '需要产生新的记录ID
        End If

        If obj出诊记录 Is Nothing Then
            Set obj出诊记录 = New 出诊记录
            With mobj出诊安排.出诊号源
                obj出诊记录.时间段 = Item.Tag
                If mobj出诊安排.所有上班时段.Exits("K" & obj出诊记录.时间段) Then
                    Set obj出诊记录.上班时段 = mobj出诊安排.所有上班时段("K" & obj出诊记录.时间段)
                Else
                    Set obj出诊记录.上班时段 = New 上班时段
                End If

                Set obj出诊记录.安排门诊诊室集 = New 分诊诊室集
                obj出诊记录.安排门诊诊室集.医生姓名 = mobj出诊安排.出诊号源.医生姓名
                '缺省出诊诊室
                If GetDefaultRoom(mobj出诊安排, byt缺省分诊方式, obj缺省分诊诊室) Then
                    obj出诊记录.安排门诊诊室集.分诊方式 = byt缺省分诊方式
                    Set obj出诊记录.安排门诊诊室集 = obj缺省分诊诊室
                End If

                Set obj出诊记录.号序信息集 = New 号序信息集
                obj出诊记录.号序信息集.时间段 = obj出诊记录.时间段
                obj出诊记录.号序信息集.出诊频次 = .出诊频次
            End With
            obj出诊记录.记录ID = 0 '需要产生新的记录ID
        End If
        obj出诊记录.是否修改 = True '新增了肯定是修改
        mobj出诊安排(1).是否修改 = True
        mobj出诊安排(1).AddItem obj出诊记录, "K" & obj出诊记录.时间段
    Else
        Set obj出诊记录集 = mobj出诊安排(1)
        For i = 1 To obj出诊记录集.Count
            If obj出诊记录集(i).时间段 = Item.Tag Then
                obj出诊记录集.是否修改 = True '删除了肯定是修改
                obj出诊记录集.Remove i: Exit For
            End If
        Next
    End If
    CPDPages.LoadData mobj出诊安排(1), mobj出诊安排.所有分诊诊室, mobj出诊安排.所有合作单位
    LoadPlanToGrid mobj出诊安排, 0
    stbThis.Panels(2).Text = "说明：字体颜色为蓝色的上班时段表示已在号源中设置了缺省安排。"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetDefaultRoom(ByVal obj出诊安排 As 出诊安排, _
    ByRef byt分诊方式 As Byte, ByRef obj分诊诊室 As 分诊诊室集) As Boolean
    '获取缺省分诊方式及分诊诊室
    
    If obj出诊安排 Is Nothing Then Exit Function
    If obj出诊安排.Count > 0 Then
        If obj出诊安排(1).Count > 0 Then
            byt分诊方式 = obj出诊安排(1)(1).安排门诊诊室集.分诊方式
            Set obj分诊诊室 = obj出诊安排(1)(1).安排门诊诊室集.Clone
            GetDefaultRoom = True: Exit Function
        End If
    End If
    If Not obj出诊安排.未保存出诊安排 Is Nothing Then
        If obj出诊安排.未保存出诊安排.Count > 0 Then
            If obj出诊安排.未保存出诊安排(1).Count > 0 Then
                byt分诊方式 = obj出诊安排.未保存出诊安排(1)(1).安排门诊诊室集.分诊方式
                Set obj分诊诊室 = obj出诊安排.未保存出诊安排(1)(1).安排门诊诊室集.Clone
                GetDefaultRoom = True: Exit Function
            End If
        End If
    End If
    If Not obj出诊安排.已保存出诊安排 Is Nothing Then
        If obj出诊安排.已保存出诊安排.Count > 0 Then
            If obj出诊安排.已保存出诊安排(1).Count > 0 Then
                byt分诊方式 = obj出诊安排.已保存出诊安排(1)(1).安排门诊诊室集.分诊方式
                Set obj分诊诊室 = obj出诊安排.已保存出诊安排(1)(1).安排门诊诊室集.Clone
                GetDefaultRoom = True: Exit Function
            End If
        End If
    End If
    If Not obj出诊安排.号源安排 Is Nothing Then
        If obj出诊安排.号源安排.Count > 0 Then
            byt分诊方式 = obj出诊安排.号源安排(1).安排门诊诊室集.分诊方式
            Set obj分诊诊室 = obj出诊安排.号源安排(1).安排门诊诊室集.Clone
            GetDefaultRoom = True: Exit Function
        End If
    End If
End Function

Private Sub picDateList_Resize()
    Err = 0: On Error Resume Next
    With picDateList
        shpItemSel.Top = .ScaleTop
        shpItemSel.Width = .ScaleWidth
        shpItemSel.Left = .ScaleLeft
        shpItemSel.Height = .ScaleHeight

        cldsCalenbarSel.Left = .ScaleLeft + 50
        cldsCalenbarSel.Top = .ScaleTop + 50
        cldsCalenbarSel.Width = .ScaleWidth - 100
        cldsCalenbarSel.Height = .ScaleHeight - 100
    End With
End Sub

Private Sub picDetailedList_Resize()
    Err = 0: On Error Resume Next
    With picDetailedList
        shpDetailedList.Top = .ScaleTop
        shpDetailedList.Width = .ScaleWidth
        shpDetailedList.Left = .ScaleLeft
        shpDetailedList.Height = .ScaleHeight

        btnLeft.Left = .ScaleLeft + 30
        btnRight.Left = btnLeft.Left + btnLeft.Width + 20
        lblTittle.Left = btnRight.Left + btnRight.Width + 50

        picApply.Left = 4500
        picApply.Top = lblTittle.Top
        picApply.Width = .ScaleWidth - picApply.Left - 20

        CPDPages.Left = .ScaleLeft + 30
        'picApply.Tag = "1"表示隐藏
        CPDPages.Top = .ScaleTop + IIf(picApply.Tag = "", picApply.Top + picApply.Height, lblTittle.Top + lblTittle.Height) + 50
        CPDPages.Width = .ScaleWidth - 40
        CPDPages.Height = .ScaleHeight - CPDPages.Top - 10
    End With
End Sub

Private Sub picWorkTimeList_Resize()
    Err = 0: On Error Resume Next
    With picWorkTimeList
        shpWorkLine.Top = .ScaleTop
        shpWorkLine.Width = .ScaleWidth
        shpWorkLine.Left = .ScaleLeft
        shpWorkLine.Height = .ScaleHeight

        lvwWorkTime.Left = lblCalendbarTittle.Left
        lvwWorkTime.Top = lblCalendbarTittle.Top + lblCalendbarTittle.Height + 50
        lvwWorkTime.Width = .ScaleWidth - lvwWorkTime.Left - 50
        lvwWorkTime.Height = .ScaleHeight - lvwWorkTime.Top - 50
    End With
End Sub

Private Sub picSouceList_Resize()
    Err = 0: On Error Resume Next
    With picSouceList
        shpSourceLine.Top = .ScaleTop
        shpSourceLine.Width = .ScaleWidth
        shpSourceLine.Left = .ScaleLeft
        shpSourceLine.Height = .ScaleHeight

        SourceInfor.Left = lblSourceTittle.Left
        SourceInfor.Top = lblSourceTittle.Top + lblSourceTittle.Height + 50
        SourceInfor.Width = .ScaleWidth - SourceInfor.Left - 50
        SourceInfor.Height = .ScaleHeight - SourceInfor.Top - 50
    End With
End Sub

Private Sub optRule_Click(index As Integer)
    '功能:设置应用于的显示
    Dim i As Integer
    Dim dtCur As Date

    Err = 0: On Error GoTo ErrHandler:
    fraLoopSkip.Visible = False
    picApplyWeek.Visible = False
    If index = 3 Then '按星期
        picApplyWeek.Visible = True
        picApplyWeek.Top = picApplyRule.Top + picApplyRule.Height
        picApply.Height = picApplyWeek.Top + picApplyWeek.Height
        mblnNotClick = True
        For i = chkWeek.LBound To chkWeek.UBound
            chkWeek(i).Value = vbUnchecked
        Next
        mblnNotClick = False
    Else
        picApply.Height = picApplyRule.Top + picApplyRule.Height
    End If
    fraLoopSkip.Visible = index = 4

    '加载轮询开始日期
    If index = 4 Then
        If cboDays.ListCount = 0 Then
            If Not mobj出诊安排 Is Nothing Then
                '肯定是排班方式为1，月排班
                dtCur = mobj出诊安排.开始时间
                Do While True
                    cboDays.AddItem Format(dtCur, "yyyy-mm-dd")
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
                Loop
                If cboDays.ListCount > 0 Then cboDays.ListIndex = 0
            End If
        Else
            cboDays.ListIndex = 0
        End If
    End If
    picDetailedList_Resize
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetTitleText()
    '设置显示标题，以及应用选项
    Dim strTemp As String, i As Integer
    Dim str停诊原因 As String

    On Error GoTo Errhand
    lblTittle.Caption = "无安排"

    picApply.Visible = False: picApply.Tag = "1" '隐藏
    If mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdateUnit _
        Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel _
        Or mbytFun = Fun_UpdatePlan Then
        '没有应用于
    Else
        Select Case mbytPlanType
        Case F_Templet '模板规则
            If mobj出诊安排.排班规则 = 1 Then
                picApply.Visible = mbytFun <> Fun_View
                picApply.Tag = IIf(mbytFun = Fun_View, "1", "")
                picApplyRule.Visible = False
                picApplyWeek.Visible = mbytFun <> Fun_View
                picApplyWeek.Top = picApplyRule.Top + 60
                picApply.Height = picApplyWeek.Top + picApplyWeek.Height
                mblnNotClick = True
                For i = chkWeek.LBound To chkWeek.UBound
                    If chkWeek(i).Caption = mstrCurDay Then
                        chkWeek(i).Tag = "1"
                        chkWeek(i).Value = vbChecked
                    Else
                        chkWeek(i).Tag = ""
                        chkWeek(i).Value = vbUnchecked
                    End If
                Next
                mblnNotClick = False
            End If
        Case F_FixedRule '固定安排
            picApply.Visible = mbytFun <> Fun_View
            picApply.Tag = IIf(mbytFun = Fun_View, "1", "")
            picApplyRule.Visible = False
            picApplyWeek.Visible = mbytFun <> Fun_View
            picApplyWeek.Top = picApplyRule.Top + 60
            picApply.Height = picApplyWeek.Top + picApplyWeek.Height
            mblnNotClick = True
            For i = chkWeek.LBound To chkWeek.UBound
                If chkWeek(i).Caption = mstrCurDay Then
                    chkWeek(i).Tag = "1"
                    chkWeek(i).Value = vbChecked
                Else
                    chkWeek(i).Tag = ""
                    chkWeek(i).Value = vbUnchecked
                End If
            Next
            mblnNotClick = False
        Case F_MonthPlan, F_WeekPlan, F_MonthTemplet '月安排/周安排
            picApply.Visible = mbytFun <> Fun_View
            picApply.Tag = IIf(mbytFun = Fun_View, "1", "")
            If optRule(0).Value Then optRule(1).Value = True
            optRule(0).Value = True
            If mbytPlanType = F_WeekPlan Then '按周,不能使用轮循
                optRule(4).Visible = False
                fraLoopSkip.Visible = False
            End If
        End Select
    End If
    Call SetButtonEnabled
    Call picDetailedList_Resize

    If mstrCurDay <> "" And mbytPlanType = F_Templet Then
        '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
        Select Case mobj出诊安排.排班规则
        Case 2
            strTemp = "按单日出诊"
        Case 3
            strTemp = "按双日出诊"
        Case 4, 5
            strTemp = "按" & Val(mstrCurDay) & "天轮循"
        Case 6
            strTemp = ""
            If mobj出诊安排.已保存出诊安排.排班规则 = 6 Then
                For i = 1 To mobj出诊安排.已保存出诊安排.Count
                    strTemp = strTemp & "," & mobj出诊安排.已保存出诊安排(i).出诊日期
                Next
            End If
            For i = 1 To mobj出诊安排.未保存出诊安排.Count
                strTemp = strTemp & "," & mobj出诊安排.未保存出诊安排(i).出诊日期
            Next
            For i = 1 To mobj出诊安排.Count
                strTemp = strTemp & "," & mobj出诊安排(i).出诊日期
            Next
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            strTemp = "按每月的" & ZlNumStrSort(strTemp, True) & "日固定出诊"
        Case Else
            strTemp = mstrCurDay
        End Select
    ElseIf mstrCurDay <> "" Then
        strTemp = mstrCurDay
    End If

    If strTemp = "" Then strTemp = "无安排"
    If IsDate(strTemp) Then
        If mbytPlanType = F_MonthTemplet Then
            lblTittle.Caption = FormatApplyToStr(strTemp)
        Else
            str停诊原因 = CurDayIsNotVisit(strTemp)
            If str停诊原因 = "" Then
                lblTittle.Caption = Format(strTemp, "yyyy-mm-dd")
            Else
                lblTittle.Caption = Format(strTemp, "yyyy-mm-dd") & "(" & str停诊原因 & ")"
            End If
        End If
    Else
        lblTittle.Caption = strTemp
    End If

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function IsClearAll(Optional ByVal intIndex As Integer = -1) As Boolean
    '是否清空了所有上班时间段
    'intIndex 排除该项
    Dim blnSelected As Boolean
    Dim i As Integer

    Err = 0: On Error GoTo ErrHandler
    For i = 1 To lvwWorkTime.ListItems.Count
        If lvwWorkTime.ListItems(i).Checked Then
            If i <> intIndex Then
                blnSelected = True: Exit For
            End If
        End If
    Next
    IsClearAll = blnSelected = False
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsValied() As Boolean
    '检查数据
    Dim blnSelected As Boolean
    Dim i As Integer, blnFind As Boolean
    Dim obj出诊记录集 As 出诊记录集
    Dim dtStart As Date, dtEnd As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset

    Err = 0: On Error GoTo ErrHandler
    If mbytPlanType = F_FixedRule Then
        dtStart = mobj出诊安排.开始时间
        dtEnd = mobj出诊安排.终止时间
        If mbytFun = Fun_AddSignalSourcePlan Or mbytFun = Fun_TempPlan Or mbytFun = Fun_Update Then
            If dtpBegin.Visible And dtpBegin.Enabled And dtpEnd.Enabled Then
                '焦点在时间控件中且调整后，直接点保存，Value的值还没有更改过来
                mblnValiedCanSave = True
                If Me.ActiveControl Is dtpBegin Then
                    dtpEnd.SetFocus: Call dtpBegin_LostFocus
                ElseIf Me.ActiveControl Is dtpEnd Then
                    dtpBegin.SetFocus: Call dtpEnd_LostFocus
                End If
                If mblnValiedCanSave = False Then Exit Function
                mblnValiedCanSave = False
                
                If dtpEnd.Value < zlDatabase.Currentdate Then
                    MsgBox "有效期的终止时间不能小于当前时间。", vbInformation, gstrSysName
                    If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
                    Exit Function
                End If
                If dtpBegin.Value >= dtpEnd.Value Then
                    MsgBox "有效期的开始时间应该小于结束时间。", vbInformation, gstrSysName
                    If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
                    Exit Function
                End If
                
                '有效期的开始日期大于等于当前日期时才检查
                If mdtToday <= CDate(Format(dtpBegin.Value, "yyyy-mm-dd")) _
                    And Format(dtpBegin.Value, "hh:mm:ss") <> "00:00:00" Then
                    If MsgBox("按照当前有效期的设置，该安排在 " & Format(dtpBegin.Value, "yyyy-mm-dd") & _
                        " 不会生效，如果你希望在 " & Format(dtpBegin.Value, "yyyy-mm-dd") & _
                        " 生效请将开始时间调整为 " & Format(dtpBegin.Value, "yyyy-mm-dd 00:00:00") & _
                        "，是否继续按当前设置保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
                        Exit Function
                    End If
                End If
            End If
            dtStart = Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss")
            dtEnd = Format(dtpEnd.Value, "yyyy-mm-dd hh:mm:ss")
        End If
        If mbytFun = Fun_AddSignalSourcePlan Then
            If CheckSignalSource(mbytPlanType = F_FixedRule, _
                mlng号源Id, mlng出诊ID, dtStart, dtEnd, True) = False Then Exit Function
        End If
        mobj出诊安排.开始时间 = Format(dtStart, "yyyy-mm-dd hh:mm:ss")
        mobj出诊安排.终止时间 = Format(dtEnd, "yyyy-mm-dd hh:mm:ss")
        
        If CheckTempFixedPlan(mobj出诊安排, dtStart, dtEnd, , , True) = False Then Exit Function
        
        If mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "所有科室") Then
            If cboFeeItem.ListIndex = -1 Then
                MsgBox "收费项目不能为空！", vbInformation, gstrSysName
                If cboFeeItem.Visible And cboFeeItem.Enabled Then cboFeeItem.SetFocus
                Exit Function
            End If
            
            '科室，医生，收费项目在号源中不能重复
            strSQL = "Select 号码 From 临床出诊号源" & _
                    " Where 科室ID=[1] And Nvl(医生ID,0)=[2] And Nvl(医生姓名,'-')=[3] And 项目ID=[4]" & _
                    "       And Nvl(是否删除,0)=0 And 号码 <> [5]"
            With mobj出诊安排.出诊号源
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查号源是否唯一", .科室ID, .医生ID, _
                    IIf(.医生姓名 = "", "-", .医生姓名), cboFeeItem.ItemData(cboFeeItem.ListIndex), .号码)
                If Not rsTemp.EOF Then
                    MsgBox .科室名称 & " " & IIf(.医生姓名 = "", "", "的医生 " & .医生姓名 & " ") & _
                        "已经存在收费项目为 " & cboFeeItem.Text & " 的号源【" & Nvl(rsTemp!号码) & "】，" & _
                        "您不能对当前号源制定收费项目为 " & cboFeeItem.Text & " 的临时安排！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End With
        
            mobj出诊安排.项目ID = cboFeeItem.ItemData(cboFeeItem.ListIndex)
            mobj出诊安排.项目名称 = cboFeeItem.Text
        End If
    End If
    
    If CPDPages.IsValied() = False Then Exit Function

    Set obj出诊记录集 = CPDPages.Get出诊记录集
    If (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) _
        And obj出诊记录集.是否修改 Then
        If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, CDate(mstrCurDay)) Then
            MsgBox Format(mstrCurDay, "yyyy-mm-dd") & " 已在其它出诊表中设置了出诊安排，不能重复安排！", vbInformation, gstrSysName
            mobj出诊安排(1).RemoveAll
            mobj出诊安排(1).是否修改 = False
            Call LoadDetailData
            Set obj出诊记录集 = Nothing
            Exit Function
        End If
    End If

    If mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan _
        Or mbytPlanType = F_MonthTemplet Then
        If CheckExistRecord(0, Replace(GetApplyToStr(), ",", "|"), mobj出诊安排) Then
            If MsgBox("注意：" & vbCrLf & _
                      "      部分被应用的日期当前已存在出诊安排，应用后这部分安排将会被覆盖！是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If mbytFun = Fun_UpdatePlan Then
        '检查是否被他人停诊或挂号预约
        Set mrsVisitedRecordByDate = GetVisitedRecordByDate(mlng安排ID, mstrCurDay) '重新获取已停诊或已使用的安排
        If Not mrsVisitedRecordByDate Is Nothing Then
            With mrsVisitedRecordByDate
                Do While Not .EOF
                    For i = 1 To obj出诊记录集.Count
                        If Val(Nvl(!ID)) = obj出诊记录集(i).记录ID And obj出诊记录集(i).是否固定 = False Then
                            If CheckPlanIsStopOrUsed(obj出诊记录集(i).记录ID) Then
                                MsgBox "上班时段 " & obj出诊记录集(i).时间段 & " 已停诊或已存在预约挂号记录，" & _
                                    "但当前该上班时段为可修改状态，请退出当前窗口重新进入进行调整！", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    Next
                    .MoveNext
                Loop
            End With
        End If
    End If

    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get出诊安排() As 出诊安排
    On Error GoTo Errhand
    Call Get当前出诊安排(mobj出诊安排)
    Call ChangeCurPlan(mobj出诊安排, mstrCurDay) '将当前安排放入未保存集合中
    Set Get出诊安排 = mobj出诊安排.Clone
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsVisitedOtherTable(ByVal lng出诊ID As Long, ByVal lng号源Id As Long, _
    ByVal dtDate As Date, Optional ByVal blnOtherTable As Boolean = True, _
    Optional ByVal lng安排ID As Long) As Boolean
    '是否其它出诊表已有出诊记录
    '参数：
    '   blnOtherTable 是否是其它出诊表中，否则是当前出诊表中
    '   lng安排ID blnOtherTable为False时传入
    Dim strFilter As String
    
    Err = 0: On Error GoTo Errhand
    If mrsVisitedRecord Is Nothing Then Exit Function
    mrsVisitedRecord.Filter = ""
    If mrsVisitedRecord.RecordCount = 0 Then Exit Function
    
    If blnOtherTable Then
        strFilter = "出诊ID<>" & lng出诊ID
    Else
        strFilter = "出诊ID=" & lng出诊ID & " And 安排ID<>" & lng安排ID
    End If
    strFilter = strFilter & " And 号源ID=" & lng号源Id & " And 出诊日期=#" & Format(dtDate, "yyyy-mm-dd") & "#"
    mrsVisitedRecord.Filter = strFilter
    IsVisitedOtherTable = mrsVisitedRecord.RecordCount > 0
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetApplyToStr() As String
    '获取应用于字符串
    Dim strApplyTo As String, i As Integer, k As Integer
    Dim dtCur As Date, blnFind As Boolean
    Dim varTemp As Variant, strWeekName As String

    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then Exit Function
    If picApply.Tag = "1" Then Exit Function

    '月排班/周排班
    If picApplyRule.Visible Then
        If optRule(1).Value Then '单日
            dtCur = mobj出诊安排.开始时间
            Do While True
                If Day(dtCur) Mod 2 = 1 Then
                    If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                    End If
                End If
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
            Loop
        ElseIf optRule(2).Value Then '双日
            dtCur = mobj出诊安排.开始时间
            Do While True
                If Day(dtCur) Mod 2 = 0 Then
                    If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                    End If
                End If
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
            Loop
        ElseIf optRule(3).Value Then '星期
            dtCur = mobj出诊安排.开始时间
            Do While True
                For i = chkWeek.LBound To chkWeek.UBound
                    If chkWeek(i).Value = vbChecked Then
                        If Weekday(dtCur, vbMonday) = i + 1 Then
                            If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                                strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                            End If
                        End If
                    End If
                Next
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
            Loop
        ElseIf optRule(4).Value Then '轮循,间隔多少天
            If Not (cboDays.ListIndex = -1 Or Val(txtSkip.Text) = 0) Then
                dtCur = CDate(cboDays) '开始时间
                Do While True
                    If IsVisitedOtherTable(mlng出诊ID, mlng号源Id, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                    End If
                    dtCur = DateAdd("d", Val(txtSkip.Text) + 1, dtCur)
                    If DateDiff("d", mobj出诊安排.终止时间, dtCur) > 0 Then Exit Do
                Loop
            End If
        End If
        If strApplyTo <> "" Then strApplyTo = Mid(strApplyTo, 2)
        GetApplyToStr = strApplyTo
        Exit Function
    End If

    '模板的星期规则或固定模板规则
    If picApplyWeek.Visible Then
        For i = chkWeek.LBound To chkWeek.UBound
            If chkWeek(i).Value = vbChecked Then
                strWeekName = GetWeekName(i)
                blnFind = False
                '不能修改的星期不能应用于
                If Not mcllFixedPlan Is Nothing And strWeekName <> mstrCurDay Then
                    'Array(出诊日期,限制项目,上班时段,开始时间,终止时间)
                    For k = 1 To mcllFixedPlan.Count
                        If mcllFixedPlan(k)(1) = strWeekName Then
                            blnFind = True: Exit For
                        End If
                    Next
                    If blnFind = False Then strApplyTo = strApplyTo & "," & strWeekName
                Else
                    strApplyTo = strApplyTo & "," & strWeekName
                End If
            End If
        Next
        If strApplyTo <> "" Then strApplyTo = Mid(strApplyTo, 2)
        GetApplyToStr = strApplyTo
        Exit Function
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CurDayIsNotVisit(ByVal Day As Date) As String
    '当前日期是否节假日或有停止安排
    '返回：停止原因或节假日名称
    Dim i As Integer

    '当前日期是否保存了安排
    With mobj出诊安排
        If Not .停诊记录 Is Nothing Then
            For i = 1 To .停诊记录.Count
                If DateDiff("d", Day, .停诊记录(i).开始时间) <= 0 And DateDiff("d", Day, .停诊记录(i).终止时间) >= 0 Then
                    CurDayIsNotVisit = .停诊记录(i).停诊原因
                    Exit Function
                End If
            Next
        End If
    End With
End Function

Private Sub txtSignal_LostFocus()
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtSkip_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub InitPage()
    '初始化页面控件
    Dim i As Long, ObjItem As TabControlItem
    Dim objUnit As 合作单位控制, lngRow As Long
    Dim intPageCount As Integer
    
    Err = 0: On Error GoTo Errhand
    tbPageSourceAndPlan.InsertItem Pg_号源信息, "号源信息", picSouceList.Hwnd, 0
    tbPageSourceAndPlan.InsertItem Pg_安排预览, "安排预览", picPlan.Hwnd, 0
    
    With tbPageSourceAndPlan
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionBottom
        If mbytFun = Fun_AddSignalSourcePlan Then
            .Item(Pg_号源信息).Selected = True
        Else
            .Item(Pg_安排预览).Selected = True
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPlan(ByVal obj出诊安排 As 出诊安排)
    Err = 0: On Error GoTo Errhand
    If mbytPlanType <> F_FixedRule Then Exit Sub
    
    With vsfPlan
        .Redraw = flexRDNone
        .Cols = 3
        Set .Cell(flexcpPicture, 0, 0, .Rows - 1, 0) = imglist16.ListImages("plan_nothing").Picture
        .Cell(flexcpText, 0, 2, .Rows - 1, .Cols - 1) = ""
        
        If obj出诊安排 Is Nothing Then Exit Sub
        
        If Not obj出诊安排.已保存出诊安排 Is Nothing Then
            LoadPlanToGrid obj出诊安排.已保存出诊安排, 1
        End If
        If Not obj出诊安排.未保存出诊安排 Is Nothing Then
            LoadPlanToGrid obj出诊安排.未保存出诊安排, 0
        End If
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPlanToGrid(ByVal obj出诊安排 As 出诊安排, ByVal bytMode As Byte)
    '加载安排预览
    '入参：
    '   bytMode 0-未保存，1-已保存
    '说明：
    '   图标索引："plan_deleted"-已删除,"plan_saved"-已保存,"plan_nosave"-未保存,"plan_nothing"-无安排
    Dim ObjItem As 出诊记录集, objRecord As 出诊记录
    Dim i As Integer, j As Integer, k As Integer, blnFindNotNull As Boolean
    Dim blnFindItem As Boolean, blnFindRecord As Boolean
    
    Err = 0: On Error GoTo Errhand
    With vsfPlan
        If obj出诊安排 Is Nothing Then Exit Sub
        
        For Each ObjItem In obj出诊安排
            blnFindItem = False
            For i = 0 To .Rows - 1
                If blnFindItem = True Then Exit For
                If .TextMatrix(i, 1) = ObjItem.出诊日期 Then
                    blnFindItem = True
                    .Cell(flexcpText, i, 2, i, .Cols - 1) = "" '先清空
                    Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_nothing").Picture
                    If ObjItem.Count = 0 And .RowData(i) = 1 Then
                        Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_deleted").Picture
                        .RowData(i) = 3 '已删除
                        Exit For
                    End If
                    For k = .Cols - 1 To 3 Step -1
                        blnFindNotNull = False
                        For j = 0 To .Rows - 1
                            If Trim(.TextMatrix(j, k)) <> "" Then blnFindNotNull = True: Exit For
                        Next
                        If blnFindNotNull = False Then
                            .Cols = .Cols - 1
                        Else
                            Exit For
                        End If
                    Next
                    If ObjItem.Count > 0 Then
                        If ObjItem.是否删除 Then
                            Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_deleted").Picture
                            .RowData(i) = 3 '已删除
                        Else
                            If bytMode = 1 Then
                                Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_saved").Picture
                                .RowData(i) = 1 '已保存
                            Else
                                Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_nosave").Picture
                                .RowData(i) = 2 '未保存
                            End If
                        End If
                        .Cell(flexcpPictureAlignment, i, 0) = flexAlignCenterCenter
                    End If
                    For Each objRecord In ObjItem
                        blnFindRecord = False
                        For j = 2 To .Cols - 1
                            If .TextMatrix(i, j) = "" Then blnFindRecord = True: Exit For
                        Next
                        If blnFindRecord = False Then
                            .Cols = .Cols + 1: j = .Cols - 1
                            .ColWidth(j) = 800: .ColAlignment(j) = flexAlignCenterCenter
                        End If
                        .TextMatrix(i, j) = objRecord.时间段 & vbCrLf & _
                            IIf(objRecord.预约控制 = 1, "-", IIf(objRecord.限约数 = 0, IIf(objRecord.限号数 = 0, "∞", objRecord.限号数), objRecord.限约数)) & _
                            "/" & IIf(objRecord.限号数 = 0, "∞", objRecord.限号数)
                    Next
                End If
            Next
        Next
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfplan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTemp As String
    
    On Error Resume Next
    If vsfPlan.MouseCol <> 0 And vsfPlan.MouseCol <> 1 Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    strTemp = vsfPlan.TextMatrix(vsfPlan.MouseRow, 1)
    Select Case vsfPlan.RowData(vsfPlan.MouseRow)
    Case 1 '已保存
        stbThis.Panels(2).Text = strTemp & " 已保存"
    Case 2 '未保存
        stbThis.Panels(2).Text = strTemp & " 未保存"
    Case 3 '已保存
        stbThis.Panels(2).Text = strTemp & " 已删除"
    Case Else
        stbThis.Panels(2).Text = strTemp & " 无安排"
    End Select
End Sub

Private Sub txtSignal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '按了右键菜单快捷键，清除粘贴板内容
        If Clipboard.GetText <> "" Then Clipboard.Clear
    End If
End Sub

Private Sub txtSignal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtSignal.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtSignal.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtSignal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtSignal.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Function CheckSignalSource(ByVal blnFixedPlan As Boolean, _
    ByVal lng号源Id As Long, ByVal lng出诊ID As Long, _
    ByVal d_开始时间 As Date, ByVal d_终止时间 As Date, _
    Optional ByVal blnSaveBefore As Boolean) As Boolean
    '号源有效性检查
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    '判断是否为新增加的号源
    strSQL = "Select 1 From 临床出诊安排 Where 号源id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng号源Id)
    If rsTemp.EOF Then CheckSignalSource = True: Exit Function
    
    '1.月/周排班 转 固定排班
    If blnFixedPlan Then
        If blnSaveBefore Then
            strSQL = "Select Nvl(Max(a.终止时间), To_Date('1900-01-01', 'yyyy-mm-dd')) As 终止时间" & _
                    " From 临床出诊安排 A, 临床出诊表 B" & _
                    " Where a.出诊id = b.Id And b.排班方式 In (1, 2) And a.号源id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng号源Id)
            If CDate(Format(d_开始时间, "yyyy-mm-dd")) < CDate(rsTemp!终止时间) Then
                ShowMsgbox "当前号源在" & Format(rsTemp!终止时间, "yyyy-mm-dd") & "及之前已制定了出诊安排，新安排的有效期不能再包含这段时间！"
                Exit Function
            End If
        End If
        CheckSignalSource = True: Exit Function
    End If
 
    '2.固定排班 转 月/周排班
    '判断在当前出诊表的开始时间之后，固定排班的出诊记录是否被使用
    strSQL = "Select 1" & _
            " From 临床出诊记录 A, 临床出诊安排 C, 临床出诊表 D" & _
            " Where A.安排ID = c.ID And c.出诊ID = d.ID" & _
            "       And d.排班方式 = 0 And A.号源ID = [1] And A.出诊日期 >= [2]" & _
            "       And Exists(Select 1 From 病人挂号记录 Where 出诊记录id = a.Id) And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng号源Id, d_开始时间)
    If Not rsTemp.EOF Then
        ShowMsgbox "当前号源在" & Format(d_开始时间, "yyyy-mm-dd") & "之后存在已被使用的安排，不能将其添加到当前出诊表中！"
        Exit Function
    End If
    
    '删除未被使用的固定安排的出诊记录
    'Zl_临床出诊记录_Delete(
    strSQL = "Zl_临床出诊记录_Delete("
    '  号源id_In   临床出诊记录.号源id%Type,
    strSQL = strSQL & "" & lng号源Id & ","
    '  开始日期_In 临床出诊记录.出诊日期%Type
    strSQL = strSQL & "To_Date('" & Format(d_开始时间, "yyyy-mm-dd") & "','yyyy-mm-dd'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '3.月排班 转 周排班、周排班 转 月排班
    '在当前出诊表的时间范围内不能有出诊记录
    strSQL = "Select a.开始时间, a.终止时间" & _
            " From 临床出诊安排 A, 临床出诊表 B" & _
            " Where a.出诊ID = b.ID And b.排班方式 In(1, 2) And a.号源ID = [1]" & _
            "       And a.开始时间 <= [2] And a.终止时间 >= [3] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng号源Id, d_开始时间, d_终止时间)
    If Not rsTemp.EOF Then
        ShowMsgbox "当前号源在有效期(" & _
            Format(d_开始时间, "yyyy-mm-dd") & "～" & Format(d_终止时间, "yyyy-mm-dd") & _
            ")内已存在出诊安排，不能将该号源添加到当前出诊表。"
        Exit Function
    End If
    
    CheckSignalSource = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

