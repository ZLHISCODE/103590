VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanEdit 
   Caption         =   "���ﰲ������"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15120
   Icon            =   "frmClinicPlanEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "������������"
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
         Caption         =   "�շ���Ŀ"
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
            Name            =   "����"
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
         ToolTipText     =   "��ǰ����������룬ҽ����������룬�������ƻ������в��ҡ�"
         Top             =   30
         Width           =   2025
      End
      Begin VB.Label lblSignal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
            Name            =   "����"
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
         Caption         =   "��Դ��Ϣ"
         BeginProperty Font 
            Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "����Ԥ��"
            BeginProperty Font 
               Name            =   "����"
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
         Caption         =   "��Ч��"
         Height          =   180
         Left            =   90
         TabIndex        =   30
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblTimeRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
            Name            =   "����"
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
            Text            =   "ʱ���"
            Object.Width           =   9596
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��ʼʱ��"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��ֹʱ��"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblCalendbarTittle 
         BackStyle       =   0  'Transparent
         Caption         =   "�ϰ�ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
                  Caption         =   "��ʼ��ѭ����"
                  Height          =   180
                  Left            =   0
                  TabIndex        =   41
                  Top             =   60
                  Width           =   1080
               End
               Begin VB.Label lblLoopSkipDays 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���       ��"
                  Height          =   180
                  Left            =   2445
                  TabIndex        =   42
                  Top             =   60
                  Width           =   1170
               End
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��ѭ"
               Height          =   240
               Index           =   4
               Left            =   3150
               TabIndex        =   33
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               Height          =   240
               Index           =   3
               Left            =   2325
               TabIndex        =   38
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "˫��"
               Height          =   240
               Index           =   2
               Left            =   1530
               TabIndex        =   37
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               Height          =   240
               Index           =   1
               Left            =   735
               TabIndex        =   36
               Top             =   75
               Width           =   735
            End
            Begin VB.OptionButton optRule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��ǰ"
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
               Caption         =   "����"
               Height          =   180
               Index           =   6
               Left            =   5685
               TabIndex        =   40
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               Height          =   180
               Index           =   5
               Left            =   4735
               TabIndex        =   48
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               Height          =   180
               Index           =   4
               Left            =   3788
               TabIndex        =   47
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               Height          =   180
               Index           =   3
               Left            =   2841
               TabIndex        =   46
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               Height          =   180
               Index           =   2
               Left            =   1894
               TabIndex        =   45
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�ܶ�"
               Height          =   180
               Index           =   1
               Left            =   947
               TabIndex        =   44
               Top             =   30
               Width           =   690
            End
            Begin VB.CheckBox chkWeek 
               BackColor       =   &H00FFFFFF&
               Caption         =   "��һ"
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
            Caption         =   "Ӧ����(&Y)"
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
         Caption         =   "���ڶ�"
         BeginProperty Font 
            Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
Option Explicit 'Ҫ���������
Private Enum m_Fun
    F_FixedRule = 0
    F_MonthPlan = 1
    F_WeekPlan = 2
    F_Templet = 3
    F_MonthTemplet = 4
End Enum
Private mbytPlanType As m_Fun '0-�̶��Ű�,1-�����Ű�,2-�����Ű�,3-ģ��,4-���찲�ų������ģ��
Private mbytFun As G_Enum_Fun

Public Enum Pancel_Index
    Pan_���� = 1001
    Pan_ʱ��� = 1002
    Pan_��Դ = 1003
    Pan_���� = 1004
End Enum
Private mblnFirst As Boolean
Private mblnTimeChanged As Boolean, mblnFeeItemChanged As Boolean

Private mobj���ﰲ�� As ���ﰲ��, mlng����ID As Long
Private mlng��ԴId As Long, mlng����ID As Long
Private mstrȱʡ���� As String
Private mlngSavedRecords As Long
Private mblnNotClick As Boolean

Private mstrCurDay As String
Private mdtToday As Date '��ǰ����������
Private mrsVisitedRecord As ADODB.Recordset '�ú�Դ�ѳ����¼
Private mobjͣ���¼�� As ͣ���¼�� '��ǰ��Դ�ĳ����¼��
Private mrsVisitedRecordByDate As ADODB.Recordset '��ǰ��Դĳ�����ڵĳ����¼
Private Type FixedPlanDateRange
    dtStart As Date
    dtEnd As Date
End Type
Private mFixedPlanDateRange As FixedPlanDateRange
Private mstrPrivs As String
Private mcllFixedPlan As Collection  '��ʱ���ţ���¼����ԤԼ�Һż�¼�ĳ����¼��Array(��������,������Ŀ,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��)

Private mblnCheckedByDay As Boolean '�Ƿ��Ѱ�����
Private mblnValiedCanSave As Boolean

Private Enum mPgIndex 'TabPage����
    Pg_��Դ��Ϣ = 0
    Pg_����Ԥ�� = 1
End Enum

Private Enum mMenuID
    M_Signal = 1
    M_ValidTime = 2
    M_FeeItem = 3
    M_Verify = 4
End Enum

Public Function ShowMe(frmParent As Object, ByVal bytPlanType As Byte, ByVal bytFun As G_Enum_Fun, ByVal lng����ID As Long, _
    Optional ByVal lng��ԴId As Long, Optional ByVal lng����ID As Long, _
    Optional ByVal strȱʡ���� As String, Optional ByVal strPrivs As String) As Boolean
    '���ܣ��������
    '��Σ�
    '   bytPlanType 0-�̶��Ű�,1-�����Ű�,2-�����Ű�,3-ģ��,4-���찲�ų������ģ��
    '   bytFun '0-�鿴,1-����,2-�޸�,3-ɾ��,4-��ʱ����(�̶������),5-����������λԤԼ�Һ�,6-������Դ,7-��������ʱ����
    '   lng��ԴID/lng����ID 6-������Դʱ�ɲ�����
    '   strȱʡ���� ȱʡѡ������
    mbytPlanType = bytPlanType: mbytFun = bytFun: mlng����ID = lng����ID
    mlng��ԴId = lng��ԴId: mlng����ID = lng����ID
    mstrȱʡ���� = strȱʡ����
    mstrPrivs = strPrivs
    mlngSavedRecords = 0
    mstrCurDay = ""
    Set mobj���ﰲ�� = New ���ﰲ��

    On Error Resume Next
    Me.Show 1, frmParent
    ShowMe = mlngSavedRecords > 0
End Function

Private Function CheckDepend(ByVal bytMode As Byte, ByVal strCurDate As String, _
    Optional ByVal dtStartTime As Date, Optional ByVal dtEndTime As Date, _
    Optional ByVal blnShowErr As Boolean = True) As Boolean
    '����:����ϰ�ʱ�����Ƿ�ɳ���
    '����:
    '   bytMode - 0,�������Ƿ�ɳ���;1,���ʱ�䷶ΧdtStart~dtEnd�Ƿ�ɳ���
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objͣ���¼ As ͣ���¼
    Dim obj�����¼ As �����¼, blnAllInvalid As Boolean

    On Error GoTo ErrHandler
    If IsDate(strCurDate) = False Then CheckDepend = True: Exit Function
    If mbytPlanType = F_MonthTemplet Then CheckDepend = True: Exit Function '���찲�ų������ģ���˳�
    If Not (mbytFun = Fun_TempPlanRecord _
        Or mbytFun = Fun_AddSignalSourcePlan And mlng��ԴId <> 0 _
        Or mbytFun = Fun_UpdatePlan) Then CheckDepend = True: Exit Function
    
    If mbytFun = Fun_UpdatePlan Then
        '���ʱ��ȫ��ͣ�����ʹ�ã����������
        If bytMode = 0 Then
            If mobj���ﰲ��(1).Count > 0 Then
                blnAllInvalid = True
                For Each obj�����¼ In mobj���ﰲ��(1)
                    If obj�����¼.�Ƿ�̶� = False Then blnAllInvalid = False
                Next
                If blnAllInvalid Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " û�пɵ����İ��ţ���Щ������ʧЧ����ͣ���������ԤԼ�Һţ��������������", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        If IsVisitedOtherTable(mlng����ID, mlng��ԴId, Format(strCurDate, "yyyy-mm-dd")) Then
            If blnShowErr Then
                If mbytFun = Fun_TempPlanRecord Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ������������������õĳ��ﰲ�ţ������ڵ�ǰ������н�����ʱ���ﰲ�ţ�", vbInformation, gstrSysName
                Else
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ��������������������˳��ﰲ�ţ������ظ����ţ�", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
        
        If IsVisitedOtherTable(mlng����ID, mlng��ԴId, Format(strCurDate, "yyyy-mm-dd"), False, mlng����ID) Then
            If blnShowErr Then
                If mbytFun = Fun_TempPlanRecord Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ���ڵ�ǰ�������������������õĳ��ﰲ�ţ������ڵ�ǰ�����н�����ʱ���ﰲ�ţ�", vbInformation, gstrSysName
                Else
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ���ڵ�ǰ�������������������õĳ��ﰲ�ţ������ظ����ţ�", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
        
        If mobj���ﰲ��.�Ű෽ʽ = 0 Then
            '����/�ܰ����еĹ̶����Ų�����ʱ����
            strSQL = "Select b.�Ű෽ʽ" & vbNewLine & _
                    " From �ٴ����ﰲ�� A, �ٴ������ B" & vbNewLine & _
                    " Where a.����id = b.Id And a.��Դid = [2] And Nvl(b.�Ű෽ʽ, 0) In (1, 2) And [3] Between a.��ʼʱ�� And a.��ֹʱ��" & vbNewLine & _
                    "       And Exists(Select 1" & vbNewLine & _
                    "           From �ٴ����ﰲ�� M, �ٴ������ N" & vbNewLine & _
                    "           Where m.����id = n.Id And Nvl(n.�Ű෽ʽ, 0) = 0 And m.��Դid = a.��Դid And n.Id = [1]) And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlng����ID, mlng��ԴId, CDate(strCurDate))
            If Not rsTemp.EOF Then
                If blnShowErr Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ����" & IIf(Val(Nvl(rsTemp!�Ű෽ʽ)) = 1, "��", "��") & "������У������ڵ�ǰ�����н�����ʱ���ﰲ�ţ�", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            
            strSQL = "Select c.�Ű෽ʽ" & vbNewLine & _
                    " From �ٴ����ﰲ�� A, �ٴ������ B, �ٴ������Դ C" & vbNewLine & _
                    " Where a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) In (1, 2) And a.��Դid = c.Id And Nvl(c.�Ű෽ʽ, 0) <> 0" & vbNewLine & _
                    "       And c.Id = [1] And a.��ʼʱ�� < [2] And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlng��ԴId, CDate(strCurDate))
            If Not rsTemp.EOF Then
                If blnShowErr Then
                    MsgBox "��ǰ��Դ�ѵ���Ϊ�˰�" & IIf(Val(Nvl(rsTemp!�Ű෽ʽ)) = 1, "��", "��") & "�Ű࣬�����ڵ�ǰ�����н�����ʱ���ﰲ�ţ�", vbInformation, gstrSysName
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

    '���ܶ���ʷ�İ��Ž��в���
    If bytMode = 0 Then
        If DateDiff("d", strCurDate, Format(zlDatabase.Currentdate, "yyyy-mm-dd")) > 0 Then
            If blnShowErr Then MsgBox "���ܶ���ʷ�ĳ������ڽ���" & _
                IIf(mbytFun = Fun_UpdatePlan, "�������ţ�", IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "���ﰲ�ţ�"), vbInformation, gstrSysName
            Exit Function
        End If
        '��ǰ�����Ƿ�����Чʱ�䷶Χ��
        If mobj���ﰲ��.�Ű෽ʽ = 0 Then
            If Not (DateDiff("d", dtEndTime, mobj���ﰲ��.��ʼʱ��) <= 0 And DateDiff("d", dtStartTime, mobj���ﰲ��.��ֹʱ��) >= 0) Then
                If blnShowErr Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ���ڸú�Դ�Ŀ�ԤԼʱ�䷶Χ�ڲ��ܽ���" & _
                        IIf(mbytFun = Fun_UpdatePlan, "�������ţ�", IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "���ﰲ�ţ�"), vbInformation, gstrSysName
                End If
                Exit Function
            End If
            If Not (DateDiff("d", dtEndTime, mFixedPlanDateRange.dtStart) <= 0 And DateDiff("d", dtStartTime, mFixedPlanDateRange.dtEnd) >= 0) Then
                If blnShowErr Then
                    MsgBox Format(strCurDate, "yyyy-mm-dd") & " ���ڸú�Դ���ﰲ�ŵ���Ч��Χ�ڲ��ܽ���" & _
                        IIf(mbytFun = Fun_UpdatePlan, "�������ţ�", IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "���ﰲ�ţ�"), vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
    Else
        '��ǰ�ϰ�ʱ���Ƿ�����Чʱ�䷶Χ��
        If mobj���ﰲ��.�Ű෽ʽ = 0 Then
            If Not (DateDiff("d", dtEndTime, mFixedPlanDateRange.dtStart) <= 0 And DateDiff("d", dtStartTime, mFixedPlanDateRange.dtEnd) >= 0) Then
                If blnShowErr Then MsgBox "��ǰ�ϰ�ʱ�β��ڸó��ﰲ�ŵ���Ч��Χ�ڲ��ܽ���" & IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "���ﰲ�ţ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mbytFun = Fun_UpdatePlan Then CheckDepend = True: Exit Function
    If mblnCheckedByDay Then CheckDepend = True: Exit Function
    If mobjͣ���¼�� Is Nothing Then
        Set mobjͣ���¼�� = GetStopVisitObjects(GetStopVisit(mlng��ԴId, mobj���ﰲ��.��ʼʱ��, mobj���ﰲ��.��ֹʱ��, False))
    End If
    If mobjͣ���¼�� Is Nothing Then CheckDepend = True: Exit Function
    If mobjͣ���¼��.Count = 0 Then CheckDepend = True: Exit Function
    For Each objͣ���¼ In mobjͣ���¼��
        '���嶼ͣ����������ã�������ѡ���ϰ�ʱ��ʱ���ж�
        If DateDiff("s", objͣ���¼.��ʼʱ��, dtStartTime) >= 0 _
            And DateDiff("s", objͣ���¼.��ֹʱ��, dtEndTime) <= 0 Then
            Select Case objͣ���¼.����
            Case 1   'ͣ�ﰲ��
                If bytMode = 0 Then
                    mblnCheckedByDay = True
                    If blnShowErr Then
                        If MsgBox("ע�⣺" & vbCrLf & _
                                  "    ��ǰ��Դ�ĳ���ҽ���ڽ�����ͣ���ȷ��Ҫ����" & IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "������", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    If blnShowErr Then
                        If MsgBox("ע�⣺" & vbCrLf & _
                                  "    ��ǰ��Դ�ĳ���ҽ���ڸ��ϰ�ʱ�ε�ʱ�䷶Χ����ͣ���ȷ��Ҫ����" & IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "������", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            Case 2   '�����ڼ���
                If bytMode = 0 Then
                    mblnCheckedByDay = True
                    If blnShowErr Then
                        If MsgBox("ע�⣺" & vbCrLf & _
                                  "    ��ǰ��Դ������Ϊ�����ڼ���(" & objͣ���¼.ͣ��ԭ�� & ")��ͣ���ȷ��Ҫ����" & IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "������", _
                                  vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    If blnShowErr Then
                        If MsgBox("ע�⣺" & vbCrLf & _
                                  "    ��ǰ�ϰ�ʱ���ڷ����ڼ���(" & objͣ���¼.ͣ��ԭ�� & ")ͣ��ʱ�䷶Χ�ڣ���ȷ��Ҫ����" & IIf(mbytFun = Fun_TempPlanRecord, "��ʱ", "") & "������", _
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
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False

    '�˵�����
    cbsThis.DeleteAll

    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched

    With cbrToolBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_Signal, "����"): cbrCustom.Handle = picSignal.Hwnd
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_ValidTime, "��Ч��"): cbrCustom.Handle = picValidTime.Hwnd
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_FeeItem, "�շ���Ŀ"): cbrCustom.Handle = picFeeItem.Hwnd
        Set cbrCustom = .Add(xtpControlCustom, mMenuID.M_Verify, "������������"): cbrCustom.Handle = picVerify.Hwnd
        
        If mbytFun = Fun_Delete Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "ȷ��"): cbrControl.flags = xtpFlagRightAlign
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "ȡ��    "): cbrControl.flags = xtpFlagRightAlign
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.flags = xtpFlagRightAlign
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�    "): cbrControl.flags = xtpFlagRightAlign
        End If
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlLabel And cbrControl.Type <> xtpControlCustom Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next

    '�����
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
    '����:��ʼ��Docking�ؼ�
    '����:���˺�
    '����:2016-01-08 14:34:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    Dim strReg As String
    Dim panThis As Pane, panLeft As Pane

    On Error GoTo Errhand
    dkpMain.SetCommandBars cbsThis
    dkpMain.VisualTheme = ThemeOffice2003 '������ʾ���
    sngWidth = 200
    sngHeight = 200
    
    Set panLeft = dkpMain.CreatePane(Pancel_Index.Pan_����, sngWidth, sngHeight, DockTopOf, Nothing)
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Title = "": panLeft.Tag = Pancel_Index.Pan_����
    panLeft.Handle = picDateList.Hwnd
    '�̶���С
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

    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_����, sngWidth, 300, DockRightOf, panLeft)
    panThis.Title = ""
    panThis.Tag = Pancel_Index.Pan_����
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picDetailedList.Hwnd
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_ʱ���, sngWidth, 250, DockBottomOf, panLeft)
    panThis.Title = "�ϰ�ʱ��"
    panThis.Tag = Pancel_Index.Pan_ʱ���
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picWorkTimeList.Hwnd

    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_��Դ, sngWidth, 300, DockBottomOf, panThis)
    panThis.Title = "��ǰ��Դ��Ϣ"
    panThis.Tag = Pancel_Index.Pan_��Դ
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    If mbytPlanType = F_FixedRule Then
        Call InitPage
        picSourceAndPlan.Visible = True
        Set picSouceList.Container = picSourceAndPlan
        panThis.Handle = picSourceAndPlan.Hwnd
    Else
        panThis.Handle = picSouceList.Hwnd
    End If
    '�̶����߶�
    If cldsCalenbarSel.ShowStyle = Show_Plan_Week Then
        panThis.MaxTrackSize.Height = 240
    Else
        panThis.MaxTrackSize.Height = 150
    End If

    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Call picDateList_Resize
    'Set dkpMain.PaintManager.CaptionFont = use.Font

    'zlRestoreDockPanceToReg Me, dkpMan, "����"
    InitPanel = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub RemovePlan(obj���ﰲ�� As ���ﰲ��, ByVal strItem As String)
    '��δ������ﰲ�ź͵�ǰ������ɾ��ĳһ������
    Dim strKey As String

    strKey = GetPlanKey(strItem)
    If obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then obj���ﰲ��.δ������ﰲ��.Remove strKey
    If obj���ﰲ��.Exits(strKey) Then obj���ﰲ��.Remove strKey
    Call ChangeCurPlan(obj���ﰲ��, strItem)
End Sub

Private Sub cboDays_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboFeeItem_Click()
    Dim obj��Դ As �����Դ
    
    Err = 0: On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If cboFeeItem.ListIndex = -1 Then Exit Sub
    If Val(cboFeeItem.Tag) = cboFeeItem.ItemData(cboFeeItem.ListIndex) Then Exit Sub
    
    mblnFeeItemChanged = True
    Set obj��Դ = mobj���ﰲ��.�����Դ.Clone
    obj��Դ.��ĿID = cboFeeItem.ItemData(cboFeeItem.ListIndex)
    obj��Դ.��Ŀ���� = cboFeeItem.Text
    Call SourceInfor.LoadData(obj��Դ)
    Set obj��Դ = Nothing
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
    Dim obj���ﰲ�� As ���ﰲ��, obj�����¼�� As �����¼��
    Dim cllPro As New Collection, ObjItem As �����¼��
    Dim strKey As String, lngCount As Long
    Dim blnHavePlan As Boolean

    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_Save '��������
        If IsValied() = False Then Control.Enabled = True: Exit Sub
        Set obj���ﰲ�� = Get���ﰲ��

        If Not obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then
            For Each ObjItem In obj���ﰲ��.�ѱ�����ﰲ��
                If ObjItem.�Ƿ�ɾ�� Then lngCount = lngCount + 1
                blnHavePlan = True
            Next
        End If
        If Not obj���ﰲ��.δ������ﰲ�� Is Nothing Then
            For Each ObjItem In obj���ﰲ��.δ������ﰲ��
                lngCount = lngCount + 1
            Next
        End If

        If lngCount = 0 Then
            If blnHavePlan And (mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel _
                Or mblnTimeChanged Or mblnFeeItemChanged _
                Or chkAotuVerify.Visible And chkAotuVerify.Value = vbChecked) Then
                '�Զ����
                If chkAotuVerify.Visible And chkAotuVerify.Value = vbChecked Then
                    If TempPlanVerifyOrCancel(obj���ﰲ��.����ID, obj���ﰲ��.�����Դ.ID, 1) Then
                        mlngSavedRecords = mlngSavedRecords + 1
                        Unload Me: Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Else
                MsgBox "��ǰû����Ҫ�������Ч���ţ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If

        Control.Enabled = False
        If SaveData(obj���ﰲ��) Then
            mlngSavedRecords = mlngSavedRecords + 1
            mlng����ID = obj���ﰲ��.����ID

            '��ʱ������޸�ԤԼ�Һſ���ʱ�˳�
            If mbytFun = Fun_TempPlanRecord _
                Or mbytFun = Fun_UpdateUnit _
                Or mbytFun = Fun_TempPlanVerify _
                Or mbytFun = Fun_TempPlanCancel _
                Or mbytFun = Fun_UpdatePlan Then
                Unload Me: Exit Sub
            End If
            
            mblnTimeChanged = False: mblnFeeItemChanged = False
            mblnCheckedByDay = False
            stbThis.Panels(2).Text = "����ɹ���"
            
            '�Զ����
            If (chkAotuVerify.Visible And chkAotuVerify.Value = vbChecked) Then
                If TempPlanVerifyOrCancel(obj���ﰲ��.����ID, obj���ﰲ��.�����Դ.ID, 1) Then
                    Unload Me: Exit Sub
                End If
            End If

            '���¼�������
            Set mobj���ﰲ�� = New ���ﰲ��
            mstrȱʡ���� = mstrCurDay '����ȱʡ����
            If InitData(mobj���ﰲ��, mlng����ID, mlng��ԴId, mlng����ID) Then
                Call CheckTempFixedPlan(mobj���ﰲ��, dtpBegin.Value, dtpEnd.Value)
                Call LoadData
            End If
        End If
        Control.Enabled = True
    Case conMenu_File_Exit '�˳�
        If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
            lngCount = 0
            Set obj���ﰲ�� = Get���ﰲ��
            '����Ƿ������޸İ���
            If Not obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then
                For Each ObjItem In obj���ﰲ��.�ѱ�����ﰲ��
                    If ObjItem.�Ƿ�ɾ�� Then lngCount = lngCount + 1
                Next
            End If
            If Not obj���ﰲ��.δ������ﰲ�� Is Nothing Then
                For Each ObjItem In obj���ﰲ��.δ������ﰲ��
                    lngCount = lngCount + 1
                Next
            End If

            If lngCount > 0 Or mblnTimeChanged Or mblnFeeItemChanged Then
                If MsgBox("���ְ��ſ����ѱ��޸ģ��Ƿ񲻱���ֱ���˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
            And mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "���п���")
        Control.Enabled = Control.Visible
    Case mMenuID.M_Verify
        Control.Visible = mbytPlanType = F_FixedRule _
            And (mbytFun = Fun_TempPlan Or mbytFun = Fun_AddSignalSourcePlan) _
            And zlStr.IsHavePrivs(mstrPrivs, "�����ʱ�̶�����") And mobj���ﰲ��.����ʱ�� <> ""
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Save '��������
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
        'Array(��������,������Ŀ,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��)
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

    '��ʱ���ﰲ�Ų����л�����
    If mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan Then Cancel = True: Exit Sub
    If Format(OldDate, "yyyy-mm-dd") = Format(NewDate, "yyyy-mm-dd") Then Cancel = True: Exit Sub

    If IsDate(NewDate) Then
        'ֻ��ѡ��ʼʱ������ֹʱ�䷶Χ�ڵ�
        If mobj���ﰲ��.�Ű෽ʽ = 0 And cldsCalenbarSel.ShowStyle = Show_Plan_Day Then
            If DateDiff("d", NewDate, mFixedPlanDateRange.dtStart) > 0 Or DateDiff("d", NewDate, mFixedPlanDateRange.dtEnd) < 0 Then
                Cancel = True: Exit Sub
            End If
        Else
            If DateDiff("d", NewDate, mobj���ﰲ��.��ʼʱ��) > 0 Or DateDiff("d", NewDate, mobj���ﰲ��.��ֹʱ��) < 0 Then
                Cancel = True: Exit Sub
            End If
        End If

        '����ϰ�ʱ�����Ƿ�ɳ���
        If mobj���ﰲ��(1).Count > 0 Then
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
        Or (mbytPlanType = F_Templet And mobj���ﰲ��.�Ű���� = 1) _
        Or mbytPlanType = F_MonthTemplet Then
        If CheckExistRecord(0, Replace(GetApplyToStr(), ",", "|"), mobj���ﰲ��) Then
            If MsgBox("ע�⣺" & vbCrLf & _
                      "      ���ֱ�Ӧ�õ����ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�Ӧ�ú��ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        End If
    End If

    '��鵱ǰ������Ϣ
    If CPDPages.IsValied() = False Then
        Cancel = True: Exit Sub
    End If

    '��ȡ��ǰ������Ϣ
    Call Get��ǰ���ﰲ��(mobj���ﰲ��)
    If (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) _
        And mobj���ﰲ��(1).�Ƿ��޸� And IsDate(mstrCurDay) _
        And mbytPlanType <> F_MonthTemplet Then
        If IsVisitedOtherTable(mlng����ID, mlng��ԴId, CDate(mstrCurDay)) Then
            MsgBox Format(mstrCurDay, "yyyy-mm-dd") & " ��������������������˳��ﰲ�ţ������ظ����ţ�", vbInformation, gstrSysName
            mobj���ﰲ��(1).RemoveAll
            mobj���ﰲ��(1).�Ƿ��޸� = False
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

Private Function CheckTempFixedPlan(obj���ﰲ�� As ���ﰲ��, _
    ByVal dtStartTime As Date, ByVal dtEndTime As Date, _
    Optional ByVal blnTimeRangeChanged As Boolean, _
    Optional ByRef blnUnloadForm As Boolean, Optional ByVal blnSaveBeforeValid As Boolean) As Boolean
    '���̶�������ָ��ʱ�䷶Χ���ܷ������ʱ����
    '��Σ�
    '   obj���ﰲ�� ���ﰲ�Ŷ���
    '   dtStartTime��dtEndTime ��ʱ���ŵ���Ч��
    '   blnTimeRangeChanged �Ƿ�ʱ�䷶Χ�����˱仯
    '˵����
    '   ���������£�
    '   1.��ǰ�ѵ���Ϊ��/���Ű����Ѱ���/���ƶ��˰��ţ���������������ʱ����
    '   2.����Ƿ��а���/���Ű�ģ����У�����������
    '   3.��ԤԼ�Һ����ݣ�����������
    '   4.��ԤԼ�Һ����ݣ���
    '     ��ͬһ�ܼ�(����һ)��ԤԼ�Һ����ݵ��ϰ�ʱ��֮��ʱ�䷶Χ�н��棬����ʾ����ֹ�����йҺ����ݵ��ϰ�ʱ��Ϊ"1��(��һ)-����"��"8��(��һ)-����"��
    '     ���򣬰���ԤԼ�Һ����ݵ��ϰ�ʱ���Զ��ӵ���������ʱ�����У��Ҳ������޸�
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim rsFixedRecord As ADODB.Recordset, rsNewFixedRecord As ADODB.Recordset
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim strPriorDay As String, strPriorItem As String
    Dim strPrior�ϰ�ʱ�� As String, dtPriorStart As Date, dtPriorEnd As Date
    Dim strMsgInfo As String, strErrorInfo As String
    Dim i As Long, j As Long, strFixedFilter As String, strKey As String
    Dim blnFindItem As Boolean, blnChangedItem As Boolean
    Dim bln��ʾ As Boolean, blnTemp As Boolean
    
    On Error GoTo ErrHandler
    Set mcllFixedPlan = New Collection 'Array(��������,������Ŀ,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��)
    blnUnloadForm = False
    If Not (mbytFun = Fun_TempPlan _
        Or mbytFun = Fun_TempPlanVerify _
        Or mbytFun = Fun_TempPlanCancel) Then CheckTempFixedPlan = True: Exit Function
    
    If dtStartTime < mdtToday Then dtStartTime = mdtToday
    'ȡ�����ʱ�ļ��
    If mbytFun = Fun_TempPlanCancel Then
        '1.������飬�Ƿ�����ȡ�����
        '2.һ�����ű�ʹ�þͲ�����ȡ�������
        strSQL = "Select 1 From �ٴ����ﰲ�� A Where a.Id = [1] And a.���ʱ�� Is Not Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������", mlng����ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ�����ѱ�����ȡ����˻�ɾ����������ȡ����ˣ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������¼ A, ���˹Һż�¼ B" & vbNewLine & _
                " Where a.Id = b.�����¼id And a.����id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ�����Ƿ���ԤԼ�Һ�����", mlng����ID)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�����Ѵ���ԤԼ�Һ����ݣ�����ȡ����ˣ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C" & vbNewLine & _
                " Where a.��Դid = b.��Դid And a.����id = c.Id And c.�Ű෽ʽ = 0 And a.Id <> b.Id And b.Id = [1] And a.�Ǽ�ʱ�� > b.�Ǽ�ʱ��" & vbNewLine & _
                "       And a.���ʱ�� Is Not Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���֮���Ƿ�������˰���", mlng����ID)
        If Not rsTemp.EOF Then
            MsgBox "�ú�Դ�ڵ�ǰ����֮�󻹴�������˵İ��ţ��㲻��ȡ����˵�ǰ���ţ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        CheckTempFixedPlan = True: Exit Function
    ElseIf mbytFun = Fun_TempPlanVerify Then
        '1.������飬�Ƿ��������
        strSQL = "Select 1 From �ٴ����ﰲ�� A Where a.Id = [1] And a.���ʱ�� Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������", mlng����ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ�����ѱ�������˻�ɾ������������ˣ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        '2.û����Ч���ŵĲ������
        strSQL = "Select 1 From �ٴ��������� A Where a.����Id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����Ч����", mlng����ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ���������κ���Ч���ţ�������ˣ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C" & vbNewLine & _
                " Where a.��Դid = b.��Դid And a.����id = c.Id And c.�Ű෽ʽ = 0 And a.Id <> b.Id And b.Id = [1] And a.�Ǽ�ʱ�� < b.�Ǽ�ʱ��" & vbNewLine & _
                "       And a.���ʱ�� Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���֮ǰ�Ƿ���δ��˰���", mlng����ID)
        If Not rsTemp.EOF Then
            MsgBox "�ú�Դ�ڵ�ǰ����֮ǰ������δ��˵İ��ţ��㲻����˵�ǰ���ţ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
    End If
    
    strSQL = "Select 1 From �ٴ������Դ A " & vbNewLine & _
            " Where Nvl(a.�Ű෽ʽ,0)<>0 and a.Id = [1] And Rownum < 2" & vbNewLine & _
            "       And Exists(Select 1" & vbNewLine & _
            "                  From �ٴ������¼ M,�ٴ����ﰲ�� N,�ٴ������ P" & vbNewLine & _
            "                  Where m.����ID=n.ID And n.����ID=p.ID And n.��ԴID=a.ID And Nvl(p.�Ű෽ʽ,0)=Nvl(a.�Ű෽ʽ,0) And Rownum<2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��а���", mlng��ԴId)
    If Not rsTemp.EOF Then
        MsgBox "��ǰ��Դ����Ϊ�������ǹ̶��Ű෽ʽ�����ƶ��˰��ţ��������ڸó�������ƶ���ʱ���ţ�", vbInformation + vbOKOnly, gstrSysName
        blnUnloadForm = True: Exit Function
    End If
    
    strSQL = "Select 1 From �ٴ������ A Where a.Id = [1] And a.����ʱ�� Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ񷢲�", mlng����ID)
    If Not rsTemp.EOF Then
        strSQL = "Select 1 From �ٴ����ﰲ�� A, �ٴ��������� B Where a.Id = b.����id And a.����id = [1] And a.��ԴID = [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��а���", mlng����ID, mlng��ԴId)
        If rsTemp.EOF Then
            MsgBox "�ú�Դ�ڵ�ǰδ�����ĳ�����л�δ�ƶ��κΰ��ţ������ƶ���ʱ���ţ�", vbInformation + vbOKOnly, gstrSysName
            blnUnloadForm = True: Exit Function
        End If
    End If
    
    '����Ƿ��а���/���Ű��
    strSQL = "Select b.�Ű෽ʽ, a.��ʼʱ��, a.��ֹʱ��" & vbNewLine & _
            " From �ٴ����ﰲ�� A, �ٴ������ B" & vbNewLine & _
            " Where a.����id = b.Id And b.�Ű෽ʽ In (1, 2) And a.��Դid = [1] And a.��ʼʱ�� < [3] And a.��ֹʱ�� > [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��а���/���Ű��", obj���ﰲ��.�����Դ.ID, dtStartTime, dtEndTime)
    If Not rsTemp.EOF Then
        dtCurStart = Nvl(rsTemp!��ʼʱ��): dtCurEnd = Nvl(rsTemp!��ֹʱ��)
        If dtCurStart < dtStartTime Then dtCurStart = dtStartTime
        If dtCurEnd > dtEndTime Then dtCurEnd = dtEndTime
        
        MsgBox "��ǰ��Դ��ʱ�䷶Χ(" & Format(dtCurStart, "yyyy-mm-dd hh:mm:ss") & _
            "-" & Format(dtCurEnd, "yyyy-mm-dd hh:mm:ss") & ")���Ѱ�" & _
            IIf(Val(rsTemp!�Ű෽ʽ) = 1, "��", "��") & "�������Ű࣬�����������ʱ�䷶Χ���ƶ���ʱ���ţ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '1.����Ƿ���ԤԼ�Һ�����
    Set rsFixedRecord = GetԤԼ�Һż�¼(obj���ﰲ��.�����Դ.ID, _
        CDate(Format(dtStartTime, "yyyy-mm-dd")), CDate(Format(dtEndTime, "yyyy-mm-dd")))
    If rsFixedRecord.EOF Then
        Call SetPlanFixed(obj���ﰲ��, False, "", "", blnChangedItem) '�����г����¼�����Ϊ�����޸ĵ�
        If blnChangedItem Then Call LoadData
        CheckTempFixedPlan = True: Exit Function
    End If
    
    '1.��ͬһ�ܼ�(����һ)��ԤԼ�Һ����ݵ��ϰ�ʱ��֮��ʱ�䷶Χ�н��棬����ʾ����ֹ
    Do While Not rsFixedRecord.EOF
        '�����ǰ�������Ŀ�Ϳ�ʼʱ��(hh:mm:ss)����ģ�
        '����ÿһ��������Ŀֻ��Ҫ��������ϰ�ʱ�ε�ʱ�䷶Χ�Ƿ��н��漴��
        dtCurStart = Format(mdtToday, "yyyy-mm-dd ") & Format(Nvl(rsFixedRecord!��ʼʱ��), "hh:mm:ss")
        dtCurEnd = Format(mdtToday, "yyyy-mm-dd ") & Format(Nvl(rsFixedRecord!��ֹʱ��), "hh:mm:ss")
        dtCurEnd = GetWorkTrueDate(dtCurStart, dtCurEnd)
        
        If strPriorItem = Nvl(rsFixedRecord!������Ŀ) Then
            If Not (DateDiff("n", dtCurStart, dtPriorEnd) <= 0 Or DateDiff("n", dtCurEnd, dtPriorStart) >= 0) Then
                '�ϰ�ʱ��ʱ�䷶Χ�н��棬��֯��ʾ��Ϣ
                strErrorInfo = strErrorInfo & vbCrLf & _
                    strPriorDay & "(" & strPriorItem & ")�ϰ�ʱ�Ρ�" & strPrior�ϰ�ʱ�� & "(" & Format(dtPriorStart, "hh:mm") & "-" & Format(dtPriorEnd, "hh:mm") & ")��" & _
                    "��" & Format(Nvl(rsFixedRecord!��������), "yyyy-mm-dd ") & "(" & Nvl(rsFixedRecord!������Ŀ) & ")�ϰ�ʱ�Ρ�" & Nvl(rsFixedRecord!�ϰ�ʱ��) & "(" & Format(dtCurStart, "hh:mm") & "-" & Format(dtCurEnd, "hh:mm") & ")��"
            End If
        End If
        
        strPriorItem = Nvl(rsFixedRecord!������Ŀ): strPriorDay = Format(Nvl(rsFixedRecord!��������), "yyyy-mm-dd ")
        strPrior�ϰ�ʱ�� = Nvl(rsFixedRecord!�ϰ�ʱ��)
        dtPriorStart = dtCurStart: dtPriorEnd = dtCurEnd
        
        '�ϰ�ʱ�β����޸Ĺ̶����룬��֯��ʾ��Ϣ
        strMsgInfo = strMsgInfo & vbCrLf & _
            strPriorDay & "(" & strPriorItem & ") �ϰ�ʱ�Ρ�" & strPrior�ϰ�ʱ�� & "(" & Format(dtPriorStart, "hh:mm") & "-" & Format(dtPriorEnd, "hh:mm") & ")��"
        
        'Array(��������,������Ŀ,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��)
        strKey = "K" & strPriorItem & "_" & strPrior�ϰ�ʱ��
        If CollExitsValue(mcllFixedPlan, strKey) = False Then
            mcllFixedPlan.Add Array(strPriorItem, strPriorItem, strPrior�ϰ�ʱ��, dtPriorStart, dtPriorEnd), strKey
        End If
        rsFixedRecord.MoveNext
    Loop
    
    If strErrorInfo <> "" Then
        MsgBox "�ڵ�ǰѡ�����Чʱ�䷶Χ�ڵ�ǰ��Դ��ͬһ��������Ŀ(����)�´��ڽ�����ϰ�ʱ�Σ�" & _
               "������Щ�ϰ�ʱ��ʱ�䷶Χ�ڶ�����ԤԼ�Һż�¼���������ڸ���Чʱ�䷶Χ���ƶ���ʱ" & _
               "���ţ�������޸���Чʱ�䷶ΧȻ�������" & vbCrLf & _
               strErrorInfo, vbInformation + vbOKOnly, gstrSysName
        If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
        Exit Function
    End If
    
    If blnTimeRangeChanged = False And obj���ﰲ��.����ID <> 0 And IsDate(obj���ﰲ��.�Ǽ�ʱ��) Then
        '2.����������Ժ��Ƿ����µĲ������а����е�ԤԼ�Һż�¼
        strSQL = "Select Decode(To_Char(��������, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) As ������Ŀ," & vbNewLine & _
                "        ID As ��¼id, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, �Ƿ��ռ" & vbNewLine & _
                " From (Select a.Id, a.��������, a.�ϰ�ʱ��, a.��ʼʱ��, a.��ֹʱ��, �Ƿ��ռ," & vbNewLine & _
                "               Row_Number() Over(Partition By To_Char(a.��������, 'D'), a.�ϰ�ʱ�� Order By a.��������) As �к�" & vbNewLine & _
                "        From �ٴ������¼ A, ���˹Һż�¼ B" & vbNewLine & _
                "        Where a.Id = b.�����¼id And a.�ϰ�ʱ�� Is Not Null And a.��Դid = [1] And a.�������� Between [2] And [3]" & vbNewLine & _
                "              And b.�Ǽ�ʱ�� > [4])" & vbNewLine & _
                " Where �к� < 2"
        Set rsNewFixedRecord = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ԤԼ�Һ����ݵĳ����¼", obj���ﰲ��.�����Դ.ID, _
            CDate(Format(dtStartTime, "yyyy-mm-dd")), CDate(Format(dtEndTime, "yyyy-mm-dd")), CDate(obj���ﰲ��.�Ǽ�ʱ��))
        
        strMsgInfo = ""
        Do While Not rsNewFixedRecord.EOF
            '�ϰ�ʱ�β����޸Ĺ̶����룬��֯��ʾ��Ϣ
            strMsgInfo = strMsgInfo & vbCrLf & _
                Format(Nvl(rsNewFixedRecord!��������), "yyyy-mm-dd ") & "(" & Nvl(rsNewFixedRecord!������Ŀ) & _
                    ") �ϰ�ʱ�Ρ�" & Nvl(rsNewFixedRecord!�ϰ�ʱ��) & "(" & Format(Nvl(rsNewFixedRecord!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsNewFixedRecord!��ֹʱ��), "hh:mm") & ")��"
            rsNewFixedRecord.MoveNext
        Loop
        
        If strMsgInfo <> "" Then
            MsgBox "�ڵ�ǰѡ�����Чʱ�䷶Χ�ڣ���ǰ��Դ���ϴ��������޸ĵ��������ʱ�䷶Χ���в��ڸð����е��ϰ�ʱ�β�������" & _
                   "��ԤԼ�Һż�¼(��Щ�ϰ�ʱ�α���������µİ�����)��" & _
                   IIf(mbytFun <> Fun_TempPlanVerify, "��Ҫ���µ������ţ�", "�������µ������ŷ�������ˣ�") & vbCrLf & _
                    strMsgInfo, vbInformation + vbOKOnly, gstrSysName
            If mbytFun = Fun_TempPlanVerify Then blnUnloadForm = True: Exit Function
        End If
    Else
        bln��ʾ = True
    End If
    
    blnChangedItem = False
    If rsFixedRecord.RecordCount > 0 Then
        Call GetFixedPlan(obj���ﰲ��, rsFixedRecord, blnTemp)
        blnChangedItem = blnChangedItem Or blnTemp
    End If
    
    'Array(��������,������Ŀ,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��)
    '������ڹ̶������޸ĵİ��ż����ڵİ���
    Call SetPlanFixed(obj���ﰲ��, False, "", "", blnTemp, mcllFixedPlan)
    blnChangedItem = blnChangedItem Or blnTemp
    
    If blnChangedItem Then
        If bln��ʾ Then
            MsgBox "�ڵ�ǰѡ�����Чʱ�䷶Χ�ڣ���ǰ��Դ���������ڴ�����ԤԼ�Һż�¼���ϰ�ʱ�Σ�" & _
                    "��Щ�ϰ�ʱ�ζ��������޸ģ��ұ�����뵽�µ���ʱ�����У�" & vbCrLf & _
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

Private Function SetPlanFixed(obj���ﰲ�� As ���ﰲ��, ByVal blnFixed As Boolean, _
    Optional ByVal str������Ŀ As String, Optional ByVal str�ϰ�ʱ�� As String, _
    Optional ByRef blnChangedItem As Boolean, Optional ByVal cllFixedPlan As Collection) As Boolean
    '���ð��ŵĹ̶���ǣ���Ҫ���ڹ̶�������ʱ����
    '��Σ�
    '   blnFixed �Ƿ�̶�
    '   str������Ŀ ���Ϊ�գ��������еĳ�������
    '   str�ϰ�ʱ�� ���Ϊ�գ��������е��ϰ�ʱ��
    '���أ�
    '   str������Ŀ��Ϊ����str�ϰ�ʱ�β�Ϊ��ʱ�����Ƿ��ҵ�
    '   blnChangedItem �Ƿ��иı����Ŀ
    '˵����
    '   cllFixedPlanԪ�ش�����ʱ������������ڼ����еĹ̶���ʶ
    Dim blnFindItem As Boolean, blnTemp As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    blnChangedItem = False
    blnFindItem = SetPlanFixedSub(obj���ﰲ��, blnFixed, str������Ŀ, str�ϰ�ʱ��, blnTemp, cllFixedPlan)
    blnChangedItem = blnTemp
    blnFindItem = blnFindItem Or SetPlanFixedSub(obj���ﰲ��.�ѱ�����ﰲ��, blnFixed, str������Ŀ, str�ϰ�ʱ��, blnTemp, cllFixedPlan)
    blnChangedItem = blnChangedItem Or blnTemp
    blnFindItem = blnFindItem Or SetPlanFixedSub(obj���ﰲ��.δ������ﰲ��, blnFixed, str������Ŀ, str�ϰ�ʱ��, blnTemp, cllFixedPlan)
    blnChangedItem = blnChangedItem Or blnTemp
    
    SetPlanFixed = blnFindItem
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetPlanFixedSub(obj���ﰲ�� As ���ﰲ��, ByVal blnFixed As Boolean, _
    Optional ByVal str������Ŀ As String, Optional ByVal str�ϰ�ʱ�� As String, _
    Optional ByRef blnChangedItem As Boolean, Optional ByVal cllFixedPlan As Collection) As Boolean
    '���ð��ŵĹ̶���ǣ���Ҫ���ڹ̶�������ʱ����
    '��Σ�
    '   blnFixed �Ƿ�̶�
    '   str������Ŀ ���Ϊ�գ��������еĳ�������
    '   str�ϰ�ʱ�� ���Ϊ�գ��������е��ϰ�ʱ��
    '���أ�
    '   str������Ŀ��Ϊ����str�ϰ�ʱ�β�Ϊ��ʱ�����Ƿ��ҵ�
    '   blnChangedItem �Ƿ��иı����Ŀ
    '˵����
    '   cllFixedPlanԪ�ش�����ʱ������������ڼ����еĹ̶���ʶ
    Dim blnFind As Boolean, blnFindItem As Boolean
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    
    blnChangedItem = False
    If obj���ﰲ�� Is Nothing Then Exit Function
    If cllFixedPlan Is Nothing Then Set cllFixedPlan = New Collection
    If cllFixedPlan.Count = 0 Then
        For Each obj�����¼�� In obj���ﰲ��
            If blnFind And str������Ŀ <> "" Then Exit For
            If obj�����¼��.�������� = str������Ŀ Or str������Ŀ = "" Then
                For Each obj�����¼ In obj�����¼��
                    If obj�����¼.ʱ��� = str�ϰ�ʱ�� Or str�ϰ�ʱ�� = "" Then
                        If obj�����¼.�Ƿ�̶� <> blnFixed Then
                            obj�����¼.�Ƿ�̶� = blnFixed
                            blnChangedItem = True
                        End If
                        blnFind = True: blnFindItem = True
                        If str�ϰ�ʱ�� <> "" Then Exit For
                    End If
                Next
            End If
        Next
    Else
        For Each obj�����¼�� In obj���ﰲ��
            For Each obj�����¼ In obj�����¼��
                If CollExitsValue(cllFixedPlan, "K" & obj�����¼��.�������� & "_" & obj�����¼.ʱ���) = False Then
                    If obj�����¼.�Ƿ�̶� <> blnFixed Then
                        obj�����¼.�Ƿ�̶� = blnFixed
                        blnChangedItem = True
                    End If
                End If
            Next
        Next
    End If
    SetPlanFixedSub = blnFindItem
End Function

Private Sub GetFixedPlan(obj���ﰲ�� As ���ﰲ��, rsRecord As ADODB.Recordset, _
    Optional ByRef blnNewAdd As Boolean)
    '���̶������޸ĵİ��ŵ���Ϊ��¼��
    '����:
    '   blnNewAdd - �Ƿ�������/�޸ĵ�
    '˵����
    '   �����ݼ���ʱ������ȫ����Ϊ���޸ĵ�ԭ���ǣ�
    '   �ƶ���ʱ����ʱ�޷��ж��Ƿ����µĲ����޸ĵ��ϰ�ʱ��(���µ�ԤԼ�Һż�¼���ϰ�ʱ��)
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim obj������λ As ������λ����, obj������λTmp As ������λ����
    Dim strItem As String, rsUnitReg As ADODB.Recordset
    Dim strKey As String
    
    '��ȡ����ĳ��ﰲ�Ŷ���
    Err = 0: On Error GoTo ErrHandler
    blnNewAdd = False
    If obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then Set obj���ﰲ��.�ѱ�����ﰲ�� = New ���ﰲ��
    If obj���ﰲ��.δ������ﰲ�� Is Nothing Then Set obj���ﰲ��.δ������ﰲ�� = New ���ﰲ��
    If rsRecord.RecordCount = 0 Then Exit Sub
    
    rsRecord.MoveFirst
    Set obj�����¼�� = New �����¼��
    obj�����¼��.�������� = Nvl(rsRecord!������Ŀ)
    
    strKey = "K" & obj�����¼��.��������
    If obj���ﰲ��.Exits(strKey) Then
        Set obj�����¼�� = obj���ﰲ��(strKey).Clone
    ElseIf obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then
        Set obj�����¼�� = obj���ﰲ��.δ������ﰲ��(strKey).Clone
    ElseIf obj���ﰲ��.�ѱ�����ﰲ��.Exits(strKey) Then
        Set obj�����¼�� = obj���ﰲ��.�ѱ�����ﰲ��(strKey).Clone
    End If
    Do While Not rsRecord.EOF
        'ת���ɶ���
        If strItem <> "" And strItem <> Nvl(rsRecord!������Ŀ) Then
            If obj�����¼��.Count > 0 Then
                '���뵽������
                strKey = "K" & obj�����¼��.��������
                If obj���ﰲ��.�ѱ�����ﰲ��.Exits(strKey) Then obj���ﰲ��.�ѱ�����ﰲ��(strKey).�Ƿ�ɾ�� = True '���Ϊɾ��
                If obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then obj���ﰲ��.δ������ﰲ��.Remove strKey '�������Ƴ�
                obj���ﰲ��.δ������ﰲ��.AddItem obj�����¼��, strKey
                
                If obj���ﰲ��.Exits(strKey) Then
                    obj���ﰲ��.Remove strKey '�������Ƴ�
                    obj���ﰲ��.AddItem obj�����¼��.Clone, strKey
                End If
            End If
            Set obj�����¼�� = New �����¼��
            obj�����¼��.�������� = Nvl(rsRecord!������Ŀ)
            
            strKey = "K" & obj�����¼��.��������
            If obj���ﰲ��.Exits(strKey) Then
                Set obj�����¼�� = obj���ﰲ��(strKey).Clone
            ElseIf obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then
                Set obj�����¼�� = obj���ﰲ��.δ������ﰲ��(strKey).Clone
            ElseIf obj���ﰲ��.�ѱ�����ﰲ��.Exits(strKey) Then
                Set obj�����¼�� = obj���ﰲ��.�ѱ�����ﰲ��(strKey).Clone
            End If
        End If
        
        '1.�����¼
        Set obj�����¼ = GetVisitTimesObject(GetVisitTime(Val(rsRecord!��¼ID), True))
        '�ϰ�ʱ��
        If obj���ﰲ��.�����ϰ�ʱ��.Exits("K" & obj�����¼.ʱ���) Then
            Set obj�����¼.�ϰ�ʱ�� = obj���ﰲ��.�����ϰ�ʱ��("K" & obj�����¼.ʱ���).Clone
        Else
            '�����¼ʱ���ϰ�ʱ�ο����ѱ�ɾ��
            Set obj�����¼.�ϰ�ʱ�� = New �ϰ�ʱ��
            With obj�����¼.�ϰ�ʱ��
                .��ʼʱ�� = obj�����¼.��ʼʱ��
                .����ʱ�� = obj�����¼.��ֹʱ��
            End With
        End If

        '2.��������
        Set obj�����¼.�����������Ҽ� = GetVisitRoomsObjects(GetVisitRooms(Val(rsRecord!��¼ID), True))
        obj�����¼.�����������Ҽ�.���﷽ʽ = obj�����¼.���﷽ʽ
        obj�����¼.�����������Ҽ�.ҽ������ = obj���ﰲ��.�����Դ.ҽ������

        '3.������Ϣ
        Set obj�����¼.������Ϣ�� = GetTimeIntervalObjects(GetTimeInterval(Val(rsRecord!��¼ID), True))
        With obj�����¼.������Ϣ��
            .����Ƶ�� = obj���ﰲ��.�����Դ.����Ƶ��
            .�Ƿ��ʱ�� = obj�����¼.�Ƿ��ʱ��
            .�Ƿ���ſ��� = obj�����¼.�Ƿ���ſ���
            .�޺��� = obj�����¼.�޺���
            .��Լ�� = obj�����¼.��Լ��
            .ԤԼ���� = obj�����¼.ԤԼ����
        End With

        '4.������λԤԼ����
        Set obj�����¼.������λ���Ƽ� = New ������λ���Ƽ�
        obj�����¼.������λ���Ƽ�.�Ƿ��ռ = Val(Nvl(rsRecord!�Ƿ��ռ)) = 1
        For Each obj������λ In obj���ﰲ��.���к�����λ
            Set rsUnitReg = GetUnitReg(Val(rsRecord!��¼ID), obj������λ.������λ����, obj������λ.����, True)
            If Not rsUnitReg.EOF Then
                Set obj������λTmp = New ������λ����
                obj������λTmp.������λ���� = obj������λ.������λ����
                obj������λTmp.���� = obj������λ.����
                obj������λTmp.ԤԼ���Ʒ�ʽ = Val(rsUnitReg!���Ʒ�ʽ)
                Set obj������λTmp.������Ϣ�� = GetTimeIntervalObjects(rsUnitReg)
                
                obj�����¼.������λ���Ƽ�.AddItem obj������λTmp, "K" & obj������λTmp.������λ����
            End If
        Next
        obj�����¼.�Ƿ�̶� = True '�ڽ�����ʱ����ʱ������ɾ��
        
        strKey = "K" & obj�����¼.ʱ���
        If obj�����¼��.Exits(strKey) Then
            If obj�����¼��(strKey).�Ƿ�̶� = False Then
                blnNewAdd = True
                obj�����¼.�Ƿ��޸� = True
            End If
            obj�����¼��.Remove strKey
        Else
            blnNewAdd = True
            obj�����¼.�Ƿ��޸� = True
        End If
        obj�����¼��.AddItem obj�����¼, "K" & obj�����¼.ʱ���
        strItem = Nvl(rsRecord!������Ŀ)
        rsRecord.MoveNext
    Loop
    
    '������󣬼�����һ���Ƿ�����Ч����
    If obj�����¼��.Count > 0 Then
        '���뵽������
        strKey = "K" & obj�����¼��.��������
        If obj���ﰲ��.�ѱ�����ﰲ��.Exits(strKey) Then obj���ﰲ��.�ѱ�����ﰲ��(strKey).�Ƿ�ɾ�� = True '���Ϊɾ��
        If obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then obj���ﰲ��.δ������ﰲ��.Remove strKey '�������Ƴ�
        obj���ﰲ��.δ������ﰲ��.AddItem obj�����¼��, strKey
        
        If obj���ﰲ��.Exits(strKey) Then
            obj���ﰲ��.Remove strKey '�������Ƴ�
            obj���ﰲ��.AddItem obj�����¼��.Clone, strKey
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
    '��ʾ��ǰ����
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim obj���ﰲ�� As ���ﰲ��
    
    Err = 0: On Error GoTo ErrHandler
    stbThis.Panels(2).Text = ""
    If mbytPlanType <> F_FixedRule Then Exit Sub

    Set obj�����¼�� = CPDPages.Get�����¼��()
    If obj�����¼�� Is Nothing Then Exit Sub
    If obj�����¼��.Count = 0 Then Exit Sub
    
    Set obj���ﰲ�� = New ���ﰲ��
    obj���ﰲ��.AddItem obj�����¼��, GetPlanKey(obj�����¼��.��������)
    
    LoadPlanToGrid obj���ﰲ��, 0
    Set obj���ﰲ�� = Nothing
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
    mblnValiedCanSave = CheckTempFixedPlan(mobj���ﰲ��, dtpBegin.Value, dtpEnd.Value, True)
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
    mblnValiedCanSave = CheckTempFixedPlan(mobj���ﰲ��, dtpBegin.Value, dtpEnd.Value, True)
    dtpEnd.Tag = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    Dim blnUnloadForm As Boolean
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    On Error GoTo ErrHandler
    If mbytFun = Fun_TempPlan _
        Or mbytFun = Fun_Update And mobj���ﰲ��.�Ƿ���ʱ���� _
        Or mbytFun = Fun_TempPlanVerify _
        Or mbytFun = Fun_TempPlanCancel Then
        
        If CheckTempFixedPlan(mobj���ﰲ��, dtpBegin.Value, dtpEnd.Value, False, blnUnloadForm) = False Then
            If blnUnloadForm Then Unload Me: Exit Sub
        End If
    End If
    
    If (mbytFun = Fun_Add Or mbytFun = Fun_Update) _
        And (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) And IsDate(mstrȱʡ����) Then
        If IsVisitedOtherTable(mlng����ID, mlng��ԴId, CDate(mstrȱʡ����)) Then
            MsgBox "ע�⣬" & Format(mstrȱʡ����, "yyyy-mm-dd") & " ��������������������˳��ﰲ�ţ�", vbInformation, gstrSysName
        End If
    End If

    If CheckDepend(0, mstrȱʡ����) = False Then Unload Me: Exit Sub

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
    cldsCalenbarSel.KeyShift = Shift        '�Ƿ���Ctrl��
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    cldsCalenbarSel.KeyShift = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    mblnFirst = True
    
    mblnTimeChanged = False: mblnFeeItemChanged = False
    mblnCheckedByDay = False
    If mlng����ID = 0 Then
        MsgBox "�������Ϣδ�ҵ�����ˢ�º����ԣ�", vbInformation + vbOKOnly, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Select Case mbytFun
    Case Fun_View   '�鿴
        Me.Caption = "�鿴����"
    Case Fun_Add   '����
        Me.Caption = "��������"
    Case Fun_Update, Fun_UpdatePlan '�༭,�����ѷ�����İ���
        Me.Caption = "��������"
    Case Fun_Delete   'ɾ��
        Me.Caption = "ɾ������"
    Case Fun_TempPlan   '��ʱ����(�̶������)
        Me.Caption = "��ʱ����"
    Case Fun_UpdateUnit   '����������λԤԼ�Һ�
        Me.Caption = "����ԤԼ�Һſ���"
    Case Fun_AddSignalSourcePlan   '������Դ
        Me.Caption = "������Դ����"
    Case Fun_TempPlanRecord  '��������ʱ����
        Me.Caption = "��ʱ����"
    Case Fun_TempPlanVerify   '��ʱ����(�̶������)���
        Me.Caption = "�����ʱ����"
    Case Fun_TempPlanCancel   '��ʱ����(�̶������)ȡ�����
        Me.Caption = "ȡ�������ʱ����"
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
    If InitData(mobj���ﰲ��, mlng����ID, mlng��ԴId, mlng����ID) = False Then Unload Me: Exit Sub

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

Private Function GetLastValidPlanFeeItem(ByVal lng��ԴId As Long) As ADODB.Recordset
    '��ȡ���һ����Ч���ŵ��շ���Ŀ
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select b.id, b.����" & vbNewLine & _
            " From (Select ��Ŀid From �ٴ����ﰲ��" & vbNewLine & _
            "       Where ��Դid = [1] And ���ʱ�� Is Not Null" & vbNewLine & _
            "       Order By �Ǽ�ʱ�� Desc) A, �շ���ĿĿ¼ B" & vbNewLine & _
            " Where a.��Ŀid = b.Id And Rownum < 2"
    Set GetLastValidPlanFeeItem = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���һ����Ч���ŵ��շ���Ŀ", lng��ԴId)
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData(ByRef obj���ﰲ�� As ���ﰲ��, ByVal lng����ID As Long, _
    ByVal lng��ԴId As Long, Optional ByVal lng����ID As Long) As Boolean
    '���ܣ�������ﰲ�Ŷ���
    Dim rsSignalSource As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim obj�����ϰ�ʱ�� As �ϰ�ʱ�μ�, obj���з������� As �������Ҽ�, obj���к�����λ As ������λ���Ƽ�
    Dim obj������λ���� As ������λ����, obj������λ As ������λ����
    Dim blnRecord As Boolean, strTemp As String
    Dim obj������Ϣ�� As ������Ϣ��

    '�������ҡ�������Ϣ��������λ������Ϣ�����ݰ���IDһ���Դ����ݿ��ж�ȡ����
    Dim rs������Ŀ As ADODB.Recordset, rs�����¼ As ADODB.Recordset
    Dim rs�������� As ADODB.Recordset, rs������Ϣ As ADODB.Recordset, rs������λ������Ϣ As ADODB.Recordset
    Dim obj�����¼ As �����¼, obj�����¼�� As �����¼��
    Dim strSQL As String, rsLastValidPlanFeeItem As ADODB.Recordset
    Dim dtNow As Date
    
    Err = 0: On Error GoTo ErrHandler
    '�����շ���Ŀ
    If mbytPlanType = F_FixedRule _
            And mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "���п���") Then
        strSQL = "Select ID,���� From �շ���ĿĿ¼ " & _
                " Where ���='1' And (����ʱ�� >=To_date('3000-01-01','yyyy-mm-dd') Or ����ʱ�� Is Null)" & _
                " And (վ��='" & gstrNodeNo & "' Or վ�� is Null) " & _
                " Order by ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        cboFeeItem.Clear
        Do While Not rsTemp.EOF
            cboFeeItem.AddItem rsTemp!����
            cboFeeItem.ItemData(cboFeeItem.NewIndex) = Val(Nvl(rsTemp!ID))
            rsTemp.MoveNext
        Loop
    End If
    
    '���ﰲ����Ϣ
    Set obj���ﰲ�� = GetVisitPlanObjects(GetVisitPlan(lng����ID, lng����ID))
    If obj���ﰲ��.����ID = 0 Then
        MsgBox "�������Ϣδ�ҵ�����ˢ�º����ԣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '��Դ��Ϣ
    Set rsSignalSource = GetSignalSource("", IIf(mbytFun = Fun_AddSignalSourcePlan And lng��ԴId = 0, -1, lng��ԴId))
    If rsSignalSource.RecordCount = 0 Then
        If mbytFun <> Fun_AddSignalSourcePlan Then
            MsgBox "�ٴ������Դ��Ϣδ�ҵ�����ˢ�º����ԣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    dtNow = zlDatabase.Currentdate
    mdtToday = CDate(Format(dtNow, "yyyy-mm-dd"))
    blnRecord = Not (mbytPlanType = F_Templet Or mbytPlanType = F_FixedRule Or mbytPlanType = F_MonthTemplet)
    If mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel Then
        mstrȱʡ���� = "��һ" '��˺�ȡ�����ȱʡΪ��һ
    End If
    
    If lng����ID = 0 Then
        '����ȱʡʱ�䷶Χ
        Select Case mbytPlanType
        Case F_FixedRule
            obj���ﰲ��.��ʼʱ�� = Format(DateAdd("d", IIf(mbytFun = Fun_TempPlan, GetԤԼ����(lng����ID, lng��ԴId) + 1, 1), mdtToday), "yyyy-MM-dd hh:mm:ss")
            If mbytFun = Fun_TempPlan Then
                'ȱʡһ������
                obj���ﰲ��.��ֹʱ�� = Format(DateAdd("d", 7, obj���ﰲ��.��ʼʱ��), "yyyy-MM-dd 23:59:59")
            Else
                obj���ﰲ��.��ֹʱ�� = "3000-01-01"
            End If
        Case F_Templet
            
        Case F_MonthTemplet
            obj���ﰲ��.��ʼʱ�� = "1900-01-01"
            obj���ﰲ��.��ֹʱ�� = "1900-01-31"
        Case Else
            Dim varDateRange As Variant
            varDateRange = GetDateRange(obj���ﰲ��.���, obj���ﰲ��.�·�, IIf(mbytPlanType = F_WeekPlan, obj���ﰲ��.����, 0))
            obj���ﰲ��.��ʼʱ�� = Format(varDateRange(0), "yyyy-mm-dd hh:mm:ss")
            obj���ﰲ��.��ֹʱ�� = Format(varDateRange(1), "yyyy-mm-dd hh:mm:ss")
        End Select
    Else
        '�޸Ĺ̶����ų����¼�Ŀ�ʼ����Ϊ��ǰ���ڣ���ֹ��������
        If obj���ﰲ��.�Ű෽ʽ = 0 And blnRecord Then
            With mFixedPlanDateRange
                .dtStart = IIf(DateDiff("d", obj���ﰲ��.��ʼʱ��, mdtToday) > 0, mdtToday, obj���ﰲ��.��ʼʱ��)
                .dtEnd = mobj���ﰲ��.��ֹʱ��
                If .dtEnd > CDate(Format(mdtToday + GetԤԼ����(lng����ID, mlng��ԴId), "yyyy-mm-dd 23:59:59")) Then
                    .dtEnd = CDate(Format(mdtToday + GetԤԼ����(lng����ID, mlng��ԴId), "yyyy-mm-dd 23:59:59"))
                End If
            End With
            
            obj���ﰲ��.��ʼʱ�� = Format(mdtToday, "yyyy-MM-dd")
            obj���ﰲ��.��ֹʱ�� = Format(mdtToday + GetԤԼ����(lng����ID, mlng��ԴId), "yyyy-mm-dd 23:59:59")
        End If
        If (mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan) And IsDate(mstrȱʡ����) Then
            obj���ﰲ��.��ʼʱ�� = mstrȱʡ����
            obj���ﰲ��.��ֹʱ�� = Format(mstrȱʡ����, "yyyy-mm-dd 23:59:59")
        End If
        If mbytPlanType = F_MonthTemplet Then
            obj���ﰲ��.��ʼʱ�� = "1900-01-01"
            obj���ﰲ��.��ֹʱ�� = "1900-01-31"
        End If
    End If
    
    '��Դ��Ϣ,���ҡ�ҽ������Ŀȡ�����е�
    Set obj���ﰲ��.�����Դ = GetSignalSourceObject(rsSignalSource)
    If obj���ﰲ��.����ID <> 0 Then
        With obj���ﰲ��.�����Դ
            .��ĿID = obj���ﰲ��.��ĿID
            .��Ŀ���� = obj���ﰲ��.��Ŀ����
            .ҽ��ID = obj���ﰲ��.ҽ��ID
            .ҽ������ = obj���ﰲ��.ҽ������
            .ҽ��ְ�� = obj���ﰲ��.ҽ��ְ��
        End With
    Else
        If mbytPlanType = F_FixedRule And mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "���п���") Then
            'ȡ���һ����Ч���ŵ��շ���Ŀ
            Set rsLastValidPlanFeeItem = GetLastValidPlanFeeItem(mlng��ԴId)
            If rsLastValidPlanFeeItem.RecordCount > 0 Then
                obj���ﰲ��.�����Դ.��ĿID = Val(Nvl(rsLastValidPlanFeeItem!ID))
                obj���ﰲ��.�����Դ.��Ŀ���� = Nvl(rsLastValidPlanFeeItem!����)
            End If
        End If
    End If
    
    mblnNotClick = True
    zlControl.CboLocate cboFeeItem, obj���ﰲ��.�����Դ.��ĿID, True
    If cboFeeItem.ListIndex = -1 Then
        cboFeeItem.AddItem obj���ﰲ��.�����Դ.��Ŀ����
        cboFeeItem.ItemData(cboFeeItem.NewIndex) = obj���ﰲ��.�����Դ.��ĿID
        cboFeeItem.ListIndex = cboFeeItem.NewIndex
    End If
    mblnNotClick = False

    '��������
    Set obj���з������� = GetVisitRoomsObjects(GetDoctorRooms(obj���ﰲ��.�����Դ.����ID))
    Set obj�����ϰ�ʱ�� = GetWorkTimesObjects(GetWorkTimes(obj���ﰲ��.�����Դ.վ��, obj���ﰲ��.�����Դ.����))
    Set obj���к�����λ = GetUnitsObjects(GetUnitAll())

    If Not (mbytFun = Fun_View Or mbytFun = Fun_UpdateUnit _
        Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel) Then
        If mbytFun <> Fun_UpdatePlan Then
            '��������ʱ���������ϰ�ʱ��
            Set obj���ﰲ��.�����ϰ�ʱ�� = obj�����ϰ�ʱ��
        End If
        Set obj���ﰲ��.���з������� = obj���з�������
    End If
    Set obj���ﰲ��.���к�����λ = obj���к�����λ

    Set obj���ﰲ��.��Դ���� = GetClinicRecordFromSignalSource(obj���ﰲ��.�����Դ.ID)
    If blnRecord Then
        '��ȡͣ�ﰲ��
        Set obj���ﰲ��.ͣ���¼ = GetStopVisitObjects(GetStopVisit(obj���ﰲ��.�����Դ.ID, obj���ﰲ��.��ʼʱ��, obj���ﰲ��.��ֹʱ��))
        If mbytFun = Fun_UpdatePlan Then
            Set mrsVisitedRecordByDate = GetVisitedRecordByDate(mlng����ID, mstrȱʡ����)
        Else
            Set mrsVisitedRecord = GetVisitedRecord(obj���ﰲ��.�����Դ.ID, obj���ﰲ��.��ʼʱ��, obj���ﰲ��.��ֹʱ��)
        End If
    End If

    obj���ﰲ��.��ʱ���� = (mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan)
    obj���ﰲ��.���º�����λ = mbytFun = Fun_UpdateUnit
    '�����ﰲ����Ϣ���Ƶ�δ������ѱ��漯����
    Set obj���ﰲ��.δ������ﰲ�� = obj���ﰲ��.Clone
    Set obj���ﰲ��.�ѱ�����ﰲ�� = obj���ﰲ��.Clone
    
    '�����ݿ��ȡ����
    If blnRecord Then
        If mbytPlanType = F_FixedRule Then
            '��������
            strSQL = "Select To_Char(b.��������,'yyyy-mm-dd') As ��������" & vbNewLine & _
                    " From �ٴ����ﰲ�� A, �ٴ������¼ B" & vbNewLine & _
                    " Where a.Id = b.����id And a.����Id = [1] And a.��ԴID = [2]" & vbNewLine & _
                    "       And b.�ϰ�ʱ�� Is Not Null And b.�������� Between [3] And [4]" & vbNewLine & _
                    " Group By To_Char(b.��������,'yyyy-mm-dd')"
            Set rs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ", mlng����ID, mlng��ԴId, _
                CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            '����ʱ�����Ϣ
            strSQL = "Select b.ID As ��¼ID,To_Char(b.��������,'yyyy-mm-dd') As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, b.��ʼʱ��, b.��ֹʱ��," & vbNewLine & _
                    "        b.�޺���, b.�ѹ���, b.��Լ��, b.��Լ��, b.���﷽ʽ, b.ԤԼ����, b.�Ƿ���ʱ����," & vbNewLine & _
                    "        b.����ҽ������, b.����ID, b.��ĿID, c.���� As ��Ŀ����, b.ҽ��ID, b.ҽ������, b.�Ƿ��ռ," & vbNewLine & _
                    "        b.ͣ�￪ʼʱ��, b.ͣ����ֹʱ��, b.ͣ��ԭ��" & vbNewLine & _
                    " From �ٴ����ﰲ�� A, �ٴ������¼ B, �շ���ĿĿ¼ C" & vbNewLine & _
                    " Where a.Id = b.����id And b.��ĿID = c.Id And a.����Id = [1] And a.��ԴID = [2]" & vbNewLine & _
                    "       And b.�ϰ�ʱ�� Is Not Null And b.�������� Between [3] And [4]"
            Set rs�����¼ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", mlng����ID, mlng��ԴId, _
                CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            '��������
            strSQL = "Select c.ID As ��¼ID,a.����ID, b.����" & vbNewLine & _
                    " From �ٴ����ﰲ�� D,�ٴ��������Ҽ�¼ A, �ٴ������¼ C, �������� B" & vbNewLine & _
                    " Where a.��¼ID = c.ID And a.����id = b.Id And d.����Id = [1] And d.��ԴID = [2]" & vbNewLine & _
                    "       And c.�������� Between [3] And [4]"
            Set rs�������� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ٴ���������", mlng����ID, mlng��ԴId, _
                CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            '������Ϣ
            '"       And (a.��ʼʱ�� <> a. ��ֹʱ�� Or a.��ʼʱ�� Is Null And a. ��ֹʱ�� Is Null)" & vbNewLine & _'��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
            strSQL = "Select b.ID As ��¼ID,a.���, a.��ʼʱ��, a. ��ֹʱ��, a.����, a.�Ƿ�ԤԼ, a.�Ƿ�ͣ��" & vbNewLine & _
                    " From �ٴ����ﰲ�� D,�ٴ�������ſ��� A,�ٴ������¼ B" & vbNewLine & _
                    " Where a.��¼ID = b.ID And d.����Id = [1] And d.��ԴID = [2]" & vbNewLine & _
                    "       And b.�������� Between [3] And [4] " & vbNewLine & _
                    "       And (a.��ʼʱ�� <> a. ��ֹʱ�� Or a.��ʼʱ�� Is Null And a. ��ֹʱ�� Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(b.�Ƿ��ʱ��,0)=1 And Nvl(b.�Ƿ���ſ���,0)=0)" & vbNewLine & _
                    "               Or Nvl(b.�Ƿ��ʱ��,0)=1 And Nvl(b.�Ƿ���ſ���,0)=0 And a.ԤԼ˳��� IS NULL)"
            Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", mlng����ID, mlng��ԴId, _
                CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            'ԤԼ�Һſ���
            strSQL = "Select c.ID As ��¼ID,a.����,a.����,a.����,a.���Ʒ�ʽ, a.���, b.��ʼʱ��, b.��ֹʱ��, a.����, b.�Ƿ�ԤԼ, b.�Ƿ�ͣ��" & vbNewLine & _
                    " From �ٴ����ﰲ�� D,�ٴ�����Һſ��Ƽ�¼ A, �ٴ�������ſ��� B,�ٴ������¼ C" & vbNewLine & _
                    " Where a.��¼id = b.��¼id(+) And a.��� = b.���(+)  And a.��¼ID=c.ID" & vbNewLine & _
                    "       And d.����Id = [1] And d.��ԴID = [2] And c.�������� Between [3] And [4]" & vbNewLine & _
                    "       And (b.��ʼʱ�� <> b. ��ֹʱ�� Or b.��ʼʱ�� Is Null And b. ��ֹʱ�� Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(c.�Ƿ��ʱ��,0)=1 And Nvl(c.�Ƿ���ſ���,0)=0) " & vbNewLine & _
                    "               Or Nvl(c.�Ƿ��ʱ��,0)=1 And Nvl(c.�Ƿ���ſ���,0)=0 And b.ԤԼ˳��� IS NULL)"
            Set rs������λ������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", mlng����ID, mlng��ԴId, _
                CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))
        ElseIf mlng����ID <> 0 Then
            '��������
            strSQL = "Select To_Char(b.��������,'yyyy-mm-dd') As ��������" & _
                    " From �ٴ������¼ B" & vbNewLine & _
                    " Where b.����id = [1] And b.�ϰ�ʱ�� Is Not Null And b.�������� Between [2] And [3]" & vbNewLine & _
                    " Group By To_Char(b.��������,'yyyy-mm-dd')"
            Set rs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ", lng����ID, _
                CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            '����ʱ�����Ϣ
            strSQL = "Select b.ID As ��¼ID,To_Char(b.��������,'yyyy-mm-dd') As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, b.��ʼʱ��, b.��ֹʱ��," & vbNewLine & _
                    "        b.�޺���, b.�ѹ���, b.��Լ��, b.��Լ��, b.���﷽ʽ, b.ԤԼ����, b.�Ƿ���ʱ����," & vbNewLine & _
                    "        b.����ҽ������, b.����ID, b.��ĿID, c.���� As ��Ŀ����, b.ҽ��ID, b.ҽ������, b.�Ƿ��ռ," & vbNewLine & _
                    "        b.ͣ�￪ʼʱ��, b.ͣ����ֹʱ��, b.ͣ��ԭ��" & vbNewLine & _
                    " From �ٴ������¼ B, �շ���ĿĿ¼ C" & vbNewLine & _
                    " Where b.��ĿID = c.Id And b.����Id = [1] And b.�ϰ�ʱ�� Is Not Null And b.�������� Between [2] And [3]"
            Set rs�����¼ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng����ID, CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            '��������
            strSQL = "Select c.ID As ��¼ID,a.����ID, b.����" & vbNewLine & _
                    " From �ٴ��������Ҽ�¼ A, �ٴ������¼ C, �������� B" & vbNewLine & _
                    " Where a.��¼ID = c.ID And a.����id = b.Id And c.����ID = [1] And c.�������� Between [2] And [3]"
            Set rs�������� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ٴ���������", lng����ID, CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            '������Ϣ
            '"       And (a.��ʼʱ�� <> a. ��ֹʱ�� Or a.��ʼʱ�� Is Null And a. ��ֹʱ�� Is Null)" & vbNewLine & _'��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
            strSQL = "Select b.ID As ��¼ID,a.���, a.��ʼʱ��, a. ��ֹʱ��, a.����, a.�Ƿ�ԤԼ, a.�Ƿ�ͣ��" & vbNewLine & _
                    " From �ٴ�������ſ��� A,�ٴ������¼ B" & vbNewLine & _
                    " Where a.��¼ID = b.ID And b.����ID=[1] And b.�������� Between [2] And [3] " & vbNewLine & _
                    "       And (a.��ʼʱ�� <> a. ��ֹʱ�� Or a.��ʼʱ�� Is Null And a. ��ֹʱ�� Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(b.�Ƿ��ʱ��,0)=1 And Nvl(b.�Ƿ���ſ���,0)=0) Or Nvl(b.�Ƿ��ʱ��,0)=1 And Nvl(b.�Ƿ���ſ���,0)=0 And a.ԤԼ˳��� IS NULL)"
            Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))

            'ԤԼ�Һſ���
            strSQL = "Select c.ID As ��¼ID,a.����,a.����,a.����,a.���Ʒ�ʽ, a.���, b.��ʼʱ��, b.��ֹʱ��, a.����, b.�Ƿ�ԤԼ, b.�Ƿ�ͣ��" & vbNewLine & _
                    " From �ٴ�����Һſ��Ƽ�¼ A, �ٴ�������ſ��� B,�ٴ������¼ C" & vbNewLine & _
                    " Where a.��¼id = b.��¼id(+) And a.��� = b.���(+)  And a.��¼ID=c.ID" & vbNewLine & _
                    "       And c.����id = [1] And c.�������� Between [2] And [3]" & vbNewLine & _
                    "       And (b.��ʼʱ�� <> b. ��ֹʱ�� Or b.��ʼʱ�� Is Null And b. ��ֹʱ�� Is Null)" & vbNewLine & _
                    "       And (Not(Nvl(c.�Ƿ��ʱ��,0)=1 And Nvl(c.�Ƿ���ſ���,0)=0) Or Nvl(c.�Ƿ��ʱ��,0)=1 And Nvl(c.�Ƿ���ſ���,0)=0 And b.ԤԼ˳��� IS NULL)"
            Set rs������λ������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, CDate(obj���ﰲ��.��ʼʱ��), CDate(obj���ﰲ��.��ֹʱ��))
        End If
    ElseIf mlng����ID <> 0 Then
        '������Ŀ
        strSQL = "Select " & IIf(mbytPlanType = F_MonthTemplet, "'1900-01-' || Replace(b.������Ŀ, '��', '')", "b.������Ŀ") & " As ��������" & _
                " From �ٴ��������� B" & vbNewLine & _
                " Where b.����id = [1] And b.�ϰ�ʱ�� Is Not Null" & vbNewLine & _
                " Group By b.������Ŀ"
        Set rs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ", lng����ID)

        'ʱ�����Ϣ
        strSQL = "Select b.ID as ��¼ID," & IIf(mbytPlanType = F_MonthTemplet, "'1900-01-' || Replace(b.������Ŀ, '��', '')", "b.������Ŀ") & " As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, NULL as ��ʼʱ��, NULL as ��ֹʱ��," & vbNewLine & _
                "        b.�޺���, 0 as �ѹ���, b.��Լ��, 0 as ��Լ��, b.���﷽ʽ, b.ԤԼ����, 0 As �Ƿ���ʱ����," & vbNewLine & _
                "        '' As ����ҽ������,0 As ����ID, 0 as ��ĿID, '' As ��Ŀ����, 0 as ҽ��ID, '' As ҽ������, b.�Ƿ��ռ," & vbNewLine & _
                "        NULL as ͣ�￪ʼʱ��, NULL as ͣ����ֹʱ��, NULL as ͣ��ԭ��" & vbNewLine & _
                " From �ٴ��������� B" & vbNewLine & _
                " Where b.����id = [1] And b.�ϰ�ʱ�� Is Not Null"
        Set rs�����¼ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng����ID)

        '��������
        strSQL = "Select c.ID As ��¼ID,a.����ID, b.����" & vbNewLine & _
                " From �ٴ��������� A, �ٴ��������� C, �������� B" & vbNewLine & _
                " Where a.����ID=c.ID And a.����id = b.Id And c.����id = [1]"
        Set rs�������� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ٴ���������", lng����ID)

        '������Ϣ
        strSQL = "Select b.ID As ��¼ID,a.���, a.��ʼʱ��, a. ��ֹʱ��, a.��������  As ����, a.�Ƿ�ԤԼ, 0 As �Ƿ�ͣ��" & vbNewLine & _
                " From �ٴ�����ʱ�� A,�ٴ��������� B" & vbNewLine & _
                " Where a.����ID=b.ID And b.����ID = [1]"
        Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID)

        'ԤԼ�Һſ���
        strSQL = "Select c.ID As ��¼ID,a.����,a.����,a.����,a.���Ʒ�ʽ, a.���, b.��ʼʱ��, b.��ֹʱ��, a.����, b.�Ƿ�ԤԼ, 0 As �Ƿ�ͣ��" & vbNewLine & _
                " From �ٴ�����Һſ��� A, �ٴ�����ʱ�� B,�ٴ��������� C" & vbNewLine & _
                " Where a.����ID = b.����ID(+) And a.��� = b.���(+) And a.����ID=c.ID " & vbNewLine & _
                "       And c.����ID = [1]"
        Set rs������λ������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������λ�Һſ���", lng����ID)
    End If

    'ת���ɶ���
    If Not rs������Ŀ Is Nothing Then
        Do While Not rs������Ŀ.EOF
            Set obj�����¼�� = New �����¼��
            rs�����¼.Filter = "��������='" & Nvl(rs������Ŀ!��������) & "'"
            Do While Not rs�����¼.EOF
                Set obj�����¼ = GetVisitTimesObject(rs�����¼)

                '�ϰ�ʱ��
                If obj�����ϰ�ʱ��.Exits("K" & obj�����¼.ʱ���) Then
                    Set obj�����¼.�ϰ�ʱ�� = obj�����ϰ�ʱ��("K" & obj�����¼.ʱ���).Clone
                Else
                    '�����¼ʱ���ϰ�ʱ�ο����ѱ�ɾ��
                    Set obj�����¼.�ϰ�ʱ�� = New �ϰ�ʱ��
                    With obj�����¼.�ϰ�ʱ��
                        .��ʼʱ�� = obj�����¼.��ʼʱ��
                        .����ʱ�� = obj�����¼.��ֹʱ��
                    End With
                End If

                '��������
                rs��������.Filter = "��¼ID=" & obj�����¼.��¼ID
                Set obj�����¼.�����������Ҽ� = GetVisitRoomsObjects(rs��������)
                obj�����¼.�����������Ҽ�.���﷽ʽ = obj�����¼.���﷽ʽ
                obj�����¼.�����������Ҽ�.ҽ������ = obj���ﰲ��.�����Դ.ҽ������

                '������Ϣ
                rs������Ϣ.Filter = "��¼ID=" & obj�����¼.��¼ID
                Set obj�����¼.������Ϣ�� = GetTimeIntervalObjects(rs������Ϣ)
                With obj�����¼.������Ϣ��
                    .����Ƶ�� = obj���ﰲ��.�����Դ.����Ƶ��
                    .�Ƿ��ʱ�� = obj�����¼.�Ƿ��ʱ��
                    .�Ƿ���ſ��� = obj�����¼.�Ƿ���ſ���
                    .�޺��� = obj�����¼.�޺���
                    .��Լ�� = obj�����¼.��Լ��
                    .ԤԼ���� = obj�����¼.ԤԼ����
                End With

                '������λԤԼ����
                Set obj�����¼.������λ���Ƽ� = New ������λ���Ƽ�
                obj�����¼.������λ���Ƽ�.�Ƿ��ռ = Val(Nvl(rs�����¼!�Ƿ��ռ))
                Set obj������Ϣ�� = Nothing
                strTemp = ""

                rs������λ������Ϣ.Filter = "��¼ID=" & obj�����¼.��¼ID
                rs������λ������Ϣ.Sort = "����,����,����,���"
                Do While Not rs������λ������Ϣ.EOF
                    If strTemp <> Nvl(rs������λ������Ϣ!����) & "-" & Nvl(rs������λ������Ϣ!����) & "-" & Nvl(rs������λ������Ϣ!����) Then
                        If Not obj������Ϣ�� Is Nothing Then
                            Set obj������λ����.������Ϣ�� = obj������Ϣ��
                            obj�����¼.������λ���Ƽ�.AddItem obj������λ����, "K" & obj������λ����.������λ����
                        End If
                        Set obj������λ���� = New ������λ����
                        obj������λ����.������λ���� = Nvl(rs������λ������Ϣ!����)
                        obj������λ����.���� = Val(Nvl(rs������λ������Ϣ!����))
                        obj������λ����.ԤԼ���Ʒ�ʽ = Val(Nvl(rs������λ������Ϣ!���Ʒ�ʽ))
                        Set obj������Ϣ�� = New ������Ϣ��

                        strTemp = Nvl(rs������λ������Ϣ!����) & "-" & Nvl(rs������λ������Ϣ!����) & "-" & Nvl(rs������λ������Ϣ!����)
                    End If

                    obj������Ϣ��.AddItem GetTimeIntervalObject(rs������λ������Ϣ)
                    rs������λ������Ϣ.MoveNext
                Loop
                If Not obj������Ϣ�� Is Nothing Then
                    Set obj������λ����.������Ϣ�� = obj������Ϣ��
                    obj�����¼.������λ���Ƽ�.AddItem obj������λ����, "K" & obj������λ����.������λ����
                End If
                
                obj�����¼.�Ƿ�̶� = False
                If mbytFun = Fun_TempPlanRecord Then '�ڽ�����ʱ����ʱ������ɾ��
                    obj�����¼.�Ƿ�̶� = True
                ElseIf mbytFun = Fun_UpdatePlan Then '��ͣ����ѱ����ڹҺ�ԤԼ�Ĳ��ܵ���
                    If CheckPlanIsStopOrUsed(obj�����¼.��¼ID) Then
                        obj�����¼.�Ƿ�̶� = True
                    ElseIf CDate(obj�����¼.��ֹʱ��) < dtNow Then '��ֹʱ����С�ڵ�ǰʱ��Ĳ��ܵ���
                        obj�����¼.�Ƿ�̶� = True
                    End If
                End If

                obj�����¼��.AddItem obj�����¼, "K" & obj�����¼.ʱ���
                rs�����¼.MoveNext
            Loop

            obj�����¼��.�������� = Format(Nvl(rs������Ŀ!��������), "yyyy-mm-dd")
            obj�����¼��.�Ƿ�ɾ�� = False

            obj���ﰲ��.�ѱ�����ﰲ��.AddItem obj�����¼��, "K" & obj�����¼��.��������
            If obj�����¼��.�������� = Format(mstrȱʡ����, "yyyy-mm-dd") Then
                obj���ﰲ��.AddItem obj�����¼��.Clone, "K" & obj�����¼��.��������
            End If

            rs������Ŀ.MoveNext
        Loop
    End If
    
    'ȱʡѡ������
    If Not (mbytFun = Fun_AddSignalSourcePlan Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel) Then
        obj���ﰲ��.ȱʡ�������� = mstrȱʡ����
        If obj���ﰲ��.Count = 0 And mbytPlanType <> F_MonthTemplet Then
            'ģ����������ʱ��ȱʡ����һ��������Ŀ
            If obj���ﰲ��.�ѱ�����ﰲ��.Count > 0 And Not (obj���ﰲ��.�Ű���� = 0 Or obj���ﰲ��.�Ű���� = 1) Then
                obj���ﰲ��.AddItem obj���ﰲ��.�ѱ�����ﰲ��(1).Clone, GetPlanKey(obj���ﰲ��.�ѱ�����ﰲ��(1).��������)
                obj���ﰲ��.ȱʡ�������� = obj���ﰲ��.�ѱ�����ﰲ��(1).��������
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
    'blnReLoadTimeRange �Ƿ����¼���ʱ�䷶Χ
    Dim obj�ϰ�ʱ�� As �ϰ�ʱ��
    Dim objListItem As ListItem
    
    On Error GoTo Errhand

    '��Чʱ�䷶Χ
    If mbytPlanType = F_FixedRule And mblnFirst Then
        dtpBegin.Value = Format(mobj���ﰲ��.��ʼʱ��, "yyyy-MM-dd hh:mm:ss")
        dtpEnd.Value = Format(mobj���ﰲ��.��ֹʱ��, "yyyy-MM-dd hh:mm:ss")
        
        dtpBegin.Tag = dtpBegin.Value: dtpEnd.Tag = dtpEnd.Value
    End If
    
    '��������
    cldsCalenbarSel.LoadData mobj���ﰲ��

    '����ʱ��
    lvwWorkTime.ListItems.Clear
    If Not mobj���ﰲ��.�����ϰ�ʱ�� Is Nothing Then
        For Each obj�ϰ�ʱ�� In mobj���ﰲ��.�����ϰ�ʱ��
            Set objListItem = lvwWorkTime.ListItems.Add(, "K" & obj�ϰ�ʱ��.ʱ���, obj�ϰ�ʱ��.ʱ��� & _
                "(" & Format(obj�ϰ�ʱ��.��ʼʱ��, "hh:mm") & "-" & Format(obj�ϰ�ʱ��.����ʱ��, "hh:mm") & ")")
            objListItem.SubItems(1) = obj�ϰ�ʱ��.��ʼʱ��
            objListItem.SubItems(2) = obj�ϰ�ʱ��.����ʱ��
            objListItem.Tag = obj�ϰ�ʱ��.ʱ���
            '����ɫ�����Ƿ��Դ�����õ�ʱ���
            If mobj���ﰲ��.��Դ����.Exits("K" & obj�ϰ�ʱ��.ʱ���) Then
                objListItem.ForeColor = vbBlue
            End If
        Next
    End If
    If Not lvwWorkTime.SelectedItem Is Nothing Then
        lvwWorkTime.SelectedItem.Selected = False 'ȥ��ѡ�����
    End If

    '��Դ��Ϣ
    SourceInfor.LoadData mobj���ﰲ��.�����Դ

    '���ﰲ��
    Call LoadDetailData

    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadDetailData() As Boolean
    Dim objTemp As �����¼��, i As Integer

    On Error GoTo Errhand

    Screen.MousePointer = vbHourglass
    '��ǰ��Ŀ
    If mobj���ﰲ��.Count = 0 Then
        Set objTemp = New �����¼��
    Else
        Set objTemp = mobj���ﰲ��(1).Clone
    End If

    mstrCurDay = objTemp.��������
    Call SetTitleText

    'ʱ���
    CheckWorkTime objTemp
    
    '����Ԥ��
    Call ShowPlan(mobj���ﰲ��)
    
    '����
    CPDPages.LoadData objTemp, mobj���ﰲ��.���з�������, mobj���ﰲ��.���к�����λ, True

    '�ָ�Ӧ��
    If mbytPlanType = F_MonthPlan Or mbytPlanType = F_MonthTemplet Then
        optRule(0).Value = True
        mblnNotClick = True
        For i = chkWeek.LBound To chkWeek.UBound
            chkWeek(i).Value = vbUnchecked
        Next
        mblnNotClick = False
    End If

    SetEnabled Not (mbytFun = Fun_View _
            Or mbytFun = Fun_UpdateUnit Or mlng��ԴId = 0 _
            Or mbytFun = Fun_TempPlanVerify _
            Or mbytFun = Fun_TempPlanCancel)
    CPDPages.EditMode = IIf(mbytFun = Fun_View Or mlng��ԴId = 0 Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel, _
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

Private Function SavePlanData(ByVal lng����ID As Long, ByVal obj���ﰲ�� As ���ﰲ��, _
    ByVal obj�����Դ As �����Դ, cllPro As Collection, ByVal blnRecord As Boolean, _
    ByVal dtCurdate As Date, ByVal obj�ѱ�����ﰲ�� As ���ﰲ��) As Boolean
    '���ܣ����氲����ϸ����
    Dim strSQL As String
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim strTemp As String, lngTemp As Long, lng��¼ID As Long
    Dim str���� As String, obj���� As ��������, byt���﷽ʽ As Byte
    Dim str���� As String, obj���� As ������Ϣ, cll���� As Collection
    Dim obj������λ As ������λ����, bln�Ƿ��ռ As Boolean
    Dim i As Long, blnPublished As Boolean
        
    Dim obj�����¼��1 As �����¼��, obj�����¼1 As �����¼
    Dim blnFind As Boolean, blnFindPlan As Boolean
    Dim lngSavedCount As Long, lngNewCount As Long

    Err = 0: On Error GoTo ErrHandler
    If obj���ﰲ�� Is Nothing Then SavePlanData = True: Exit Function
    If obj���ﰲ��.Count = 0 Then SavePlanData = True: Exit Function

    If cllPro Is Nothing Then Set cllPro = New Collection
    
    If Not obj�ѱ�����ﰲ�� Is Nothing Then
        '�ѱ��氲�����ȫ����ɾ������û���µİ��ţ���ɾ���ٴ����ﰲ��
        For Each obj�����¼�� In obj�ѱ�����ﰲ��
            lngSavedCount = lngSavedCount + obj�����¼��.Count
        Next
        For Each obj�����¼�� In obj���ﰲ��
            lngNewCount = lngNewCount + obj�����¼��.Count
        Next
        
        '������ѱ����еĳ����¼�������޸ĺ��еĳ����¼
        'ע�⣺����������ѱ����У�������δ�����У����ʾδ���й��鿴���϶�û���޸ģ�����ɾ��
        For Each obj�����¼�� In obj�ѱ�����ﰲ��
            blnFindPlan = False
            If obj�����¼��.�Ƿ�ɾ�� = False Then ' ȫ��ɾ��������ǰ�洦��
                For Each obj�����¼ In obj�����¼��
                    blnFind = False
                    For Each obj�����¼��1 In obj���ﰲ��
                        If blnFind Then Exit For
                        If obj�����¼��1.�������� = obj�����¼��.�������� Then
                            blnFindPlan = True 'δ������δ�ҵ���ʾû���޸�
                            For Each obj�����¼1 In obj�����¼��1
                                '�ü�¼ID�ж�
                                'If obj�����¼1.ʱ��� = obj�����¼.ʱ��� Then
                                If obj�����¼1.��¼ID = obj�����¼.��¼ID Then
                                    blnFind = True: Exit For
                                End If
                            Next
                        End If
                    Next
                    
                    If blnFindPlan And blnFind = False Then
                        lngSavedCount = lngSavedCount - 1
                        'Zl_�ٴ������ϰ�ʱ��_Delete(
                        strSQL = "Zl_�ٴ������ϰ�ʱ��_Delete("
                        '����id_In       �ٴ���������.����id%Type,
                        strSQL = strSQL & "" & lng����ID & ","
                        '��Ŀ_In         �ٴ���������.������Ŀ%Type,
                        strSQL = strSQL & "'" & IIf(mbytPlanType = F_MonthTemplet And blnRecord = False, _
                            FormatApplyToStr(obj�����¼��.��������), obj�����¼��.��������) & "',"
                        '�����¼_In     Number := 0,
                        strSQL = strSQL & "" & IIf(blnRecord, 1, 0) & ","
                        '�ϰ�ʱ��_In     �ٴ���������.�ϰ�ʱ��%Type,
                        strSQL = strSQL & "'" & obj�����¼.ʱ��� & "',"
                        'ɾ�����ﰲ��_In Number:=0
                        strSQL = strSQL & "" & IIf(lngSavedCount = 0 And lngNewCount = 0, 1, 0) & ")"
                        cllPro.Add strSQL
                    End If
                Next
            End If
        Next
    End If
    
    '�����¼
    If blnRecord Then
        For Each obj�����¼�� In obj���ﰲ��
            '��������¼
            For Each obj�����¼ In obj�����¼��
                '�̶���δ�޸ģ�������
                If obj�����¼.�Ƿ�̶� = False Then
                    lng��¼ID = obj�����¼.��¼ID
                    If lng��¼ID = 0 Then lng��¼ID = zlDatabase.GetNextId("�ٴ������¼")
                    bln�Ƿ��ռ = obj�����¼.������λ���Ƽ�.�Ƿ��ռ
                    obj�����¼.��ʼʱ�� = Format(obj�����¼��.��������, "yyyy-mm-dd ") & Format(obj�����¼.�ϰ�ʱ��.��ʼʱ��, "hh:mm:ss")
                    obj�����¼.��ֹʱ�� = GetWorkTrueDate(obj�����¼.��ʼʱ��, obj�����¼.�ϰ�ʱ��.����ʱ��)

                    '��������
                    byt���﷽ʽ = obj�����¼.�����������Ҽ�.���﷽ʽ
                    str���� = ""
                    For Each obj���� In obj�����¼.�����������Ҽ�
                        '����_In:����1,����2,...
                        str���� = str���� & "," & obj����.����ID
                    Next
                    If str���� <> "" Then str���� = Mid(str����, 2)

                    '����ʱ��
                    Set cll���� = New Collection: str���� = ""
                    For Each obj���� In obj�����¼.������Ϣ��
                        strTemp = obj����.��� & "," & _
                            GetWorkTrueDate(obj�����¼.��ʼʱ��, ZDate(obj����.��ʼʱ��, obj�����¼.��ʼʱ��, False), , False) & "," & _
                            GetWorkTrueDate(obj�����¼.��ʼʱ��, ZDate(obj����.��ֹʱ��, obj�����¼.��ֹʱ��, False)) & "," & _
                            obj����.���� & "," & IIf(obj����.�Ƿ�ԤԼ, 1, 0)
                        If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                            'ʱ��_In:���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
                            str���� = Mid(str����, 2)
                            cll����.Add str����
                            str���� = ""
                        End If
                        str���� = str���� & "|" & strTemp
                    Next
                    If str���� <> "" Then
                        str���� = Mid(str����, 2)
                        cll����.Add str����
                    End If
                    For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                        'Zl_�ٴ������¼_Insert(
                        strSQL = "Zl_�ٴ������¼_Insert("
                        'Id_In           �ٴ������¼.Id%Type,
                        strSQL = strSQL & "" & lng��¼ID & ","
                        '����id_In       �ٴ���������.����id%Type,
                        strSQL = strSQL & "" & lng����ID & ","
                        '��Դid_In       �ٴ������¼.��Դid%Type,
                        strSQL = strSQL & "" & obj�����Դ.ID & ","
                        '��������_In     �ٴ������¼.��������%Type,
                        strSQL = strSQL & "To_Date('" & obj�����¼��.�������� & "','yyyy-mm-dd'),"
                        '�ϰ�ʱ��_In     �ٴ������¼.�ϰ�ʱ��%Type,
                        strSQL = strSQL & "'" & obj�����¼.ʱ��� & "',"
                        '��ʼʱ��_In     �ٴ������¼.��ʼʱ��%Type,
                        strSQL = strSQL & "" & ZDate(obj�����¼.��ʼʱ��) & ","
                        '��ֹʱ��_In     �ٴ������¼.��ֹʱ��%Type,
                        strSQL = strSQL & "" & ZDate(obj�����¼.��ֹʱ��) & ","
                        'ȱʡԤԼʱ��_In �ٴ������¼.ȱʡԤԼʱ��%Type,
                        strSQL = strSQL & "" & ZDate(GetWorkTrueDate(obj�����¼.��ʼʱ��, obj�����¼.�ϰ�ʱ��.ȱʡԤԼʱ��)) & ","
                        '��ǰ�Һ�ʱ��_In �ٴ������¼.��ǰ�Һ�ʱ��%Type,
                        strSQL = strSQL & "" & ZDate(GetWorkTrueDate(obj�����¼.��ʼʱ��, obj�����¼.�ϰ�ʱ��.��ǰ�Һ�ʱ��, False)) & ","
                        '�޺���_In       �ٴ������¼.�޺���%Type,
                        strSQL = strSQL & "" & ZVal(obj�����¼.�޺���) & ","
                        '��Լ��_In       �ٴ������¼.��Լ��%Type,
                        strSQL = strSQL & "" & ZVal(obj�����¼.��Լ��) & ","
                        '�Ƿ���ſ���_In �ٴ������¼.�Ƿ���ſ���%Type,
                        strSQL = strSQL & "" & IIf(obj�����¼.�Ƿ���ſ���, 1, 0) & ","
                        '�Ƿ��ʱ��_In   �ٴ������¼.�Ƿ��ʱ��%Type,
                        strSQL = strSQL & "" & IIf(obj�����¼.�Ƿ��ʱ��, 1, 0) & ","
                        'ԤԼ����_In     �ٴ������¼.ԤԼ����%Type,
                        strSQL = strSQL & "" & obj�����¼.ԤԼ���� & ","
                        '�Ƿ��ռ_In     �ٴ������¼.�Ƿ��ռ%Type,
                        strSQL = strSQL & "" & IIf(bln�Ƿ��ռ, 1, 0) & ","
                        '��Ŀid_In       �ٴ������¼.��Ŀid%Type,
                        strSQL = strSQL & "" & ZVal(obj�����Դ.��ĿID) & ","
                        '����id_In       �ٴ������¼.����id%Type,
                        strSQL = strSQL & "" & obj�����Դ.����ID & ","
                        'ҽ��id_In       �ٴ������¼.ҽ��id%Type,
                        strSQL = strSQL & "" & ZVal(obj�����Դ.ҽ��ID) & ","
                        'ҽ������_In     �ٴ������¼.ҽ������%Type,
                        strTemp = obj�����Դ.ҽ������
                        strSQL = strSQL & "" & IIf(strTemp = "", "NULL", "'" & strTemp & "'") & ","
                        '���﷽ʽ_In     �ٴ������¼.���﷽ʽ%Type,
                        strSQL = strSQL & "" & byt���﷽ʽ & ","
                        '�Ƿ���ʱ����_In �ٴ������¼.�Ƿ���ʱ����%Type,
                        strSQL = strSQL & "" & IIf(mbytFun = Fun_UpdatePlan, "NULL", IIf(mbytFun = Fun_TempPlanRecord, 1, 0)) & ","
                        '�Ǽ���_In       �ٴ������¼.�Ǽ���%Type,
                        strSQL = strSQL & IIf(mbytFun = Fun_UpdatePlan, "NULL", "'" & UserInfo.���� & "'") & ","
                        '�Ǽ�ʱ��_In     �ٴ������¼.�Ǽ�ʱ��%Type,
                        strSQL = strSQL & IIf(mbytFun = Fun_UpdatePlan, "NULL", ZDate(dtCurdate)) & ","
                        '�Ƿ񷢲�_In     �ٴ������¼.�Ƿ񷢲�%Type,
                        '1.��ʱ����
                        '2.�°��Ż��ܰ������Ӻ�Դ
                        blnPublished = mbytFun = Fun_TempPlanRecord _
                                    Or mbytFun = Fun_AddSignalSourcePlan And (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan)
                        strSQL = strSQL & IIf(mbytFun = Fun_UpdatePlan, "NULL", IIf(blnPublished, 1, 0)) & ","
                        '����_In         Varchar2 := Null,
                        strSQL = strSQL & "'" & str���� & "',"
                        'ʱ��_In         Varchar2 := Null,
                        str���� = ""
                        If cll����.Count > 0 Then str���� = cll����(i)
                        strSQL = strSQL & "'" & str���� & "',"
                        'ɾ�����_In Number:=0
                        strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                        cllPro.Add strSQL
                    Next
                    '����Һſ���
                    For Each obj������λ In obj�����¼.������λ���Ƽ�
                        'ԤԼ����:0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                        '����:1-��������;2-ԤԼ��ʽ
                        Set cll���� = New Collection: str���� = ""
                        For Each obj���� In obj������λ.������Ϣ��
                            strTemp = obj����.��� & "," & obj����.����
                            If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                                '���ſ���_in:���1,����|���2,����|...
                                str���� = Mid(str����, 2)
                                cll����.Add str����
                                str���� = ""
                            End If
                            str���� = str���� & "|" & strTemp
                        Next
                        If str���� <> "" Then
                            str���� = Mid(str����, 2)
                            cll����.Add str����
                        End If
                        For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                            'Zl_�ٴ�����Һſ��Ƽ�¼_Insert(
                            strSQL = "Zl_�ٴ�����Һſ��Ƽ�¼_Insert("
                            '��¼id_In   �ٴ�����Һſ��Ƽ�¼.��¼id%Type,
                            strSQL = strSQL & "" & lng��¼ID & ","
                            '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                            strSQL = strSQL & "" & obj������λ.���� & ","
                            '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                            strSQL = strSQL & "" & 1 & ","
                            '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                            strSQL = strSQL & "'" & obj������λ.������λ���� & "',"
                            '���Ʒ�ʽ_In �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type,
                            strSQL = strSQL & "" & obj������λ.ԤԼ���Ʒ�ʽ & ","
                            '�Ƿ��ռ_In �ٴ������¼.�Ƿ��ռ%Type,
                            strSQL = strSQL & "" & IIf(obj�����¼.������λ���Ƽ�.�Ƿ��ռ, 1, 0) & ","
                            '���ſ���_In Varchar2,
                            str���� = ""
                            If cll����.Count > 0 Then str���� = cll����(i)
                            strSQL = strSQL & "'" & str���� & "',"
                            'ɾ��_In Number:=0
                            strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                            cllPro.Add strSQL
                        Next
                    Next
                End If
            Next
        Next
    Else  'ģ��/�̶�����
        For Each obj�����¼�� In obj���ﰲ��
            '����������Ŀ
            For Each obj�����¼ In obj�����¼��
                lng��¼ID = obj�����¼.��¼ID
                If lng��¼ID = 0 Then lng��¼ID = zlDatabase.GetNextId("�ٴ���������")
                bln�Ƿ��ռ = obj�����¼.������λ���Ƽ�.�Ƿ��ռ
                '��������
                byt���﷽ʽ = obj�����¼.�����������Ҽ�.���﷽ʽ
                str���� = ""
                For Each obj���� In obj�����¼.�����������Ҽ�
                    '����_In:����1,����2,...
                    str���� = str���� & "," & obj����.����ID
                Next
                If str���� <> "" Then str���� = Mid(str����, 2)
    
                '����ʱ��
                Set cll���� = New Collection: str���� = ""
                For Each obj���� In obj�����¼.������Ϣ��
                    strTemp = obj����.��� & ","
                    strTemp = strTemp & GetWorkTrueDate(obj�����¼.��ʼʱ��, ZDate(obj����.��ʼʱ��, obj�����¼.��ʼʱ��, False), , False) & ","
                    strTemp = strTemp & GetWorkTrueDate(obj�����¼.��ʼʱ��, ZDate(obj����.��ֹʱ��, obj�����¼.��ֹʱ��, False)) & ","
                    strTemp = strTemp & obj����.���� & "," & IIf(obj����.�Ƿ�ԤԼ, 1, 0)
    
                    If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                        'ʱ��_In:���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
                        str���� = Mid(str����, 2)
                        cll����.Add str����
                        str���� = ""
                    End If
                    str���� = str���� & "|" & strTemp
                Next
                If str���� <> "" Then
                    str���� = Mid(str����, 2)
                    cll����.Add str����
                End If
                For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                    'Zl_�ٴ���������_Insert(
                    strSQL = "Zl_�ٴ���������_Insert("
                    'Id_In           �ٴ���������.Id%Type,
                    strSQL = strSQL & "" & lng��¼ID & ","
                    '����id_In       �ٴ���������.����id%Type,
                    strSQL = strSQL & "" & lng����ID & ","
                    '������Ŀ_In     �ٴ���������.������Ŀ%Type,
                    strSQL = strSQL & "'" & IIf(mbytPlanType = F_MonthTemplet, FormatApplyToStr(obj�����¼��.��������), obj�����¼��.��������) & "',"
                    '�ϰ�ʱ��_In     �ٴ���������.�ϰ�ʱ��%Type,
                    strSQL = strSQL & "'" & obj�����¼.ʱ��� & "',"
                    '�޺���_In       �ٴ���������.�޺���%Type,
                    strSQL = strSQL & "" & ZVal(obj�����¼.�޺���) & ","
                    '��Լ��_In       �ٴ���������.��Լ��%Type,
                    strSQL = strSQL & "" & ZVal(obj�����¼.��Լ��) & ","
                    '�Ƿ��ʱ��_In   �ٴ���������.�Ƿ��ʱ��%Type,
                    strSQL = strSQL & "" & IIf(obj�����¼.�Ƿ��ʱ��, 1, 0) & ","
                    '�Ƿ���ſ���_In �ٴ���������.�Ƿ���ſ���%Type,
                    strSQL = strSQL & "" & IIf(obj�����¼.�Ƿ���ſ���, 1, 0) & ","
                    'ԤԼ����_In     �ٴ���������.ԤԼ����%Type,
                    strSQL = strSQL & "" & obj�����¼.ԤԼ���� & ","
                    '�Ƿ��ռ_In     �ٴ���������.�Ƿ��ռ%Type,
                    strSQL = strSQL & "" & IIf(bln�Ƿ��ռ, 1, 0) & ","
                    '���﷽ʽ_In     �ٴ���������.���﷽ʽ%Type := Null,
                    strSQL = strSQL & "" & byt���﷽ʽ & ","
                    '����_In         Varchar2 := Null,
                    strSQL = strSQL & "'" & str���� & "',"
                    'ʱ��_In         Varchar2 := Null,
                    str���� = ""
                    If cll����.Count > 0 Then str���� = cll����(i)
                    strSQL = strSQL & "'" & str���� & "',"
                    'ɾ�����_In Number:=0
                    strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                    cllPro.Add strSQL
                Next
                '����Һſ���
                For Each obj������λ In obj�����¼.������λ���Ƽ�
                    'ԤԼ����:0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                    '����:1-��������;2-ԤԼ��ʽ
                    Set cll���� = New Collection: str���� = ""
                    For Each obj���� In obj������λ.������Ϣ��
                        strTemp = obj����.��� & "," & obj����.����
                        If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                            '���ſ���_in:���1,����|���2,����|...
                            str���� = Mid(str����, 2)
                            cll����.Add str����
                            str���� = ""
                        End If
                        str���� = str���� & "|" & strTemp
                    Next
                    If str���� <> "" Then
                        str���� = Mid(str����, 2)
                        cll����.Add str����
                    End If
                    For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                        'Zl_�ٴ�����Һſ���_Insert(
                        strSQL = "Zl_�ٴ�����Һſ���_Insert("
                        '����id_In   �ٴ�����Һſ���.����id%Type,
                        strSQL = strSQL & "" & lng��¼ID & ","
                        '����_In     �ٴ�����Һſ���.����%Type,
                        strSQL = strSQL & "" & obj������λ.���� & ","
                        '����_In     �ٴ�����Һſ���.����%Type,
                        strSQL = strSQL & "" & 1 & ","
                        '����_In     �ٴ�����Һſ���.����%Type,
                        strSQL = strSQL & "'" & obj������λ.������λ���� & "',"
                        '���Ʒ�ʽ_In �ٴ�����Һſ���.���Ʒ�ʽ%Type,
                        strSQL = strSQL & "" & obj������λ.ԤԼ���Ʒ�ʽ & ","
                        '�Ƿ��ռ_In �ٴ���������.�Ƿ��ռ%Type,
                        strSQL = strSQL & "" & IIf(obj�����¼.������λ���Ƽ�.�Ƿ��ռ, 1, 0) & ","
                        '���ſ���_In Varchar2,
                        str���� = ""
                        If cll����.Count > 0 Then str���� = cll����(i)
                        strSQL = strSQL & "'" & str���� & "',"
                        'ɾ��_In Number:=0
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
    If mbytFun = Fun_AddSignalSourcePlan And mlng��ԴId <> 0 And mlngSavedRecords = 0 Then
        If mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan Then
            '�������ɳ����¼���ڿ�ʼʱɾ����δ��ʹ�õĹ̶����ŵĳ����¼��
            strSQL = "Zl1_Auto_Buildingregisterplan(Null)"
            zlDatabase.ExecuteProcedure strSQL, "�ָ������¼"
        End If
    End If
    
    Set mobj���ﰲ�� = Nothing
    Set mrsVisitedRecord = Nothing
    Set mobjͣ���¼�� = Nothing
    Set mrsVisitedRecordByDate = Nothing
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveUpdateUnit(ByRef obj���ﰲ�� As ���ﰲ��) As Boolean
    '���ܣ����������λ
    Dim strSQL As String, cllPro As Collection, strTemp As String
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim str���� As String, obj���� As ������Ϣ, cll���� As Collection
    Dim obj������λ As ������λ����, i As Integer
    Dim blnTrans As Boolean

    Err = 0: On Error GoTo ErrHandler
    Set cllPro = New Collection
    For Each obj�����¼�� In obj���ﰲ��
        Select Case mbytPlanType
        Case F_Templet, F_FixedRule, F_MonthTemplet
            For Each obj�����¼ In obj�����¼��
                If obj�����¼.������λ���Ƽ�.�Ƿ��޸� Then
                    '����Һſ���
                    For Each obj������λ In obj�����¼.������λ���Ƽ�
                        'ԤԼ����:0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                        '����:1-��������;2-ԤԼ��ʽ
                        If Not (obj�����¼.ԤԼ���� = 1 _
                            Or (obj�����¼.ԤԼ���� = 2 And obj������λ.���� = 1)) Then

                            Set cll���� = New Collection: str���� = ""
                            For Each obj���� In obj������λ.������Ϣ��
                                strTemp = obj����.��� & "," & obj����.����
                                If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                                    '���ſ���_in:���1,����|���2,����|...
                                    str���� = Mid(str����, 2)
                                    cll����.Add str����
                                    str���� = ""
                                End If
                                str���� = str���� & "|" & strTemp
                            Next
                            If str���� <> "" Then
                                str���� = Mid(str����, 2)
                                cll����.Add str����
                            End If
                            For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                                'Zl_�ٴ�����Һſ���_Insert(
                                strSQL = "Zl_�ٴ�����Һſ���_Insert("
                                '����id_In   �ٴ�����Һſ���.����id%Type,
                                strSQL = strSQL & "" & obj�����¼.��¼ID & ","
                                '����_In     �ٴ�����Һſ���.����%Type,
                                strSQL = strSQL & "" & obj������λ.���� & ","
                                '����_In     �ٴ�����Һſ���.����%Type,
                                strSQL = strSQL & "" & 1 & ","
                                '����_In     �ٴ�����Һſ���.����%Type,
                                strSQL = strSQL & "'" & obj������λ.������λ���� & "',"
                                '���Ʒ�ʽ_In �ٴ�����Һſ���.���Ʒ�ʽ%Type,
                                strSQL = strSQL & "" & obj������λ.ԤԼ���Ʒ�ʽ & ","
                                '�Ƿ��ռ_In �ٴ������¼.�Ƿ��ռ%Type,
                                strSQL = strSQL & "" & IIf(obj�����¼.������λ���Ƽ�.�Ƿ��ռ, 1, 0) & ","
                                '���ſ���_In Varchar2,
                                str���� = ""
                                If cll����.Count > 0 Then str���� = cll����(i)
                                strSQL = strSQL & "'" & str���� & "',"
                                'ɾ��_In Number:=0
                                strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                                cllPro.Add strSQL
                            Next
                        End If
                    Next
                End If
            Next
        Case Else
            For Each obj�����¼ In obj�����¼��
                If obj�����¼.������λ���Ƽ�.�Ƿ��޸� Then
                    For Each obj������λ In obj�����¼.������λ���Ƽ�
                        'ԤԼ����:0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                        '����:1-��������;2-ԤԼ��ʽ
                        If Not (obj�����¼.ԤԼ���� = 1 _
                            Or (obj�����¼.ԤԼ���� = 2 And obj������λ.���� = 1)) Then

                            Set cll���� = New Collection: str���� = ""
                            For Each obj���� In obj������λ.������Ϣ��
                                strTemp = obj����.��� & "," & obj����.����
                                If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                                    '���ſ���_in:���1,����|���2,����|...
                                    str���� = Mid(str����, 2)
                                    cll����.Add str����
                                    str���� = ""
                                End If
                                str���� = str���� & "|" & strTemp
                            Next
                            If str���� <> "" Then
                                str���� = Mid(str����, 2)
                                cll����.Add str����
                            End If
                            For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                                'Zl_�ٴ�����Һſ��Ƽ�¼_Insert(
                                strSQL = "Zl_�ٴ�����Һſ��Ƽ�¼_Insert("
                                '��¼id_In   �ٴ�����Һſ��Ƽ�¼.��¼id%Type,
                                strSQL = strSQL & "" & obj�����¼.��¼ID & ","
                                '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                                strSQL = strSQL & "" & obj������λ.���� & ","
                                '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                                strSQL = strSQL & "" & 1 & ","
                                '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                                strSQL = strSQL & "'" & obj������λ.������λ���� & "',"
                                '���Ʒ�ʽ_In �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type,
                                strSQL = strSQL & "" & obj������λ.ԤԼ���Ʒ�ʽ & ","
                                '�Ƿ��ռ_In �ٴ������¼.�Ƿ��ռ%Type,
                                strSQL = strSQL & "" & IIf(obj�����¼.������λ���Ƽ�.�Ƿ��ռ, 1, 0) & ","
                                '���ſ���_In Varchar2,
                                str���� = ""
                                If cll����.Count > 0 Then str���� = cll����(i)
                                strSQL = strSQL & "'" & str���� & "',"
                                'ɾ��_In Number:=0
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

Private Function TempPlanVerifyOrCancel(ByVal lng����ID As Long, ByVal lng��ԴId As Long, _
    ByVal blnVerify As Boolean) As Boolean
    '���ܣ���˻�ȡ�������ʱ����
    Dim strSQL As String

    Err = 0: On Error GoTo ErrHandler
    If lng����ID = 0 Then Exit Function
    
    If blnVerify Then
        'Zl_�ٴ�������ʱ����_Verify(
        strSQL = "Zl_�ٴ�������ʱ����_Verify("
        '����id_In In �ٴ����ﰲ��.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '�����_in in �ٴ����ﰲ��.�����%type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '���ʱ��_in in �ٴ����ﰲ��.���ʱ��%type
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    Else
        'Zl_�ٴ�������ʱ����_Cancel(
        strSQL = "Zl_�ٴ�������ʱ����_Cancel("
        '����id_In In �ٴ����ﰲ��.Id%Type
        strSQL = strSQL & "" & lng����ID & ")"
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    TempPlanVerifyOrCancel = True
    
    '�������ɳ����¼
    strSQL = "Zl1_Auto_Buildingregisterplan(Null," & lng��ԴId & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData(obj���ﰲ�� As ���ﰲ��) As Boolean
    '��������
    Dim cllPro As Collection, blnTrans As Boolean
    Dim dtCurdate As Date, strTemp As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnRecord As Boolean, blnDeletePlan As Boolean
    Dim obj���ﰲ��Temp As ���ﰲ��, obj�����¼�� As �����¼��
    Dim ObjItem As �����¼��, blnUpdatePlan As Boolean
    Dim blnExistPlan As Boolean

    On Error GoTo ErrHandler
    If mbytFun = Fun_UpdateUnit Then
        '����ԤԼ�Һſ���
        SaveData = SaveUpdateUnit(obj���ﰲ��.δ������ﰲ��)
        Exit Function
    ElseIf mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel Then
        SaveData = TempPlanVerifyOrCancel(obj���ﰲ��.����ID, obj���ﰲ��.�����Դ.ID, mbytFun = Fun_TempPlanVerify)
        Exit Function
    End If

    Set cllPro = New Collection
    If obj���ﰲ��.����ID = 0 Then
        '��ȡ����ID
        obj���ﰲ��.����ID = zlDatabase.GetNextId("�ٴ����ﰲ��")
    End If
    dtCurdate = zlDatabase.Currentdate
    blnRecord = mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan

    '����ѱ����б�ɾ����δ���������У����ʱҪ�����޸İ��ţ�
    blnUpdatePlan = False
    If mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan Then
        '��������ʱ�����������Ų��ܵ���ԭ�ٴ����ﰲ�ż�¼����Ϊ�ѷ���
    Else
        For Each obj�����¼�� In obj���ﰲ��.�ѱ�����ﰲ��
            If obj�����¼��.�Ƿ�ɾ�� Then
                For Each ObjItem In obj���ﰲ��.δ������ﰲ��
                    If obj�����¼��.�������� = ObjItem.�������� Then
                        blnUpdatePlan = True: Exit For
                    End If
                Next
            End If
        Next
    End If
    
    '1.����ģ���Ҫ���°���
    '2.�ı�̶�����ʱ�䷶Χ��Ҫ�������ŵĿ�ʼʱ�����ֹʱ��
    '3.�ı��շ���Ŀ��Ҫ�������ŵ���ĿID
    If mbytPlanType = F_Templet Or mblnTimeChanged Or mblnFeeItemChanged Then
        blnUpdatePlan = True
    End If

    'ģ�����仯��ɾ��ԭ�г�����Ŀ��Ϣ
    'ģ���Ű����Ϊ2(����),3(˫��),4(������ѭ),5(������ѭ)ʱɾ���������
    If obj���ﰲ��.�ѱ�����ﰲ��.Count > 0 And mbytPlanType = F_Templet _
        And (obj���ﰲ��.�ѱ�����ﰲ��.�Ű���� <> obj���ﰲ��.�Ű���� _
            Or InStr(",2,3,4,5,", obj���ﰲ��.�Ű����) > 0) Then
        'Zl_�ٴ����ﰲ��_Batchdelete
        strSQL = "Zl_�ٴ����ﰲ��_Batchdelete("
        '����id_In �ٴ������.Id%Type,
        strSQL = strSQL & "" & obj���ﰲ��.����ID & ","
        '��Աid_In ��Ա��.Id%Type := 0,--������0��ɾ����Ա���ڿ��ҵ����к�Դ����
        strSQL = strSQL & "" & "NULL" & ","
        'վ��_In   ���ű�.վ��%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '��Դid_In �ٴ����ﰲ��.��Դid%Type := 0--������0��ɾ���ú�Դ�����а���
        strSQL = strSQL & "" & obj���ﰲ��.�����Դ.ID & ")"
        cllPro.Add strSQL
    End If

    'ɾ��δ����ʱ�εİ���
    blnDeletePlan = mbytFun = Fun_Update Or mbytFun = Fun_TempPlan
    If DeletePlan(obj���ﰲ��.�ѱ�����ﰲ��, cllPro, blnRecord, blnExistPlan, blnDeletePlan) = False Then Exit Function
    
    If blnUpdatePlan Or obj���ﰲ��.δ������ﰲ��.Count > 0 Then
        If (blnUpdatePlan Or blnExistPlan = False And obj���ﰲ��.δ������ﰲ��.Count > 0) _
            And Not (mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdatePlan) Then
            'Zl_�ٴ����ﰲ��_Insert(
            strSQL = "Zl_�ٴ����ﰲ��_Insert("
            'Id_In           �ٴ����ﰲ��.Id%Type,
            strSQL = strSQL & "" & obj���ﰲ��.����ID & ","
            '����id_In       �ٴ����ﰲ��.����id%Type,
            strSQL = strSQL & "" & obj���ﰲ��.����ID & ","
            '��Դid_In       �ٴ����ﰲ��.��Դid%Type,
            strSQL = strSQL & "" & obj���ﰲ��.�����Դ.ID & ","
            '��Ŀid_In       �ٴ����ﰲ��.��Ŀid%Type,
            If obj���ﰲ��.��ĿID = 0 Then
                strTemp = obj���ﰲ��.�����Դ.��ĿID
            Else
                strTemp = obj���ﰲ��.��ĿID
            End If
            strSQL = strSQL & "" & ZVal(strTemp) & ","
            'ҽ��id_In       �ٴ����ﰲ��.ҽ��id%Type,
            If obj���ﰲ��.ҽ��ID = 0 Then
                strTemp = obj���ﰲ��.�����Դ.ҽ��ID
            Else
                strTemp = obj���ﰲ��.ҽ��ID
            End If
            strSQL = strSQL & "" & ZVal(strTemp) & ","
            'ҽ������_In     �ٴ����ﰲ��.ҽ������%Type,
            If obj���ﰲ��.ҽ������ = "" Then
                strTemp = obj���ﰲ��.�����Դ.ҽ������
            Else
                strTemp = obj���ﰲ��.ҽ������
            End If
            strSQL = strSQL & "" & IIf(strTemp = "", "NULL", "'" & strTemp & "'") & ","
            '�Ű����_In     �ٴ����ﰲ��.�Ű����%Type,
            If mbytPlanType = F_Templet Then
                strTemp = obj���ﰲ��.�Ű����
            Else
                strTemp = IIf(mbytPlanType = F_MonthTemplet, 6, "NULL")
            End If
            strSQL = strSQL & "" & strTemp & ","
            '�Ƿ���������_In �ٴ����ﰲ��.�Ƿ���������%Type,
            If mbytPlanType = F_Templet And InStr("2,3,4,5", obj���ﰲ��.�Ű����) > 0 Then
                strTemp = IIf(obj���ﰲ��.����������, "0", "1")
            Else
                strTemp = "NULL"
            End If
            strSQL = strSQL & "" & strTemp & ","
            '�Ƿ����ճ���_In �ٴ����ﰲ��.�Ƿ����ճ���%Type,
            If mbytPlanType = F_Templet And InStr("2,3,4,5", obj���ﰲ��.�Ű����) > 0 Then
                strTemp = IIf(obj���ﰲ��.���ղ�����, "0", "1")
            Else
                strTemp = "NULL"
            End If
            strSQL = strSQL & "" & strTemp & ","
            '��ʼʱ��_In     �ٴ����ﰲ��.��ʼʱ��%Type,
            strSQL = strSQL & "" & IIf(mbytPlanType = F_MonthTemplet, "NULL", ZDate(obj���ﰲ��.��ʼʱ��)) & ","
            '��ֹʱ��_In     �ٴ����ﰲ��.��ֹʱ��%Type,
            strSQL = strSQL & "" & IIf(mbytPlanType = F_MonthTemplet, "NULL", ZDate(obj���ﰲ��.��ֹʱ��)) & ","
            '����Ա����_In   �ٴ����ﰲ��.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '�Ǽ�ʱ��_In     �ٴ����ﰲ��.�Ǽ�ʱ��%Type
            strSQL = strSQL & "" & ZDate(dtCurdate) & ","
            '�Ƿ����_In     number
            If mbytFun = Fun_AddSignalSourcePlan _
                And (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            strSQL = strSQL & "" & strTemp & ","
            '�Ƿ���ʱ����_In �ٴ����ﰲ��.�Ƿ���ʱ����%Type := 0
            strSQL = strSQL & "" & IIf(mbytFun = Fun_TempPlan, 1, 0) & ")"
            cllPro.Add strSQL
        End If
        
        '����δ����ĳ��ﰲ��
        If SavePlanData(obj���ﰲ��.����ID, obj���ﰲ��.δ������ﰲ��, obj���ﰲ��.�����Դ, _
            cllPro, blnRecord, dtCurdate, obj���ﰲ��.�ѱ�����ﰲ��) = False Then Exit Function
    Else
        obj���ﰲ��.����ID = 0 '������ް�����
    End If

    If cllPro.Count = 0 Then
        MsgBox "��ǰû����Ҫ�������Ч���ţ�", vbInformation, gstrSysName
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

Private Function DeletePlan(ByVal obj���ﰲ�� As ���ﰲ��, _
    ByRef cllPro As Collection, _
    ByVal blnRecord As Boolean, _
    ByRef blnExistPlan As Boolean, _
    ByVal blnDeletePlan As Boolean) As Boolean
    '���ܣ�ɾ�����ﰲ��
    '���Σ�
    '   blnExistPlan ɾ�����Ƿ񻹴����ѱ��氲��
    Dim strSQL As String
    Dim ObjItem As �����¼��

    Err = 0: On Error GoTo ErrHandler
    blnExistPlan = False
    If obj���ﰲ�� Is Nothing Then DeletePlan = True: Exit Function
    If obj���ﰲ��.Count = 0 Then DeletePlan = True: Exit Function
    
    If cllPro Is Nothing Then Set cllPro = New Collection
    For Each ObjItem In obj���ﰲ��
        If ObjItem.�Ƿ�ɾ�� Then
            'Zl_�ٴ������ϰ�ʱ��_Delete(
            strSQL = "Zl_�ٴ������ϰ�ʱ��_Delete("
            '����id_In   �ٴ���������.����id%Type,
            strSQL = strSQL & "" & obj���ﰲ��.����ID & ","
            'Id_In       �ٴ���������.������Ŀ%Type := 0,
            strSQL = strSQL & "'" & IIf(mbytPlanType = F_MonthTemplet, FormatApplyToStr(ObjItem.��������), ObjItem.��������) & "',"
            '�����¼_In Number:=0,
            strSQL = strSQL & "" & IIf(blnRecord, 1, 0) & ","
            '�ϰ�ʱ��_In     �ٴ���������.�ϰ�ʱ��%Type := Null,
            strSQL = strSQL & "" & "NULL" & ","
            'ɾ�����ﰲ��_In Number:=0
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
    stbThis.Panels(2).Text = "˵����������ɫΪ��ɫ���ϰ�ʱ�α�ʾ���ں�Դ��������ȱʡ���š�"
End Sub

Private Sub lvwWorkTime_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Item.Selected = False 'ȥ��ѡ�����
    stbThis.Panels(2).Text = "˵����������ɫΪ��ɫ���ϰ�ʱ�α�ʾ���ں�Դ��������ȱʡ���š�"
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
    stbThis.Panels(2).Text = "��ǰ����������룬ҽ����������룬�������ƻ������в��ҡ�"
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
    '���ݺ��롢�������ơ�ҽ����������ģ������
    If strText <> "" Then
        strKey = gstrLike & UCase(strText) & "%"
        If IsNumeric(strText) Then   '�������ȫ����
            strWhere = " And a.���� Like [2]"
        ElseIf zlCommFun.IsCharAlpha(strText) Then  '�������ȫ��ĸ
            strWhere = " And (Upper(c.����) Like [2] Or Upper(d.����) Like [2])"
        ElseIf zlCommFun.IsCharChinese(strText) Then '�Ƿ��к���,'���к���,�϶���������
            strWhere = " And (c.���� Like [2] Or a.ҽ������ Like [2])"
        Else
            strWhere = " And (a.���� Like [2] Or c.���� Like [2] Or Upper(c.����) Like [2] Or a.ҽ������ Like [2] Or Upper(d.����) Like [2])"
        End If

        strSQL = "Select a.Id, a.����, a.����, b.���� As ��Ŀ, c.���� As ����, a.ҽ������" & vbNewLine & _
                " From �ٴ������Դ A, �շ���ĿĿ¼ B, ���ű� C, ��Ա�� D" & vbNewLine & _
                " Where a.��Ŀid = b.Id And a.����id = c.Id And a.ҽ��id = d.Id(+)" & vbNewLine & _
                "       And a.�Ű෽ʽ In (Select �Ű෽ʽ From �ٴ������ Where ID = [1])" & vbNewLine & _
                "       And Not Exists (Select 1 From �ٴ����ﰲ�� Where ��Դid = a.Id And ����id = [1])" & vbNewLine & _
                "       And Not Exists (Select 1 From �ٴ����ﰲ�� P,�ٴ������ Q,�ٴ������ H" & vbNewLine & _
                "                       Where p.����ID = q.ID And q.�Ű෽ʽ = h.�Ű෽ʽ And q.������� = h.�������" & vbNewLine & _
                "                             And p.��Դid = a.Id And h.id = [1])" & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And Nvl(Nvl(c.վ��,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
                strWhere
        vRect = zlControl.GetControlRect(txtSignal.Hwnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�����Դ", False, "", "��ѡ������Դ", False, False, True, _
            vRect.Left, vRect.Top, txtSignal.Height, blnCancel, True, False, mlng����ID, strKey, UserInfo.ID, _
            gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
        If blnCancel Then zlControl.TxtSelAll txtSignal: Exit Sub
        If rsTemp Is Nothing Then
            MsgBox "û���ҵ���Ҫ��������Ч�����Դ�����飡", vbInformation, gstrSysName
            zlControl.TxtSelAll txtSignal
            Exit Sub
        End If
        If rsTemp.EOF Then
            MsgBox "û���ҵ���Ҫ��������Ч�����Դ�����飡", vbInformation, gstrSysName
            zlControl.TxtSelAll txtSignal
            Exit Sub
        End If
        mlng��ԴId = Nvl(rsTemp!ID)
        txtSignal.Text = Nvl(rsTemp!����)
        zlControl.TxtSelAll txtSignal

        If CheckSignalSource(mbytPlanType = F_FixedRule, mlng��ԴId, mlng����ID, _
            mobj���ﰲ��.��ʼʱ��, mobj���ﰲ��.��ֹʱ��) = False Then
            mlng��ԴId = 0
            Exit Sub
        End If

        If InitData(mobj���ﰲ��, mlng����ID, mlng��ԴId, mlng����ID) = False Then Exit Sub
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
    '���ÿؼ�����״̬
    lvwWorkTime.Enabled = blnEnabled
    If lvwWorkTime.Enabled Then lvwWorkTime.BackColor = lvwWorkTime.BackColor
    picApply.Enabled = blnEnabled
    cldsCalenbarSel.Enabled = blnEnabled
End Sub


Private Function CheckWorkTime(ByVal obj�����¼�� As �����¼��) As Boolean
    'ѡ��ʱ���
    Dim objListItem As ListItem, i As Integer
    Dim ObjItem As �����¼, blnFind As Boolean
    Dim strTmp As String, blnClearAllWorkTime As Boolean
    Dim strSetedTime As String

    On Error GoTo Errhand
    Call LockWindowUpdate(lvwWorkTime.Hwnd) '��ֹ��˸
    '��������ϰ�ʱ��
    If mobj���ﰲ��.�����ϰ�ʱ�� Is Nothing Then
        blnClearAllWorkTime = True
    ElseIf mobj���ﰲ��.�����ϰ�ʱ��.Count = 0 Then
        blnClearAllWorkTime = True
    End If
    
    If blnClearAllWorkTime Then
        lvwWorkTime.ListItems.Clear
    Else
        'ȡ�����е�ѡ��
        For i = 1 To lvwWorkTime.ListItems.Count
            lvwWorkTime.ListItems(i).Checked = False
        Next
    End If
    If obj�����¼�� Is Nothing Then Exit Function
    lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width - 600
    For Each ObjItem In obj�����¼��
        blnFind = False
        If ObjItem.ʱ��� <> "" Then
            For i = 1 To lvwWorkTime.ListItems.Count
                If InStr(";" & strSetedTime & ";", ";" & lvwWorkTime.ListItems(i).Tag & ";") = 0 Then
                    '���ͣ��״̬
                    strTmp = lvwWorkTime.ListItems(i).Text
                    lvwWorkTime.ListItems(i).Text = Left(strTmp, IIf(InStr(strTmp, ")(") = 0, Len(strTmp), InStr(strTmp, ")(")))
                End If
                
                If ObjItem.ʱ��� = lvwWorkTime.ListItems(i).Tag Then
                    lvwWorkTime.ListItems(i).Checked = True
                    '��ʾͣ��״̬
                    lvwWorkTime.ListItems(i).Text = lvwWorkTime.ListItems(i).Text & _
                        Decode(CheckRecordStopVisit(ObjItem), 1, "(����ͣ��)", 2, "(��ͣ��)", "")
                    '103078,�������п��
                    If InStr(lvwWorkTime.ListItems(i).Text, "(����ͣ��)") > 0 Then
                        lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width + 200
                    ElseIf InStr(lvwWorkTime.ListItems(i).Text, "(��ͣ��)") > 0 Then
                        lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width
                    End If
                    blnFind = True
                    strSetedTime = strSetedTime & ";" & lvwWorkTime.ListItems(i).Tag
                    Exit For
                End If
            Next
            If blnFind = False Then
                Set objListItem = lvwWorkTime.ListItems.Add(, "K" & ObjItem.ʱ���, ObjItem.ʱ��� & _
                    "(" & Format(ObjItem.��ʼʱ��, "hh:mm") & "-" & Format(ObjItem.��ֹʱ��, "hh:mm") & ")")
                '��ʾͣ��״̬
                objListItem.Text = objListItem.Text & Decode(CheckRecordStopVisit(ObjItem), 1, "(����ͣ��)", 2, "(��ͣ��)", "")
                objListItem.SubItems(1) = ObjItem.��ʼʱ��
                objListItem.SubItems(2) = ObjItem.��ֹʱ��
                objListItem.Tag = ObjItem.ʱ���
                objListItem.Checked = True
                '103078,�������п��
                If InStr(objListItem.Text, "(����ͣ��)") > 0 Then
                    lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width + 200
                ElseIf InStr(objListItem.Text, "(��ͣ��)") > 0 Then
                    lvwWorkTime.ColumnHeaders(1).Width = lvwWorkTime.Width
                End If
                strSetedTime = strSetedTime & ";" & objListItem.Tag
                '����ɫ�����Ƿ��Դ�����õ�ʱ���
                If mobj���ﰲ��.��Դ����.Exits("K" & ObjItem.ʱ���) Then
                    objListItem.ForeColor = vbBlue
                End If
            End If
        End If
    Next
    If Not lvwWorkTime.SelectedItem Is Nothing Then
        lvwWorkTime.SelectedItem.Selected = False 'ȥ��ѡ�����
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

Private Function CheckRecordStopVisit(ByVal obj�����¼ As �����¼) As Integer
    '�жϳ����¼��ͣ�����
    '��Σ�
    '   obj�����¼ - �����¼
    '���Σ�
    '   ͣ�����ͣ�0-û��ͣ�1-����ʱ��ͣ�2-ȫ��ʱ��ͣ��
    Err = 0: On Error GoTo Errhand
    If obj�����¼ Is Nothing Then Exit Function
    If obj�����¼.ʱ��� = "" _
        Or obj�����¼.ͣ�￪ʼʱ�� = "" Or obj�����¼.ͣ����ֹʱ�� = "" Then Exit Function

    If DateDiff("s", obj�����¼.��ʼʱ��, obj�����¼.ͣ�￪ʼʱ��) = 0 _
        And DateDiff("s", obj�����¼.��ֹʱ��, obj�����¼.ͣ����ֹʱ��) = 0 Then
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

Private Function CheckPlanIsStopOrUsed(ByVal lng��¼ID As Long) As Boolean
    '�������¼�Ƿ���ͣ����ѱ�ʹ��
    Err = 0: On Error GoTo Errhand
    If lng��¼ID = 0 Then Exit Function
    If mrsVisitedRecordByDate Is Nothing Then Exit Function
    
    mrsVisitedRecordByDate.Filter = "ID=" & lng��¼ID
    If mrsVisitedRecordByDate.EOF Then Exit Function
    
    If Val(Nvl(mrsVisitedRecordByDate!��ʹ��)) = 1 _
        Or Val(Nvl(mrsVisitedRecordByDate!��ͣ��)) = 1 Then
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
    If mobj���ﰲ�� Is Nothing Then Exit Sub

    strOldDate = mstrCurDay
    If IsDate(strOldDate) Then
        strNewDate = Format(DateAdd("d", -1, strOldDate), "yyyy-mm-dd")
    Else
        If mobj���ﰲ��.�Ű෽ʽ = 3 And mobj���ﰲ��.�Ű���� <> 1 Then
            If mobj���ﰲ��.�Ű���� = 6 Then
                If Val(strOldDate) - 1 > 0 Then strNewDate = Val(strOldDate) - 1 & "��"
            End If
        Else '����
            intIndex = GetWeekIndex(strOldDate)
            If intIndex - 1 >= 0 And intIndex - 1 <= 6 Then
                strNewDate = GetWeekName(intIndex - 1)
            End If
        End If
    End If
    Call cldsCalenbarSel_SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then Exit Sub

    Call ChangeCurPlan(mobj���ﰲ��, strNewDate)
    Call cldsCalenbarSel_SelectedChanged(strOldDate, strNewDate)

    cldsCalenbarSel.LoadData mobj���ﰲ��
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
    If mobj���ﰲ�� Is Nothing Then Exit Sub

    strOldDate = mstrCurDay
    If IsDate(strOldDate) Then
        strNewDate = Format(DateAdd("d", 1, strOldDate), "yyyy-mm-dd")
    Else
        If mobj���ﰲ��.�Ű෽ʽ = 3 And mobj���ﰲ��.�Ű���� <> 1 Then
            If mobj���ﰲ��.�Ű���� = 6 Then
                If Val(strOldDate) + 1 <= 31 Then strNewDate = Val(strOldDate) + 1 & "��"
            End If
        Else '����
            intIndex = GetWeekIndex(strOldDate)
            If intIndex + 1 >= 0 And intIndex + 1 <= 6 Then
                strNewDate = GetWeekName(intIndex + 1)
            End If
        End If
    End If
    Call cldsCalenbarSel_SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then Exit Sub

    Call ChangeCurPlan(mobj���ﰲ��, strNewDate)
    Call cldsCalenbarSel_SelectedChanged(strOldDate, strNewDate)

    cldsCalenbarSel.LoadData mobj���ﰲ��
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnabled()
    '������һ������һ����ť����״̬
    Dim intIndex As Integer

    On Error GoTo Errhand
    If mobj���ﰲ�� Is Nothing Then
        btnLeft.Enabled = False
        btnRight.Enabled = False
    ElseIf mobj���ﰲ��.��ʱ���� Then
        btnLeft.Enabled = False
        btnRight.Enabled = False
    Else
        If cldsCalenbarSel.ShowStyle = Show_Plan_Rule And mobj���ﰲ��.�Ű���� <> 1 Then
            '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
            btnLeft.Enabled = False
            btnRight.Enabled = False
            '��ѡʱ�����л�����ѡʱ���л�
            If mobj���ﰲ��.�Ű���� = 6 Then
                If mobj���ﰲ��.Count = 1 Then
                    btnLeft.Enabled = Val(mstrCurDay) > 1
                    btnRight.Enabled = Val(mstrCurDay) < 31
                End If
            End If
        Else
            btnLeft.Enabled = True
            btnRight.Enabled = True
            If IsDate(mstrCurDay) Then '����
                If DateDiff("d", mobj���ﰲ��.��ʼʱ��, mstrCurDay) <= 0 Then
                    btnLeft.Enabled = False
                End If
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, mstrCurDay) >= 0 Then
                    btnRight.Enabled = False
                End If
            Else '����
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

Private Sub Get��ǰ���ﰲ��(obj���ﰲ�� As ���ﰲ��, Optional ByVal blnApply As Boolean = True)
    '��ȡ��ǰ���ﰲ��
    Dim obj�����¼�� As �����¼��
    Dim ObjItem As �����¼��, objRecord As �����¼
    Dim strApplyTo As String, varApplyTo As Variant
    Dim i As Integer, blnAdd As Boolean
    Dim objӦ�ó����¼�� As �����¼��
    Dim blnFindCur As Boolean '�Ƿ�Ӧ���ڵ�ǰѡ����Ŀ

    On Error GoTo Errhand
    Set obj�����¼�� = CPDPages.Get�����¼��
    If obj�����¼�� Is Nothing Then Exit Sub

    '�����ж�����ﰲ�ţ�����-�ƶ����ڣ�������ǰ���ﰲ��Ӧ�������г��ﰲ��
    For Each ObjItem In obj���ﰲ��
        '�Ƴ���������
        ObjItem.RemoveAll
        ObjItem.�Ƿ��޸� = ObjItem.�Ƿ��޸� Or obj�����¼��.�Ƿ��޸�
        For Each objRecord In obj�����¼��
            ObjItem.AddItem objRecord.Clone, "K" & objRecord.ʱ���
        Next
    Next

    If blnApply = False Then Exit Sub
    'Ӧ������������
    blnFindCur = False
    varApplyTo = Split(GetApplyToStr, ",")
    For i = 0 To UBound(varApplyTo)
        blnAdd = True
'        If IsDate(varApplyTo(i)) Then
'            'С�ڵ�ǰ���ڵĲ��������
'            If DateDiff("d", varApplyTo(i), mdtToday) > 0 Then blnAdd = False
'        End If
        If blnAdd Then
            If IsDate(varApplyTo(i)) Then
                If GetPlanKey(mstrCurDay) = GetPlanKey(varApplyTo(i)) Then blnFindCur = True
            Else
                blnFindCur = True
            End If
            Set objӦ�ó����¼�� = New �����¼��
            With objӦ�ó����¼��
                .�������� = IIf(IsDate(varApplyTo(i)), Format(varApplyTo(i), "yyyy-mm-dd"), varApplyTo(i))
                .�Ƿ��޸� = True '�����Ŀ϶����޸�
                For Each objRecord In obj�����¼��
                    If GetPlanKey(mstrCurDay) <> GetPlanKey(varApplyTo(i)) Then
                        '��Ӧ���ڵ�����������Ҫ�����µļ�¼ID
                        objRecord.��¼ID = 0
                        objRecord.�Ƿ�̶� = False
                        objRecord.�Ƿ��޸� = True
                    End If
                    .AddItem objRecord.Clone, "K" & objRecord.ʱ���
                Next
            End With
            If obj���ﰲ��.Exits(GetPlanKey(varApplyTo(i))) = False Then
                obj���ﰲ��.AddItem objӦ�ó����¼��, GetPlanKey(varApplyTo(i))
            End If
        End If
    Next

    '��ǰѡ�����ڲ���Ӧ�����У����Ƴ�
    If UBound(varApplyTo) > -1 And blnFindCur = False Then
        If obj���ﰲ��.Exits(GetPlanKey(mstrCurDay)) Then
            obj���ﰲ��.Remove GetPlanKey(mstrCurDay)
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
    Case Pancel_Index.Pan_����
        Item.Handle = picDateList.Hwnd
    Case Pancel_Index.Pan_ʱ���
        Item.Handle = picWorkTimeList.Hwnd
    Case Pancel_Index.Pan_��Դ
        If mbytPlanType = F_FixedRule Then
            Item.Handle = picSourceAndPlan.Hwnd
        Else
            Item.Handle = picSouceList.Hwnd
        End If
    Case Pancel_Index.Pan_����
        Item.Handle = picDetailedList.Hwnd
    End Select
End Sub

Private Sub lvwWorkTime_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, obj���ﰲ�� As ���ﰲ��
    Dim blnChecked As Boolean, ObjItem As ���ﰲ��, objTemp As ���ﰲ��
    Dim lngFindIndex As Long, objItemTemp As �����¼
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim dtStart As Date, dtEnd As Date
    Dim dtStopStart As Date, dtStopEnd As Date
    Dim blnNotCheck As Boolean
    Dim bytȱʡ���﷽ʽ As Byte, objȱʡ�������� As �������Ҽ�

    On Error GoTo Errhand
    blnChecked = Item.Checked
    Item.Checked = Not blnChecked
    If mobj���ﰲ��.Count = 0 Then
        MsgBox IIf(cldsCalenbarSel.ShowStyle = Show_Plan_Rule, "�������δѡ��", "��������δѡ��"), vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If blnChecked Then
        '�����ϰ�ʱ��Σ������н��棬ע���ų�ͣ���ʱ��
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

                '�ų�ͣ���ʱ�䷶Χ
                If mobj���ﰲ��.�ѱ�����ﰲ��.Exits(GetPlanKey(mstrCurDay)) Then
                    If mobj���ﰲ��.�ѱ�����ﰲ��(GetPlanKey(mstrCurDay)).Exits("K" & lvwWorkTime.ListItems(i).Tag) Then
                        With mobj���ﰲ��.�ѱ�����ﰲ��(GetPlanKey(mstrCurDay))("K" & lvwWorkTime.ListItems(i).Tag)
                            If .ͣ�￪ʼʱ�� <> "" Then
                                dtStopStart = Format(dtCurStart, "yyyy-mm-dd ") & Format(.ͣ�￪ʼʱ��, "hh:mm:ss")
                                dtStopEnd = Format(dtCurStart, "yyyy-mm-dd ") & Format(.ͣ����ֹʱ��, "hh:mm:ss")
                                dtStopEnd = GetWorkTrueDate(dtStopStart, dtStopEnd)
                                If DateDiff("n", dtCurStart, dtStopStart) <= 0 And DateDiff("n", dtCurEnd, dtStopEnd) >= 0 Then
                                    'ȫ��ͣ�ﲻ�ü��
                                    blnNotCheck = True
                                Else
                                    '����ͣ�ȡ��δͣ���ʱ�䷶Χ
                                    If DateDiff("n", dtCurStart, dtStopStart) = 0 Then '1.ͣ��ǰ����[dtStopEnd,dtCurEnd]
                                        dtCurStart = dtStopEnd
                                    ElseIf DateDiff("n", dtCurEnd, dtStopEnd) = 0 Then '2.ͣ��󲿷�[dtCurStart,dtStopStart]
                                        dtCurEnd = dtStopStart
                                    Else '3.ͣ���м䲿��
                                        'ǰ,�ȼ��[dtCurStart,dtStopStart]
                                        If CheckTimeBucketIsCross(dtCurStart, dtStopStart, dtStart, dtEnd) Then
                                            MsgBox "��ǰ�ϰ�ʱ�ε�ʱ�䷶Χ����ѡ�����Ч�ϰ�ʱ�Ρ�" & lvwWorkTime.ListItems(i).Tag & "����ʱ�䷶Χ�н��棬����ͬʱѡ��", vbInformation + vbOKOnly, gstrSysName
                                            Exit Sub
                                        End If
                                        '��[dtStopEnd,dtCurEnd]
                                        dtCurStart = dtStopEnd
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
                
                
                If blnNotCheck = False And CheckTimeBucketIsCross(dtCurStart, dtCurEnd, dtStart, dtEnd) Then
                    MsgBox "��ǰ�ϰ�ʱ�ε�ʱ�䷶Χ����ѡ�����Ч�ϰ�ʱ�Ρ�" & lvwWorkTime.ListItems(i).Tag & "����ʱ�䷶Χ�н��棬����ͬʱѡ��", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        Next

        '��ǰ����+ѡ��ʱ��ο�ʼʱ�������ڵ�ǰʱ��
        If IsDate(mstrCurDay) And (mbytFun = Fun_AddSignalSourcePlan Or mbytFun = Fun_TempPlanRecord _
            Or mbytFun = Fun_UpdatePlan) Then
            If DateDiff("s", dtEnd, zlDatabase.Currentdate) > 0 Then
                MsgBox "��ǰ�ϰ�ʱ�ε���ֹʱ��С���˵�ǰʱ�䣬���ܰ��ų��", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If

            '����ϰ�ʱ�����Ƿ�ɳ���
            If CheckDepend(1, mstrCurDay, dtStart, dtEnd) = False Then
                Exit Sub
            End If
        End If
    Else
        '��ʱ����ʱ����ʱ�β���ɾ��
        If mobj���ﰲ��.��ʱ���� Or mbytFun = Fun_TempPlan Then
            If mobj���ﰲ��(1).Exits("K" & Item.Tag) Then
                If mobj���ﰲ��(1)("K" & Item.Tag).�Ƿ�̶� Then
                    If mbytFun = Fun_TempPlan Then
                        MsgBox "��ǰ�ϰ�ʱ������ʱ�����б������������ȡ����", vbInformation, gstrSysName
                    ElseIf mbytFun = Fun_TempPlanRecord Then
                        If mobj���ﰲ��(1)("K" & Item.Tag).�Ƿ���ʱ���� Then
                            MsgBox "��ʱ����ʱֻ�������ϰ�ʱ�ΰ��ţ�����Ҫȡ�����ϰ�ʱ�ΰ��ţ���ʹ�õ������Ź��ܲ�����", vbInformation, gstrSysName
                        Else
                            MsgBox "��ʱ����ʱֻ�������ϰ�ʱ�ΰ��ţ�����Ҫȡ�����ϰ�ʱ�ΰ��ţ���ʹ��ͣ�﹦�ܲ�����", vbInformation, gstrSysName
                        End If
                    ElseIf mbytFun = Fun_UpdatePlan Then
                        MsgBox "��ǰ�ϰ�ʱ�ΰ�����ͣ����Ѵ���ԤԼ�Һż�¼�������������", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    If mbytFun = Fun_UpdatePlan And mobj���ﰲ��(1)("K" & Item.Tag).�Ƿ���ʱ���� = False Then
                        MsgBox "��������ʱ����ȡ������ʱ������ϰ�ʱ�ΰ��ţ�����Ҫȡ�����ϰ�ʱ�ΰ��ţ���ʹ��ͣ�﹦�ܲ�����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If

        'ɾ������ʱ�α�ʾɾ������
        If mobj���ﰲ��.�ѱ�����ﰲ��.Exits(GetPlanKey(mobj���ﰲ��(1).��������)) _
            Or mobj���ﰲ��.δ������ﰲ��.Exits(GetPlanKey(mobj���ﰲ��(1).��������)) Then
            If IsClearAll(Item.index) Then
                If MsgBox("��ȷ��Ҫɾ�� " & IIf(mbytPlanType = F_MonthTemplet, FormatApplyToStr(mobj���ﰲ��(1).��������), mobj���ﰲ��(1).��������) & " �ĳ��ﰲ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    Item.Checked = blnChecked

    Get��ǰ���ﰲ�� mobj���ﰲ��, False

    If Item.Checked Then
        'ȡ�ѱ����е�
        If mobj���ﰲ��.�ѱ�����ﰲ��.Exits("K" & Format(mobj���ﰲ��(1).��������, "yyyy-mm-dd")) Then
            If mobj���ﰲ��.�ѱ�����ﰲ��("K" & Format(mobj���ﰲ��(1).��������, "yyyy-mm-dd")).Exits("K" & Item.Tag) Then
                Set obj�����¼ = mobj���ﰲ��.�ѱ�����ﰲ��("K" & Format(mobj���ﰲ��(1).��������, "yyyy-mm-dd"))("K" & Item.Tag).Clone
            End If
        End If

        'ȡ��Դ�����˵�
        If obj�����¼ Is Nothing And mobj���ﰲ��.��Դ����.Exits("K" & Item.Tag) Then
            Set obj�����¼ = mobj���ﰲ��.��Դ����("K" & Item.Tag).Clone
            obj�����¼.��¼ID = 0 '��Ҫ�����µļ�¼ID
        End If

        If obj�����¼ Is Nothing Then
            Set obj�����¼ = New �����¼
            With mobj���ﰲ��.�����Դ
                obj�����¼.ʱ��� = Item.Tag
                If mobj���ﰲ��.�����ϰ�ʱ��.Exits("K" & obj�����¼.ʱ���) Then
                    Set obj�����¼.�ϰ�ʱ�� = mobj���ﰲ��.�����ϰ�ʱ��("K" & obj�����¼.ʱ���)
                Else
                    Set obj�����¼.�ϰ�ʱ�� = New �ϰ�ʱ��
                End If

                Set obj�����¼.�����������Ҽ� = New �������Ҽ�
                obj�����¼.�����������Ҽ�.ҽ������ = mobj���ﰲ��.�����Դ.ҽ������
                'ȱʡ��������
                If GetDefaultRoom(mobj���ﰲ��, bytȱʡ���﷽ʽ, objȱʡ��������) Then
                    obj�����¼.�����������Ҽ�.���﷽ʽ = bytȱʡ���﷽ʽ
                    Set obj�����¼.�����������Ҽ� = objȱʡ��������
                End If

                Set obj�����¼.������Ϣ�� = New ������Ϣ��
                obj�����¼.������Ϣ��.ʱ��� = obj�����¼.ʱ���
                obj�����¼.������Ϣ��.����Ƶ�� = .����Ƶ��
            End With
            obj�����¼.��¼ID = 0 '��Ҫ�����µļ�¼ID
        End If
        obj�����¼.�Ƿ��޸� = True '�����˿϶����޸�
        mobj���ﰲ��(1).�Ƿ��޸� = True
        mobj���ﰲ��(1).AddItem obj�����¼, "K" & obj�����¼.ʱ���
    Else
        Set obj�����¼�� = mobj���ﰲ��(1)
        For i = 1 To obj�����¼��.Count
            If obj�����¼��(i).ʱ��� = Item.Tag Then
                obj�����¼��.�Ƿ��޸� = True 'ɾ���˿϶����޸�
                obj�����¼��.Remove i: Exit For
            End If
        Next
    End If
    CPDPages.LoadData mobj���ﰲ��(1), mobj���ﰲ��.���з�������, mobj���ﰲ��.���к�����λ
    LoadPlanToGrid mobj���ﰲ��, 0
    stbThis.Panels(2).Text = "˵����������ɫΪ��ɫ���ϰ�ʱ�α�ʾ���ں�Դ��������ȱʡ���š�"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetDefaultRoom(ByVal obj���ﰲ�� As ���ﰲ��, _
    ByRef byt���﷽ʽ As Byte, ByRef obj�������� As �������Ҽ�) As Boolean
    '��ȡȱʡ���﷽ʽ����������
    
    If obj���ﰲ�� Is Nothing Then Exit Function
    If obj���ﰲ��.Count > 0 Then
        If obj���ﰲ��(1).Count > 0 Then
            byt���﷽ʽ = obj���ﰲ��(1)(1).�����������Ҽ�.���﷽ʽ
            Set obj�������� = obj���ﰲ��(1)(1).�����������Ҽ�.Clone
            GetDefaultRoom = True: Exit Function
        End If
    End If
    If Not obj���ﰲ��.δ������ﰲ�� Is Nothing Then
        If obj���ﰲ��.δ������ﰲ��.Count > 0 Then
            If obj���ﰲ��.δ������ﰲ��(1).Count > 0 Then
                byt���﷽ʽ = obj���ﰲ��.δ������ﰲ��(1)(1).�����������Ҽ�.���﷽ʽ
                Set obj�������� = obj���ﰲ��.δ������ﰲ��(1)(1).�����������Ҽ�.Clone
                GetDefaultRoom = True: Exit Function
            End If
        End If
    End If
    If Not obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then
        If obj���ﰲ��.�ѱ�����ﰲ��.Count > 0 Then
            If obj���ﰲ��.�ѱ�����ﰲ��(1).Count > 0 Then
                byt���﷽ʽ = obj���ﰲ��.�ѱ�����ﰲ��(1)(1).�����������Ҽ�.���﷽ʽ
                Set obj�������� = obj���ﰲ��.�ѱ�����ﰲ��(1)(1).�����������Ҽ�.Clone
                GetDefaultRoom = True: Exit Function
            End If
        End If
    End If
    If Not obj���ﰲ��.��Դ���� Is Nothing Then
        If obj���ﰲ��.��Դ����.Count > 0 Then
            byt���﷽ʽ = obj���ﰲ��.��Դ����(1).�����������Ҽ�.���﷽ʽ
            Set obj�������� = obj���ﰲ��.��Դ����(1).�����������Ҽ�.Clone
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
        'picApply.Tag = "1"��ʾ����
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
    '����:����Ӧ���ڵ���ʾ
    Dim i As Integer
    Dim dtCur As Date

    Err = 0: On Error GoTo ErrHandler:
    fraLoopSkip.Visible = False
    picApplyWeek.Visible = False
    If index = 3 Then '������
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

    '������ѯ��ʼ����
    If index = 4 Then
        If cboDays.ListCount = 0 Then
            If Not mobj���ﰲ�� Is Nothing Then
                '�϶����Ű෽ʽΪ1�����Ű�
                dtCur = mobj���ﰲ��.��ʼʱ��
                Do While True
                    cboDays.AddItem Format(dtCur, "yyyy-mm-dd")
                    dtCur = DateAdd("d", 1, dtCur)
                    If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
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
    '������ʾ���⣬�Լ�Ӧ��ѡ��
    Dim strTemp As String, i As Integer
    Dim strͣ��ԭ�� As String

    On Error GoTo Errhand
    lblTittle.Caption = "�ް���"

    picApply.Visible = False: picApply.Tag = "1" '����
    If mbytFun = Fun_TempPlanRecord Or mbytFun = Fun_UpdateUnit _
        Or mbytFun = Fun_TempPlanVerify Or mbytFun = Fun_TempPlanCancel _
        Or mbytFun = Fun_UpdatePlan Then
        'û��Ӧ����
    Else
        Select Case mbytPlanType
        Case F_Templet 'ģ�����
            If mobj���ﰲ��.�Ű���� = 1 Then
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
        Case F_FixedRule '�̶�����
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
        Case F_MonthPlan, F_WeekPlan, F_MonthTemplet '�°���/�ܰ���
            picApply.Visible = mbytFun <> Fun_View
            picApply.Tag = IIf(mbytFun = Fun_View, "1", "")
            If optRule(0).Value Then optRule(1).Value = True
            optRule(0).Value = True
            If mbytPlanType = F_WeekPlan Then '����,����ʹ����ѭ
                optRule(4).Visible = False
                fraLoopSkip.Visible = False
            End If
        End Select
    End If
    Call SetButtonEnabled
    Call picDetailedList_Resize

    If mstrCurDay <> "" And mbytPlanType = F_Templet Then
        '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
        Select Case mobj���ﰲ��.�Ű����
        Case 2
            strTemp = "�����ճ���"
        Case 3
            strTemp = "��˫�ճ���"
        Case 4, 5
            strTemp = "��" & Val(mstrCurDay) & "����ѭ"
        Case 6
            strTemp = ""
            If mobj���ﰲ��.�ѱ�����ﰲ��.�Ű���� = 6 Then
                For i = 1 To mobj���ﰲ��.�ѱ�����ﰲ��.Count
                    strTemp = strTemp & "," & mobj���ﰲ��.�ѱ�����ﰲ��(i).��������
                Next
            End If
            For i = 1 To mobj���ﰲ��.δ������ﰲ��.Count
                strTemp = strTemp & "," & mobj���ﰲ��.δ������ﰲ��(i).��������
            Next
            For i = 1 To mobj���ﰲ��.Count
                strTemp = strTemp & "," & mobj���ﰲ��(i).��������
            Next
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            strTemp = "��ÿ�µ�" & ZlNumStrSort(strTemp, True) & "�չ̶�����"
        Case Else
            strTemp = mstrCurDay
        End Select
    ElseIf mstrCurDay <> "" Then
        strTemp = mstrCurDay
    End If

    If strTemp = "" Then strTemp = "�ް���"
    If IsDate(strTemp) Then
        If mbytPlanType = F_MonthTemplet Then
            lblTittle.Caption = FormatApplyToStr(strTemp)
        Else
            strͣ��ԭ�� = CurDayIsNotVisit(strTemp)
            If strͣ��ԭ�� = "" Then
                lblTittle.Caption = Format(strTemp, "yyyy-mm-dd")
            Else
                lblTittle.Caption = Format(strTemp, "yyyy-mm-dd") & "(" & strͣ��ԭ�� & ")"
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
    '�Ƿ�����������ϰ�ʱ���
    'intIndex �ų�����
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
    '�������
    Dim blnSelected As Boolean
    Dim i As Integer, blnFind As Boolean
    Dim obj�����¼�� As �����¼��
    Dim dtStart As Date, dtEnd As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset

    Err = 0: On Error GoTo ErrHandler
    If mbytPlanType = F_FixedRule Then
        dtStart = mobj���ﰲ��.��ʼʱ��
        dtEnd = mobj���ﰲ��.��ֹʱ��
        If mbytFun = Fun_AddSignalSourcePlan Or mbytFun = Fun_TempPlan Or mbytFun = Fun_Update Then
            If dtpBegin.Visible And dtpBegin.Enabled And dtpEnd.Enabled Then
                '������ʱ��ؼ����ҵ�����ֱ�ӵ㱣�棬Value��ֵ��û�и��Ĺ���
                mblnValiedCanSave = True
                If Me.ActiveControl Is dtpBegin Then
                    dtpEnd.SetFocus: Call dtpBegin_LostFocus
                ElseIf Me.ActiveControl Is dtpEnd Then
                    dtpBegin.SetFocus: Call dtpEnd_LostFocus
                End If
                If mblnValiedCanSave = False Then Exit Function
                mblnValiedCanSave = False
                
                If dtpEnd.Value < zlDatabase.Currentdate Then
                    MsgBox "��Ч�ڵ���ֹʱ�䲻��С�ڵ�ǰʱ�䡣", vbInformation, gstrSysName
                    If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
                    Exit Function
                End If
                If dtpBegin.Value >= dtpEnd.Value Then
                    MsgBox "��Ч�ڵĿ�ʼʱ��Ӧ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                    If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
                    Exit Function
                End If
                
                '��Ч�ڵĿ�ʼ���ڴ��ڵ��ڵ�ǰ����ʱ�ż��
                If mdtToday <= CDate(Format(dtpBegin.Value, "yyyy-mm-dd")) _
                    And Format(dtpBegin.Value, "hh:mm:ss") <> "00:00:00" Then
                    If MsgBox("���յ�ǰ��Ч�ڵ����ã��ð����� " & Format(dtpBegin.Value, "yyyy-mm-dd") & _
                        " ������Ч�������ϣ���� " & Format(dtpBegin.Value, "yyyy-mm-dd") & _
                        " ��Ч�뽫��ʼʱ�����Ϊ " & Format(dtpBegin.Value, "yyyy-mm-dd 00:00:00") & _
                        "���Ƿ��������ǰ���ñ��棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
                mlng��ԴId, mlng����ID, dtStart, dtEnd, True) = False Then Exit Function
        End If
        mobj���ﰲ��.��ʼʱ�� = Format(dtStart, "yyyy-mm-dd hh:mm:ss")
        mobj���ﰲ��.��ֹʱ�� = Format(dtEnd, "yyyy-mm-dd hh:mm:ss")
        
        If CheckTempFixedPlan(mobj���ﰲ��, dtStart, dtEnd, , , True) = False Then Exit Function
        
        If mbytFun = Fun_TempPlan And zlStr.IsHavePrivs(mstrPrivs, "���п���") Then
            If cboFeeItem.ListIndex = -1 Then
                MsgBox "�շ���Ŀ����Ϊ�գ�", vbInformation, gstrSysName
                If cboFeeItem.Visible And cboFeeItem.Enabled Then cboFeeItem.SetFocus
                Exit Function
            End If
            
            '���ң�ҽ�����շ���Ŀ�ں�Դ�в����ظ�
            strSQL = "Select ���� From �ٴ������Դ" & _
                    " Where ����ID=[1] And Nvl(ҽ��ID,0)=[2] And Nvl(ҽ������,'-')=[3] And ��ĿID=[4]" & _
                    "       And Nvl(�Ƿ�ɾ��,0)=0 And ���� <> [5]"
            With mobj���ﰲ��.�����Դ
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Դ�Ƿ�Ψһ", .����ID, .ҽ��ID, _
                    IIf(.ҽ������ = "", "-", .ҽ������), cboFeeItem.ItemData(cboFeeItem.ListIndex), .����)
                If Not rsTemp.EOF Then
                    MsgBox .�������� & " " & IIf(.ҽ������ = "", "", "��ҽ�� " & .ҽ������ & " ") & _
                        "�Ѿ������շ���ĿΪ " & cboFeeItem.Text & " �ĺ�Դ��" & Nvl(rsTemp!����) & "����" & _
                        "�����ܶԵ�ǰ��Դ�ƶ��շ���ĿΪ " & cboFeeItem.Text & " ����ʱ���ţ�", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End With
        
            mobj���ﰲ��.��ĿID = cboFeeItem.ItemData(cboFeeItem.ListIndex)
            mobj���ﰲ��.��Ŀ���� = cboFeeItem.Text
        End If
    End If
    
    If CPDPages.IsValied() = False Then Exit Function

    Set obj�����¼�� = CPDPages.Get�����¼��
    If (mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan) _
        And obj�����¼��.�Ƿ��޸� Then
        If IsVisitedOtherTable(mlng����ID, mlng��ԴId, CDate(mstrCurDay)) Then
            MsgBox Format(mstrCurDay, "yyyy-mm-dd") & " ��������������������˳��ﰲ�ţ������ظ����ţ�", vbInformation, gstrSysName
            mobj���ﰲ��(1).RemoveAll
            mobj���ﰲ��(1).�Ƿ��޸� = False
            Call LoadDetailData
            Set obj�����¼�� = Nothing
            Exit Function
        End If
    End If

    If mbytPlanType = F_MonthPlan Or mbytPlanType = F_WeekPlan _
        Or mbytPlanType = F_MonthTemplet Then
        If CheckExistRecord(0, Replace(GetApplyToStr(), ",", "|"), mobj���ﰲ��) Then
            If MsgBox("ע�⣺" & vbCrLf & _
                      "      ���ֱ�Ӧ�õ����ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�Ӧ�ú��ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If mbytFun = Fun_UpdatePlan Then
        '����Ƿ�����ͣ���Һ�ԤԼ
        Set mrsVisitedRecordByDate = GetVisitedRecordByDate(mlng����ID, mstrCurDay) '���»�ȡ��ͣ�����ʹ�õİ���
        If Not mrsVisitedRecordByDate Is Nothing Then
            With mrsVisitedRecordByDate
                Do While Not .EOF
                    For i = 1 To obj�����¼��.Count
                        If Val(Nvl(!ID)) = obj�����¼��(i).��¼ID And obj�����¼��(i).�Ƿ�̶� = False Then
                            If CheckPlanIsStopOrUsed(obj�����¼��(i).��¼ID) Then
                                MsgBox "�ϰ�ʱ�� " & obj�����¼��(i).ʱ��� & " ��ͣ����Ѵ���ԤԼ�Һż�¼��" & _
                                    "����ǰ���ϰ�ʱ��Ϊ���޸�״̬�����˳���ǰ�������½�����е�����", vbInformation, gstrSysName
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

Private Function Get���ﰲ��() As ���ﰲ��
    On Error GoTo Errhand
    Call Get��ǰ���ﰲ��(mobj���ﰲ��)
    Call ChangeCurPlan(mobj���ﰲ��, mstrCurDay) '����ǰ���ŷ���δ���漯����
    Set Get���ﰲ�� = mobj���ﰲ��.Clone
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsVisitedOtherTable(ByVal lng����ID As Long, ByVal lng��ԴId As Long, _
    ByVal dtDate As Date, Optional ByVal blnOtherTable As Boolean = True, _
    Optional ByVal lng����ID As Long) As Boolean
    '�Ƿ�������������г����¼
    '������
    '   blnOtherTable �Ƿ�������������У������ǵ�ǰ�������
    '   lng����ID blnOtherTableΪFalseʱ����
    Dim strFilter As String
    
    Err = 0: On Error GoTo Errhand
    If mrsVisitedRecord Is Nothing Then Exit Function
    mrsVisitedRecord.Filter = ""
    If mrsVisitedRecord.RecordCount = 0 Then Exit Function
    
    If blnOtherTable Then
        strFilter = "����ID<>" & lng����ID
    Else
        strFilter = "����ID=" & lng����ID & " And ����ID<>" & lng����ID
    End If
    strFilter = strFilter & " And ��ԴID=" & lng��ԴId & " And ��������=#" & Format(dtDate, "yyyy-mm-dd") & "#"
    mrsVisitedRecord.Filter = strFilter
    IsVisitedOtherTable = mrsVisitedRecord.RecordCount > 0
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetApplyToStr() As String
    '��ȡӦ�����ַ���
    Dim strApplyTo As String, i As Integer, k As Integer
    Dim dtCur As Date, blnFind As Boolean
    Dim varTemp As Variant, strWeekName As String

    On Error GoTo Errhand
    If mobj���ﰲ�� Is Nothing Then Exit Function
    If picApply.Tag = "1" Then Exit Function

    '���Ű�/���Ű�
    If picApplyRule.Visible Then
        If optRule(1).Value Then '����
            dtCur = mobj���ﰲ��.��ʼʱ��
            Do While True
                If Day(dtCur) Mod 2 = 1 Then
                    If IsVisitedOtherTable(mlng����ID, mlng��ԴId, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                    End If
                End If
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
            Loop
        ElseIf optRule(2).Value Then '˫��
            dtCur = mobj���ﰲ��.��ʼʱ��
            Do While True
                If Day(dtCur) Mod 2 = 0 Then
                    If IsVisitedOtherTable(mlng����ID, mlng��ԴId, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                    End If
                End If
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
            Loop
        ElseIf optRule(3).Value Then '����
            dtCur = mobj���ﰲ��.��ʼʱ��
            Do While True
                For i = chkWeek.LBound To chkWeek.UBound
                    If chkWeek(i).Value = vbChecked Then
                        If Weekday(dtCur, vbMonday) = i + 1 Then
                            If IsVisitedOtherTable(mlng����ID, mlng��ԴId, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                                strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                            End If
                        End If
                    End If
                Next
                dtCur = DateAdd("d", 1, dtCur)
                If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
            Loop
        ElseIf optRule(4).Value Then '��ѭ,���������
            If Not (cboDays.ListIndex = -1 Or Val(txtSkip.Text) = 0) Then
                dtCur = CDate(cboDays) '��ʼʱ��
                Do While True
                    If IsVisitedOtherTable(mlng����ID, mlng��ԴId, dtCur) = False And CheckDepend(0, dtCur, , , False) Then
                        strApplyTo = strApplyTo & "," & Format(dtCur, "yyyy-mm-dd")
                    End If
                    dtCur = DateAdd("d", Val(txtSkip.Text) + 1, dtCur)
                    If DateDiff("d", mobj���ﰲ��.��ֹʱ��, dtCur) > 0 Then Exit Do
                Loop
            End If
        End If
        If strApplyTo <> "" Then strApplyTo = Mid(strApplyTo, 2)
        GetApplyToStr = strApplyTo
        Exit Function
    End If

    'ģ������ڹ����̶�ģ�����
    If picApplyWeek.Visible Then
        For i = chkWeek.LBound To chkWeek.UBound
            If chkWeek(i).Value = vbChecked Then
                strWeekName = GetWeekName(i)
                blnFind = False
                '�����޸ĵ����ڲ���Ӧ����
                If Not mcllFixedPlan Is Nothing And strWeekName <> mstrCurDay Then
                    'Array(��������,������Ŀ,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��)
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
    '��ǰ�����Ƿ�ڼ��ջ���ֹͣ����
    '���أ�ֹͣԭ���ڼ�������
    Dim i As Integer

    '��ǰ�����Ƿ񱣴��˰���
    With mobj���ﰲ��
        If Not .ͣ���¼ Is Nothing Then
            For i = 1 To .ͣ���¼.Count
                If DateDiff("d", Day, .ͣ���¼(i).��ʼʱ��) <= 0 And DateDiff("d", Day, .ͣ���¼(i).��ֹʱ��) >= 0 Then
                    CurDayIsNotVisit = .ͣ���¼(i).ͣ��ԭ��
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
    '��ʼ��ҳ��ؼ�
    Dim i As Long, ObjItem As TabControlItem
    Dim objUnit As ������λ����, lngRow As Long
    Dim intPageCount As Integer
    
    Err = 0: On Error GoTo Errhand
    tbPageSourceAndPlan.InsertItem Pg_��Դ��Ϣ, "��Դ��Ϣ", picSouceList.Hwnd, 0
    tbPageSourceAndPlan.InsertItem Pg_����Ԥ��, "����Ԥ��", picPlan.Hwnd, 0
    
    With tbPageSourceAndPlan
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionBottom
        If mbytFun = Fun_AddSignalSourcePlan Then
            .Item(Pg_��Դ��Ϣ).Selected = True
        Else
            .Item(Pg_����Ԥ��).Selected = True
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPlan(ByVal obj���ﰲ�� As ���ﰲ��)
    Err = 0: On Error GoTo Errhand
    If mbytPlanType <> F_FixedRule Then Exit Sub
    
    With vsfPlan
        .Redraw = flexRDNone
        .Cols = 3
        Set .Cell(flexcpPicture, 0, 0, .Rows - 1, 0) = imglist16.ListImages("plan_nothing").Picture
        .Cell(flexcpText, 0, 2, .Rows - 1, .Cols - 1) = ""
        
        If obj���ﰲ�� Is Nothing Then Exit Sub
        
        If Not obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then
            LoadPlanToGrid obj���ﰲ��.�ѱ�����ﰲ��, 1
        End If
        If Not obj���ﰲ��.δ������ﰲ�� Is Nothing Then
            LoadPlanToGrid obj���ﰲ��.δ������ﰲ��, 0
        End If
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPlanToGrid(ByVal obj���ﰲ�� As ���ﰲ��, ByVal bytMode As Byte)
    '���ذ���Ԥ��
    '��Σ�
    '   bytMode 0-δ���棬1-�ѱ���
    '˵����
    '   ͼ��������"plan_deleted"-��ɾ��,"plan_saved"-�ѱ���,"plan_nosave"-δ����,"plan_nothing"-�ް���
    Dim ObjItem As �����¼��, objRecord As �����¼
    Dim i As Integer, j As Integer, k As Integer, blnFindNotNull As Boolean
    Dim blnFindItem As Boolean, blnFindRecord As Boolean
    
    Err = 0: On Error GoTo Errhand
    With vsfPlan
        If obj���ﰲ�� Is Nothing Then Exit Sub
        
        For Each ObjItem In obj���ﰲ��
            blnFindItem = False
            For i = 0 To .Rows - 1
                If blnFindItem = True Then Exit For
                If .TextMatrix(i, 1) = ObjItem.�������� Then
                    blnFindItem = True
                    .Cell(flexcpText, i, 2, i, .Cols - 1) = "" '�����
                    Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_nothing").Picture
                    If ObjItem.Count = 0 And .RowData(i) = 1 Then
                        Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_deleted").Picture
                        .RowData(i) = 3 '��ɾ��
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
                        If ObjItem.�Ƿ�ɾ�� Then
                            Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_deleted").Picture
                            .RowData(i) = 3 '��ɾ��
                        Else
                            If bytMode = 1 Then
                                Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_saved").Picture
                                .RowData(i) = 1 '�ѱ���
                            Else
                                Set .Cell(flexcpPicture, i, 0) = imglist16.ListImages("plan_nosave").Picture
                                .RowData(i) = 2 'δ����
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
                        .TextMatrix(i, j) = objRecord.ʱ��� & vbCrLf & _
                            IIf(objRecord.ԤԼ���� = 1, "-", IIf(objRecord.��Լ�� = 0, IIf(objRecord.�޺��� = 0, "��", objRecord.�޺���), objRecord.��Լ��)) & _
                            "/" & IIf(objRecord.�޺��� = 0, "��", objRecord.�޺���)
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
    Case 1 '�ѱ���
        stbThis.Panels(2).Text = strTemp & " �ѱ���"
    Case 2 'δ����
        stbThis.Panels(2).Text = strTemp & " δ����"
    Case 3 '�ѱ���
        stbThis.Panels(2).Text = strTemp & " ��ɾ��"
    Case Else
        stbThis.Panels(2).Text = strTemp & " �ް���"
    End Select
End Sub

Private Sub txtSignal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '�����Ҽ��˵���ݼ������ճ��������
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
    ByVal lng��ԴId As Long, ByVal lng����ID As Long, _
    ByVal d_��ʼʱ�� As Date, ByVal d_��ֹʱ�� As Date, _
    Optional ByVal blnSaveBefore As Boolean) As Boolean
    '��Դ��Ч�Լ��
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    '�ж��Ƿ�Ϊ�����ӵĺ�Դ
    strSQL = "Select 1 From �ٴ����ﰲ�� Where ��Դid = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԴId)
    If rsTemp.EOF Then CheckSignalSource = True: Exit Function
    
    '1.��/���Ű� ת �̶��Ű�
    If blnFixedPlan Then
        If blnSaveBefore Then
            strSQL = "Select Nvl(Max(a.��ֹʱ��), To_Date('1900-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & _
                    " From �ٴ����ﰲ�� A, �ٴ������ B" & _
                    " Where a.����id = b.Id And b.�Ű෽ʽ In (1, 2) And a.��Դid = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԴId)
            If CDate(Format(d_��ʼʱ��, "yyyy-mm-dd")) < CDate(rsTemp!��ֹʱ��) Then
                ShowMsgbox "��ǰ��Դ��" & Format(rsTemp!��ֹʱ��, "yyyy-mm-dd") & "��֮ǰ���ƶ��˳��ﰲ�ţ��°��ŵ���Ч�ڲ����ٰ������ʱ�䣡"
                Exit Function
            End If
        End If
        CheckSignalSource = True: Exit Function
    End If
 
    '2.�̶��Ű� ת ��/���Ű�
    '�ж��ڵ�ǰ�����Ŀ�ʼʱ��֮�󣬹̶��Ű�ĳ����¼�Ƿ�ʹ��
    strSQL = "Select 1" & _
            " From �ٴ������¼ A, �ٴ����ﰲ�� C, �ٴ������ D" & _
            " Where A.����ID = c.ID And c.����ID = d.ID" & _
            "       And d.�Ű෽ʽ = 0 And A.��ԴID = [1] And A.�������� >= [2]" & _
            "       And Exists(Select 1 From ���˹Һż�¼ Where �����¼id = a.Id) And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԴId, d_��ʼʱ��)
    If Not rsTemp.EOF Then
        ShowMsgbox "��ǰ��Դ��" & Format(d_��ʼʱ��, "yyyy-mm-dd") & "֮������ѱ�ʹ�õİ��ţ����ܽ�����ӵ���ǰ������У�"
        Exit Function
    End If
    
    'ɾ��δ��ʹ�õĹ̶����ŵĳ����¼
    'Zl_�ٴ������¼_Delete(
    strSQL = "Zl_�ٴ������¼_Delete("
    '  ��Դid_In   �ٴ������¼.��Դid%Type,
    strSQL = strSQL & "" & lng��ԴId & ","
    '  ��ʼ����_In �ٴ������¼.��������%Type
    strSQL = strSQL & "To_Date('" & Format(d_��ʼʱ��, "yyyy-mm-dd") & "','yyyy-mm-dd'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '3.���Ű� ת ���Űࡢ���Ű� ת ���Ű�
    '�ڵ�ǰ������ʱ�䷶Χ�ڲ����г����¼
    strSQL = "Select a.��ʼʱ��, a.��ֹʱ��" & _
            " From �ٴ����ﰲ�� A, �ٴ������ B" & _
            " Where a.����ID = b.ID And b.�Ű෽ʽ In(1, 2) And a.��ԴID = [1]" & _
            "       And a.��ʼʱ�� <= [2] And a.��ֹʱ�� >= [3] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԴId, d_��ʼʱ��, d_��ֹʱ��)
    If Not rsTemp.EOF Then
        ShowMsgbox "��ǰ��Դ����Ч��(" & _
            Format(d_��ʼʱ��, "yyyy-mm-dd") & "��" & Format(d_��ֹʱ��, "yyyy-mm-dd") & _
            ")���Ѵ��ڳ��ﰲ�ţ����ܽ��ú�Դ��ӵ���ǰ�����"
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

