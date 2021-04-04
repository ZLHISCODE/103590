VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISApply 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "电子病历访问申请"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20475
   Icon            =   "frmCISApply.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   20475
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   15015
      TabIndex        =   10
      Top             =   5040
      Width           =   15015
      Begin VB.Frame picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "授权信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   12840
         TabIndex        =   17
         Top             =   840
         Width           =   5175
         Begin VSFlex8Ctl.VSFlexGrid vsInfo 
            Height          =   6915
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4750
            _cx             =   1989550266
            _cy             =   1989554085
            Appearance      =   0
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
            MouseIcon       =   "frmCISApply.frx":6852
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   10000
            ColWidthMin     =   4650
            ColWidthMax     =   10000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCISApply.frx":712C
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.Frame fraFillter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "查询过滤"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   17055
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   277
            Width           =   1365
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "查询(&F)"
            Height          =   375
            Left            =   13080
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已撤消"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   12150
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已拒绝"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   10920
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已作废"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   9675
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已审批"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   8445
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "待审批"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   7200
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   3
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2595
            TabIndex        =   1
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   217448451
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4890
            TabIndex        =   2
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   217448451
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "frmCISApply.frx":71C7
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   4
            Left            =   11850
            Picture         =   "frmCISApply.frx":7751
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   10605
            Picture         =   "frmCISApply.frx":DFA3
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   9375
            Picture         =   "frmCISApply.frx":147F5
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   8130
            Picture         =   "frmCISApply.frx":1B047
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   6900
            Picture         =   "frmCISApply.frx":21899
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            X1              =   4605
            X2              =   4805
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "申请时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   330
            Width           =   855
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   7275
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   12645
         _cx             =   1989564192
         _cy             =   1989554720
         Appearance      =   0
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
         MouseIcon       =   "frmCISApply.frx":280EB
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISApply.frx":289C5
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   14
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":28A60
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":2F2C2
            Key             =   "boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":35B24
            Key             =   "访问时限"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":360BE
            Key             =   "访问内容"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":36658
            Key             =   "访问医生"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":36BF2
            Key             =   "访问病人"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":3718C
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":372E6
            Key             =   "unCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMec 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   6480
      ScaleHeight     =   4575
      ScaleWidth      =   9615
      TabIndex        =   11
      Top             =   1200
      Width           =   9615
      Begin VB.PictureBox picVLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9300
         Left            =   5920
         MousePointer    =   9  'Size W E
         ScaleHeight     =   9300
         ScaleWidth      =   30
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   120
         Width           =   30
      End
      Begin VB.Frame fraPatiFilter 
         Appearance      =   0  'Flat
         Caption         =   "病人查找"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   9495
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   5895
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   5535
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   5655
            _Version        =   589884
            _ExtentX        =   9975
            _ExtentY        =   9763
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1800
            TabIndex        =   35
            Top             =   1250
            Width           =   2535
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "按手术查找"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   3
            Left            =   4320
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "按诊断查找"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   2
            Left            =   2960
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "按标识查找"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   1
            Left            =   1600
            TabIndex        =   30
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "按科室查找"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1200
            Width           =   2565
         End
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   600
            ScaleHeight     =   240
            ScaleWidth      =   1140
            TabIndex        =   21
            Top             =   1250
            Width           =   1170
            Begin VB.ComboBox cboFind 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   300
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   -30
               Width           =   1215
            End
         End
         Begin VB.ComboBox cboSelectTime 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   750
            Width           =   2565
         End
         Begin VB.CommandButton cmdPatiFind 
            Caption         =   "查找(&F)"
            Height          =   375
            Left            =   4440
            TabIndex        =   37
            Top             =   1150
            Width           =   1215
         End
         Begin VB.Label lblDept 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "↓病人科室"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   720
            TabIndex        =   23
            Top             =   1290
            Width           =   900
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "时间范围(&T)"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   720
            TabIndex        =   22
            Top             =   840
            Width           =   990
         End
      End
      Begin VB.PictureBox picMecInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   6120
         ScaleHeight     =   3735
         ScaleWidth      =   3855
         TabIndex        =   19
         Top             =   240
         Width           =   3855
         Begin VB.PictureBox picShow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   8175
            Left            =   0
            ScaleHeight     =   8175
            ScaleWidth      =   11775
            TabIndex        =   25
            Top             =   480
            Width           =   11775
            Begin VB.PictureBox PicNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   4695
               Left            =   1320
               ScaleHeight     =   4665
               ScaleWidth      =   7665
               TabIndex        =   27
               Top             =   2400
               Width           =   7695
               Begin VB.PictureBox picNoUse 
                  BorderStyle     =   0  'None
                  Height          =   2535
                  Left            =   0
                  Picture         =   "frmCISApply.frx":37440
                  ScaleHeight     =   2535
                  ScaleWidth      =   7815
                  TabIndex        =   28
                  Top             =   960
                  Width           =   7815
               End
            End
            Begin XtremeSuiteControls.TabControl tbcMec 
               Height          =   6420
               Left            =   120
               TabIndex        =   26
               Top             =   480
               Width           =   8610
               _Version        =   589884
               _ExtentX        =   15187
               _ExtentY        =   11324
               _StockProps     =   64
            End
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5580
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   8130
      _Version        =   589884
      _ExtentX        =   14340
      _ExtentY        =   9842
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   10560
      Width           =   20475
      _ExtentX        =   36116
      _ExtentY        =   635
      SimpleText      =   $"frmCISApply.frx":3CD45
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCISApply.frx":3CD8C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   31036
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   600
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCISApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdtBegin As Date, mdtEnd As Date
Private mintPreTime As Integer
Private mobjArchiveView As frmArchiveView
Private mrsTmp As ADODB.Recordset     '当前医生信息缓存数据集
Private mstrDeptIds As String      '当前医生科室IDs

Private Enum colList
    COL_申请ID = 1
    COL_访问内容 = 2
    COL_内容时限 = 3
    COL_撤消时间 = 4
    COL_撤消人 = 5
    COL_申请人 = 6

    COL_申请时间 = 7
    COL_申请访问病人 = 8
    COL_访问开始时间 = 9
    COL_访问结束时间 = 10
    COL_申请原因 = 11
    COL_审批状态 = 12
End Enum

Private Enum RowInfo
    Row_访问病人标题 = 0
    Row_访问病人 = 1
    Row_内容时限标题 = 3
    Row_内容时限 = 4
    Row_访问内容标题 = 6
    Row_访问内容 = 7
End Enum


Private Enum colPati
    col_选择 = 0
    col_病人Id = 1
    col_姓名 = 2
    col_性别 = 3
    col_年龄 = 4
    COL_标识号 = 5
    col_科室 = 6
    COL_当前状态 = 7
    col_科室ID = 8
    col_就诊ID = 9
End Enum

Private Enum CmdIndex
    Cmd_所有科室 = 9999991
    Cmd_门诊科室 = 9999992
    Cmd_住院科室 = 9999993
End Enum


Private Sub cboSelectTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    If cboSelectTime.ListIndex = mintPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtBegin, mdtEnd, cboSelectTime) Then
            '取消时恢复原来的选择
            Call Cbo.SetIndex(cboSelectTime.hwnd, mintPreTime)
            Exit Sub
        End If
    Else
        mdtEnd = Format(datCurr, "yyyy-MM-dd 23:59:59")
        mdtBegin = datCurr - intDateCount
    End If
    If mdtBegin = CDate(0) Or mdtEnd = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "范围：" & Format(mdtBegin, "yyyy-MM-dd") & " 至 " & Format(mdtEnd, "yyyy-MM-dd")
    End If
    mintPreTime = cboSelectTime.ListIndex
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate

    dtpTime(0).MaxDate = curDate + 1
    dtpTime(1).MaxDate = curDate + 1

    
    Select Case cboTime.ListIndex
    Case 0 '今日
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '最近二天
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '最近三天
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '最近一周
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '最近一月
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 '指  定
        If Me.Visible Then
            dtpTime(0).SetFocus
        End If
    End Select
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngApplyID As Long
    
    Select Case Control.ID
        Case conMenu_Edit_ApplyAdd
            If frmCISApplyEdit.ShowEdit(Me, 0, lngApplyID, IIf(tbcSub.Selected.Tag = "访问电子病历", GetPatiRs, Nothing)) Then
                Call LoadList(lngApplyID)
            End If
        Case conMenu_Edit_ApplyEdit
            If Val(vsList.TextMatrix(vsList.Row, COL_申请ID)) = 0 Then Exit Sub
            lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_申请ID))
            If frmCISApplyEdit.ShowEdit(Me, 1, lngApplyID) Then
                Call LoadList(lngApplyID)
            End If
        Case conMenu_Edit_ApplyBack
            If Val(vsList.TextMatrix(vsList.Row, COL_申请ID)) = 0 And vsList.TextMatrix(vsList.Row, COL_审批状态) <> "待审批" Then Exit Sub
            lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_申请ID))
            If ApplyBack(lngApplyID) Then
                Call LoadList(lngApplyID)
            End If
        Case conMenu_View_Refresh
            If tbcSub.Selected.Tag = "申请记录" Then
                Call LoadList
            Else
                Call LoadPati
            End If
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
            
        Case Cmd_所有科室, Cmd_门诊科室, Cmd_住院科室
            lblDept.Tag = Control.Parameter
            lblDept.Caption = Decode(lblDept.Tag, "", "↓病人科室", "门诊", "↓门诊科室", "住院", "↓住院科室")
            Call LoadDept
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Function ApplyBack(lngApplyID As Long) As Boolean
    Dim strSql As String
    Dim curDate As Date
    Dim blnTran As Boolean
    
    On Error GoTo errH
    If MsgBox("确定要撤消选中的授权申请记录吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    curDate = zlDatabase.Currentdate
    strSql = "Zl_电子病历访问申请_审批状态(" & lngApplyID & ",4,'" & UserInfo.姓名 & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    ApplyBack = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lblDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, True)
End Sub



Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_ApplyEdit
        If vsList.Row <= 0 And tbcSub.Selected.Tag = "申请记录" Then Control.Enabled = False: Exit Sub
        Control.Visible = tbcSub.Selected.Tag = "申请记录" And vsList.TextMatrix(vsList.Row, COL_审批状态) = "待审批"
    Case conMenu_Edit_ApplyBack
        If vsList.Row <= 0 And tbcSub.Selected.Tag = "申请记录" Then Control.Enabled = False: Exit Sub
        Control.Visible = tbcSub.Selected.Tag = "申请记录" And vsList.TextMatrix(vsList.Row, COL_审批状态) = "待审批"
    Case Cmd_所有科室, Cmd_门诊科室, Cmd_住院科室
         Control.Checked = Control.Parameter = lblDept.Tag
    End Select
End Sub

Private Sub chkFilter_Click(Index As Integer)
    Dim i As Long
    Dim blnCheck As Boolean
    
    For i = 0 To 4
        If chkFilter(i).Value = 1 Then
            blnCheck = True
            Exit For
        End If
    Next
    If Not blnCheck Then
        MsgBox "请至少选择一种分类用于过滤。", vbInformation, gstrSysName
        chkFilter(Index).Value = 1
        Exit Sub
    End If
End Sub

Private Sub cmdFind_Click()
    Call LoadList
End Sub

Public Function GetRs病人姓名(rsTmp As ADODB.Recordset) As Boolean
    Dim str病人ids As String
    Dim arrTmp As Variant
    Dim colPati As Collection
    Dim i As Long, j As Long
    Dim str姓名 As String, colValue As Collection
    
    If rsTmp Is Nothing Then Exit Function
    If rsTmp.EOF Then Exit Function
    
    
    '加载病人信息
    str病人ids = ""
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
             If rsTmp!病人ids & "" <> "" Then
                arrTmp = Split(rsTmp!病人ids & "", ",")
                For j = LBound(arrTmp) To UBound(arrTmp)
                    If InStr("," & str病人ids & ",", "," & Val(arrTmp(j)) & ",") = 0 Then
                       str病人ids = str病人ids & "," & Val(arrTmp(j))
                    End If
                Next
             End If
             rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    End If

    
    If str病人ids <> "" Then
        str病人ids = Mid(str病人ids, 2)
        Set colPati = PatiSvrGetpatiinfo(1, 0, 1241, 0, 2, "", "", "", "", str病人ids)
        
        If Not colPati Is Nothing Then
            Set rsTmp = zlDatabase.CopyNewRec(rsTmp)
            Do While Not rsTmp.EOF
               If rsTmp!病人ids & "" <> "" Then
                    arrTmp = Split(rsTmp!病人ids & "", ",")
                    str姓名 = ""
                    For j = LBound(arrTmp) To UBound(arrTmp)
                        If Val(arrTmp(j)) <> 0 Then
                            Set colValue = GetColObj(colPati, "_" & arrTmp(j))
                            If Not colValue Is Nothing Then
                                If GetColVal(colValue, "_pati_name") <> "" Then
                                    str姓名 = str姓名 & "," & GetColVal(colValue, "_pati_name")
                                End If
                            End If
                        End If
                    Next
                End If
                
                If str姓名 <> "" Then
                    str姓名 = Mid(str姓名, 2)
                    rsTmp!病人姓名 = str姓名
                End If
                
                rsTmp.MoveNext
            Loop
            rsTmp.MoveFirst
        End If
    End If
End Function

Private Sub LoadList(Optional lng申请id As Long)
    Dim strSql As String
    Dim strFilter As String
    Dim str已撤消 As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date
    
    For i = 0 To 3
        If chkFilter(i).Value = 1 Then strFilter = strFilter & "," & i
    Next
    strFilter = Mid(strFilter, 2)
    
    '过滤已撤消记录
    If chkFilter(4).Value = 0 Then
        str已撤消 = " And A.撤消时间 is null"
    Else
        If strFilter = "" Then
            str已撤消 = " And A.撤消时间 is not null"
        Else
            str已撤消 = " Or A.撤消时间 is not null)"
        End If
    End If
    
    On Error GoTo errH
    If cboTime.ListIndex <> 5 Then
        curDate = zlDatabase.Currentdate
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSql = "Select a.Id, a.访问开始时间, a.访问结束时间, a.内容时限, a.申请原因, a.审批状态, a.申请人, a.申请时间,A.撤消时间,A.撤消人," & vbNewLine & _
                "       f_List2str(Cast(Collect(b.病人id || '') As t_Strlist)) As 病人ids,null as 病人姓名" & vbNewLine & _
                "From 电子病历访问申请 A, 电子病历申请访问病人 B" & vbNewLine & _
                "Where a.Id = b.申请id  And a.申请时间 Between [1] And [2] And a.申请人 = [3]" & vbNewLine & _
                IIf(strFilter <> "", IIf(chkFilter(4).Value = 1, " And (Instr([4], a.审批状态) > 0", " And Instr([4], a.审批状态) > 0"), "") & str已撤消 & vbNewLine & _
                "Group By a.Id, a.访问开始时间, a.访问结束时间, a.内容时限, a.申请原因, a.审批状态, a.申请人, a.申请时间,A.撤消时间,A.撤消人" & vbNewLine & _
                "Order by a.审批状态,A.id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtpTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboTime.ListIndex <> 5, dtpTime(1).Value + 1, dtpTime(1).Value), "yyyy-MM-dd hh:mm")), UserInfo.姓名, strFilter)
    
    Call GetRs病人姓名(rsTmp)
    With vsList
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '隐藏列
                .TextMatrix(i, COL_申请ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_内容时限) = Val(rsTmp!内容时限 & "")
                .TextMatrix(i, COL_撤消时间) = Format(rsTmp!撤消时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_撤消人) = rsTmp!撤消人 & ""
                '显示列
                .TextMatrix(i, COL_申请人) = rsTmp!申请人 & ""
                .TextMatrix(i, COL_申请时间) = Format(rsTmp!申请时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_申请访问病人) = rsTmp!病人姓名 & ""
                .TextMatrix(i, COL_访问开始时间) = Format(rsTmp!访问开始时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_访问结束时间) = Format(rsTmp!访问结束时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_申请原因) = rsTmp!申请原因 & ""
                
                If rsTmp!撤消时间 & "" <> "" Then
                    .TextMatrix(i, COL_审批状态) = "已撤消"
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(4).Picture
                Else
                    .TextMatrix(i, COL_审批状态) = Decode(Val(rsTmp!审批状态 & ""), 0, "待审批", 1, "已审批", 2, "已作废", 3, "已拒绝")
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(Val(rsTmp!审批状态 & "")).Picture
                End If

                If Val(rsTmp!ID & "") = lng申请id Then
                    .Row = i
                End If
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             If tbcSub.Selected.Tag = "申请记录" Then stbThis.Panels(2).Text = "当前过滤查找到 " & rsTmp.RecordCount & " 份申请信息"
        Else
            .Rows = .FixedRows + 1
            If tbcSub.Selected.Tag = "申请记录" Then stbThis.Panels(2).Text = "当前过滤没有查找到申请信息"
        End If
        
        If .Row <= 0 Then .Row = .Rows - 1
        If Me.Tag = "1" And .Visible Then .SetFocus
        .WordWrap = True
        '自动调整行高
        .AutoSize COL_申请访问病人, COL_申请原因
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub cmdPatiFind_Click()
    Call LoadPati
End Sub

Private Sub fraPatiFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call SetLBLFace(lblDept, False)
End Sub

Private Sub optFind_Click(Index As Integer)
    Call SetFindCtl
    If optFind(0).Value Then
        If cboDept.Visible Then cboDept.SetFocus
    Else
        If txtFind.Visible Then txtFind.SetFocus
    End If
End Sub

Private Sub SetFindCtl()
    txtFind.Text = ""
    cboDept.Visible = optFind(0).Value
    txtFind.Visible = Not optFind(0).Value
    lblDept.Visible = Not optFind(1).Value
    picTmp(1).Visible = optFind(1).Value

    lblDept.Caption = IIf(optFind(0).Value, "↓病人科室", IIf(optFind(2).Value, "病人诊断(&D)", IIf(optFind(3).Value, "病人手术(&O)", "")))
    lblDept.Tag = ""
End Sub


Private Sub picVLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picMec_Resize
End Sub


Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    stbThis.Panels(2).Text = IIf(tbcSub.Selected.Tag = "访问电子病历", "访问电子病历", "查看申请记录")
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Or NewCol < 0 Then Exit Sub
    If vsList.Col >= vsList.FixedCols Then
        vsList.ForeColorSel = vsList.Cell(flexcpForeColor, NewRow, NewCol)
    End If
    With vsInfo
        If Val(vsList.TextMatrix(NewRow, COL_申请ID)) <> 0 Then
            '访问病人
            .TextMatrix(Row_访问病人, 0) = vsList.TextMatrix(NewRow, COL_申请访问病人) & ""
            
            '内容时限
            .TextMatrix(Row_内容时限, 0) = "于 " & Format(vsList.TextMatrix(NewRow, COL_访问开始时间), "yyyy-mm-dd hh:mm") & vbCrLf & "至 " & _
                                        Format(vsList.TextMatrix(NewRow, COL_访问结束时间), "yyyy-mm-dd hh:mm") & "期间" & vbCrLf & "访问病人" & Decode(Val(vsList.TextMatrix(NewRow, COL_内容时限)), 0, "所有病历内容", 1, "未归档的病历", "已归档的病历")
                             
            '访问内容
            .TextMatrix(Row_访问内容, 0) = GetXmlInfo(NewRow)
        Else
            .TextMatrix(Row_访问病人, 0) = ""
            .TextMatrix(Row_内容时限, 0) = ""
            .TextMatrix(Row_访问内容, 0) = ""
        End If
        .WordWrap = True
        '自动调整行高
        .AutoSize 0
    End With
End Sub


Private Function GetXmlString(objXML As Object, ByVal strNode As String, ByRef strValue As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strValue = ""
    If objXML.GetMultiNodeRecord(strNode, rsTmp) Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!node_value
                rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
    End If
    GetXmlString = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetXmlInfo(lngRow As Long) As String
    '获取申请内容的Xml并解析
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String
    Dim strOut As String
    Dim strTmp As String
    
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function
    If Val(vsList.TextMatrix(lngRow, COL_申请ID)) = 0 Then Exit Function
    
    '读取缓存
    If vsList.TextMatrix(lngRow, COL_访问内容) <> "" Then GetXmlInfo = vsList.TextMatrix(lngRow, COL_访问内容): Exit Function
    
    strXML = Sys.ReadXML("电子病历访问申请", "访问内容", "ID=[1]", strErr, Val(vsList.TextMatrix(lngRow, COL_申请ID)))
    If Err.Number = 0 And strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        Exit Function
    End If
    
    If objXML.OpenXMLDocument(strXML) = False Then Exit Function

    '所有内容
    strValue = "": Call objXML.GetSingleNodeValue("all_files", strValue, xsNumber)
    If Val(strValue) = 1 Then
        strOut = "无限制访问所有内容"
    Else
        '病案首页、医嘱、临床路径
        strValue = "": Call objXML.GetSingleNodeValue("medical_record", strValue, xsNumber): If Val(strValue) = 1 Then strOut = "病案首页、" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("advice", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "病人医嘱、" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("cispath", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "临床路径、" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("patipeis", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "体检报告、" & vbCrLf & vbCrLf
        
        '护理记录
        strValue = "": Call objXML.GetSingleNodeValue("nursing_record", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/nursing_all", strValue, xsNumber)
            If Val(strValue) = 1 Then
                strOut = strOut & "护理记录(所有护理记录)" & vbCrLf & vbCrLf
            Else
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/thermometer", strValue, xsNumber): If Val(strValue) = 1 Then strTmp = "体温单、"
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/record_file", strValue, xsNumber)
                If Val(strValue) = 1 Then
                    Call GetXmlString(objXML, "nursing_info/file_name", strValue)
                    strValue = Replace(strValue, ",", "、")
                    strTmp = strTmp & strValue
                Else
                    strTmp = Replace(strTmp, "、", "")
                End If
                strOut = strOut & "护理记录" & vbCrLf & "(记录范围：" & strTmp & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '检查报告
        strValue = "": Call objXML.GetSingleNodeValue("pacs_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("pacs_info/pacs_type", strValue, xsNumber)
            'pacs_type =0所有检查报告 =1指定类型的检查报告
            If Val(strValue) = 0 Then
                strOut = strOut & "检查报告(所有检查报告)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "pacs_info/pacs_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "检查报告" & vbCrLf & "(类型范围：" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '检验报告
        strValue = "": Call objXML.GetSingleNodeValue("lis_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("lis_info/lis_type", strValue, xsNumber)
            'lis_type =0 所有检验报告 =1指定类型的检验报告
            If Val(strValue) = 0 Then
                strOut = strOut & "检验报告(所有检验报告)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "lis_info/lis_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "检验报告" & vbCrLf & "(类型范围：" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '电子病历
        strValue = "": Call objXML.GetSingleNodeValue("emr", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("emr_info/emr_type", strValue, xsNumber)
            'emr_type =0 所有电子病历  =1指定类型的电子病历  =1指定种类的电子病历
            If Val(strValue) = 0 Then
                strOut = strOut & "电子病历(所有电子病历)" & vbCrLf & vbCrLf
            ElseIf Val(strValue) = 1 Then
                Call GetXmlString(objXML, "emr_info/standard_class/class_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "电子病历" & vbCrLf & "(病历类型范围：" & strValue & ")" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "emr_info/antetype_class/class_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "电子病历" & vbCrLf & "(病历范围：" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
    End If
    
    If Right(strOut, 5) = "、" & vbCrLf & vbCrLf Then strOut = Left(strOut, Len(strOut) - 5)
    
    '缓存内容数据
    vsList.TextMatrix(lngRow, COL_访问内容) = strOut
    GetXmlInfo = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Load()
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    Call MainDefCommandBar
    
    '初始化拖动定位
    Me.picVLine.Left = 5895
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "申请记录", picApply.hwnd, 0).Tag = "申请记录"
        .InsertItem(1, "访问电子病历", picMec.hwnd, 0).Tag = "访问电子病历"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call InitListTable
    
    '初始化详情表格
    With vsInfo
        '访问病人
        .TextMatrix(Row_访问病人标题, 0) = "访问病人："
        .Cell(flexcpForeColor, Row_访问病人标题, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_访问病人标题, 0) = img16.ListImages("访问病人").Picture
        .Cell(flexcpFontBold, Row_访问病人标题, 0) = True
        
        '内容时限
        .TextMatrix(Row_内容时限标题, 0) = "内容时限："
        .Cell(flexcpForeColor, Row_内容时限标题, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_内容时限标题, 0) = img16.ListImages("访问时限").Picture
        .Cell(flexcpFontBold, Row_内容时限标题, 0) = True

        '访问内容
        .TextMatrix(Row_访问内容标题, 0) = "访问内容："
        .Cell(flexcpForeColor, Row_访问内容标题, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_访问内容标题, 0) = img16.ListImages("访问内容").Picture
        .Cell(flexcpFontBold, Row_访问内容标题, 0) = True

        .WordWrap = True
        '自动调整行高
        .AutoSize 0
    End With
    
    '---cboTime
    cboTime.AddItem "今    日"
    cboTime.AddItem "最近二天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指  定]"
    cboTime.ListIndex = 3
    
    Call InitSelectTime
    
    Call LoadDept
    
    Call InitReportColumn
    
    Call SetFindCtl
    
    '执行结果下拉菜单初始化
    cboFind.Clear
    cboFind.AddItem "姓名"
    cboFind.AddItem "身份证号"
    cboFind.AddItem "门诊号"
    cboFind.AddItem "住院号"
    cboFind.AddItem "病人ID"
    cboFind.ListIndex = 0
    
    Call GetFrom
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Call LoadList
    Me.Tag = "1"
End Sub
'


Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "新增申请(&A)")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "调整申请(&E)")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyBack, "撤消申请(&Q)")
            objControl.IconId = 5019
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "新增申请")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "调整申请")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyBack, "撤消申请")
            objControl.IconId = 5019
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call cbsMain_Resize
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = Me.Height - stbThis.Height - 1500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mobjArchiveView
    Set mobjArchiveView = Nothing
    Set mrsTmp = Nothing
    mstrDeptIds = ""
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    mintPreTime = -1
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtBegin = datCurr
    mdtEnd = mdtBegin - 7
    
    cboSelectTime.Clear '出院
    With cboSelectTime
        .AddItem "今天内"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天内"
        .ItemData(.NewIndex) = 1
        .AddItem "前天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 3
    mintPreTime = cboSelectTime.ListIndex
End Sub

Private Sub picApply_Resize()
    On Error Resume Next
    '固定详细信息4000长度
    picInfo.Width = 5000

    fraFillter.Top = 100: fraFillter.Left = 30
    fraFillter.Width = picApply.Width - 60
    
    vsList.Top = fraFillter.Top + fraFillter.Height + 150: vsList.Height = picApply.Height - fraFillter.Height - 260

    
    vsList.Left = fraFillter.Left
    vsList.Width = fraFillter.Width - 5000 - 30
    
    picInfo.Top = vsList.Top - 70: picInfo.Left = vsList.Left + vsList.Width + 50
    picInfo.Height = vsList.Height + 70
    vsInfo.Height = picInfo.Height - 300
End Sub


Private Sub picVline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picVLine.Left = Me.picVLine.Left + X
    End If
End Sub

Private Sub picMec_Resize()

    Me.picVLine.Top = 0
    Me.picVLine.Height = Me.picMec.Height
    If Me.picVLine.Left < 100 Then Me.picVLine.Left = 100
    If Me.picVLine.Left > Me.picMec.Width - 100 Then Me.picVLine.Left = Me.picMec.Width - 100


    On Error Resume Next
    fraPatiFilter.Top = 100: fraPatiFilter.Left = 30
    fraPatiFilter.Width = Me.picVLine.Left - Me.fraPatiFilter.Left
    rptPati.Width = fraPatiFilter.Width - 200
    fraPatiFilter.Height = picMec.Height - 100
    
    rptPati.Height = fraPatiFilter.Height - rptPati.Top - 100
    
    picMecInfo.Top = 180: picMecInfo.Height = fraPatiFilter.Height - 80
    picMecInfo.Left = fraPatiFilter.Width + 60
    picMecInfo.Width = picMec.Width - picMecInfo.Left - 60
    
    '隐藏tab标签
    picShow.Top = -360: picShow.Left = 0
    picShow.Width = picMecInfo.Width: picShow.Height = picMecInfo.Height + 360
    
    tbcMec.Top = 0: tbcMec.Left = 0
    tbcMec.Width = picShow.Width: tbcMec.Height = picShow.Height
    
End Sub


Private Sub InitListTable()
'功能：初始化列表清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "申请id;访问内容;内容时限;撤消时间;撤消人;申请人;" & _
                "申请时间,2000,1;申请访问病人,4000,1;访问开始时间,2000,1;访问结束时间,2000,1;申请原因,3800,1;审批状态,1050,4"

    arrHead = Split(strHead, ";")
    With vsList
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDNone
    End With
End Sub

Private Sub LoadDept()
'加载查询科室
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long
    Dim strTmp As String
    
    strSql = "Select B.ID,B.编码,B.名称 From " & _
            " 部门表 B, 部门性质说明 C" & vbNewLine & _
            " Where B.Id = C.部门id " & _
            "  And C.工作性质 = '临床' " & Decode(lblDept.Tag, "", " And C.服务对象 <> 0 ", "门诊", " And C.服务对象 in (1,3) ", "住院", " And C.服务对象 in (2,3) ") & "  And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) Order By B.编码"
    On Error GoTo errH
    cboDept.Clear
    '所有部门
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID & ""
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
    
    '缓存操作员科室id字符串
    
    strSql = "Select b.Id" & vbNewLine & _
        "From 部门人员 A, 部门表 B, 部门性质说明 C" & vbNewLine & _
        "Where b.Id = c.部门id And a.部门id = b.Id And a.人员id = [1] And c.工作性质 = '临床' And c.服务对象 <> 0 And" & vbNewLine & _
        "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
        "Order By b.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            strTmp = strTmp & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        strTmp = strTmp & ","
    End If
    mstrDeptIds = strTmp
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        Set objCol = .Columns.Add(col_选择, "", 18, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False)
        Set objCol = .Columns.Add(col_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(col_性别, "性别", 30, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 30, True)
        Set objCol = .Columns.Add(COL_标识号, "标识号", 100, True)
        Set objCol = .Columns.Add(col_科室, "科室", 80, True)
        Set objCol = .Columns.Add(COL_当前状态, "当前状态", 150, True)
        Set objCol = .Columns.Add(col_科室ID, "科室ID", 0, False)
        Set objCol = .Columns.Add(col_就诊ID, "就诊ID", 0, False)
        
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
            .HighlightBackColor = &HFFEDCA
            .HighlightForeColor = vbBlack
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub


Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call LoadPati
    Else
        If cboFind.Visible Then
        Select Case cboFind.Text
            Case "住院号", "门诊号", "病人ID"
                If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "身份证号"
                If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "姓名"
        End Select
        End If
    End If
End Sub


Private Sub LoadPati()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Dim colPati As Collection, str病人ids As String, i As Long
    
    On Error GoTo errH
    
    
    If (optFind(1).Value Or optFind(2).Value Or optFind(3).Value) And txtFind.Text = "" Then Exit Sub
    '按科室查找、按标识查找
    If optFind(0).Value = True Or optFind(1).Value = True Then
        If cboFind.Text = "门诊号" Then
            strSql = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室,g.ID As 科室ID, d.标识号,D.就诊ID,d.当前状态" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号, a.id As 就诊ID,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                        "              From 病人挂号记录 A" & vbNewLine & _
                        "              Where A.记录状态=1 And a.执行时间 Between [2] And [3]" & IIf(txtFind.Text = "", "", " And A.门诊号=[4]") & ") C) D, 部门表 G" & vbNewLine & _
                        "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.科室=[1]") & vbNewLine & _
                        "Order By d.操作时间 Desc"
        ElseIf cboFind.Text = "住院号" Then
            strSql = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室,g.ID As 科室ID, d.标识号,D.就诊ID,d.当前状态" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号, B.主页ID As 就诊ID,decode(B.出院日期,null,'在院','第'||B.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                        "              From 病案主页 B" & vbNewLine & _
                        "              Where  b.入院日期 Between [2] And [3]" & IIf(txtFind.Text = "", "", " And B.住院号=[4]") & ") C) D, 部门表 G" & vbNewLine & _
                        "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.科室=[1]") & vbNewLine & _
                        "Order By d.操作时间 Desc"
        Else
        
        
            If cboFind.Text = "身份证号" Then
                Set colPati = PatiSvrGetpatiinfo(1, 0, 1240, 0, 2, txtFind.Text)
            End If
        
            If Not colPati Is Nothing Then
                If colPati.Count > 0 Then
                    For i = 1 To colPati.Count
                        If InStr("," & str病人ids & ",", "," & Val(GetColVal(colPati(i), "_pati_id")) & ",") = 0 Then
                           str病人ids = str病人ids & "," & Val(GetColVal(colPati(i), "_pati_id"))
                        End If
                    Next
                End If
            End If
            If str病人ids <> "" Then str病人ids = Mid(str病人ids, 2)
            If (optFind(0).Value = True And lblDept.Tag = "") Or optFind(1).Value = True Then
                strSql = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室,g.ID As 科室ID, d.标识号,D.就诊ID,d.当前状态" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号, a.id As 就诊ID,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                            "              From 病人挂号记录 A" & vbNewLine & _
                            "              Where A.记录状态=1  And a.执行时间 Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "身份证号", " A.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "病人ID", "A.病人ID =[4]", "姓名", "A.姓名 like [4]")) & vbNewLine & _
                            "              Union All" & vbNewLine & _
                            "              Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号, B.主页ID As 就诊ID,decode(B.出院日期,null,'在院','第'||B.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                            "              From 病案主页 B" & vbNewLine & _
                            "              Where b.入院日期 Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "身份证号", " B.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "病人ID", "B.病人ID =[4]", "姓名", "B.姓名 like [4]")) & ") C) D, 部门表 G" & vbNewLine & _
                            "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.科室=[1]") & vbNewLine & _
                            "Order By d.操作时间 Desc"
            ElseIf optFind(0).Value = True And lblDept.Tag = "门诊" Then
                strSql = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室,g.ID As 科室ID, d.标识号,D.就诊ID,d.当前状态" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号, a.id As 就诊ID,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                            "              From 病人挂号记录 A" & vbNewLine & _
                            "              Where A.记录状态=1  And a.执行时间 Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "身份证号", " A.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "病人ID", "A.病人ID =[4]", "姓名", "A.姓名 like [4]")) & ") C) D, 部门表 G" & vbNewLine & _
                            "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.科室=[1]") & vbNewLine & _
                            "Order By d.操作时间 Desc"
            ElseIf optFind(0).Value = True And lblDept.Tag = "住院" Then
                strSql = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室,g.ID As 科室ID, d.标识号,d.当前状态" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号, B.主页ID As 就诊ID,decode(B.出院日期,null,'在院','第'||B.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                            "              From 病案主页 B" & vbNewLine & _
                            "              Where b.入院日期 Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "身份证号", " B.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "病人ID", "B.病人ID =[4]", "姓名", "B.姓名 like [4]")) & ") C) D, 部门表 G" & vbNewLine & _
                            "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.科室=[1]") & vbNewLine & _
                            "Order By d.操作时间 Desc"
            End If
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, cboDept.ItemData(cboDept.ListIndex), mdtBegin, mdtEnd, IIf(InStr(",门诊号,住院号,病人ID,", cboFind.Text) > 0, Val(txtFind.Text), IIf(cboFind.Text = "姓名", txtFind.Text & "%", txtFind.Text)), str病人ids)
    ElseIf optFind(2).Value = True Then '按诊断查找
        strSql = "Select d.Id, d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, g.Id As 科室id, d.标识号,D.就诊ID, d.当前状态" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号, a.id As 就诊ID, Decode(a.执行状态, 1, '在' || To_Char(a.执行时间, 'yyyy-mm-dd') || '门诊就诊离院', '门诊正在就诊') As 当前状态" & vbNewLine & _
                "              From 病人挂号记录 A, 病人诊断记录 M" & vbNewLine & _
                "              Where  a.病人id = m.病人id And a.Id = m.主页id And a.记录状态 = 1 And a.执行时间 Between [1] And [2] And" & vbNewLine & _
                "                    (Exists" & vbNewLine & _
                "                     (Select 1 From 疾病编码目录 N Where n.Id = m.疾病id And (n.编码 Like [3] Or n.名称 Like [3] Or UPPER(n.简码) Like [3])) Or Exists" & vbNewLine & _
                "                     (Select 1 From 疾病诊断目录 I,疾病诊断别名 Z  Where i.Id = m.诊断id AND i.ID=Z.诊断ID And (i.编码 Like [3] Or i.名称 Like [3] or UPPER(Z.简码) like [3])))" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号, B.主页ID As 就诊ID," & vbNewLine & _
                "                     Decode(b.出院日期, Null, '在院', '第' || B.主页id || '次住院离院') As 当前状态" & vbNewLine & _
                "              From 病案主页 B, 病人诊断记录 O" & vbNewLine & _
                "              Where  b.病人id = o.病人id And b.主页id = o.主页id And b.入院日期 Between [1] And [2] And" & vbNewLine & _
                "                    (Exists (Select 1 From 疾病编码目录 N Where n.Id = o.疾病id And (n.编码 Like [3] Or n.名称 Like [3] Or UPPER(n.简码) Like [3])) Or Exists" & vbNewLine & _
                "                     (Select 1 From 疾病诊断目录 I,疾病诊断别名 Z Where i.Id = O.诊断id AND i.ID=Z.诊断ID And (i.编码 Like [3] Or i.名称 Like [3] or UPPER(Z.简码) like [3]))) ) C) D, 部门表 G" & vbNewLine & _
                "Where g.Id = d.科室 And d.Top = 1" & vbNewLine & _
                "Order By d.操作时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mdtBegin, mdtEnd, UCase(txtFind.Text) & "%")
    ElseIf optFind(3).Value = True Then '按手术查找
        strSql = "Select d.Id, d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, g.Id As 科室id, d.标识号,D.就诊ID, d.当前状态" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号, a.id As 就诊ID," & vbNewLine & _
                "                     Decode(a.执行状态, 1, '在' || To_Char(a.执行时间, 'yyyy-mm-dd') || '门诊就诊离院', '门诊正在就诊') As 当前状态" & vbNewLine & _
                "              From 病人挂号记录 A, 病人手麻记录 M, 疾病编码目录 N" & vbNewLine & _
                "              Where m.手术操作id = n.Id And a.病人id = m.病人id And a.Id = m.主页id And a.记录状态 = 1 And" & vbNewLine & _
                "                    a.执行时间 Between [1] And [2] And (Upper(n.编码) Like [3] Or n.名称 Like [3] Or Upper(n.简码) Like [3])" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号, B.主页ID As 就诊ID," & vbNewLine & _
                "                     Decode(b.出院日期, Null, '在院', '第' || b.主页id || '次住院离院') As 当前状态" & vbNewLine & _
                "              From 病案主页 B, 病人手麻记录 O, 疾病编码目录 V" & vbNewLine & _
                "              Where o.手术操作id = v.Id And b.病人id = o.病人id And b.主页id = o.主页id And b.入院日期 Between [1] And [2] And" & vbNewLine & _
                "                    (Upper(v.编码) Like [3] Or v.名称 Like [3] Or Upper(v.简码) Like [3])) C) D, 部门表 G" & vbNewLine & _
                "Where g.Id = d.科室 And d.Top = 1" & vbNewLine & _
                "Order By d.操作时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mdtBegin, mdtEnd, UCase(txtFind.Text) & "%")
    End If

    rptPati.Records.DeleteAll

    With rptPati
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                Set objRecord = .Records.Add()
                Set objItem = objRecord.AddItem("")
                    objItem.Icon = img16.ListImages("unCheck").Index - 1
                Set objItem = objRecord.AddItem(rsTmp!ID & "")
                Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                    objItem.Icon = img16.ListImages.Item(IIf(rsTmp!性别 & "" = "女", "girl", "boy")).Index - 1
                Set objItem = objRecord.AddItem(rsTmp!性别 & "")
                Set objItem = objRecord.AddItem(rsTmp!年龄 & "")
                Set objItem = objRecord.AddItem(rsTmp!标识号 & "")
                Set objItem = objRecord.AddItem(rsTmp!科室 & "")
                Set objItem = objRecord.AddItem(rsTmp!当前状态 & "")
                Set objItem = objRecord.AddItem(rsTmp!科室ID & "")
                Set objItem = objRecord.AddItem(rsTmp!就诊id & "")
                rsTmp.MoveNext
            Loop
            stbThis.Panels(2).Text = "在当前过滤查找到 " & rsTmp.RecordCount & " 位" & lblDept.Tag & "病人"
        End If
        .Populate
    End With
    Exit Sub
errH:
    MsgBox "在当前科室未查找到病人!", vbInformation, gstrSysName
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetFrom()
'功能：调用电子病案查阅功能，嵌入式获取窗体对象
    Set mobjArchiveView = New frmArchiveView
    mobjArchiveView.BorderStyle = FormBorderStyleConstants.vbBSNone '设置为无边框
    mobjArchiveView.Caption = mobjArchiveView.Caption       '重点是这一句
    SetParent mobjArchiveView.hwnd, picMecInfo.hwnd
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcMec
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "电子病历查阅", mobjArchiveView.hwnd, 0).Tag = "电子病历查阅"
        .InsertItem(1, "暂无授权", PicNo.hwnd, 0).Tag = "暂无授权"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub


Private Sub PicNo_Resize()
    On Error Resume Next
    picNoUse.Top = PicNo.Height / 2 - picNoUse.Height / 2
    picNoUse.Left = PicNo.Width / 2 - picNoUse.Width / 2
End Sub



Private Function CheckUse(ByVal lng病人ID As Long, ByVal lng科室ID As Long, ByRef intTime As Integer) As String
    Dim strSql As String
    Dim strTmp As String
    Dim blnALLTime As Boolean
    Dim blnTmp As Boolean
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    
    '自定义权限检查
    strSql = "Select Zl_Fun_Checkpatimec([1],[2],[3]) as 结果 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Zl_Fun_Checkpatimec", lng病人ID, lng科室ID, UserInfo.ID)
    
    If Val(rsTmp!结果 & "") = 1 Then
        CheckUse = ""
        Exit Function
    End If
    
    If mrsTmp Is Nothing Then
        strSql = "Select a.Id As 授权id, a.授权类型, a.访问病人, a.方案名, a.访问病人, a.内容时限, f_List2str(Cast(Collect(c.授权内容 || '') As t_Strlist)) As 授权范围" & vbNewLine & _
                " From 电子病历访问授权 A, 电子病历授权访问人员 B, 电子病历授权访问病人 C" & vbNewLine & _
                " Where a.Id = b.授权id And a.Id = c.授权id(+) And b.人员id = [1] And a.访问开始时间 <= Sysdate And a.访问结束时间 >= Sysdate And" & vbNewLine & _
                " a.作废时间 Is Null" & vbNewLine & _
                " Group By a.Id, a.授权类型, a.访问病人, a.方案名, a.访问病人, a.内容时限"
        Set mrsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    End If
    If Not mrsTmp Is Nothing Then
        If mrsTmp.RecordCount > 0 Then mrsTmp.MoveFirst
        Do While Not mrsTmp.EOF
            blnTmp = False
            Select Case Val(mrsTmp!访问病人) '0-全院病人，1-本科病人，2-指定科室病人，3-指定病人，4-诊断为指定疾病的病人，5-指定手术的病人。2-4的多项内容通过子表存储';
                Case 0 '全院病人
                    strTmp = strTmp & ";" & Val(mrsTmp!授权id & "") & "," & Val(mrsTmp!内容时限 & "")
                    blnTmp = True
                Case 1 '本科病人
                    If InStr(mstrDeptIds, ";" & lng科室ID & ",") > 0 Then
                        strTmp = strTmp & ";" & Val(mrsTmp!授权id & "") & "," & Val(mrsTmp!内容时限 & "")
                        blnTmp = True
                    End If
                Case 2 '指定科室病人
                    If InStr("," & mrsTmp!授权范围 & ",", "," & lng科室ID & ",") > 0 Then
                        strTmp = strTmp & ";" & Val(mrsTmp!授权id & "") & "," & Val(mrsTmp!内容时限 & "")
                        blnTmp = True
                    End If
                Case 3 '指定病人
                    If InStr("," & mrsTmp!授权范围 & ",", "," & lng病人ID & ",") > 0 Then
                        strTmp = strTmp & ";" & Val(mrsTmp!授权id & "") & "," & Val(mrsTmp!内容时限 & "")
                        blnTmp = True
                    End If
            End Select
            
            '计算综合时限
            If blnTmp Then
                If Val(mrsTmp!内容时限 & "") = 0 Then
                    blnALLTime = True
                Else
                    If intTime <> 0 And intTime <> Val(mrsTmp!内容时限 & "") Then
                        blnALLTime = True
                    End If
                End If
                intTime = Val(mrsTmp!内容时限 & "")
            End If
            
            mrsTmp.MoveNext
        Loop
        If blnALLTime Then intTime = 0
        CheckUse = Mid(strTmp, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub rptPati_SelectionChanged()
    Dim strIDs As String
    Dim lngApplyID As Long
    Dim intTime As Integer
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '非正常情况

    If Not mobjArchiveView Is Nothing And Val(rptPati.Tag) <> Val(rptPati.SelectedRows(0).Record(col_病人Id).Value) Then
    
        strIDs = CheckUse(Val(rptPati.SelectedRows(0).Record(col_病人Id).Value), Val(rptPati.SelectedRows(0).Record(col_科室ID).Value), intTime)
        If strIDs <> "" Then
            rptPati.Tag = Val(rptPati.SelectedRows(0).Record(col_病人Id).Value)
            Me.tbcMec.Item(0).Selected = True
        Else
            rptPati.Tag = Val(rptPati.SelectedRows(0).Record(col_病人Id).Value)
            Me.tbcMec.Item(1).Selected = True
            Exit Sub
        End If
    
    
        If Val(rptPati.SelectedRows(0).Record(col_病人Id).Value) <> 0 Then
            rptPati.Tag = Val(rptPati.SelectedRows(0).Record(col_病人Id).Value)
            Call mobjArchiveView.zlRefresh(Val(rptPati.SelectedRows(0).Record(col_病人Id).Value), Val(rptPati.SelectedRows(0).Record(col_就诊ID).Value), strIDs, intTime)
            If strIDs = "" Then
                Me.tbcMec.Item(1).Selected = True
            End If
        End If
    End If
    rptPati.SetFocus
End Sub



Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptPati.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptPatiCheck(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(col_选择))
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim objHitTest As ReportHitTestInfo
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(col_选择).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(col_选择).Icon = img16.ListImages("AllCheck").Index - 1
                            rptPati.Records(i).Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(col_选择).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(col_选择).Icon = img16.ListImages("unCheck").Index - 1
                            rptPati.Records(i).Tag = "0"
                        Next
                    End If
                End If
            End If
        ElseIf rptPati.HitTest(X, Y).ht = xtpHitTestReportArea Then
            Set objHitTest = rptPati.HitTest(X, Y)
            If Not objHitTest.Column Is Nothing And Not objHitTest.Row Is Nothing Then
                If objHitTest.Column.Index = col_选择 Then
                    If rptPati.SelectedRows.Count > 0 Then
                        Call rptPatiCheck(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(col_选择))
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPatiCheck(Row As XtremeReportControl.IReportRow, Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(col_选择).Icon = img16.ListImages("unCheck").Index - 1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(col_选择).Icon = img16.ListImages.Item("AllCheck").Index - 1
        Row.Record.Tag = "1"
    End If
    rptPati.Populate
End Sub


    
Private Function GetPatiRs() As ADODB.Recordset
    '获取勾选病人的记录集
    Dim rsCurr As New ADODB.Recordset
    Dim i As Long
    '先获取当前已经设置好值
    rsCurr.Fields.Append "ID", adInteger, , adFldIsNullable
    rsCurr.Fields.Append "姓名", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "性别", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "年龄", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "科室", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "标识号", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "当前状态", adVarChar, 4000, adFldIsNullable

    rsCurr.CursorLocation = adUseClient
    rsCurr.LockType = adLockOptimistic
    rsCurr.CursorType = adOpenStatic
    rsCurr.Open
    
    For i = 0 To rptPati.Records.Count - 1
        If rptPati.Records(i).Tag = "1" And Val(rptPati.Records(i)(col_病人Id).Value) <> 0 Then
            rsCurr.AddNew
            rsCurr!ID = Val(rptPati.Records(i)(col_病人Id).Value)
            rsCurr!姓名 = rptPati.Records(i)(col_姓名).Value
            rsCurr!性别 = rptPati.Records(i)(col_性别).Value
            rsCurr!年龄 = rptPati.Records(i)(col_年龄).Value
            rsCurr!科室 = rptPati.Records(i)(col_科室).Value
            rsCurr!标识号 = rptPati.Records(i)(COL_标识号).Value
            rsCurr!当前状态 = rptPati.Records(i)(COL_当前状态).Value
            rsCurr.Update
        End If
    Next
    If (Not rsCurr Is Nothing) And (Not rsCurr.EOF) Then
        rsCurr.MoveFirst
    Else
        Set rsCurr = Nothing
    End If
    Set GetPatiRs = rsCurr
End Function



Private Sub lblDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vRect As RECT, strSql As String
    Dim str单位 As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, Cmd_所有科室, "所有科室")
        objControl.Parameter = ""
        Set objControl = .Add(xtpControlButton, Cmd_住院科室, "住院科室")
        objControl.Parameter = "住院"
        Set objControl = .Add(xtpControlButton, Cmd_门诊科室, "门诊科室")
        objControl.Parameter = "门诊"
    End With
    GetWindowRect fraPatiFilter.hwnd, vRect
    objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + lblDept.Left + lblDept.Width, vRect.Top * Screen.TwipsPerPixelY + lblDept.Top
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub SetLBLFace(ByRef objCtl As Object, ByVal blnOver As Boolean)
    If blnOver Then
        If objCtl.BorderStyle = 0 Then
            objCtl.BorderStyle = 1
            objCtl.BackStyle = 1
        End If
    Else
        If objCtl.BorderStyle = 1 Then
            objCtl.BorderStyle = 0
            objCtl.BackStyle = 0
        End If
    End If
End Sub

