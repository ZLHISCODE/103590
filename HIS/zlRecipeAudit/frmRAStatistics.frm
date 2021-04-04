VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmRAStatistics 
   Caption         =   "处方审查统计"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmRAStatistics.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1575
      Begin XtremeSuiteControls.TabControl tbcTab 
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   480
      ScaleHeight     =   3135
      ScaleWidth      =   6975
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3720
      Width           =   6975
      Begin VB.PictureBox picMX_LR 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   4800
         ScaleHeight     =   1695
         ScaleWidth      =   75
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   75
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Index           =   1
         Left            =   0
         ScaleHeight     =   2895
         ScaleWidth      =   4695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   4695
         Begin VB.PictureBox picWhere 
            BorderStyle     =   0  'None
            Height          =   2055
            Index           =   1
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   4455
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   480
            Width           =   4455
            Begin VB.OptionButton optClass2 
               Caption         =   "住院"
               Height          =   180
               Index           =   1
               Left            =   2280
               TabIndex        =   33
               Top             =   120
               Width           =   735
            End
            Begin VB.OptionButton optClass2 
               Caption         =   "门诊"
               Height          =   180
               Index           =   0
               Left            =   1440
               TabIndex        =   32
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.ComboBox cboDoctor 
               Height          =   300
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ComboBox cboClinic 
               Height          =   300
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   1320
               Width           =   2895
            End
            Begin MSComCtl2.DTPicker dtpDate 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1440
               TabIndex        =   23
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Format          =   293535745
               CurrentDate     =   42115
               MaxDate         =   401769
               MinDate         =   36526
            End
            Begin MSComCtl2.DTPicker dtpDate 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1440
               TabIndex        =   24
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Format          =   293535745
               CurrentDate     =   42115
               MaxDate         =   401769
               MinDate         =   36526
            End
            Begin VB.Label lblDoctor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医生(&C)"
               Height          =   180
               Left            =   240
               TabIndex        =   25
               Top             =   1710
               Width           =   990
            End
            Begin VB.Label lblClass 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "审查分类(&C)"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   21
               Top             =   120
               Width           =   990
            End
            Begin VB.Label lblDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "审查日期(&D)"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   20
               Top             =   510
               Width           =   990
            End
            Begin VB.Label lblClinic 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床科室(&R)"
               Height          =   180
               Left            =   240
               TabIndex        =   19
               Top             =   1350
               Width           =   990
            End
         End
         Begin VB.Label lblFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过滤条件"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   855
         Left            =   5280
         TabIndex        =   29
         Top             =   480
         Width           =   1575
         _cx             =   2778
         _cy             =   1508
         Appearance      =   2
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   Begin VB.PictureBox picStatistics 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   2280
      ScaleHeight     =   3375
      ScaleWidth      =   6855
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6855
      Begin VB.PictureBox picTJ_TB_S 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   4080
         MousePointer    =   7  'Size N S
         ScaleHeight     =   75
         ScaleWidth      =   1395
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1395
      End
      Begin VB.PictureBox picTJ_LR 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   3840
         ScaleHeight     =   1695
         ScaleWidth      =   75
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   75
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   0
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   3615
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   3615
         Begin VB.PictureBox picWhere 
            BorderStyle     =   0  'None
            Height          =   1815
            Index           =   0
            Left            =   120
            ScaleHeight     =   1815
            ScaleWidth      =   3375
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   480
            Width           =   3375
            Begin VB.PictureBox picElement 
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   1440
               ScaleHeight     =   495
               ScaleWidth      =   1815
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1200
               Width           =   1815
               Begin VB.OptionButton optElement 
                  Caption         =   "医生"
                  Height          =   180
                  Index           =   1
                  Left            =   840
                  TabIndex        =   36
                  Top             =   120
                  Width           =   735
               End
               Begin VB.OptionButton optElement 
                  Caption         =   "科室"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   35
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin MSComCtl2.DTPicker dtpDate 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1440
               TabIndex        =   9
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Format          =   293535745
               CurrentDate     =   42115
               MaxDate         =   401769
               MinDate         =   36526
            End
            Begin VB.OptionButton optClass1 
               Caption         =   "住院"
               Height          =   180
               Index           =   1
               Left            =   2280
               TabIndex        =   8
               Top             =   120
               Width           =   735
            End
            Begin VB.OptionButton optClass1 
               Caption         =   "门诊"
               Height          =   180
               Index           =   0
               Left            =   1440
               TabIndex        =   7
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin MSComCtl2.DTPicker dtpDate 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1440
               TabIndex        =   10
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Format          =   294191105
               CurrentDate     =   42115
               MaxDate         =   401769
               MinDate         =   36526
            End
            Begin VB.Label lblElement 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "统计(&T)"
               Height          =   180
               Left            =   240
               TabIndex        =   6
               Top             =   1320
               Width           =   990
            End
            Begin VB.Label lblDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "审查日期(&D)"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   5
               Top             =   510
               Width           =   990
            End
            Begin VB.Label lblClass 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "审查分类(&C)"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   4
               Top             =   120
               Width           =   990
            End
         End
         Begin VB.Label lblFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过滤条件"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStatistics 
         Height          =   495
         Left            =   4080
         TabIndex        =   11
         Top             =   480
         Width           =   2055
         _cx             =   3625
         _cy             =   873
         Appearance      =   2
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfStatisticsDetail 
         Height          =   495
         Left            =   4080
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
         _cx             =   3625
         _cy             =   873
         Appearance      =   2
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   8070
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRAStatistics.frx":57E2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmRAStatistics.frx":6074
      Left            =   1560
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   1080
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmRAStatistics.frx":6088
   End
End
Attribute VB_Name = "frmRAStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_VSF_STATI As String = _
        "审查项目,,3,2000|科室,,3,2000|不合格药嘱数量,,3,1500,n|药嘱数量,,3,1500,n|不合格百分比,,3,1500,n" & _
        "|审查项目id,,0,0|提交科室id,,0,0"
Private Const MSTR_VSF_STATI_DETAIL As String = _
        "审查时间,,3,1600,dt|病人,,3,1000|医嘱ID,,0,0|相关ID,,0,0|诊断,,3,2000|药品名称,,3,2000|规格,,3,1500|单位,,3,1000|数量,,3,1000,n|单量,,3,1000" & _
        "|用法,,3,1000|频次,,3,1000"
Private Const MSTR_VSF_DETAIL As String = _
        "审查时间,,3,1600,dt|临床科室,,3,2000|医生,,3,1000|病人,,3,1000|医嘱ID,,0,0|相关ID,,0,0|诊断,,3,2000|审查项目,,3,1500|审查人,,3,1000" & _
        "|药品名称,,3,2000|规格,,3,1500|单位,,3,1000|数量,,3,1000,n|单量,,3,1000|用法,,3,1000|频次,,3,1000"

Private mobjPubAdvice As zlPublicAdvice.clsPublicAdvice
Private mlngModule As Long
Private mstrPrivs As String
Private mblnMemory As Boolean
Private msngY As Single

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim dtpBegin As DTPicker, dtpEnd As DTPicker

    Select Case Control.ID
        Case enuMenus.打印设置
            Call zlPrintSet
        Case enuMenus.打印预览, enuMenus.打印, enuMenus.输出Excel
            Dim objTmp As Object
            Dim strTitle As String
            
            Set objTmp = Me.ActiveControl
            If TypeName(objTmp) = "VSFlexGrid" Then
                objTmp.Redraw = False
                If UCase(objTmp.Name) = "VSFSTATISTICS" Then
                    strTitle = zlStr.FormatString("不合格审查项目统计（[1]）", IIf(optClass1(0).Value, "门诊", "住院"))
                ElseIf UCase(objTmp.Name) = "VSFSTATISTICSDETAIL" Then
                    strTitle = zlStr.FormatString("不合格审查项目明细（[1]）", IIf(optClass1(0).Value, "门诊", "住院"))
                ElseIf UCase(objTmp.Name) = "VSFDETAIL" Then
                    strTitle = zlStr.FormatString("审查不合格药嘱明细（[1]）", IIf(optClass2(0).Value, "门诊", "住院"))
                End If
                If strTitle <> "" Then
                    If Control.ID = enuMenus.打印预览 Then
                        RptExport 0, objTmp, strTitle
                    ElseIf Control.ID = enuMenus.打印 Then
                        RptExport 1, objTmp, strTitle
                    Else
                        RptExport 3, objTmp, strTitle
                    End If
                End If
                objTmp.Redraw = True
            End If
        Case enuMenus.退出
            Unload Me
        Case enuMenus.刷新
            '检查日期范围
            If UCase(TypeName(Me.ActiveControl)) = "DTPICKER" Then
                Me.picTab.SetFocus
            End If
            If tbcTab.Item(0).Selected Then
                Set dtpBegin = dtpDate(0)
                Set dtpEnd = dtpDate(1)
            Else
                Set dtpBegin = dtpDate(2)
                Set dtpEnd = dtpDate(3)
            End If
            If dtpBegin.Value > dtpEnd.Value Then
                MsgBox "“开始日期”大于“结束日期”！", vbInformation, gstrSysName
                dtpBegin.SetFocus
                Exit Sub
            End If
            If dtpEnd.Value - dtpBegin.Value > 31 Then
                MsgBox "日期范围不能超过31天！", vbInformation, gstrSysName
                dtpBegin.SetFocus
                Exit Sub
            End If
        
            Call FS.ShowFlash
            If tbcTab.Item(0).Selected Then
                Call FillVSFData(1)
                Call FillVSFData(2)
            Else
                Call FillVSFData(3)
            End If
            Call SetStatusbar
            Call FS.StopFlash
        Case enuMenus.标准按钮
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case enuMenus.文本标签
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case enuMenus.大图标
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            cbsMain.RecalcLayout
        Case enuMenus.状态栏
            stbThis.Visible = Not Control.Checked
            cbsMain.RecalcLayout
        Case enuMenus.帮助主题
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case enuMenus.中联主页
            Call zlHomePage(Me.hwnd)
        Case enuMenus.中联论坛
            Call zlWebForum(Me.hwnd)
        Case enuMenus.发送反馈
            Call zlMailTo(Me.hwnd)
        Case enuMenus.关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            '报表
            If Between(Control.ID, enuMenus.报表 * 100# + 1, enuMenus.报表 * 100# + 99) And Control.Parameter <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = stbThis.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case enuMenus.标准按钮
            Control.Checked = Me.cbsMain(2).Visible
        Case enuMenus.文本标签
            Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
        Case enuMenus.大图标
            Control.Checked = cbsMain.Options.LargeIcons
        Case enuMenus.状态栏
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picTab.hwnd
    End If
End Sub

Private Sub Form_Load()
    Dim datCurrent As Date

    mlngModule = glngModule
    mstrPrivs = zlStr.FormatString(";[1];", GetPrivFunc(glngSys, mlngModule))
    mblnMemory = Val(zlDatabase.GetPara("使用个性化风格")) = 1
    
    '临床公共方法
    On Error Resume Next
    Set mobjPubAdvice = New zlPublicAdvice.clsPublicAdvice
    If Not mobjPubAdvice Is Nothing Then
        Call mobjPubAdvice.InitCommon(gcnOracle, glngSys)
    End If
    Err.Clear: On Error GoTo 0
    
    '初始化控件
    Call InitCommandbars
    Call InitDockPane
    Call InitTBCTab
    
    Call InitVSF(vsfStatistics)
    Call InitVSF(vsfStatisticsDetail)
    Call InitVSF(vsfDetail)
    
    Call mdlDefine.SetVSFHead(vsfStatistics, MSTR_VSF_STATI)
    Call mdlDefine.SetVSFHead(vsfStatisticsDetail, MSTR_VSF_STATI_DETAIL)
    Call mdlDefine.SetVSFHead(vsfDetail, MSTR_VSF_DETAIL)
    
    '恢复上次界面
    RestoreWinState Me, App.ProductName
    If mblnMemory Then
        picTJ_TB_S.Visible = True
    End If
    
    '加载数据
    datCurrent = SYS.Currentdate
    dtpDate(0).Value = Format(datCurrent, "yyyy-mm-01")
    dtpDate(1).Value = Format(DateAdd("m", 1, datCurrent) - Day(DateAdd("m", 1, datCurrent)), "yyyy-mm-dd")
    dtpDate(2).Value = dtpDate(0).Value
    dtpDate(3).Value = dtpDate(1).Value
    
End Sub

Private Sub InitCommandbars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbrTmp As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003 'xtpthemeoffice2000有凹凸感
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        .EnableCustomization False
        Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
    '文件
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.文件, "文件(&F)", -1, False)
    With cbpTmp
        .ID = enuMenus.文件
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印设置, "打印设置(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印预览, "打印预览(&V)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印, "打印")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.输出Excel, "输出到&Excel...")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.退出, "退出")
        cbcTmp.BeginGroup = True
    End With
    
    '查看
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.查看, "查看(&V)", -1, False)
    With cbpTmp
        .ID = enuMenus.查看
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.工具栏, "工具栏(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.标准按钮, "标准按钮(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.文本标签, "文本标签(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.大图标, "大图标(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.状态栏, "状态栏(&S)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.刷新, "刷新")
        cbcTmp.BeginGroup = True
    End With
    
    '帮助
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.帮助, "帮助(&H)", -1, False)
    With cbpTmp
        .ID = enuMenus.帮助
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助主题")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB上的中联, "&WEB上的中联")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联主页, "中联主页(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联论坛, "中联论坛(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.发送反馈, "发送反馈(&K)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.关于, "关于(&A)")
        cbcTmp.BeginGroup = True
    End With
    
    '报表接口
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    
    '菜单项的快键绑定
    With cbsMain.KeyBindings
        .Add 8, vbKeyP, enuMenus.打印
        .Add 8, vbKeyX, enuMenus.退出
        .Add 0, vbKeyF5, enuMenus.刷新
        .Add 0, vbKeyF1, enuMenus.帮助主题
    End With
    
    '定义工具栏
    Set cbrTmp = cbsMain.Add("工具栏", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.打印预览, "打印预览")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.打印, "打印")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.刷新, "刷新")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助主题")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.退出, "退出")
    End With
    
    '有图标，无文本的按钮风格
    For Each cbcTmp In cbsMain(2).Controls
        If cbcTmp.Type <> xtpControlLabel Then
            cbcTmp.Style = xtpButtonIcon
        End If
    Next
    
End Sub

Private Sub InitTBCTab()
    With tbcTab.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ClientFrame = xtpTabFrameSingleLine
        .BoldSelected = True
        .OneNoteColors = True
        .ShowIcons = False
    End With
    
    With tbcTab
        .InsertItem 0, "不合格审查项目统计(&1)", picStatistics.hwnd, 0
        .InsertItem 1, "审查不合格药嘱明细(&2)", picDetail.hwnd, 0
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    If Me.Width <= 8000 Then Me.Width = 8000
    If Me.Height <= 6000 Then Me.Height = 6000
End Sub

Private Sub InitDockPane()
    Dim panClient As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeGrippered
        
        Set panClient = .CreatePane(1, 250, 0, DockLeftOf)
        With panClient
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            .Title = ""
        End With
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjPubAdvice Is Nothing Then
        Set mobjPubAdvice = Nothing
    End If
    SaveWinState Me, App.ProductName
End Sub

Private Sub optClass2_Click(Index As Integer)
    Call SetComboxItem(cboClinic)
    Call SetComboxItem(cboDoctor)
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    
    With picFilter(1)
        .Top = 0
        .Left = 0
        .Width = 4600
        .Height = picDetail.ScaleHeight
    End With
    
    With picMX_LR
        .Top = 0
        .Left = picFilter(1).Width
        .Height = picFilter(1).Height
    End With
    
    With vsfDetail
        .Top = 0
        .Left = picMX_LR.Left + picMX_LR.Width
        .Width = picDetail.ScaleWidth - picMX_LR.Left - picMX_LR.Width
        .Height = picDetail.ScaleHeight
    End With
End Sub

Private Sub picFilter_Resize(Index As Integer)
    On Error Resume Next
    
    If Index = 0 Then
        With lblFilter(0)
            .Top = 120
            .Left = 60
        End With
        
        With picWhere(0)
            .Top = lblFilter(0).Top + lblFilter(0).Height + 120
            .Left = 0
            .Width = picFilter(0).ScaleWidth
            .Height = picFilter(0).ScaleHeight - lblFilter(0).Height + 120 * 2
        End With
    Else
        With lblFilter(1)
            .Top = 120
            .Left = 60
        End With
        
        With picWhere(1)
            .Top = lblFilter(1).Top + lblFilter(1).Height + 120
            .Left = 0
            .Width = picFilter(1).ScaleWidth
            .Height = picFilter(1).ScaleHeight - lblFilter(1).Height + 120 * 2
        End With
    End If
End Sub

Private Sub picStatistics_Resize()
    On Error Resume Next
    
    With picFilter(0)
        .Top = 0
        .Left = 0
        .Width = 3500
        .Height = picStatistics.ScaleHeight
    End With
    
    With picTJ_LR
        .Top = 0
        .Left = picFilter(0).Width
        .Height = picFilter(0).Height
    End With
    
    With vsfStatistics
        .Top = 0
        .Left = picTJ_LR.Left + picTJ_LR.Width
        .Width = picStatistics.ScaleWidth - picTJ_LR.Left - picTJ_LR.Width
        .Height = picTJ_TB_S.Top
    End With
    
    With picTJ_TB_S
        .Left = vsfStatistics.Left
        .Width = vsfStatistics.Width
    End With
    
    With vsfStatisticsDetail
        .Top = picTJ_TB_S.Top + picTJ_TB_S.Height
        .Left = vsfStatistics.Left
        .Width = vsfStatistics.Width
        .Height = picStatistics.ScaleHeight - picTJ_TB_S.Top - picTJ_TB_S.Height
    End With
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    
    With tbcTab
        .Top = 0
        .Left = 0
        .Width = picTab.ScaleWidth
        .Height = picTab.ScaleHeight
    End With
End Sub

Private Sub InitVSF(ByRef vsfVar As VSFlexGrid)
'功能：初始化窗体的VSFlexGrid控件的风格
'参数：
'  vsfVar：要初始化的VSFlexGrid控件

    With vsfVar
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .BackColorBkg = .BackColor
        .AutoResize = True
    End With
End Sub

Private Sub picTJ_TB_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    msngY = Y
End Sub

Private Sub picTJ_TB_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    With picTJ_TB_S
        If .Top + Y < ScaleHeight * 0.3 Then
            .Top = ScaleHeight * 0.3
            Exit Sub
        End If
        If .Top + Y > ScaleHeight * 0.7 Then
            .Top = ScaleHeight * 0.7
            Exit Sub
        End If
        .Move .Left, .Top + Y - msngY
    End With
End Sub

Private Sub picTJ_TB_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picStatistics_Resize
    msngY = 0
End Sub

Private Sub tbcTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Index = 1 Then
        If Me.cboClinic.ListCount <= 0 Then
            Call SetComboxItem(cboClinic)
        End If
        If Me.cboDoctor.ListCount <= 0 Then
            Call SetComboxItem(cboDoctor)
        End If
    End If
End Sub

Private Sub SetComboxItem(ByRef cboVar As ComboBox)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIdx As Long
    
    With cboVar
        lngIdx = IIf(.ListIndex < 0, 0, .ListIndex)
        .Clear
        
        On Error GoTo errHandle
        
        If .Name = "cboClinic" Then
            strSQL = "Select a.Id, '【' || a.编码 || '】' || a.名称 名称 " & vbNewLine & _
                     "From 部门表 A, 部门性质说明 B " & vbNewLine & _
                     "Where a.Id = b.部门id And b.服务对象 In ([1], 3) And b.工作性质 In ('临床') " & vbNewLine & _
                     "    And (a.撤档时间 Is Null Or To_Char(a.撤档时间, 'yyyy') = '3000')" & vbNewLine & _
                     "Order By a.编码 "
        Else
            strSQL = "Select Distinct a.Id, '【' || a.编号 || '】' || a.姓名 名称, a.编号 " & vbNewLine & _
                     "From 人员表 A, 人员性质说明 B, 部门人员 C, 部门性质说明 D " & vbNewLine & _
                     "Where a.Id = b.人员id And a.Id = c.人员id And c.部门id = d.部门id And b.人员性质 = '医生' " & vbNewLine & _
                     "    And d.工作性质 In ('临床') And d.服务对象 In ([1], 3) " & vbNewLine & _
                     "Order By a.编号 "
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取部门信息", IIf(Me.optClass2(0).Value, 1, 2))
        
        .AddItem "全部"
        Do While rsTemp.EOF = False
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!ID
            
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        
        If lngIdx >= .ListIndex Then
            .ListIndex = lngIdx
        Else
            .ListIndex = .ListCount - 1
        End If
    End With
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FillVSFData(ByVal bytFun As Byte)
'功能：填充数据
'参数：
'  bytFun：功能号；1-统计查询；2-统计明细查询；3-明细查询

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngItemID As Long, lngClinicID As Long, l As Long, lngDeptID As Long
    Dim strDoctor As String, strDiagnose As String
    
    Select Case bytFun
        Case 1
            If Me.optElement(0).Value Then
                '科室
                strSQL = "Select d.审查项目id, b.提交科室id, Sum(Decode(d.药师审查, 2, 1, 0)) 不合格药嘱数量, Count(c.医嘱id) 药嘱数量 " & vbNewLine & _
                         "From 病人医嘱记录 A, 处方审查记录 B, 处方审查明细 C, 处方审查结果 D " & vbNewLine & _
                         "Where a.Id = c.医嘱id And c.审方id = b.Id And c.审方id = d.审方id(+) And c.医嘱id = d.医嘱id(+) And a.病人来源 = [1] " & vbNewLine & _
                         "    And a.诊疗类别 In ('5', '6', '7') And b.审查时间 Between [2] And [3] And d.医嘱id Is Not Null " & vbNewLine & _
                         "Group By d.审查项目id, b.提交科室id "
                         
                strSQL = "Select a.*, " & _
                         "    Decode(a.药嘱数量, 0, 0, Round(a.不合格药嘱数量 / a.药嘱数量 * 100, 2)) 不合格百分比, " & _
                         "    '【' || b.编码 || '】' || b.名称 科室," & _
                         "    '【' || c.编码 || '】' || c.简称 审查项目 " & vbNewLine & _
                         "From (" & strSQL & ") A, 部门表 B, 处方审查项目 C " & vbNewLine & _
                         "Where a.提交科室id = b.Id(+) And a.审查项目id = c.Id(+) " & vbNewLine & _
                         "Order By c.编码, b.编码 "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询科室的不合格数量", _
                                IIf(optClass1(0).Value, 1, 2), _
                                dtpDate(0).Value, _
                                dtpDate(1).Value + 1 - 1 / 24 / 60 / 60)
                
                vsfStatistics.TextMatrix(0, vsfStatistics.ColIndex("科室")) = "科室"
            Else
                '医生
                strSQL = "Select d.审查项目id, b.提交人, Sum(Decode(d.药师审查, 2, 1, 0)) 不合格药嘱数量, Count(c.医嘱id) 药嘱数量 " & vbNewLine & _
                         "From 病人医嘱记录 A, 处方审查记录 B, 处方审查明细 C, 处方审查结果 D " & vbNewLine & _
                         "Where a.Id = c.医嘱id And c.审方id = b.Id And c.审方id = d.审方id(+) And c.医嘱id = d.医嘱id(+) And a.病人来源 = [1] " & vbNewLine & _
                         "    And a.诊疗类别 In ('5', '6', '7') And b.审查时间 Between [2] And [3] And d.医嘱id Is Not Null " & vbNewLine & _
                         "Group By d.审查项目id, b.提交人 "
                         
                strSQL = "Select a.审查项目id, a.不合格药嘱数量, a.药嘱数量, a.提交人 科室, " & _
                         "    Decode(a.药嘱数量, 0, 0, Round(a.不合格药嘱数量 / a.药嘱数量 * 100, 2)) 不合格百分比," & _
                         "    '【' || b.编码 || '】' || b.简称 审查项目" & vbNewLine & _
                         "From (" & strSQL & ") A, 处方审查项目 B " & vbNewLine & _
                         "Where a.审查项目id = b.Id(+) " & vbNewLine & _
                         "Order By b.编码, a.提交人 "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询医生的不合格数量", _
                                IIf(optClass1(0).Value, 1, 2), _
                                dtpDate(0).Value, _
                                dtpDate(1).Value + 1 - 1 / 24 / 60 / 60)
                
                vsfStatistics.TextMatrix(0, vsfStatistics.ColIndex("科室")) = "医生"
            End If
            Call mdlDefine.FillVSFData(vsfStatistics, rsTemp)
            
            If vsfStatistics.Rows > 1 Then
                vsfStatistics.Row = 1
            End If
            
        Case 2
            
            If vsfStatistics.Rows > 1 Then
                lngItemID = Val(vsfStatistics.TextMatrix(vsfStatistics.Row, vsfStatistics.ColIndex("审查项目id")))
                If optElement(0).Value Then
                    '科室
                    lngDeptID = Val(vsfStatistics.TextMatrix(vsfStatistics.Row, vsfStatistics.ColIndex("提交科室id")))
                Else
                    '医生
                    strDoctor = Trim(vsfStatistics.TextMatrix(vsfStatistics.Row, vsfStatistics.ColIndex("科室")))
                End If
            End If
            
            If Me.optClass1(0).Value = True Then
                strSQL = "Select b.审查时间, A1.姓名 病人, A1.ID 医嘱ID, '????' 诊断, d.名称 药品名称, d.规格, d.计算单位 单位,  " & _
                         "    A1.总给予量 数量, A1.单次用量 || Nvl(e.计算单位, '') 单量, A2.医嘱内容 用法, A1.执行频次 频次, " & vbNewLine & _
                         "    A1.相关ID " & vbNewLine & _
                         "From 病人医嘱记录 A1, 病人医嘱记录 A2, 处方审查记录 B, 处方审查结果 C, 收费项目目录 D, 诊疗项目目录 E " & vbNewLine & _
                         "Where A1.相关id = A2.Id And A1.Id = c.医嘱id And c.审方id = b.Id And A1.收费细目id = d.Id And A1.诊疗项目id = e.ID(+) " & _
                         "    And A1.病人来源 = [1] And c.审查项目Id = [2] And c.药师审查 = 2 " & vbNewLine & _
                         "    And b.审查时间 between [3] and [4] " & vbNewLine & _
                         IIf(optElement(0).Value, " And b.提交科室id = [5] ", " And b.提交人 = [5] ") & vbNewLine & _
                         "Order By b.审查时间 "
            Else
                strSQL = "Select b.审查时间, A1.姓名 病人, A1.ID 医嘱ID, '????' 诊断, e.名称 药品名称, d.规格, d.计算单位 单位,  " & _
                         "    A1.单次用量 || Nvl(e.计算单位, '') 单量, A2.医嘱内容 用法, A1.执行频次 频次, " & vbNewLine & _
                         "    A1.相关ID " & vbNewLine & _
                         "From 病人医嘱记录 A1, 病人医嘱记录 A2, 处方审查记录 B, 处方审查结果 C, 收费项目目录 D, 诊疗项目目录 E " & vbNewLine & _
                         "Where A1.相关id = A2.Id And A1.Id = c.医嘱id And c.审方id = b.Id And A1.收费细目id = d.Id And A1.诊疗项目id = e.Id(+) " & _
                         "    And A1.病人来源 = [1] And c.审查项目Id = [2] And c.药师审查 = 2 " & vbNewLine & _
                         "    And b.审查时间 between [3] and [4] " & vbNewLine & _
                         IIf(optElement(0).Value, " And b.提交科室id = [5] ", " And b.提交人 = [5] ") & vbNewLine & _
                         "Order By b.审查时间 "
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询不合格明细", _
                            IIf(optClass1(0).Value, 1, 2), _
                            lngItemID, _
                            dtpDate(0).Value, _
                            dtpDate(1).Value + 1 - 1 / 24 / 60 / 60, _
                            IIf(optElement(0).Value, lngDeptID, strDoctor))
            Call mdlDefine.FillVSFData(vsfStatisticsDetail, rsTemp)
            
            If vsfStatisticsDetail.Rows > 1 Then
                With vsfStatisticsDetail
                    If Not mobjPubAdvice Is Nothing Then
                        '获取药嘱的诊断信息
                        For l = 1 To .Rows - 1
                            strDiagnose = ""
                            Call mobjPubAdvice.GetAdviceDiag(Val(.TextMatrix(l, .ColIndex("相关ID"))), strDiagnose)
                            .TextMatrix(l, .ColIndex("诊断")) = strDiagnose
                        Next
                    End If
                    .Row = 1
                End With
            End If
            
        Case 3
            
            lngClinicID = cboClinic.ItemData(cboClinic.ListIndex)
            If cboDoctor.Text Like "*】*" Then
                strDoctor = Mid(cboDoctor.Text, InStr(cboDoctor.Text, "】") + 1)
            End If
            
            If Me.optClass2(0).Value = True Then
                strSQL = "Select b.审查时间, e.名称 临床科室, A1.姓名 病人, A1.ID 医嘱ID, A1.相关ID, '????' 诊断, f.简称 审查项目, b.提交人 医生, b.审查人, " & _
                         "    d.名称 药品名称, d.规格, d.计算单位 单位, A1.总给予量 数量, A1.单次用量 || Nvl(g.计算单位, '') 单量, A2.医嘱内容 用法, A1.执行频次 频次 " & vbNewLine & _
                         "From 病人医嘱记录 A1, 病人医嘱记录 A2, 处方审查记录 B, 处方审查结果 C, 收费项目目录 D, 部门表 E, 处方审查项目 F, 诊疗项目目录 G " & vbNewLine & _
                         "Where A1.相关id = A2.Id And A1.Id = c.医嘱id And c.审方id = b.Id And A1.收费细目id = d.Id And A1.诊疗项目id = g.id And b.提交科室id = e.Id " & _
                         "    And c.审查项目id = f.Id And c.药师审查 = 2 And A1.病人来源 = [1] And b.审查时间 Between [2] And [3] " & _
                         IIf(lngClinicID > 0, " And b.提交科室id = [4] ", "") & _
                         IIf(strDoctor <> "", " And b.提交人 = [5] ", "") & vbNewLine & _
                         "Order By b.审查时间 "
            Else
                strSQL = "Select b.审查时间, e.名称 临床科室, A1.姓名 病人, A1.ID 医嘱ID, A1.相关ID, '????' 诊断, f.简称 审查项目, b.提交人 医生, b.审查人, " & _
                         "    d.名称 药品名称, d.规格, d.计算单位 单位, A1.单次用量 || Nvl(g.计算单位, '') 单量, A2.医嘱内容 用法, A1.执行频次 频次 " & vbNewLine & _
                         "From 病人医嘱记录 A1, 病人医嘱记录 A2, 处方审查记录 B, 处方审查结果 C, 收费项目目录 D, 部门表 E, 处方审查项目 F, 诊疗项目目录 G " & vbNewLine & _
                         "Where A1.相关id = A2.Id And A1.Id = c.医嘱id And c.审方id = b.Id And A1.收费细目id = d.Id And A1.诊疗项目id = g.id And b.提交科室id = e.Id " & _
                         "    And c.审查项目id = f.Id And c.药师审查 = 2 And A1.病人来源 = [1] And b.审查时间 Between [2] And [3] " & _
                         IIf(lngClinicID > 0, " And b.提交科室id = [4] ", "") & _
                         IIf(strDoctor <> "", " And b.提交人 = [5] ", "") & vbNewLine & _
                         "Order By b.审查时间 "
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询所有不合格明细", _
                            IIf(optClass2(0).Value, 1, 2), _
                            dtpDate(2).Value, _
                            dtpDate(3).Value + 1 - 1 / 24 / 60 / 60, _
                            lngClinicID, _
                            strDoctor)
            Call mdlDefine.FillVSFData(vsfDetail, rsTemp)
            
            If vsfDetail.Rows > 1 Then
                With vsfDetail
                    If Not mobjPubAdvice Is Nothing Then
                        '获取药嘱的诊断信息
                        For l = 1 To .Rows - 1
                            strDiagnose = ""
                            Call mobjPubAdvice.GetAdviceDiag(Val(.TextMatrix(l, .ColIndex("相关ID"))), strDiagnose)
                            .TextMatrix(l, .ColIndex("诊断")) = strDiagnose
                        Next
                    End If
                    .Row = 1
                End With
            End If
    End Select

End Sub

Private Sub vsfStatistics_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And vsfStatistics.Rows > 1 Then
        Call FillVSFData(2)
        Call SetStatusbar
    End If
End Sub

Private Sub RptExport(ByVal bytMode As Byte, ByVal vsfVar As VSFlexGrid, ByVal strTitle As String)
'功能：报表输出
'参数：
'  bytMode：输出方式；1-打印；2-预览；3-输出到EXCEL

    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim lngRow As Long
    Dim lngColor As Long

    If vsfVar.Rows <= 1 Then Exit Sub

    lngColor = vsfVar.GridColor
    vsfVar.GridColor = vbBlack

    lngRow = vsfVar.Row
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTitle
    
    objRow.Add strRange
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    
    If UCase(vsfVar.Name) = "VSFSTATISTICS" Then
        objRow.Add "审查日期：" & Format(dtpDate(0).Value, "yyyy年mm月dd日") & " - " & Format(dtpDate(1).Value, "yyyy年mm月dd日")
    ElseIf UCase(vsfVar.Name) = "VSFSTATISTICSDETAIL" Then
        objRow.Add IIf(optElement(0).Value, "科室", "医生") & "：" & vsfStatistics.TextMatrix(vsfStatistics.Row, vsfStatistics.ColIndex("科室"))
        objRow.Add "审查项目：" & vsfStatistics.TextMatrix(vsfStatistics.Row, vsfStatistics.ColIndex("审查项目"))
    ElseIf UCase(vsfVar.Name) = "VSFDETAIL" Then
        objRow.Add "审查日期：" & Format(dtpDate(2).Value, "yyyy年mm月dd日") & " - " & Format(dtpDate(3).Value, "yyyy年mm月dd日")
    Else
        Exit Sub
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(SYS.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfVar
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    vsfVar.Row = lngRow
    vsfVar.GridColor = lngColor
End Sub

Private Sub SetStatusbar()
    Dim strText As String
    
    If tbcTab.Item(0).Selected Then
        strText = "统计数量：" & vsfStatistics.Rows - 1 & "条； "
        strText = strText & "明细数量：" & vsfStatisticsDetail.Rows - 1 & "条"
    Else
        strText = "明细数量：" & vsfDetail.Rows - 1 & "条"
    End If
    
    stbThis.Panels(2).Text = strText
End Sub
