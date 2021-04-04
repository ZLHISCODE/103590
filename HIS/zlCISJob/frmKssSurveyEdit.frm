VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKssSurveyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人抗菌药物使用情况调查表"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12900
   Icon            =   "frmKssSurveyEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picReasonable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   4620
      ScaleHeight     =   1680
      ScaleWidth      =   12795
      TabIndex        =   57
      Top             =   2370
      Width           =   12795
      Begin VB.PictureBox picPjTim 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   8040
         ScaleHeight     =   360
         ScaleWidth      =   3855
         TabIndex        =   200
         Top             =   930
         Width           =   3855
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   28
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   202
            Top             =   0
            Width           =   1245
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   56
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   201
            Top             =   0
            Width           =   1260
         End
         Begin VB.Line Line1 
            Index           =   36
            X1              =   2520
            X2              =   3735
            Y1              =   225
            Y2              =   225
         End
         Begin VB.Line Line1 
            Index           =   91
            X1              =   840
            X2              =   2115
            Y1              =   210
            Y2              =   210
         End
         Begin VB.Label lblInfo 
            Caption         =   "抽样时间：               到"
            Height          =   195
            Index           =   53
            Left            =   0
            TabIndex        =   203
            Top             =   15
            Width           =   3780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPJB 
         Height          =   2625
         Left            =   2880
         TabIndex        =   148
         Top             =   1395
         Width           =   6750
         _cx             =   11906
         _cy             =   4630
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
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
      End
      Begin VB.Label lblYJ 
         AutoSize        =   -1  'True
         Caption         =   "手术用药合理性评价意见表"
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
         Left            =   4065
         TabIndex        =   147
         Top             =   405
         Width           =   3060
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   570
      Left            =   1785
      TabIndex        =   5
      Top             =   105
      Width           =   4575
      _Version        =   589884
      _ExtentX        =   8070
      _ExtentY        =   1005
      _StockProps     =   64
   End
   Begin VB.PictureBox picUse 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8940
      Left            =   75
      ScaleHeight     =   8940
      ScaleWidth      =   12855
      TabIndex        =   59
      Top             =   690
      Width           =   12855
      Begin VB.VScrollBar vsc 
         Height          =   8470
         LargeChange     =   100
         Left            =   12585
         Max             =   300
         SmallChange     =   50
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picDCB 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13000
         Left            =   0
         ScaleHeight     =   13005
         ScaleWidth      =   12630
         TabIndex        =   60
         Top             =   0
         Width           =   12630
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   400
            Index           =   34
            Left            =   4170
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   205
            Top             =   2670
            Width           =   8160
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   400
            Index           =   27
            Left            =   1590
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   204
            Top             =   1980
            Width           =   10725
         End
         Begin VB.PictureBox picDiff 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1690
            Index           =   1
            Left            =   2535
            ScaleHeight     =   1695
            ScaleWidth      =   12690
            TabIndex        =   61
            Top             =   1320
            Width           =   12690
            Begin VB.PictureBox picOperate 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1575
               Left            =   1455
               ScaleHeight     =   1575
               ScaleWidth      =   10920
               TabIndex        =   198
               Top             =   60
               Width           =   10920
               Begin VSFlex8Ctl.VSFlexGrid vsOperate 
                  Height          =   1530
                  Left            =   0
                  TabIndex        =   199
                  Top             =   0
                  Width           =   10875
                  _cx             =   19182
                  _cy             =   2699
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
                  BackColorSel    =   16764057
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483636
                  GridColorFixed  =   -2147483636
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
                  Rows            =   18
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   250
                  RowHeightMax    =   2000
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmKssSurveyEdit.frx":6852
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
            Begin VB.Label lblInfo 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   41
               Left            =   150
               TabIndex        =   63
               Top             =   735
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   85
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   2570
            End
            Begin VB.Line Line1 
               Index           =   86
               X1              =   330
               X2              =   330
               Y1              =   0
               Y2              =   2000
            End
            Begin VB.Line Line1 
               Index           =   87
               X1              =   1320
               X2              =   1320
               Y1              =   0
               Y2              =   2000
            End
            Begin VB.Line Line1 
               Index           =   88
               X1              =   12405
               X2              =   12405
               Y1              =   0
               Y2              =   2000
            End
            Begin VB.Line Line1 
               Index           =   89
               X1              =   0
               X2              =   12405
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line1 
               Index           =   90
               X1              =   0
               X2              =   12405
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Label lblInfo 
               Caption         =   "手术情况"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   49
               Left            =   420
               TabIndex        =   62
               Top             =   690
               Width           =   840
            End
         End
         Begin VB.PictureBox picDiff 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1030
            Index           =   0
            Left            =   1080
            ScaleHeight     =   1035
            ScaleWidth      =   12840
            TabIndex        =   84
            Top             =   7815
            Width           =   12840
            Begin VB.CommandButton cmdInfect 
               Caption         =   "…"
               Height          =   225
               Index           =   0
               Left            =   11820
               TabIndex        =   208
               Top             =   660
               Width           =   390
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   35
               Left            =   3585
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   206
               Top             =   675
               Width           =   8625
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   47
               Left            =   8805
               TabIndex        =   46
               Top             =   135
               Width           =   3390
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   46
               Left            =   5025
               TabIndex        =   45
               Top             =   150
               Width           =   2955
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "磁共振"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   20
               Left            =   3210
               TabIndex        =   44
               Top             =   135
               Width           =   915
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               Caption         =   "CT"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   19
               Left            =   2520
               TabIndex        =   43
               Top             =   135
               Width           =   660
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               Caption         =   "X线"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   18
               Left            =   1740
               TabIndex        =   42
               Top             =   135
               Width           =   660
            End
            Begin VB.Line Line1 
               Index           =   37
               X1              =   12195
               X2              =   3555
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Line Line1 
               Index           =   66
               X1              =   1320
               X2              =   1320
               Y1              =   -75
               Y2              =   1200
            End
            Begin VB.Line Line1 
               Index           =   67
               X1              =   330
               X2              =   330
               Y1              =   0
               Y2              =   1185
            End
            Begin VB.Line Line1 
               Index           =   71
               X1              =   8805
               X2              =   12195
               Y1              =   345
               Y2              =   345
            End
            Begin VB.Line Line1 
               Index           =   70
               X1              =   5025
               X2              =   7965
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label lblInfo 
               Caption         =   "3.结论："
               Height          =   195
               Index           =   106
               Left            =   8085
               TabIndex        =   92
               Top             =   135
               Width           =   795
            End
            Begin VB.Label lblInfo 
               Caption         =   "2.部位："
               Height          =   195
               Index           =   105
               Left            =   4350
               TabIndex        =   91
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Caption         =   "1."
               Height          =   195
               Index           =   104
               Left            =   1515
               TabIndex        =   90
               Top             =   150
               Width           =   225
            End
            Begin VB.Label lblInfo 
               Caption         =   "影像学诊断"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   103
               Left            =   360
               TabIndex        =   89
               Top             =   150
               Width           =   1005
            End
            Begin VB.Label lblInfo 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   102
               Left            =   135
               TabIndex        =   88
               Top             =   150
               Width           =   225
            End
            Begin VB.Line Line1 
               Index           =   69
               X1              =   15
               X2              =   12420
               Y1              =   495
               Y2              =   495
            End
            Begin VB.Label lblInfo 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   101
               Left            =   120
               TabIndex        =   87
               Top             =   675
               Width           =   225
            End
            Begin VB.Label lblInfo 
               Caption         =   "临床症状"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   96
               Left            =   435
               TabIndex        =   86
               Top             =   675
               Width           =   795
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "与感染有关的主要症状："
               Height          =   180
               Index           =   91
               Left            =   1485
               TabIndex        =   85
               Top             =   675
               Width           =   1980
            End
            Begin VB.Line Line1 
               Index           =   63
               X1              =   0
               X2              =   12400
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line1 
               Index           =   64
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   4725
            End
            Begin VB.Line Line1 
               Index           =   68
               X1              =   12405
               X2              =   12405
               Y1              =   -105
               Y2              =   1200
            End
            Begin VB.Line Line1 
               Index           =   72
               X1              =   0
               X2              =   12405
               Y1              =   1020
               Y2              =   1020
            End
         End
         Begin VB.PictureBox picComm 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   -165
            ScaleHeight     =   480
            ScaleWidth      =   12615
            TabIndex        =   93
            Top             =   5985
            Width           =   12615
            Begin VB.CommandButton cmdInfect 
               Caption         =   "…"
               Height          =   225
               Index           =   1
               Left            =   11820
               TabIndex        =   209
               Top             =   135
               Width           =   390
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   36
               Left            =   6225
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   207
               Top             =   135
               Width           =   6000
            End
            Begin VB.PictureBox picOpt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   12
               Left            =   1485
               ScaleHeight     =   330
               ScaleWidth      =   3750
               TabIndex        =   191
               Top             =   90
               Width           =   3750
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "3.治疗(□)"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   21
                  Left            =   2415
                  TabIndex        =   194
                  Top             =   15
                  Width           =   1230
               End
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "2.预防(△)"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   18
                  Left            =   1140
                  TabIndex        =   193
                  Top             =   15
                  Width           =   1260
               End
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "1.未用药"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   17
                  Left            =   0
                  TabIndex        =   192
                  Top             =   15
                  Value           =   -1  'True
                  Width           =   1080
               End
            End
            Begin VB.Line Line1 
               Index           =   92
               X1              =   12195
               X2              =   6195
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line1 
               Index           =   77
               X1              =   330
               X2              =   330
               Y1              =   0
               Y2              =   570
            End
            Begin VB.Label lblInfo 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   60
               Left            =   120
               TabIndex        =   96
               Top             =   150
               Width           =   225
            End
            Begin VB.Label lblInfo 
               Caption         =   "用药目的"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   67
               Left            =   450
               TabIndex        =   95
               Top             =   120
               Width           =   840
            End
            Begin VB.Label lblInfo 
               Caption         =   "感染诊断"
               Height          =   195
               Index           =   68
               Left            =   5400
               TabIndex        =   94
               Top             =   135
               Width           =   750
            End
            Begin VB.Line Line1 
               Index           =   65
               X1              =   0
               X2              =   12405
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line1 
               Index           =   73
               X1              =   15
               X2              =   12420
               Y1              =   465
               Y2              =   465
            End
            Begin VB.Line Line1 
               Index           =   74
               X1              =   12405
               X2              =   12405
               Y1              =   -60
               Y2              =   1125
            End
            Begin VB.Line Line1 
               Index           =   75
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   1185
            End
            Begin VB.Line Line1 
               Index           =   76
               X1              =   1320
               X2              =   1320
               Y1              =   0
               Y2              =   570
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Index           =   5
            Left            =   1590
            ScaleHeight     =   690
            ScaleWidth      =   10200
            TabIndex        =   139
            Top             =   4995
            Width           =   10200
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   42
               Left            =   8475
               TabIndex        =   41
               Top             =   420
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   43
               Left            =   7695
               TabIndex        =   40
               Top             =   420
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   44
               Left            =   5370
               TabIndex        =   39
               Top             =   420
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   45
               Left            =   4590
               TabIndex        =   38
               Top             =   420
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   48
               Left            =   2100
               TabIndex        =   37
               Top             =   405
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   49
               Left            =   1380
               TabIndex        =   36
               Top             =   405
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   50
               Left            =   9360
               TabIndex        =   35
               Top             =   45
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   51
               Left            =   8550
               TabIndex        =   34
               Top             =   45
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   52
               Left            =   5370
               TabIndex        =   33
               Top             =   45
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   53
               Left            =   4590
               TabIndex        =   32
               Top             =   45
               Width           =   600
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   54
               Left            =   1740
               TabIndex        =   31
               Top             =   45
               Width           =   675
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   55
               Left            =   765
               MaxLength       =   5
               TabIndex        =   30
               Top             =   45
               Width           =   600
            End
            Begin VB.Line Line1 
               Index           =   51
               X1              =   8460
               X2              =   9090
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Line Line1 
               Index           =   52
               X1              =   7680
               X2              =   8310
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Label lblInfo 
               Caption         =   "肌酐(Cr)：      (        )"
               Height          =   195
               Index           =   19
               Left            =   6885
               TabIndex        =   145
               Top             =   420
               Width           =   2595
            End
            Begin VB.Line Line1 
               Index           =   53
               X1              =   5355
               X2              =   5985
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Line Line1 
               Index           =   54
               X1              =   4575
               X2              =   5205
               Y1              =   615
               Y2              =   615
            End
            Begin VB.Label lblInfo 
               Caption         =   "谷丙转氨酶(ATL)：       (        )"
               Height          =   195
               Index           =   20
               Left            =   3075
               TabIndex        =   144
               Top             =   420
               Width           =   3135
            End
            Begin VB.Line Line1 
               Index           =   55
               X1              =   2085
               X2              =   2715
               Y1              =   615
               Y2              =   615
            End
            Begin VB.Line Line1 
               Index           =   56
               X1              =   1365
               X2              =   1995
               Y1              =   615
               Y2              =   615
            End
            Begin VB.Label lblInfo 
               Caption         =   "C反应蛋白(CPR)：      (       )"
               Height          =   195
               Index           =   22
               Left            =   0
               TabIndex        =   143
               Top             =   420
               Width           =   2790
            End
            Begin VB.Line Line1 
               Index           =   57
               X1              =   9360
               X2              =   9970
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Line Line1 
               Index           =   58
               X1              =   8535
               X2              =   9165
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "中性粒细胞(NEUT%)：       (        )"
               Height          =   195
               Index           =   24
               Left            =   6885
               TabIndex        =   142
               Top             =   45
               Width           =   3720
            End
            Begin VB.Line Line1 
               Index           =   59
               X1              =   5355
               X2              =   5985
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Line Line1 
               Index           =   60
               X1              =   4575
               X2              =   5205
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "白细胞计数(WBC)：       (        )"
               Height          =   195
               Index           =   26
               Left            =   3075
               TabIndex        =   141
               Top             =   45
               Width           =   3405
            End
            Begin VB.Line Line1 
               Index           =   61
               X1              =   1740
               X2              =   2415
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Line Line1 
               Index           =   62
               X1              =   750
               X2              =   1380
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "体温(t)：    　 ℃(        )"
               Height          =   195
               Index           =   28
               Left            =   0
               TabIndex        =   140
               Top             =   30
               Width           =   2580
            End
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   0
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   435
            Width           =   1245
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   1
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   435
            Width           =   1245
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   870
            Width           =   1950
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   3
            Left            =   5175
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   870
            Width           =   1950
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   4
            Left            =   11625
            Locked          =   -1  'True
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   435
            Width           =   870
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   5
            Left            =   11625
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   870
            Width           =   885
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   6
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   1455
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   7
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   1455
            Width           =   585
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   8
            Left            =   5370
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1455
            Width           =   645
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   9
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1455
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   10
            Left            =   11190
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1455
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   11
            Left            =   2340
            MaxLength       =   5
            TabIndex        =   7
            Top             =   3345
            Width           =   630
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   12
            Left            =   3330
            TabIndex        =   8
            Top             =   3345
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   13
            Left            =   6180
            TabIndex        =   9
            Top             =   3345
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   14
            Left            =   6960
            TabIndex        =   10
            Top             =   3345
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   15
            Left            =   10140
            TabIndex        =   11
            Top             =   3345
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   16
            Left            =   10995
            TabIndex        =   12
            Top             =   3345
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   17
            Left            =   2970
            TabIndex        =   13
            Top             =   3705
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   18
            Left            =   3690
            TabIndex        =   14
            Top             =   3720
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   19
            Left            =   6180
            TabIndex        =   15
            Top             =   3720
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   20
            Left            =   6960
            TabIndex        =   16
            Top             =   3720
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   21
            Left            =   9285
            TabIndex        =   17
            Top             =   3720
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   22
            Left            =   10065
            TabIndex        =   18
            Top             =   3720
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   23
            Left            =   4470
            TabIndex        =   20
            Top             =   4215
            Width           =   1080
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   24
            Left            =   4470
            TabIndex        =   27
            Top             =   4530
            Width           =   1080
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   25
            Left            =   10260
            TabIndex        =   24
            Top             =   4230
            Width           =   1350
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   26
            Left            =   6435
            TabIndex        =   21
            Top             =   4215
            Width           =   1845
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1515
            ScaleHeight     =   315
            ScaleWidth      =   1125
            TabIndex        =   82
            Top             =   2715
            Width           =   1125
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "无"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   45
               TabIndex        =   6
               Top             =   30
               Value           =   -1  'True
               Width           =   510
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "有"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   585
               TabIndex        =   83
               Top             =   45
               Width           =   510
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2640
            ScaleHeight     =   285
            ScaleWidth      =   1755
            TabIndex        =   80
            Top             =   4200
            Width           =   1755
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "1.未做"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Value           =   -1  'True
               Width           =   900
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "2.做"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   975
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   0
               Width           =   720
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2640
            ScaleHeight     =   270
            ScaleWidth      =   1740
            TabIndex        =   79
            Top             =   4500
            Width           =   1740
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "1.未做"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   25
               Top             =   -15
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "2.做"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   975
               TabIndex        =   26
               Top             =   0
               Width           =   700
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   3
            Left            =   8550
            ScaleHeight     =   435
            ScaleWidth      =   1710
            TabIndex        =   78
            Top             =   4185
            Width           =   1710
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "未检出/"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "检出-"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   930
               TabIndex        =   23
               Top             =   0
               Width           =   795
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   4
            Left            =   6120
            ScaleHeight     =   360
            ScaleWidth      =   1740
            TabIndex        =   77
            Top             =   4500
            Width           =   1740
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "相符/"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Value           =   -1  'True
               Width           =   765
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "不相符"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   750
               TabIndex        =   29
               Top             =   0
               Width           =   945
            End
         End
         Begin VB.PictureBox picComm 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5280
            Index           =   0
            Left            =   180
            ScaleHeight     =   5280
            ScaleWidth      =   12630
            TabIndex        =   64
            Top             =   4485
            Width           =   12630
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   360
               Index           =   57
               Left            =   1470
               MultiLine       =   -1  'True
               TabIndex        =   58
               Top             =   4020
               Width           =   10800
            End
            Begin VB.PictureBox picOpt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   570
               Index           =   11
               Left            =   1365
               ScaleHeight     =   570
               ScaleWidth      =   11025
               TabIndex        =   172
               Top             =   3315
               Width           =   11025
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "适应证(如无适应证，不再评价余下各项)"
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   37
                  Left            =   120
                  TabIndex        =   184
                  Top             =   30
                  Width           =   3630
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "药物选择"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   36
                  Left            =   3825
                  TabIndex        =   183
                  Top             =   30
                  Width           =   1050
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "单次剂量"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   35
                  Left            =   4965
                  TabIndex        =   182
                  Top             =   30
                  Width           =   1035
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "每日给药频次"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   34
                  Left            =   6105
                  TabIndex        =   181
                  Top             =   30
                  Width           =   1530
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "溶 剂"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   33
                  Left            =   7650
                  TabIndex        =   180
                  Top             =   30
                  Width           =   795
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "用药途径"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   32
                  Left            =   8505
                  TabIndex        =   179
                  Top             =   30
                  Width           =   1065
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "用药疗程"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   31
                  Left            =   9600
                  TabIndex        =   178
                  Top             =   30
                  Width           =   1155
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "更换药品"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   30
                  Left            =   120
                  TabIndex        =   177
                  Top             =   315
                  Width           =   1110
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "联合用药"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   29
                  Left            =   1245
                  TabIndex        =   176
                  Top             =   315
                  Width           =   1215
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "术前"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   4320
                  TabIndex        =   175
                  Top             =   315
                  Width           =   795
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "术中"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   5175
                  TabIndex        =   174
                  Top             =   315
                  Width           =   915
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "术后"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   6090
                  TabIndex        =   173
                  Top             =   315
                  Width           =   915
               End
               Begin VB.Label lblInfo 
                  Caption         =   "围手术期用药时间："
                  Height          =   195
                  Index           =   70
                  Left            =   2655
                  TabIndex        =   185
                  Top             =   345
                  Width           =   1635
               End
            End
            Begin VB.PictureBox picOpt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   570
               Index           =   8
               Left            =   1365
               ScaleHeight     =   570
               ScaleWidth      =   11025
               TabIndex        =   158
               Top             =   2625
               Width           =   11025
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "术后"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   6090
                  TabIndex        =   170
                  Top             =   315
                  Width           =   915
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "术中"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   5175
                  TabIndex        =   169
                  Top             =   315
                  Width           =   915
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "术前"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   4320
                  TabIndex        =   168
                  Top             =   315
                  Width           =   915
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "联合用药"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   1245
                  TabIndex        =   167
                  Top             =   315
                  Width           =   1215
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "更换药品"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   120
                  TabIndex        =   166
                  Top             =   315
                  Width           =   1110
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "用药疗程"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   9600
                  TabIndex        =   165
                  Top             =   30
                  Width           =   1155
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "用药途径"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   8505
                  TabIndex        =   164
                  Top             =   30
                  Width           =   1035
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "溶 剂"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   7650
                  TabIndex        =   163
                  Top             =   30
                  Width           =   795
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "每日给药频次"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   6105
                  TabIndex        =   162
                  Top             =   30
                  Width           =   1530
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "单次剂量"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   4965
                  TabIndex        =   161
                  Top             =   30
                  Width           =   1050
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "药物选择"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   3825
                  TabIndex        =   160
                  Top             =   30
                  Width           =   1035
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "适应证(如无适应证，不再评价余下各项)"
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   120
                  TabIndex        =   159
                  Top             =   30
                  Width           =   3630
               End
               Begin VB.Label lblInfo 
                  Caption         =   "围手术期用药时间："
                  Height          =   195
                  Index           =   59
                  Left            =   2655
                  TabIndex        =   171
                  Top             =   330
                  Width           =   1635
               End
            End
            Begin VB.PictureBox picOpt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   8580
               ScaleHeight     =   315
               ScaleWidth      =   1275
               TabIndex        =   149
               Top             =   2205
               Width           =   1275
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "有"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   20
                  Left            =   0
                  TabIndex        =   55
                  Top             =   0
                  Width           =   570
               End
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "无"
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   720
                  TabIndex        =   56
                  Top             =   30
                  Value           =   -1  'True
                  Width           =   645
               End
            End
            Begin VB.PictureBox picUseDrug 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1665
               Index           =   0
               Left            =   1440
               ScaleHeight     =   1665
               ScaleWidth      =   10920
               TabIndex        =   67
               Top             =   60
               Width           =   10920
               Begin VB.TextBox txtInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  BorderStyle     =   0  'None
                  Height          =   200
                  Index           =   30
                  Left            =   2220
                  Locked          =   -1  'True
                  TabIndex        =   197
                  Text            =   "0"
                  Top             =   1425
                  Width           =   600
               End
               Begin VB.TextBox txtInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  BorderStyle     =   0  'None
                  Height          =   200
                  Index           =   29
                  Left            =   1335
                  Locked          =   -1  'True
                  TabIndex        =   196
                  Text            =   "0"
                  Top             =   1410
                  Width           =   600
               End
               Begin VSFlex8Ctl.VSFlexGrid vsDrugUse 
                  Height          =   1395
                  Left            =   0
                  TabIndex        =   195
                  Top             =   0
                  Width           =   10875
                  _cx             =   19182
                  _cy             =   2461
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
                  BackColorSel    =   16764057
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483636
                  GridColorFixed  =   -2147483636
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
                  Rows            =   18
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   250
                  RowHeightMax    =   2000
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmKssSurveyEdit.frx":693A
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
                  AllowUserFreezing=   0
                  BackColorFrozen =   0
                  ForeColorFrozen =   0
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VB.Label lblInfo 
                  Caption         =   "累计使用抗菌药        种        天"
                  Height          =   195
                  Index           =   55
                  Left            =   30
                  TabIndex        =   68
                  Top             =   1425
                  Width           =   3375
               End
               Begin VB.Line Line1 
                  Index           =   42
                  X1              =   1350
                  X2              =   1920
                  Y1              =   1620
                  Y2              =   1620
               End
               Begin VB.Line Line1 
                  Index           =   43
                  X1              =   2215
                  X2              =   2805
                  Y1              =   1620
                  Y2              =   1620
               End
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   31
               Left            =   2565
               Locked          =   -1  'True
               TabIndex        =   47
               Text            =   "0"
               Top             =   1845
               Width           =   885
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   32
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   48
               Text            =   "0"
               Top             =   1845
               Width           =   750
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   33
               Left            =   8385
               Locked          =   -1  'True
               TabIndex        =   49
               Text            =   "0"
               Top             =   1860
               Width           =   750
            End
            Begin VB.PictureBox picOpt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   6
               Left            =   1515
               ScaleHeight     =   285
               ScaleWidth      =   2610
               TabIndex        =   66
               Top             =   2235
               Width           =   2610
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "无效"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   14
                  Left            =   1560
                  TabIndex        =   52
                  Top             =   0
                  Width           =   795
               End
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "好转"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   13
                  Left            =   780
                  TabIndex        =   51
                  Top             =   0
                  Width           =   795
               End
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "治愈"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   12
                  Left            =   0
                  TabIndex        =   50
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   795
               End
            End
            Begin VB.PictureBox picOpt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   4845
               ScaleHeight     =   315
               ScaleWidth      =   1275
               TabIndex        =   65
               Top             =   2220
               Width           =   1275
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "无"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   16
                  Left            =   705
                  TabIndex        =   54
                  Top             =   15
                  Value           =   -1  'True
                  Width           =   645
               End
               Begin VB.OptionButton optInfo 
                  Appearance      =   0  'Flat
                  Caption         =   "有"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   15
                  Left            =   0
                  TabIndex        =   53
                  Top             =   0
                  Width           =   645
               End
            End
            Begin VB.Label lblInfo 
               Caption         =   $"frmKssSurveyEdit.frx":6A22
               Height          =   375
               Index           =   74
               Left            =   1440
               TabIndex        =   189
               Top             =   4545
               Width           =   10830
            End
            Begin VB.Label lblInfo 
               Caption         =   $"frmKssSurveyEdit.frx":6AF3
               Height          =   375
               Index           =   73
               Left            =   1440
               TabIndex        =   188
               Top             =   4560
               Width           =   10890
            End
            Begin VB.Label lblInfo 
               Caption         =   "说明"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   72
               Left            =   570
               TabIndex        =   187
               Top             =   4635
               Width           =   435
            End
            Begin VB.Label lblInfo 
               Caption         =   "备注"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   71
               Left            =   570
               TabIndex        =   186
               Top             =   4125
               Width           =   435
            End
            Begin VB.Label lblInfo 
               Caption         =   "14"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   69
               Left            =   90
               TabIndex        =   157
               Top             =   4635
               Width           =   210
            End
            Begin VB.Label lblInfo 
               Caption         =   "13"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   66
               Left            =   90
               TabIndex        =   156
               Top             =   4110
               Width           =   210
            End
            Begin VB.Label lblInfo 
               Caption         =   "12"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   65
               Left            =   90
               TabIndex        =   155
               Top             =   3450
               Width           =   195
            End
            Begin VB.Label lblInfo 
               Caption         =   "11"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   64
               Left            =   90
               TabIndex        =   154
               Top             =   2790
               Width           =   195
            End
            Begin VB.Label lblInfo 
               Caption         =   "用药合理性评价"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1080
               Index           =   63
               Left            =   480
               TabIndex        =   153
               Top             =   2700
               Width           =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "中心  或 分网"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Index           =   62
               Left            =   870
               TabIndex        =   152
               Top             =   3300
               Width           =   420
            End
            Begin VB.Label lblInfo 
               Caption         =   "本院"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   61
               Left            =   870
               TabIndex        =   151
               Top             =   2820
               Width           =   435
            End
            Begin VB.Line Line1 
               Index           =   99
               X1              =   810
               X2              =   810
               Y1              =   2565
               Y2              =   3945
            End
            Begin VB.Line Line1 
               Index           =   98
               X1              =   330
               X2              =   0
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Line Line1 
               Index           =   97
               X1              =   810
               X2              =   12405
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Line Line1 
               Index           =   96
               X1              =   0
               X2              =   12405
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Line Line1 
               Index           =   95
               X1              =   0
               X2              =   12405
               Y1              =   3945
               Y2              =   3945
            End
            Begin VB.Line Line1 
               Index           =   94
               X1              =   8385
               X2              =   8385
               Y1              =   2130
               Y2              =   2560
            End
            Begin VB.Line Line1 
               Index           =   93
               X1              =   4680
               X2              =   4680
               Y1              =   2130
               Y2              =   2560
            End
            Begin VB.Label lblInfo 
               Caption         =   "使用抗真菌药"
               Height          =   195
               Index           =   54
               Left            =   9990
               TabIndex        =   150
               Top             =   2250
               Width           =   1515
            End
            Begin VB.Line Line1 
               Index           =   83
               X1              =   1320
               X2              =   1320
               Y1              =   0
               Y2              =   5000
            End
            Begin VB.Line Line1 
               Index           =   81
               X1              =   330
               X2              =   330
               Y1              =   15
               Y2              =   4980
            End
            Begin VB.Label lblInfo 
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   45
               Left            =   120
               TabIndex        =   76
               Top             =   795
               Width           =   225
            End
            Begin VB.Label lblInfo 
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   46
               Left            =   120
               TabIndex        =   75
               Top             =   1845
               Width           =   225
            End
            Begin VB.Label lblInfo 
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   47
               Left            =   105
               TabIndex        =   74
               Top             =   2250
               Width           =   225
            End
            Begin VB.Label lblInfo 
               Caption         =   "用药情况 (注射用药请同写清溶剂名称及用量)"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   945
               Index           =   50
               Left            =   375
               TabIndex        =   73
               Top             =   465
               Width           =   915
            End
            Begin VB.Label lblInfo 
               Caption         =   "费用(元)"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   51
               Left            =   495
               TabIndex        =   72
               Top             =   1845
               Width           =   960
            End
            Begin VB.Label lblInfo 
               Caption         =   "治疗结果"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   52
               Left            =   480
               TabIndex        =   71
               Top             =   2265
               Width           =   960
            End
            Begin VB.Label lblInfo 
               Caption         =   "住院总费用：          元   住院药品总费用：        元    住院抗菌药物总费用：        元"
               Height          =   195
               Index           =   57
               Left            =   1530
               TabIndex        =   70
               Top             =   1860
               Width           =   9990
            End
            Begin VB.Line Line1 
               Index           =   44
               X1              =   2550
               X2              =   3450
               Y1              =   2055
               Y2              =   2055
            End
            Begin VB.Line Line1 
               Index           =   45
               X1              =   5325
               X2              =   6075
               Y1              =   2055
               Y2              =   2055
            End
            Begin VB.Line Line1 
               Index           =   46
               X1              =   8385
               X2              =   9135
               Y1              =   2070
               Y2              =   2070
            End
            Begin VB.Label lblInfo 
               Caption         =   "继发(医院)感染"
               Height          =   195
               Index           =   58
               Left            =   6150
               TabIndex        =   69
               Top             =   2250
               Width           =   1515
            End
            Begin VB.Line Line1 
               Index           =   78
               X1              =   0
               X2              =   12405
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line1 
               Index           =   39
               X1              =   0
               X2              =   12405
               Y1              =   2565
               Y2              =   2565
            End
            Begin VB.Line Line1 
               Index           =   40
               X1              =   0
               X2              =   12405
               Y1              =   2130
               Y2              =   2130
            End
            Begin VB.Line Line1 
               Index           =   41
               X1              =   0
               X2              =   12405
               Y1              =   1725
               Y2              =   1725
            End
            Begin VB.Line Line1 
               Index           =   79
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   4980
            End
            Begin VB.Line Line1 
               Index           =   82
               X1              =   12405
               X2              =   12405
               Y1              =   0
               Y2              =   5000
            End
            Begin VB.Line Line1 
               Index           =   84
               X1              =   0
               X2              =   12405
               Y1              =   4980
               Y2              =   4980
            End
         End
         Begin VB.Line Line1 
            Index           =   35
            X1              =   12300
            X2              =   4155
            Y1              =   3105
            Y2              =   3105
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   12300
            X2              =   1590
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label lblInfo 
            Caption         =   "临床微生物检  查"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   75
            Left            =   960
            TabIndex        =   190
            Top             =   4095
            Width           =   405
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   6900
            X2              =   8145
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   5310
            X2              =   6555
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "手术病人抗菌药物使用情况调查表"
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
            Left            =   4290
            TabIndex        =   138
            Top             =   0
            Width           =   3825
         End
         Begin VB.Label lblInfo 
            Caption         =   "抽样时间："
            Height          =   195
            Index           =   0
            Left            =   4455
            TabIndex        =   137
            Top             =   450
            Width           =   900
         End
         Begin VB.Label lblInfo 
            Caption         =   "               到"
            Height          =   210
            Index           =   1
            Left            =   5310
            TabIndex        =   136
            Top             =   450
            Width           =   2850
         End
         Begin VB.Label lblInfo 
            Caption         =   "病人所属科室："
            Height          =   195
            Index           =   2
            Left            =   100
            TabIndex        =   135
            Top             =   870
            Width           =   1260
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   1320
            X2              =   3270
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lblInfo 
            Caption         =   "病历号："
            Height          =   195
            Index           =   3
            Left            =   4455
            TabIndex        =   134
            Top             =   870
            Width           =   765
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   5175
            X2              =   7125
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "手术病人出院人数："
            Height          =   195
            Index           =   4
            Left            =   9720
            TabIndex        =   133
            Top             =   450
            Width           =   1965
         End
         Begin VB.Label lblInfo 
            Caption         =   "序号："
            Height          =   195
            Index           =   5
            Left            =   11130
            TabIndex        =   132
            Top             =   870
            Width           =   600
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   11625
            X2              =   12510
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   100
            X2              =   12500
            Y1              =   1275
            Y2              =   1275
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   105
            X2              =   12505
            Y1              =   1815
            Y2              =   1815
         End
         Begin VB.Label lblInfo 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   131
            Top             =   1470
            Width           =   90
         End
         Begin VB.Label lblInfo 
            Caption         =   "基本情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   555
            TabIndex        =   130
            Top             =   1470
            Width           =   795
         End
         Begin VB.Label lblInfo 
            Caption         =   "性别"
            Height          =   195
            Index           =   8
            Left            =   1590
            TabIndex        =   129
            Top             =   1455
            Width           =   510
         End
         Begin VB.Label lblInfo 
            Caption         =   "年龄        "
            Height          =   195
            Index           =   9
            Left            =   3315
            TabIndex        =   128
            Top             =   1455
            Width           =   1365
         End
         Begin VB.Label lblInfo 
            Caption         =   "体重        "
            Height          =   195
            Index           =   10
            Left            =   4965
            TabIndex        =   127
            Top             =   1455
            Width           =   1755
         End
         Begin VB.Label lblInfo 
            Caption         =   "入院时间"
            Height          =   195
            Index           =   11
            Left            =   8100
            TabIndex        =   126
            Top             =   1455
            Width           =   795
         End
         Begin VB.Label lblInfo 
            Caption         =   "出院时间"
            Height          =   195
            Index           =   12
            Left            =   10425
            TabIndex        =   125
            Top             =   1455
            Width           =   795
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2070
            X2              =   2745
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   3735
            X2              =   4305
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   5370
            X2              =   6015
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   8880
            X2              =   10020
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   11190
            X2              =   12330
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   105
            X2              =   12505
            Y1              =   2490
            Y2              =   2490
         End
         Begin VB.Label lblInfo 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   124
            Top             =   2100
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "诊断"
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
            Index           =   14
            Left            =   720
            TabIndex        =   123
            Top             =   2100
            Width           =   390
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   100
            X2              =   12500
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Label lblInfo 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   122
            Top             =   2790
            Width           =   90
         End
         Begin VB.Label lblInfo 
            Caption         =   "过敏史"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   600
            TabIndex        =   121
            Top             =   2760
            Width           =   645
         End
         Begin VB.Label lblInfo 
            Caption         =   "抗菌物品通用名"
            Height          =   195
            Index           =   17
            Left            =   2775
            TabIndex        =   120
            Top             =   2775
            Width           =   1260
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   900
            X2              =   12490
            Y1              =   4050
            Y2              =   4050
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   90
            X2              =   12490
            Y1              =   5760
            Y2              =   5760
         End
         Begin VB.Label lblInfo 
            Caption         =   "体温(t)：    　 ℃(        )"
            Height          =   195
            Index           =   18
            Left            =   1590
            TabIndex        =   119
            Top             =   3330
            Width           =   2580
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   2340
            X2              =   2970
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   3330
            X2              =   4005
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Label lblInfo 
            Caption         =   "白细胞计数(WBC)：       (        )"
            Height          =   195
            Index           =   21
            Left            =   4665
            TabIndex        =   118
            Top             =   3345
            Width           =   3405
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   6165
            X2              =   6795
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   6945
            X2              =   7575
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Label lblInfo 
            Caption         =   "中性粒细胞(NEUT%)：       (        )"
            Height          =   195
            Index           =   23
            Left            =   8490
            TabIndex        =   117
            Top             =   3345
            Width           =   3720
         End
         Begin VB.Line Line1 
            Index           =   22
            X1              =   10140
            X2              =   10740
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Line Line1 
            Index           =   23
            X1              =   10995
            X2              =   11600
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Label lblInfo 
            Caption         =   "C反应蛋白(CPR)：      (       )"
            Height          =   195
            Index           =   25
            Left            =   1590
            TabIndex        =   116
            Top             =   3720
            Width           =   2790
         End
         Begin VB.Line Line1 
            Index           =   24
            X1              =   2955
            X2              =   3585
            Y1              =   3915
            Y2              =   3915
         End
         Begin VB.Line Line1 
            Index           =   25
            X1              =   3675
            X2              =   4305
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Label lblInfo 
            Caption         =   "谷丙转氨酶(ATL)：       (        )"
            Height          =   195
            Index           =   27
            Left            =   4665
            TabIndex        =   115
            Top             =   3720
            Width           =   3135
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   6165
            X2              =   6795
            Y1              =   3915
            Y2              =   3915
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   6945
            X2              =   7575
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Label lblInfo 
            Caption         =   "肌酐(Cr)：      (        )"
            Height          =   195
            Index           =   29
            Left            =   8475
            TabIndex        =   114
            Top             =   3720
            Width           =   2595
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   9270
            X2              =   9900
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   10050
            X2              =   10680
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Label lblInfo 
            Caption         =   "病源学检测："
            Height          =   195
            Index           =   30
            Left            =   1590
            TabIndex        =   113
            Top             =   4215
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Caption         =   "药敏试验："
            Height          =   195
            Index           =   31
            Left            =   1590
            TabIndex        =   112
            Top             =   4500
            Width           =   885
         End
         Begin VB.Label lblInfo 
            Caption         =   "(             )"
            Height          =   195
            Index           =   32
            Left            =   4350
            TabIndex        =   111
            Top             =   4215
            Width           =   1530
         End
         Begin VB.Line Line1 
            Index           =   30
            X1              =   4470
            X2              =   5535
            Y1              =   4425
            Y2              =   4425
         End
         Begin VB.Label lblInfo 
            Caption         =   "(             )"
            Height          =   195
            Index           =   33
            Left            =   4350
            TabIndex        =   110
            Top             =   4530
            Width           =   1560
         End
         Begin VB.Line Line1 
            Index           =   31
            X1              =   4455
            X2              =   5565
            Y1              =   4740
            Y2              =   4740
         End
         Begin VB.Label lblInfo 
            Caption         =   "(                                   菌)"
            Height          =   195
            Index           =   34
            Left            =   8400
            TabIndex        =   109
            Top             =   4215
            Width           =   3720
         End
         Begin VB.Line Line1 
            Index           =   32
            X1              =   10260
            X2              =   11610
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Label lblInfo 
            Caption         =   "(                    )"
            Height          =   195
            Index           =   35
            Left            =   5985
            TabIndex        =   108
            Top             =   4530
            Width           =   2010
         End
         Begin VB.Label lblInfo 
            Caption         =   "标本"
            Height          =   195
            Index           =   36
            Left            =   6015
            TabIndex        =   107
            Top             =   4215
            Width           =   510
         End
         Begin VB.Line Line1 
            Index           =   33
            X1              =   6435
            X2              =   8280
            Y1              =   4425
            Y2              =   4425
         End
         Begin VB.Line Line1 
            Index           =   34
            X1              =   900
            X2              =   12500
            Y1              =   4860
            Y2              =   4860
         End
         Begin VB.Label lblInfo 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   240
            TabIndex        =   106
            Top             =   4395
            Width           =   90
         End
         Begin VB.Line Line1 
            Index           =   38
            X1              =   420
            X2              =   420
            Y1              =   1275
            Y2              =   5760
         End
         Begin VB.Line Line1 
            Index           =   47
            X1              =   1410
            X2              =   1410
            Y1              =   1290
            Y2              =   5760
         End
         Begin VB.Line Line1 
            Index           =   48
            X1              =   90
            X2              =   90
            Y1              =   1275
            Y2              =   5760
         End
         Begin VB.Line Line1 
            Index           =   49
            X1              =   12495
            X2              =   12495
            Y1              =   1290
            Y2              =   5760
         End
         Begin VB.Line Line1 
            Index           =   50
            X1              =   900
            X2              =   900
            Y1              =   3180
            Y2              =   5760
         End
         Begin VB.Label lblInfo 
            Caption         =   "实验室检查"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Index           =   40
            Left            =   555
            TabIndex        =   105
            Top             =   4110
            Width           =   180
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   11625
            X2              =   12495
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Label lblInfo 
            Caption         =   "用药前"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   42
            Left            =   1050
            TabIndex        =   104
            Top             =   3315
            Width           =   225
         End
         Begin VB.Label lblInfo 
            Caption         =   "用药后"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   43
            Left            =   1050
            TabIndex        =   103
            Top             =   5025
            Width           =   210
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   10485
      Top             =   30
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
            Picture         =   "frmKssSurveyEdit.frx":6B94
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":712E
            Key             =   "PatiMan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":76C8
            Key             =   "PatiWoman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":7C62
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":81FC
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":8796
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":EFF8
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssSurveyEdit.frx":1585A
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      Index           =   80
      X1              =   1185
      X2              =   930
      Y1              =   75
      Y2              =   390
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   330
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKssSurveyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Enum mCtlID
    
    '表格外 e_F;e_P用药评价;e_D调查表格明细
    e_F_txtInfo_抽样时间_起_0 = 0
    e_F_txtInfo_抽样时间_止_1 = 1
    e_F_lblInfo_出院人数标签_4 = 4
    e_F_txtInfo_出院人数_4 = 4
    e_F_txtInfo_所属科室_2 = 2
    e_F_txtInfo_病历号_3 = 3
    e_F_txtInfo_序号_5 = 5
    
    e_P_txtInfo_抽样时间_起_28 = 28
    e_P_txtInfo_抽样时间_止_56 = 56
    
    e_D_lblInfo_基本_序号_6 = 6
    e_D_txtInfo_基本_性别_6 = 6
    e_D_txtInfo_基本_年龄_7 = 7
    e_D_txtInfo_基本_体重_8 = 8
    e_D_txtInfo_基本_入院时间_9 = 9
    e_D_txtInfo_基本_出院时间_10 = 10
    
    e_D_lblInfo_诊断_序号_13 = 13
    e_D_txtInfo_诊断_诊断_27 = 27
    
    e_D_lblInfo_过敏_序号_15 = 15
    e_D_optInfo_过敏_无_0 = 0
    e_D_optInfo_过敏_有_1 = 1
    e_D_txtInfo_过敏_通用名_34 = 34
    
    e_D_optInfo_检查_病原学检测_做_3 = 3
    e_D_optInfo_检查_病原学检测_未做_2 = 2
    e_D_txtInfo_检查_病原学检测日期_23 = 23
    e_D_txtInfo_检查_病原学检测标本_26 = 26
    e_D_optInfo_检查_病原学检测_检出_4 = 4
    e_D_optInfo_检查_病原学检测_未检出_5 = 5
    e_D_txtInfo_检查_病原学检测检出细菌名_25 = 25
    e_D_optInfo_检查_药敏试验_未做_6 = 6
    e_D_optInfo_检查_药敏试验_做_7 = 7
    e_D_txtInfo_检查_药敏试验日期_24 = 24
    e_D_optInfo_检查_药敏试验_相符_8 = 8
    e_D_optInfo_检查_药敏试验_不相符_9 = 9
    e_D_txtInfo_检查_用药前体温_11 = 11
    e_D_txtInfo_检查_用药前白细胞计数_13 = 13
    e_D_txtInfo_检查_用药前中性粒细胞_15 = 15
    e_D_txtInfo_检查_用药前C反应蛋白_17 = 17
    e_D_txtInfo_检查_用药前丙谷转氨酶_19 = 19
    e_D_txtInfo_检查_用药前肌酐_21 = 21
    e_D_txtInfo_检查_用药后体温_55 = 55
    e_D_txtInfo_检查_用药后白细胞计数_53 = 53
    e_D_txtInfo_检查_用药后中性粒细胞_51 = 51
    e_D_txtInfo_检查_用药后C反应蛋白_49 = 49
    e_D_txtInfo_检查_用药后丙谷转氨酶_45 = 45
    e_D_txtInfo_检查_用药后肌酐_43 = 43
    e_D_txtInfo_检查_用药前体温日期_12 = 12
    e_D_txtInfo_检查_用药前白细胞计数日期_14 = 14
    e_D_txtInfo_检查_用药前中性粒细胞日期_16 = 16
    e_D_txtInfo_检查_用药前C反应蛋白日期_18 = 18
    e_D_txtInfo_检查_用药前丙谷转氨酶日期_20 = 20
    e_D_txtInfo_检查_用药前肌酐日期_22 = 22
    e_D_txtInfo_检查_用药后体温日期_54 = 54
    e_D_txtInfo_检查_用药后白细胞计数日期_52 = 52
    e_D_txtInfo_检查_用药后中性粒细胞日期_50 = 50
    e_D_txtInfo_检查_用药后C反应蛋白日期_48 = 48
    e_D_txtInfo_检查_用药后丙谷转氨酶日期_44 = 44
    e_D_txtInfo_检查_用药后肌酐日期_42 = 42
    
    e_D_lblInfo_影像_序号_102 = 102
    e_D_chkInfo_影像_X线_18 = 18
    e_D_chkInfo_影像_CT_19 = 19
    e_D_chkInfo_影像_磁共振_20 = 20
    e_D_txtInfo_影像_部位_46 = 46
    e_D_txtInfo_影像_结论_47 = 47
    
    e_D_lblInfo_症状_序号_101 = 101
    e_D_txtInfo_症状_感染症状_35 = 35
    
    e_D_picDiff_影像And症状容器_0 = 0
    e_D_picDiff_手术情况容器_1 = 1
    
    e_D_lblInfo_用药目的_序号_60 = 60
    e_D_lblInfo_用药目的_感染诊断_68 = 68
    e_D_txtInfo_用药目的_感染诊断_36 = 36
    e_D_optInfo_用药目的_未用_17 = 17
    e_D_optInfo_用药目的_预防_18 = 18
    e_D_optInfo_用药目的_治疗_21 = 21
    
    e_D_picComm_用药目的容器_1 = 1
    e_D_picComm_用药明细及费用容器_0 = 0
    e_D_lblInfo_用药情况_序号_45 = 45
    e_D_txtInfo_用药情况_种数_29 = 29
    e_D_txtInfo_用药情况_天数_30 = 30
    
    e_D_lblInfo_费用_序号_46 = 46
    e_D_txtInfo_费用_总费_31 = 31
    e_D_txtInfo_费用_药费_32 = 32
    e_D_txtInfo_费用_抗药费_33 = 33
    
    e_D_lblInfo_结果_序号_47 = 47
    e_D_optInfo_结果_治愈_12 = 12
    e_D_optInfo_结果_好转_13 = 13
    e_D_optInfo_结果_无效_14 = 14
    e_D_optInfo_结果_感染_有_15 = 15
    e_D_optInfo_结果_感染_无_16 = 16
    
    e_D_optInfo_结果_真菌药_无_19 = 19
    e_D_optInfo_结果_真菌药_有_20 = 20
    
    e_D_lblInfo_本院评价_序号_64 = 64
    e_D_chkInfo_评价_本院_适应症_5 = 5
    e_D_chkInfo_评价_本院_药物选择_12 = 12
    e_D_chkInfo_评价_本院_单次剂量_13 = 13
    e_D_chkInfo_评价_本院_每日给药频次_14 = 14
    e_D_chkInfo_评价_本院_溶剂_15 = 15
    e_D_chkInfo_评价_本院_给药途径_16 = 16
    e_D_chkInfo_评价_本院_用药疗程_17 = 17
    e_D_chkInfo_评价_本院_更换药物_21 = 21
    e_D_chkInfo_评价_本院_联合用药_22 = 22
    e_D_lblInfo_评价_本院_围手术标签_59 = 59
    e_D_chkInfo_评价_本院_术前_23 = 23
    e_D_chkInfo_评价_本院_术中_24 = 24
    e_D_chkInfo_评价_本院_术后_25 = 25
    
    e_D_lblInfo_中心评价_序号_65 = 65
    e_D_chkInfo_评价_中心_适应症_37 = 37
    e_D_chkInfo_评价_中心_药物选择_36 = 36
    e_D_chkInfo_评价_中心_单次剂量_35 = 35
    e_D_chkInfo_评价_中心_每日给药频次_34 = 34
    e_D_chkInfo_评价_中心_溶剂_33 = 33
    e_D_chkInfo_评价_中心_给药途径_32 = 32
    e_D_chkInfo_评价_中心_用药疗程_31 = 31
    e_D_chkInfo_评价_中心_更换药物_30 = 30
    e_D_chkInfo_评价_中心_联合用药_29 = 29
    e_D_lblInfo_评价_中心_围手术标签_70 = 70
    e_D_chkInfo_评价_中心_术前_28 = 28
    e_D_chkInfo_评价_中心_术中_27 = 27
    e_D_chkInfo_评价_中心_术后_26 = 26
    
    e_D_lblInfo_备注_序号_66 = 66
    e_D_txtInfo_备注_57 = 57
    
    e_D_lblInfo_说明_序号_69 = 69
    e_D_lblInfo_说明_手术说明文本_73 = 73
    e_D_lblInfo_说明_非手术说明文本_74 = 74
    
    Type日期 = 0
    Type文本 = 1
    Type数字 = 2
End Enum

Private Enum COL_VSF '用药评表格
    COL_评价项目 = 0
    COL_合理内容
    COL_合理本院
    COL_合理中心
    COL_不合理内容
    COL_不合理本院
    COL_不合理中心
    COL_合理编码
    COL_不合理编码
    COL_上级编码
    COL_上级名称
End Enum

Private Enum COL_VSDRUGUSE '抗菌药使用明细表格
    COL_DRUG_图标 = 0
    COL_DRUG_ID = 1
    COL_DRUG_相关ID
    COL_DRUG_药名ID
    
    COL_DRUG_药品名称
    COL_DRUG_单次剂量
    COL_DRUG_给药频次
    COL_DRUG_途径
    COL_DRUG_总用量
    COL_DRUG_起止时间
End Enum

Private Enum COL_VSOPERATE '手术病人手明细
    COL_OPE_手术ID = 1
    COL_OPE_手术名称
    COL_OPE_切口
    COL_OPE_开始时间
    COL_OPE_结束时间
    COL_OPE_用药期间
    COL_OPE_给药情况
End Enum

Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mblnChange As Boolean
Private mlng抽样ID As Long
Private mlng病人ID As Long
Private mlng主页ID  As Long
Private mlng序号  As Long
Private mbln手术 As Boolean
Private mblnReturn As Boolean
Private mbln用药评价 As Boolean
Private mstr科室 As String
Private mlng人数  As Long
Private mOldwinproc As Long
Private mbln编辑 As Boolean
Private mbln打印 As Boolean
Private mlngYear As Long
Private mblnInitSaved As Boolean '表示是否进入过编辑界面，一但进入之后会保存一次相关数据，用药，手术，界面部分基本信息
Private mrsCtl As ADODB.Recordset

Public Function ShowMe(frmParent As Object, ByVal lng抽样id As Long, ByVal lng病人ID, ByVal lng主页ID, ByVal lng序号 As Long, ByVal lng人数 As Long, ByVal str科室 As String, _
            ByVal bln手术 As Boolean, ByRef bln编辑 As Boolean, ByRef bln打印 As Boolean) As Boolean
'功能：
'参数：bln手术 是否是手术病人
    mlng抽样ID = lng抽样id
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mbln手术 = bln手术
    mlng序号 = lng序号
     
    mstr科室 = str科室
    mlng人数 = lng人数
    mbln编辑 = bln编辑
    mbln打印 = bln打印
    
    frmKssSurveyEdit.Show 1, frmParent
    
    bln编辑 = mbln编辑
    bln打印 = mbln打印
    
    ShowMe = True
End Function

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " 保存(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " 保存退出(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit
            Control.Enabled = mblnChange
    End Select
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Dim blnEdit As Boolean
    Dim i As Integer
    Select Case Index
        Case e_D_chkInfo_评价_本院_适应症_5
            blnEdit = chkInfo(Index).Value
            For i = e_D_chkInfo_评价_本院_药物选择_12 To e_D_chkInfo_评价_本院_术后_25
                If InStr(",18,19,20,", "," & i & ",") = 0 Then
                    chkInfo(i).Enabled = blnEdit
                    chkInfo(i).Value = 0
                End If
            Next
        Case e_D_chkInfo_评价_中心_适应症_37
            blnEdit = chkInfo(Index).Value
            For i = e_D_chkInfo_评价_中心_术后_26 To e_D_chkInfo_评价_中心_药物选择_36
                chkInfo(i).Enabled = blnEdit
                chkInfo(i).Value = 0
            Next
    End Select
    If mbln用药评价 Then Call Cls评价明细(Index)
    If Visible Then mblnChange = True
End Sub

Private Sub cmdInfect_Click(Index As Integer)
    '功能：抽样记录选择器---------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnCanle As Boolean
    Dim X As Long, Y As Long
    Dim vRect As RECT
    Dim lngHwnd As Long
    
    If Index = 0 Then
        lngHwnd = txtInfo(e_D_txtInfo_症状_感染症状_35).hwnd
    Else
        lngHwnd = txtInfo(e_D_txtInfo_用药目的_感染诊断_36).hwnd
    End If
    GetWindowRect lngHwnd, vRect
    X = vRect.Left * Screen.TwipsPerPixelX
    Y = vRect.Top * Screen.TwipsPerPixelY
    strSQL = "Select id,诊断描述 as 感染诊断 From 病人诊断记录 Where 诊断类型 =5 and 病人id=[1] and 主页id=[2] order by 诊断次序"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "抗菌药物抽样记录", False, "", "", False, False, True, X, Y, 200, blnCanle, False, True, mlng病人ID, mlng主页ID)
    If blnCanle Then Exit Sub
    If rsTmp Is Nothing Then
        MsgBox "该出院病人首页中填写未填写感染诊断。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Index = 0 Then
        txtInfo(e_D_txtInfo_症状_感染症状_35).Text = rsTmp!感染诊断
        txtInfo(e_D_txtInfo_症状_感染症状_35).Tag = rsTmp!ID
    Else
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Text = rsTmp!感染诊断
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Tag = rsTmp!ID
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strB As String
    Dim strE As String
    Dim bln用抗菌药 As Boolean
    
    On Error GoTo errH
    
    'TabControl
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
        .InsertItem(0, "使用情况调查表", picUse.hwnd, 0).Tag = "使用情况调查表"
        .InsertItem(1, "合理性评价表", picReasonable.hwnd, 0).Tag = "合理性评价表"
        .Item(0).Selected = True
    End With
    
    Call InitCommandBar
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    strSQL = "select 范围开始时间 as 开始时间,范围结束时间 as 结束时间 from 抗菌药物抽样记录 where id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng抽样ID)
    
    
    
    txtInfo(e_F_txtInfo_抽样时间_起_0).Text = Format(rsTmp!开始时间, "yyyy-mm-dd")
    txtInfo(e_F_txtInfo_抽样时间_止_1).Text = Format(rsTmp!结束时间, "yyyy-mm-dd")
    txtInfo(e_F_txtInfo_所属科室_2).Text = mstr科室
    txtInfo(e_F_txtInfo_出院人数_4).Text = mlng人数
    txtInfo(e_F_txtInfo_序号_5).Text = mlng序号
    
    strB = Format(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, "yyyy-06-01")
    strE = Format(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, "yyyy-06-30")
    
    mbln用药评价 = False
    
    If Not mbln用药评价 Then mbln用药评价 = Between(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, strB, strE)
    
    If Not mbln用药评价 Then mbln用药评价 = Between(txtInfo(e_F_txtInfo_抽样时间_止_1).Text, strB, strE)
    
    If Not mbln用药评价 Then
        mbln用药评价 = (strB >= txtInfo(e_F_txtInfo_抽样时间_起_0).Text And strE <= txtInfo(e_F_txtInfo_抽样时间_止_1).Text)
    End If
    
    strB = Format(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, "yyyy-12-01")
    strE = Format(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, "yyyy-12-31")
    
    If Not mbln用药评价 Then mbln用药评价 = Between(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, strB, strE)
    
    If Not mbln用药评价 Then mbln用药评价 = Between(txtInfo(e_F_txtInfo_抽样时间_止_1).Text, strB, strE)
    
    If Not mbln用药评价 Then
        mbln用药评价 = (strB >= txtInfo(e_F_txtInfo_抽样时间_起_0).Text And strE <= txtInfo(e_F_txtInfo_抽样时间_止_1).Text)
    End If
    
    tbcSub.Item(1).Visible = mbln用药评价
    If mbln用药评价 Then
        txtInfo(e_P_txtInfo_抽样时间_起_28).Text = txtInfo(e_F_txtInfo_抽样时间_起_0).Text
        txtInfo(e_P_txtInfo_抽样时间_止_56).Text = txtInfo(e_F_txtInfo_抽样时间_止_1).Text
        Call Init用药评价表
    End If
    '初始化抗菌药表格
    strTmp = "ID;相关ID;药名ID;药品名称,3500,1;单次剂量,1200,7;给药频次,1200,1;途径,1200,1;总用量,800,7;起止时间(月日时分),2530,4"
    Call InitTable(vsDrugUse, strTmp)
    
    Call LoadData
    
    If mblnInitSaved Then
        bln用抗菌药 = UseKssDrug
    Else
        bln用抗菌药 = Not optInfo(e_D_optInfo_用药目的_未用_17).Value
    End If
    
    If bln用抗菌药 Then Call LoadDrugUse
 
    If mbln手术 Then
        strTmp = "手术ID;手术名称,3000,1;切口,600,4;手术开始时间,1700,4;手术结束时间,1700,4;术前初次预防用药时间,2500,1;术中给药,1000,1"
        Call InitTable(vsOperate, strTmp)
        Call LoadOperate
    End If
    Call LoadFee
    Call SetCtlProperty
    Call optInfo_Click(e_D_optInfo_检查_病原学检测_未做_2)
    Call optInfo_Click(e_D_optInfo_检查_病原学检测_未检出_5)
    Call optInfo_Click(e_D_optInfo_检查_药敏试验_未做_6)
    Call optInfo_Click(e_D_optInfo_用药目的_未用_17)
    Call optInfo_Click(e_D_optInfo_用药目的_预防_18)
    If mbln用药评价 Then Call Load用药评价
    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    
    With Me.tbcSub
    
        .Left = lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
        .Width = lngRight - lngLeft
    
    End With
    vsc.Height = picUse.Height - 20
    Me.Refresh
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strRPTName As String
    
    If mbln手术 Then
        strRPTName = "ZL1_INSIDE_1269_2"
    Else
        strRPTName = "ZL1_INSIDE_1269_1"
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet
        SwitchPrintSet glngSys & "\" & 1269
        vsPJB.Redraw = flexRDNone
        Call zlPrintSet
        vsPJB.Redraw = flexRDDirect
        SwitchPrintSet glngSys & "\" & 1269, True
    Case conMenu_File_Preview '预览
        Call Print调查表(1)
    Case conMenu_File_Print '打印
        Call Print调查表(2)
    Case conMenu_File_Exit '退出
       Call Unload(Me)
    Case conMenu_Edit_Save
        If CheckData Then Call SaveData
    Case conMenu_Edit_SaveExit
        If CheckData Then
            Call SaveData
            Call Unload(Me)
        End If
    End Select
End Sub

Private Sub Print调查表(ByVal bytType As Byte)
'功能：打印或预览报表。 bytType=1预览，=2打印
'说明打印之前要先保存一些数据。
    Dim strRPTName As String
    
    If mbln手术 Then
        strRPTName = "ZL1_INSIDE_1269_2"
    Else
        strRPTName = "ZL1_INSIDE_1269_1"
    End If
    
    If mblnChange Then
        mbln编辑 = True
        If Not CheckData Then Exit Sub
        Call SaveData
    ElseIf mblnInitSaved Then
        If Not CheckData Then Exit Sub
        Call SaveData(True)
    End If
    vsPJB.Redraw = flexRDNone
    If bytType = 1 Then
        Call mobjReport.ReportOpen(gcnOracle, 100, strRPTName, Me, "抽样ID=" & mlng抽样ID, "序号=" & mlng序号, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "ReportFormat=" & IIf(tbcSub.Selected.Index = 0, 1, 2), 1)
    ElseIf bytType = 2 Then
        Call mobjReport.ReportOpen(gcnOracle, 100, strRPTName, Me, "抽样ID=" & mlng抽样ID, "序号=" & mlng序号, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "ReportFormat=" & IIf(tbcSub.Selected.Index = 0, 1, 2), 2)
    End If
    vsPJB.Redraw = flexRDDirect
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsc.Value
    lngMin = vsc.Min
    lngMax = vsc.Max
    
    If KeyCode = vbKeyPageDown Then '下
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '上
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur - (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMin
        End If
    End If
End Sub

Private Sub Form_Activate()
'鼠标滚轮
    Call Form_Resize
    glngPreHWnd = GetWindowLong(picDCB.hwnd, GWL_WNDPROC)
    SetWindowLong picDCB.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
'鼠标滚轮
    SetWindowLong picDCB.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_Resize()
    Call picDCB_Resize
    Call picReasonable_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("当前界面所填内容已经进行了调整尚未保存，是否要继续退出？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'改变打印状态
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    If mbln打印 Then Exit Sub '如果是已经打过的就不再重复执行
    mbln打印 = True
    On Error GoTo errH
    strSQL = Get抽样明细SQL
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optInfo_Click(Index As Integer)
    Dim blnTmp As Boolean
    Select Case Index
    Case e_D_optInfo_检查_病原学检测_未做_2, e_D_optInfo_检查_病原学检测_做_3  '病原学检测 做 未做
        blnTmp = optInfo(e_D_optInfo_检查_病原学检测_未做_2).Value  '未做
        optInfo(e_D_optInfo_检查_病原学检测_检出_4).Enabled = Not blnTmp
        optInfo(e_D_optInfo_检查_病原学检测_未检出_5).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_检查_病原学检测日期_23).Enabled = Not blnTmp '未做
        txtInfo(e_D_txtInfo_检查_病原学检测标本_26).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_检查_病原学检测日期_23).Text = ""
            txtInfo(e_D_txtInfo_检查_病原学检测标本_26).Text = ""
            txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Text = ""
            txtInfo(e_D_txtInfo_检查_病原学检测日期_23).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
            txtInfo(e_D_txtInfo_检查_病原学检测标本_26).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
            txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
        Else
            txtInfo(e_D_txtInfo_检查_病原学检测日期_23).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
            txtInfo(e_D_txtInfo_检查_病原学检测标本_26).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
            txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
        End If
        Call optInfo_Click(4)
    Case e_D_optInfo_检查_病原学检测_检出_4, e_D_optInfo_检查_病原学检测_未检出_5  '检测细菌 未检出 检出
        blnTmp = optInfo(e_D_optInfo_检查_病原学检测_检出_4).Value  '未检出
        txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Text = ""
            txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
        Else
            txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
        End If
    Case e_D_optInfo_检查_药敏试验_未做_6, e_D_optInfo_检查_药敏试验_做_7  ' 药敏试验 做 未做
        blnTmp = optInfo(e_D_optInfo_检查_药敏试验_未做_6).Value  '未做
        optInfo(e_D_optInfo_检查_药敏试验_相符_8).Enabled = Not blnTmp
        optInfo(e_D_optInfo_检查_药敏试验_不相符_9).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_检查_药敏试验日期_24).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_检查_药敏试验日期_24).Text = ""
            txtInfo(e_D_txtInfo_检查_药敏试验日期_24).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
        Else
            txtInfo(e_D_txtInfo_检查_药敏试验日期_24).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
        End If
    Case e_D_optInfo_用药目的_未用_17, e_D_optInfo_用药目的_预防_18, e_D_optInfo_用药目的_治疗_21 '用药目的， 未用药，预防，治疗
        blnTmp = optInfo(e_D_optInfo_用药目的_未用_17).Value  '未用药
        optInfo(e_D_optInfo_用药目的_未用_17).Enabled = blnTmp
        optInfo(e_D_optInfo_用药目的_预防_18).Enabled = Not blnTmp
        optInfo(e_D_optInfo_用药目的_治疗_21).Enabled = Not blnTmp
        cmdInfect(1).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Text = ""
            txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Tag = ""
            txtInfo(e_D_txtInfo_用药目的_感染诊断_36).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
        Else
            txtInfo(e_D_txtInfo_用药目的_感染诊断_36).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
        End If
        
        If Index = e_D_optInfo_用药目的_预防_18 Or Index = e_D_optInfo_用药目的_治疗_21 Then
            blnTmp = optInfo(e_D_optInfo_用药目的_预防_18).Value
            cmdInfect(1).Enabled = Not blnTmp
            txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Enabled = Not blnTmp
            If blnTmp Then
                txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Text = ""
                txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Tag = ""
                txtInfo(e_D_txtInfo_用药目的_感染诊断_36).BackColor = txtInfo(e_F_txtInfo_病历号_3).BackColor
            Else
                txtInfo(e_D_txtInfo_用药目的_感染诊断_36).BackColor = txtInfo(e_D_txtInfo_备注_57).BackColor
            End If
        End If
    End Select
    If Visible Then mblnChange = True
End Sub

Private Sub picReasonable_Resize()
    On Error Resume Next
    lblYJ.Top = 400
    lblYJ.Left = (picReasonable.Width - lblYJ.Width) / 2
    picPjTim.Top = lblYJ.Top + 400
    picPjTim.Left = picReasonable.Width - picPjTim.Width
    
    vsPJB.Move 20, picPjTim.Top + picPjTim.Height, picReasonable.Width, picReasonable.Height - (picPjTim.Top + picPjTim.Height)
End Sub

Private Sub picDCB_Resize()
    Dim lngW As Long
    Dim lngL As Long
    Dim lngT As Long
    
    lngW = 12415
    lngL = 90
    lngT = 5760
    
    On Error Resume Next
    
    lblN.Top = 0
    
    lblN.Left = (picDCB.Width - lblN.Width) / 2
    
    If mbln手术 Then
        picComm(e_D_picComm_用药目的容器_1).Move lngL, lngT, lngW, picComm(e_D_picComm_用药目的容器_1).Height
        picDiff(e_D_picDiff_手术情况容器_1).Move lngL, lngT + picComm(e_D_picComm_用药目的容器_1).Height - 10, lngW, picDiff(e_D_picDiff_手术情况容器_1).Height
        picComm(e_D_picComm_用药明细及费用容器_0).Move lngL, picDiff(e_D_picDiff_手术情况容器_1).Height + picDiff(e_D_picDiff_手术情况容器_1).Top - 10, lngW, picComm(e_D_picComm_用药明细及费用容器_0).Height
    Else
        picDiff(e_D_picDiff_影像And症状容器_0).Move lngL, lngT, lngW, picDiff(e_D_picDiff_影像And症状容器_0).Height
        picComm(e_D_picComm_用药目的容器_1).Move lngL, lngT + picDiff(e_D_picDiff_影像And症状容器_0).Height - 10, lngW, picComm(e_D_picComm_用药目的容器_1).Height
        picComm(e_D_picComm_用药明细及费用容器_0).Move lngL, picComm(e_D_picComm_用药目的容器_1).Height + picComm(e_D_picComm_用药目的容器_1).Top - 10, lngW, picComm(e_D_picComm_用药明细及费用容器_0).Height
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
''''获得焦点
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
    Case e_D_txtInfo_检查_病原学检测日期_23, e_D_txtInfo_检查_药敏试验日期_24, e_D_txtInfo_检查_用药前体温日期_12, e_D_txtInfo_检查_用药前白细胞计数日期_14, e_D_txtInfo_检查_用药前中性粒细胞日期_16, _
    e_D_txtInfo_检查_用药前C反应蛋白日期_18, e_D_txtInfo_检查_用药前丙谷转氨酶日期_20, e_D_txtInfo_检查_用药前肌酐日期_22, e_D_txtInfo_检查_用药后体温日期_54, e_D_txtInfo_检查_用药后白细胞计数日期_52, _
    e_D_txtInfo_检查_用药后中性粒细胞日期_50, e_D_txtInfo_检查_用药后C反应蛋白日期_48, e_D_txtInfo_检查_用药后丙谷转氨酶日期_44, e_D_txtInfo_检查_用药后肌酐日期_42
        If InStr("0123456789/" & Chr(vbKeyBack) & Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Case e_D_txtInfo_检查_用药前体温_11, e_D_txtInfo_检查_用药后体温_55
        If InStr("0123456789." & Chr(vbKeyBack) & Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End Select
    
    If KeyAscii = 13 Then Call ZLCommFun.PressKey(vbKeyTab)
    
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim strTmp As String
    Dim intTmp As Integer
    Dim intMM As Integer
    Dim intDD As Integer
    Dim strDate As String
    Dim strMsg As String
    strTmp = Trim(txtInfo(Index).Text)
    
    If strTmp = "" Then
        txtInfo(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
    Case e_D_txtInfo_检查_病原学检测日期_23, e_D_txtInfo_检查_药敏试验日期_24, e_D_txtInfo_检查_用药前体温日期_12, e_D_txtInfo_检查_用药前白细胞计数日期_14, e_D_txtInfo_检查_用药前中性粒细胞日期_16, _
    e_D_txtInfo_检查_用药前C反应蛋白日期_18, e_D_txtInfo_检查_用药前丙谷转氨酶日期_20, e_D_txtInfo_检查_用药前肌酐日期_22, e_D_txtInfo_检查_用药后体温日期_54, e_D_txtInfo_检查_用药后白细胞计数日期_52, _
    e_D_txtInfo_检查_用药后中性粒细胞日期_50, e_D_txtInfo_检查_用药后C反应蛋白日期_48, e_D_txtInfo_检查_用药后丙谷转氨酶日期_44, e_D_txtInfo_检查_用药后肌酐日期_42
        If InStr(strTmp, "/") = 0 Then
            strMsg = "日期格式不对，正确格式：12/22"
        Else
            intMM = Val(Split(strTmp, "/")(0))
            intDD = Val(Split(strTmp, "/")(1))
            If intMM = 0 Or intMM > 12 Then
                strMsg = "填写的月份不正确，只能是1－12"
            Else
                If InStr(",1,3,5,7,8,10,12,", "," & intMM & ",") > 0 Then
                    intTmp = 31
                ElseIf InStr(",4,6,9,11,", "," & intMM & ",") > 0 Then
                    intTmp = 30
                ElseIf intMM = 2 Then
                    strTmp = Split(txtInfo(e_F_txtInfo_抽样时间_起_0).Text, "-")(0) & "02"
                    intTmp = GetMonthMaxDay(strTmp)
                End If
                
                If intDD > intTmp Or intDD = 0 Then
                    strMsg = "填写的日期号数不正确，" & intMM & "月最共有" & intTmp & "天。"
                End If
                If strMsg = "" Then '日期应该在住院时间范围内
                    strDate = mlngYear & "/" & strTmp
                    strDate = Format(strDate, "YYYY-MM-DD")
                    If Not Between(strDate, txtInfo(e_D_txtInfo_基本_入院时间_9).Text, txtInfo(e_D_txtInfo_基本_出院时间_10).Text) Then
                        strMsg = "填写的日期应在病人住院时间范围内：" & txtInfo(e_D_txtInfo_基本_入院时间_9).Text & "到" & txtInfo(e_D_txtInfo_基本_出院时间_10).Text
                    End If
                End If
            End If
        End If
    End Select
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        txtInfo(Index).Text = DateForShow(txtInfo(Index).Tag)
        Cancel = True
        Call txtInfo_GotFocus(Index)
        Exit Sub
    End If
    txtInfo(Index).Tag = strDate
End Sub

Private Function GetMonthMaxDay(ByVal strMonth As String) As Integer
'功能：2月份的天数

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim intMaxDay As Integer
    On Error GoTo errH
    
    strSQL = "Select 开始日期, 终止日期 From 期间表 Where 期间 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMonth)
    If Not rsTmp.EOF Then
        strTmp = Format(rsTmp!终止日期 & "", "yyyy-mm-dd")
        GetMonthMaxDay = Val(Split(strTmp, "-")(2))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    picDCB.Top = (-1) * vsc.Value * Screen.TwipsPerPixelY + 400
End Sub

Private Sub SetCtlProperty()
'功能：设置控件的可见性，标题等
    Dim objCtl As Object
    
    '界面表格布局
    If mbln手术 Then
        picDiff(e_D_picDiff_手术情况容器_1).Visible = True
        picDiff(e_D_picDiff_影像And症状容器_0).Visible = False
        picComm(e_D_picComm_用药目的容器_1).Visible = True
        picComm(e_D_picComm_用药明细及费用容器_0).Visible = True
        
        lblInfo(e_D_lblInfo_说明_手术说明文本_73).Visible = True
        lblInfo(e_D_lblInfo_说明_非手术说明文本_74).Visible = False
        
        lblInfo(e_D_lblInfo_用药目的_序号_60).Caption = "5"
        lblInfo(e_D_lblInfo_用药情况_序号_45).Caption = "7"
        lblInfo(e_D_lblInfo_费用_序号_46).Caption = "8"
        lblInfo(e_D_lblInfo_结果_序号_47).Caption = "9"
        lblInfo(e_D_lblInfo_本院评价_序号_64).Caption = "10"
        lblInfo(e_D_lblInfo_中心评价_序号_65).Caption = "11"
        lblInfo(e_D_lblInfo_备注_序号_66).Caption = "12"
        lblInfo(e_D_lblInfo_说明_序号_69).Caption = "13"
        vsc.Max = 300
    Else
        picDiff(e_D_picDiff_手术情况容器_1).Visible = False
        picDiff(e_D_picDiff_影像And症状容器_0).Visible = True
        picComm(e_D_picComm_用药目的容器_1).Visible = True
        picComm(e_D_picComm_用药明细及费用容器_0).Visible = True
        
        chkInfo(e_D_chkInfo_评价_本院_术前_23).Visible = False
        chkInfo(e_D_chkInfo_评价_本院_术中_24).Visible = False
        chkInfo(e_D_chkInfo_评价_本院_术后_25).Visible = False
    
        chkInfo(e_D_chkInfo_评价_中心_术前_28).Visible = False
        chkInfo(e_D_chkInfo_评价_中心_术中_27).Visible = False
        chkInfo(e_D_chkInfo_评价_中心_术后_26).Visible = False
        
        lblInfo(e_D_lblInfo_评价_本院_围手术标签_59).Visible = False
        lblInfo(e_D_lblInfo_评价_中心_围手术标签_70).Visible = False
        
        lblInfo(e_D_lblInfo_说明_手术说明文本_73).Visible = False
        lblInfo(e_D_lblInfo_说明_非手术说明文本_74).Visible = True
        
        lblN.Caption = "非手术病人抗菌药物使用情况调查表"
        lblInfo(e_F_lblInfo_出院人数标签_4).Caption = "非手术病人出院人数："
        lblYJ.Caption = "非手术用药合理性评价意见表"
        vsc.Max = 250
    End If
    
    '控件背景色
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
            Case "OPTIONBUTTON", "CHECKBOX", "LABEL"
                objCtl.BackColor = picDCB.BackColor
        End Select
    Next
    '容器初始高度
    picDCB.Top = 400
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    Dim rs诊断 As ADODB.Recordset
    Dim lng感染诊断ID As Long
    Dim str感染诊断 As String
    Dim strTmp As String
    Dim str诊断 As String
    Dim lngTmp As Long
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select a.治疗结果, a.适应症, a.药物选择, a.单次剂量, a.每日给药频次, a.溶剂, a.给药途径, a.用药疗程, a.术前用药时间, a.术中用药, a.术后用药, a.联合用药, a.更换药物, a.备注," & vbNewLine & _
        "a.是否打印, a.是否编辑, a.用药天数, a.抗菌药种数, a.是否手术, a.病原学检测, a.病原学检测日期, a.病原学检测标本, a.病原学检测检出细菌名, a.药敏试验, a.药敏试验日期," & vbNewLine & _
        "a.药敏试验是否相符, a.用药前体温, a.用药前白细胞计数, a.用药前中性粒细胞, a.用药前c反应蛋白, a.用药前丙谷转氨酶, a.用药前肌酐, a.用药后体温, a.用药后白细胞计数, a.用药后中性粒细胞," & vbNewLine & _
        "a.用药后c反应蛋白, a.用药后丙谷转氨酶, a.用药后肌酐, a.用药前体温日期, a.用药前白细胞计数日期, a.用药前中性粒细胞日期, a.用药前c反应蛋白日期, a.用药前丙谷转氨酶日期, a.用药前肌酐日期," & vbNewLine & _
        "a.用药后体温日期, a.用药后白细胞计数日期, a.用药后中性粒细胞日期, a.用药后c反应蛋白日期, a.用药后丙谷转氨酶日期, a.用药后肌酐日期, a.影像学诊断, a.影像学诊断部位, a.影像学诊断结论," & vbNewLine & _
        "a.临床症状, a.用药目的, a.感染诊断,a.是否用抗真菌药,b.性别,b.年龄,b.住院号 as 病历号,b.入院日期 as 入院时间,b.出院日期 as 出院时间,b.体重" & vbNewLine & _
        "From 抗菌药物抽样明细 A,病案主页 b" & vbNewLine & _
        "Where a.病人id=b.病人id and a.主页id=b.主页id and a.抽样id = [1] And a.序号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng抽样ID, mlng序号)
    
    '保存过一次之后，该字段要么是1或0，如果从未保存过这个字段为Null
    mblnInitSaved = rsTmp!是否手术 & "" = ""
    
    txtInfo(e_F_txtInfo_病历号_3).Text = rsTmp!病历号 & ""
    
    txtInfo(e_D_txtInfo_基本_性别_6).Text = rsTmp!性别 & ""
    txtInfo(e_D_txtInfo_基本_年龄_7).Text = rsTmp!年龄 & ""
    
    '体重信息
    If Val(rsTmp!体重 & "") > 0 Then
        txtInfo(e_D_txtInfo_基本_体重_8).Text = Val(rsTmp!体重 & "") & "Kg"
    Else
        strSQL = "Select b.项目单位 as 单位, b.项目名称 as 信息名, b.记录内容 as 信息值" & _
            " From 病人护理记录 A, 病人护理内容 B Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2] and b.项目名称='体重'"
        Set rsOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsOther.EOF Then txtInfo(e_D_txtInfo_基本_体重_8).Text = rsOther!信息值 & rsOther!单位 & ""
        Set rsOther = Nothing
    End If
    
    txtInfo(e_D_txtInfo_基本_入院时间_9).Text = Format(rsTmp!入院时间 & "", "yyyy-MM-dd")
    txtInfo(e_D_txtInfo_基本_入院时间_9).Tag = Format(rsTmp!入院时间 & "", "yyyy-MM-dd HH:mm")
    txtInfo(e_D_txtInfo_基本_出院时间_10).Text = Format(rsTmp!出院时间 & "", "yyyy-MM-dd")
    txtInfo(e_D_txtInfo_基本_出院时间_10).Tag = Format(rsTmp!出院时间 & "", "yyyy-MM-dd HH:mm")
    
    mlngYear = Val(Split(txtInfo(e_D_txtInfo_基本_入院时间_9).Text, "-")(0))
    
    optInfo(e_D_optInfo_检查_病原学检测_未做_2).Value = Val(rsTmp!病原学检测 & "") = 0
    optInfo(e_D_optInfo_检查_病原学检测_未做_2 + 1).Value = Val(rsTmp!病原学检测 & "") = 1
    
    If rsTmp!病原学检测日期 & "" <> "" Then
        strTmp = Format(rsTmp!病原学检测日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_病原学检测日期_23).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_病原学检测日期_23).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_病原学检测标本_26).Text = rsTmp!病原学检测标本 & ""
    
    optInfo(e_D_optInfo_检查_病原学检测_检出_4).Value = True
    optInfo(e_D_optInfo_检查_病原学检测_检出_4 + 1).Value = False
    If rsTmp!病原学检测检出细菌名 & "" <> "" Then
        optInfo(e_D_optInfo_检查_病原学检测_检出_4).Value = False
        optInfo(e_D_optInfo_检查_病原学检测_检出_4 + 1).Value = True
        txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Text = rsTmp!病原学检测检出细菌名 & ""
    End If
    optInfo(e_D_optInfo_检查_药敏试验_未做_6).Value = Val(rsTmp!药敏试验 & "") = 0
    optInfo(e_D_optInfo_检查_药敏试验_未做_6 + 1).Value = Val(rsTmp!药敏试验 & "") = 1
    
    If rsTmp!药敏试验日期 & "" <> "" Then
        strTmp = Format(rsTmp!药敏试验日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_药敏试验日期_24).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_药敏试验日期_24).Tag = strTmp
    End If
    
    optInfo(e_D_optInfo_检查_药敏试验_相符_8).Value = Val(rsTmp!药敏试验是否相符 & "") = 0
    optInfo(e_D_optInfo_检查_药敏试验_相符_8 + 1).Value = Val(rsTmp!药敏试验是否相符 & "") = 1
    
    txtInfo(e_D_txtInfo_检查_用药前体温_11).Text = rsTmp!用药前体温 & ""
    If rsTmp!用药前体温日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药前体温日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药前体温日期_12).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药前体温日期_12).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药前白细胞计数_13).Text = rsTmp!用药前白细胞计数 & ""
    If rsTmp!用药前白细胞计数日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药前白细胞计数日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药前白细胞计数日期_14).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药前白细胞计数日期_14).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药前中性粒细胞_15).Text = rsTmp!用药前中性粒细胞 & ""
    If rsTmp!用药前中性粒细胞日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药前中性粒细胞日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药前中性粒细胞日期_16).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药前中性粒细胞日期_16).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药前C反应蛋白_17).Text = rsTmp!用药前c反应蛋白 & ""
    If rsTmp!用药前c反应蛋白日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药前c反应蛋白日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药前C反应蛋白日期_18).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药前C反应蛋白日期_18).Tag = strTmp
    End If
     
    txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶_19).Text = rsTmp!用药前丙谷转氨酶 & ""
    If rsTmp!用药前丙谷转氨酶日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药前丙谷转氨酶日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶日期_20).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶日期_20).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药前肌酐_21).Text = rsTmp!用药前肌酐 & ""
    If rsTmp!用药前肌酐日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药前肌酐日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药前肌酐日期_22).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药前肌酐日期_22).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药后体温_55).Text = rsTmp!用药后体温 & ""
    If rsTmp!用药后体温日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药后体温日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药后体温日期_54).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药后体温日期_54).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药后白细胞计数_53).Text = rsTmp!用药后白细胞计数 & ""
    If rsTmp!用药后白细胞计数日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药后白细胞计数日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药后白细胞计数日期_52).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药后白细胞计数日期_52).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药后中性粒细胞_51).Text = rsTmp!用药后中性粒细胞 & ""
    If rsTmp!用药后中性粒细胞日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药后中性粒细胞日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药后中性粒细胞日期_50).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药后中性粒细胞日期_50).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药后C反应蛋白_49).Text = rsTmp!用药后C反应蛋白 & ""
    If rsTmp!用药后C反应蛋白日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药后C反应蛋白日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药后C反应蛋白日期_48).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药后C反应蛋白日期_48).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶_45).Text = rsTmp!用药后丙谷转氨酶 & ""
    If rsTmp!用药后丙谷转氨酶日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药后丙谷转氨酶日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶日期_44).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶日期_44).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_检查_用药后肌酐_43).Text = rsTmp!用药后肌酐 & ""
    If rsTmp!用药后肌酐日期 & "" <> "" Then
        strTmp = Format(rsTmp!用药后肌酐日期, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_检查_用药后肌酐日期_42).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_检查_用药后肌酐日期_42).Tag = strTmp
    End If
    
    strTmp = rsTmp!影像学诊断 & ""
    If Len(strTmp) = 3 Then
        chkInfo(e_D_chkInfo_影像_X线_18).Value = Mid(strTmp, 1, 1)
        chkInfo(e_D_chkInfo_影像_CT_19).Value = Mid(strTmp, 2, 1)
        chkInfo(e_D_chkInfo_影像_磁共振_20).Value = Mid(strTmp, 3, 1)
    End If
    
    txtInfo(e_D_txtInfo_影像_部位_46).Text = rsTmp!影像学诊断部位 & ""
    txtInfo(e_D_txtInfo_影像_结论_47).Text = rsTmp!影像学诊断结论 & ""
    
    '用药目的
    lngTmp = Val(rsTmp!用药目的 & "")
    optInfo(e_D_optInfo_用药目的_未用_17 + lngTmp).Value = True
    
    '用药天数
    txtInfo(e_D_txtInfo_用药情况_天数_30).Text = Val(rsTmp!用药天数 & "")
    
    '抗菌药总数
    txtInfo(e_D_txtInfo_用药情况_种数_29).Text = Val(rsTmp!抗菌药种数 & "")
    
    optInfo(e_D_optInfo_结果_真菌药_有_20).Value = Val(rsTmp!是否用抗真菌药 & "") <> 0
    
    '用药评价
    strTmp = rsTmp!适应症 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_适应症_5).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_适应症_37).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!药物选择 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_药物选择_12).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_药物选择_36).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!单次剂量 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_单次剂量_13).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_单次剂量_35).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!每日给药频次 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_每日给药频次_14).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_每日给药频次_34).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!溶剂 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_溶剂_15).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_溶剂_33).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!给药途径 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_给药途径_16).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_给药途径_32).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!用药疗程 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_用药疗程_17).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_用药疗程_31).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!术前用药时间 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_术前_23).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_术前_28).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!术中用药 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_术中_24).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_术中_27).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!术后用药 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_术后_25).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_术后_26).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!联合用药 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_联合用药_22).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_联合用药_29).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!更换药物 & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_评价_本院_更换药物_21).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_评价_中心_更换药物_30).Value = Val(Split(strTmp, "|")(1))
    End If
    
    txtInfo(e_D_txtInfo_备注_57).Text = rsTmp!备注 & ""
    
    '过敏史
    strSQL = "Select 药物名 From 病人过敏记录 Where 病人id = [1] And 主页id = [2]"
    Set rsOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    optInfo(e_D_optInfo_过敏_无_0).Value = rsOther.EOF
    optInfo(e_D_optInfo_过敏_无_0 + 1).Value = Not rsOther.EOF
    
    If Not rsOther.EOF Then
        For i = 1 To rsOther.RecordCount
            strTmp = strTmp & "," & rsOther!药物名
            rsOther.MoveNext
        Next
        txtInfo(e_D_txtInfo_过敏_通用名_34).Text = Mid(strTmp, 2): strTmp = ""
    End If
    
'----------------------------------------------------------------------------------------------------------------------------------------------
    
    '诊断 先取出所有诊断 取首页整理中的诊断
    strSQL = "Select ID,记录来源,诊断描述,诊断类型,出院情况 From 病人诊断记录 Where 病人id = [1] and 主页id =[2]  And NVL(编码序号,1) = 1 order by 记录来源,诊断次序"
    Set rs诊断 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    '1.诊断  首页中的所有出院诊断
    rs诊断.Filter = "(诊断类型=3 and 记录来源=3) or (诊断类型=13 and 记录来源=3)"
    For i = 1 To rs诊断.RecordCount
        str诊断 = str诊断 & "," & i & "." & rs诊断!诊断描述
        If InStr("," & strTmp & ",", "," & rs诊断!出院情况 & ",") = 0 Then
            strTmp = strTmp & "," & rs诊断!出院情况
        End If
        rs诊断.MoveNext
    Next
    '诊断
    txtInfo(e_D_txtInfo_诊断_诊断_27).Text = IIf("" = str诊断, "无", Mid(str诊断, 2))
    '治疗结果
    optInfo(e_D_optInfo_结果_好转_13).Value = False
    optInfo(e_D_optInfo_结果_治愈_12).Value = False
    optInfo(e_D_optInfo_结果_无效_14).Value = False
    If mbln编辑 Then
        If InStr(",1,2,3,", "," & rsTmp!治疗结果 & ",") > 0 Then
            lngTmp = 11 + Val(rsTmp!治疗结果 & "")
            optInfo(lngTmp).Value = True
        End If
    Else
        If InStr(strTmp, "好转") > 0 Then
            optInfo(e_D_optInfo_结果_好转_13).Value = True
        ElseIf InStr(strTmp, "治愈") > 0 Then
            optInfo(e_D_optInfo_结果_治愈_12).Value = True
        ElseIf InStr(strTmp, "治愈") > 0 Then
            optInfo(e_D_optInfo_结果_无效_14).Value = True
        End If
    End If
    strTmp = ""
    
    '感染诊断   有无
    rs诊断.Filter = "诊断类型=5"
    optInfo(e_D_optInfo_结果_感染_无_16).Value = rs诊断.EOF
    optInfo(e_D_optInfo_结果_感染_有_15).Value = Not rs诊断.EOF
    If Not rs诊断.EOF Then
        str感染诊断 = rs诊断!诊断描述 & ""
        lng感染诊断ID = Val(rs诊断!ID & "")
    End If
    
    '先默认设置，再跟据是否编辑进行设置
    '临床症状-与感染有关'用药目的－感染诊断
    If mbln编辑 Then
        lng感染诊断ID = 0: str感染诊断 = ""
        lng感染诊断ID = Val(rsTmp!临床症状 & "")
        rs诊断.Filter = "id=" & lng感染诊断ID
        If Not rs诊断.EOF Then
            str感染诊断 = rs诊断!诊断描述 & ""
            lng感染诊断ID = Val(rs诊断!ID & "")
        End If
        txtInfo(e_D_txtInfo_症状_感染症状_35).Tag = lng感染诊断ID
        txtInfo(e_D_txtInfo_症状_感染症状_35).Text = str感染诊断
        lng感染诊断ID = 0: str感染诊断 = ""
        
        lng感染诊断ID = Val(rsTmp!感染诊断 & "")
        rs诊断.Filter = "id=" & lng感染诊断ID
        If Not rs诊断.EOF Then
            str感染诊断 = rs诊断!诊断描述 & ""
            lng感染诊断ID = Val(rs诊断!ID & "")
        End If
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Tag = lng感染诊断ID
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Text = str感染诊断
        If lng感染诊断ID <> 0 Then lblInfo(e_D_lblInfo_用药目的_感染诊断_68).Tag = lng感染诊断ID & "," & str感染诊断
    Else
        '未编辑过就默认指定一个
        txtInfo(e_D_txtInfo_症状_感染症状_35).Tag = lng感染诊断ID
        txtInfo(e_D_txtInfo_症状_感染症状_35).Text = str感染诊断
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Tag = lng感染诊断ID
        txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Text = str感染诊断
        If lng感染诊断ID <> 0 Then lblInfo(e_D_lblInfo_用药目的_感染诊断_68).Tag = lng感染诊断ID & "," & str感染诊断
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadFee()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Sum(Decode(Nvl(e.抗生素, 0), 0, 0, a.结帐金额)) As 抗菌药费," & vbNewLine & _
        "Sum(Decode(a.收费类别, '5', a.结帐金额, '6', a.结帐金额, '7', a.结帐金额, 0)) As 总药费, Sum(a.结帐金额) As 住院费用" & vbNewLine & _
        "From 住院费用记录 A, 药品规格 D, 药品特性 E" & vbNewLine & _
        "Where a.病人id = [1] And a.主页id = [2] and a.记录状态<>0 And a.收费细目id = d.药品id(+) And d.药名id = e.药名id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        txtInfo(e_D_txtInfo_费用_总费_31).Text = Format(Val(rsTmp!住院费用 & ""), "0.00")
        txtInfo(e_D_txtInfo_费用_药费_32).Text = Format(Val(rsTmp!总药费 & ""), "0.00")
        txtInfo(e_D_txtInfo_费用_抗药费_33).Text = Format(Val(rsTmp!抗菌药费 & ""), "0.00")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDrugUse()
'功能：加载用药情况和费用，先从现有数据中读一次，如果没有数据再从医嘱记录中取
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim rs用量 As ADODB.Recordset
    Dim str医嘱IDs As String
    Dim arrTmp As Variant
    Dim str抗菌药名IDs As String
    Dim str天数 As String
    Dim str用药目的 As String
    Dim lngRow As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim dblTmp As Double
    Dim dbl首次用量 As Double
    Dim dbl单次用量 As Double
    Dim dbl总量 As Double
    Dim lng相关ID As Long
    Dim lng次数 As Long
    Dim int用药目的 As Integer
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = "select 药名id,图标,药名,单量,频次,途径,总量,起止日期 from 抗菌药物抽样用药 where 抽样id=[1] and 序号=[2] order by 药品序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng抽样ID, mlng序号)
    
    With vsDrugUse
        .Redraw = False
        .Rows = .FixedRows
        lng相关ID = 1
        For i = 1 To rsTmp.RecordCount
            .AddItem "": lngRow = .Rows - 1
            
            If rsTmp!图标 & "" <> "" Then
                If rsTmp!图标 & "" = "┏" Then lng相关ID = lng相关ID + 1
            Else
                lng相关ID = lng相关ID + 1
            End If
            
            .TextMatrix(lngRow, COL_DRUG_相关ID) = lng相关ID
            .TextMatrix(lngRow, COL_DRUG_药名ID) = Val(rsTmp!药名id & "")
            .TextMatrix(lngRow, COL_DRUG_图标) = rsTmp!图标 & ""
            .TextMatrix(lngRow, COL_DRUG_药品名称) = rsTmp!药名 & ""
            .TextMatrix(lngRow, COL_DRUG_单次剂量) = rsTmp!单量 & ""
            .TextMatrix(lngRow, COL_DRUG_给药频次) = rsTmp!频次 & ""
            .TextMatrix(lngRow, COL_DRUG_途径) = rsTmp!途径 & ""
            .TextMatrix(lngRow, COL_DRUG_总用量) = rsTmp!总量 & ""
            .TextMatrix(lngRow, COL_DRUG_起止时间) = rsTmp!起止日期 & ""
            rsTmp.MoveNext
        Next
        .Redraw = True
    End With
    
    If rsTmp.RecordCount > 0 Then Exit Sub
    Set rsTmp = Nothing
    
    strSQL = "Select d.Id,d.相关id,e.Id As 药名id,Decode(e.Id,b.药名id,b.抗生素,0) As 抗生素,d.医嘱内容 As 药品名称,c.医嘱内容 As 途径," & vbNewLine & _
        "d.首次用量, d.单次用量,d.执行频次 As 给药频次, e.计算单位 As 单位, d.用药目的,to_char(d.开始执行时间, 'MM-DD HH24:MI')||' - '||" & vbNewLine & _
        "to_char(Nvl(Nvl(d.上次执行时间,d.执行终止时间),d.停嘱时间),'MM-DD HH24:MI') As 起止时间" & vbNewLine & _
        "From 病人医嘱记录 A, 药品特性 B, 病人医嘱记录 C, 病人医嘱记录 D, 诊疗项目目录 E" & vbNewLine & _
        "Where a.诊疗项目id = b.药名id And a.相关id = c.Id And a.诊疗类别 = '5' And Nvl(b.抗生素, 0) <> 0 And c.Id = d.相关id And d.诊疗项目id = e.Id And" & vbNewLine & _
        "d.诊疗类别 In ('5','6') And a.医嘱状态 in (8,9) And a.病人id =[1] And a.主页id =[2] Order By d.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.EOF Then Exit Sub
    For i = 1 To rsTmp.RecordCount
        If InStr("," & str医嘱IDs & ",", "," & rsTmp!相关ID & ",") = 0 Then
            str医嘱IDs = str医嘱IDs & "," & rsTmp!相关ID
        End If
        rsTmp.MoveNext
    Next
    str医嘱IDs = Mid(str医嘱IDs, 2)
    rsTmp.MoveFirst
    
    strSQL = "select id as 组医嘱id,Zl_Adviceexetimes(Id,开始执行时间,Nvl(Nvl(上次执行时间,执行终止时间),停嘱时间)," & _
        "执行时间方案,开始执行时间,开始执行时间-1,频率间隔,间隔单位,医嘱期效) as 分解时间 From 病人医嘱记录 where Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    'rs用量 这个集录记用来算医嘱的总量和用药天数
    Set rs用量 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str医嘱IDs)
    
    With vsDrugUse
        .Redraw = False
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .AddItem "": lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_DRUG_ID) = Val(rsTmp!ID & "")
            .TextMatrix(lngRow, COL_DRUG_相关ID) = Val(rsTmp!相关ID & "")
            .TextMatrix(lngRow, COL_DRUG_药名ID) = Val(rsTmp!药名id & "")
            .TextMatrix(lngRow, COL_DRUG_药品名称) = rsTmp!药品名称 & ""
            .TextMatrix(lngRow, COL_DRUG_给药频次) = rsTmp!给药频次 & ""
            .TextMatrix(lngRow, COL_DRUG_途径) = rsTmp!途径 & ""
            .TextMatrix(lngRow, COL_DRUG_起止时间) = rsTmp!起止时间 & ""
            strTmp = ""
            dblTmp = Val(rsTmp!首次用量 & "")
            dbl首次用量 = dblTmp
            If dblTmp > 0 Then
                If Mid(dblTmp, 1, 1) = "." Then
                    strTmp = "0" & dblTmp
                Else
                    strTmp = dblTmp
                End If
            End If
            
            dblTmp = Val(rsTmp!单次用量 & "")
            dbl单次用量 = dblTmp
            If dblTmp > 0 Then
                If Mid(dblTmp, 1, 1) = "." Then
                    strTmp = IIf(strTmp = "", "", strTmp & ":") & "0" & dblTmp
                Else
                    strTmp = IIf(strTmp = "", "", strTmp & ":") & dblTmp
                End If
            End If
            
            .TextMatrix(lngRow, COL_DRUG_单次剂量) = strTmp & rsTmp!单位: strTmp = "": dblTmp = 0
            
            If lng相关ID <> Val(rsTmp!相关ID & "") Then
                lng相关ID = Val(rsTmp!相关ID & "")
                '总量计算，确定次数和天数
                rs用量.Filter = "组医嘱id=" & Val(rsTmp!相关ID & "")
                If Not rs用量.EOF Then
                    strTmp = rs用量!分解时间 & ""
                    lng次数 = 0
                    If strTmp <> "" Then
                        If InStr(strTmp, ",") = 0 Then
                            strTmp = Format(rs用量!分解时间, "YYYY-MM-DD HH:MM:SS")
                        End If
                        arrTmp = Split(strTmp, ",")
                        
                        lng次数 = UBound(arrTmp) + 1
                        
                        For j = 0 To UBound(arrTmp)
                            strTmp = Format(arrTmp(j), "YYYY-MM-DD")
                            If InStr("," & str天数 & ",", "," & strTmp & ",") = 0 Then
                                str天数 = str天数 & "," & strTmp
                            End If
                        Next
                    End If
                End If
            End If
            
            strTmp = "0"
            If lng次数 > 0 Then
                If dbl首次用量 <> 0 Then
                    dbl总量 = dbl首次用量 + dbl单次用量 * (lng次数 - 1)
                Else
                    dbl总量 = dbl单次用量 * lng次数
                End If
                
                If Mid(dbl总量, 1, 1) = "." Then
                    strTmp = "0" & dbl总量
                Else
                    strTmp = dbl总量
                End If
            End If
            
            .TextMatrix(lngRow, COL_DRUG_总用量) = strTmp & rsTmp!单位: strTmp = "": dblTmp = 0
             
             
            If InStr("," & str抗菌药名IDs & ",", "," & rsTmp!药名id & ",") = 0 And Val(rsTmp!抗生素 & "") <> 0 Then
                str抗菌药名IDs = str抗菌药名IDs & "," & rsTmp!药名id
            End If
            
            If InStr("," & str用药目的 & ",", "," & Val(rsTmp!用药目的 & "") & ",") = 0 And Val(rsTmp!用药目的 & "") <> 0 Then
                str用药目的 = str用药目的 & "," & Val(rsTmp!用药目的 & "")
            End If
            rsTmp.MoveNext
        Next
        
        '用药天数
        str天数 = Mid(str天数, 2)
        If str天数 <> "" Then txtInfo(e_D_txtInfo_用药情况_天数_30).Text = UBound(Split(str天数, ",")) + 1
        
        '抗菌药总数
        str抗菌药名IDs = Mid(str抗菌药名IDs, 2)
        If str抗菌药名IDs <> "" Then txtInfo(e_D_txtInfo_用药情况_种数_29).Text = UBound(Split(str抗菌药名IDs, ",")) + 1
        
        '用药目的
        str用药目的 = Mid(str用药目的, 2)
        If str用药目的 <> "" Then
            If InStr("," & str用药目的 & ",", ",1,") > 0 And InStr("," & str用药目的 & ",", ",2,") = 0 Then
                int用药目的 = 1 '预防
            ElseIf InStr("," & str用药目的 & ",", ",1,") = 0 And InStr("," & str用药目的 & ",", ",2,") > 0 Then
                int用药目的 = 2 '治疗
            End If
        End If
        
        optInfo(e_D_optInfo_用药目的_未用_17 + 1).Value = True
        If int用药目的 <> 0 Then optInfo(e_D_optInfo_用药目的_未用_17 + int用药目的).Value = True
        '加图标
        For i = 1 To .Rows - 1
            Call Get一并给药范围(Val(.TextMatrix(i, COL_DRUG_相关ID)), lngBegin, lngEnd)
            For j = lngBegin To lngEnd
                If lngBegin <> lngEnd Then
                    If j = lngBegin Then
                        .TextMatrix(j, 0) = "┏"
                    ElseIf j = lngEnd Then
                        .TextMatrix(j, 0) = "┗"
                    Else
                        .TextMatrix(j, 0) = "┃"
                    End If
                End If
            Next
        Next
        .Redraw = True
    End With
    Call Insert用药
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Insert用药()
'功能：插入用药情况
    Dim blnTrans As Boolean
    Dim blnDo As Boolean
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    arrSQL = Array()
    strSQL = "Zl_抗菌药物抽样用药_Delete(" & mlng抽样ID & "," & mlng序号 & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    With vsDrugUse
        j = 1
        For i = .FixedRows To .Rows - 1
            blnDo = True
            strSQL = "Zl_抗菌药物抽样用药_Insert(" & mlng抽样ID & "," & mlng序号 & "," & j & "," & Val(.TextMatrix(i, COL_DRUG_药名ID)) & "," & _
                IIf(.TextMatrix(i, COL_DRUG_图标) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_图标) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_药品名称) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_药品名称) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_单次剂量) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_单次剂量) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_给药频次) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_给药频次) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_途径) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_途径) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_总用量) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_总用量) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_起止时间) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_起止时间) & "'") & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            j = j + 1
        Next
    End With
    If blnDo Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTable(ByRef vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 1
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
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub LoadOperate()
'功能：加载手术情况 先从现有数据中读一次，如果没有数据再从 病人手麻记录 中取
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim str用药期间 As String
    Dim lngRow As Long
    Dim i As Long
    
    On Error GoTo errH

    strSQL = "select 编码,名称 from 抗菌预防用药期间 order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        strTmp = ""
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "|#" & rsTmp!编码 & ";" & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If
    str用药期间 = Mid(strTmp, 2)
    
    With vsOperate
        .Rows = .FixedRows
        .ColData(COL_OPE_切口) = "Ⅰ|Ⅱ|Ⅲ"
        .ColData(COL_OPE_给药情况) = "未追加|已追加"
        .ColData(COL_OPE_用药期间) = str用药期间
    End With
    
    strSQL = "Select a.手术id,a.手术名称,a.切口,b.名称 as 用药期间,a.开始时间," & vbNewLine & _
        " a.结束时间,decode(nvl(a.给药情况,0),0,'未追加',1,'已追加',null) as 给药情况,b.编码" & vbNewLine & _
        " From 抗菌药物抽样手术 A,抗菌预防用药期间 b" & vbNewLine & _
        " Where a.预防用药期间 = b.编码(+) And a.抽样id = [1] And a.序号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng抽样ID, mlng序号)
    
    If rsTmp.EOF Then
        Set rsTmp = Nothing
        strSQL = "Select a.Id As 手术id, a.已行手术 As 手术名称, a.切口,Null As 用药期间, a.手术开始时间 As 开始时间,a.手术结束时间 As 结束时间,'未追加' As 给药情况,null as 编码" & vbNewLine & _
            "From 病人手麻记录 A Where a.病人id =[1] And a.主页id =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    End If
    
    With vsOperate
        For i = 1 To rsTmp.RecordCount
            .AddItem "": lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_OPE_手术ID) = Val(rsTmp!手术id & "")
            .TextMatrix(lngRow, COL_OPE_手术名称) = rsTmp!手术名称 & ""
            .TextMatrix(lngRow, COL_OPE_切口) = rsTmp!切口 & ""
            If rsTmp!开始时间 & "" <> "" Then
                .TextMatrix(lngRow, COL_OPE_开始时间) = Format(rsTmp!开始时间, "YYYY-MM-DD HH:MM")
            End If
            If rsTmp!结束时间 & "" <> "" Then
                .TextMatrix(lngRow, COL_OPE_结束时间) = Format(rsTmp!结束时间, "YYYY-MM-DD HH:MM")
            End If
            .TextMatrix(lngRow, COL_OPE_用药期间) = rsTmp!用药期间 & ""
            .Cell(flexcpData, lngRow, COL_OPE_用药期间) = Val(rsTmp!编码 & "")
            .TextMatrix(lngRow, COL_OPE_给药情况) = rsTmp!给药情况 & ""
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Get一并给药范围(ByVal lng相关ID As Long, lngBegin As Long, lngEnd As Long)
'功能：根据相关的给药途径医嘱ID,确定一并给药的一组药品的起止行号
    Dim i As Long
    lngBegin = vsDrugUse.FindRow(CStr(lng相关ID), , COL_DRUG_相关ID)
    lngEnd = lngBegin
    If lngBegin = -1 Then Exit Sub
    For i = lngBegin To vsDrugUse.Rows - 1
        If Not vsDrugUse.RowHidden(i) Then
            If Val(vsDrugUse.TextMatrix(i, COL_DRUG_相关ID)) = lng相关ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Function SaveData(Optional ByVal blnFirst As Boolean) As Boolean
'参数：blnFirst ＝true 从未做过修改至少保存一次，保存一部份初始值 =false正常的修改的保存
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim strTmp As String
    Dim str项目编码 As String
    Dim str项目值 As String
    Dim lngTmp As String
    Dim strResult As String
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    Dim i As Integer, j As Integer
    Dim blnInit As Boolean
    
    blnInit = mblnInitSaved And blnFirst
    
    On Error GoTo errH
    If Not blnFirst Then mbln编辑 = True      '修改后被保存则认为是已经编辑过的
    strSQL = Get抽样明细SQL
    arrSQL = Array()
    If mblnChange Or blnInit Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    
    '手术情况
    If mbln手术 And (mblnChange Or blnInit) Then
        strSQL = "Zl_抗菌药物抽样手术_Delete(" & mlng抽样ID & "," & mlng序号 & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        With vsOperate
            For i = .FixedRows To .Rows - 1
                lngTmp = Val(.Cell(flexcpData, i, COL_OPE_用药期间))
                strSQL = "Zl_抗菌药物抽样手术_Insert(" & mlng抽样ID & "," & mlng序号 & "," & Val(.TextMatrix(i, COL_OPE_手术ID)) & ",'" & .TextMatrix(i, COL_OPE_手术名称) & "'," & _
                   IIf(.TextMatrix(i, COL_OPE_切口) = "", "NULL", "'" & .TextMatrix(i, COL_OPE_切口) & "'") & "," & _
                   IIf(.TextMatrix(i, COL_OPE_开始时间) = "", "NULL", "to_date('" & .TextMatrix(i, COL_OPE_开始时间) & "','YYYY-MM-DD HH24:MI')") & "," & _
                   IIf(.TextMatrix(i, COL_OPE_结束时间) = "", "NULL", "to_date('" & .TextMatrix(i, COL_OPE_结束时间) & "','YYYY-MM-DD HH24:MI')") & "," & _
                   IIf(lngTmp = 0, "NULL", lngTmp) & "," & IIf(.TextMatrix(i, COL_OPE_给药情况) = "已追加", 1, 0) & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Next
        End With
    End If
    
    '用药评价
    If mbln用药评价 And (mblnChange Or blnInit) Then
        strSQL = "Zl_抗菌药物抽样评价_Delete(" & mlng抽样ID & "," & mlng序号 & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        With vsPJB
            For i = .FixedRows To .Rows - 1
                
                str项目编码 = .TextMatrix(i, COL_合理编码)
                If Val(str项目编码) = 0 Then str项目编码 = "0"
                
                str项目值 = .TextMatrix(i, COL_合理本院)
                If Not .Cell(flexcpPicture, i, COL_合理本院) Is Nothing Then '合理   本院
                    str项目值 = Val(.Cell(flexcpData, i, COL_合理本院))
                End If
                
                strSQL = "Zl_抗菌药物抽样评价_Insert(" & mlng抽样ID & "," & mlng序号 & ",0,'" & str项目编码 & "'," & i & ",1,'" & str项目值 & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
    
                str项目编码 = .TextMatrix(i, COL_合理编码)
                If Val(str项目编码) = 0 Then str项目编码 = "0"
                
                str项目值 = .TextMatrix(i, COL_合理中心)
                If Not .Cell(flexcpPicture, i, COL_合理中心) Is Nothing Then '合理   中心
                    str项目值 = Val(.Cell(flexcpData, i, COL_合理中心))
                End If
                
                strSQL = "Zl_抗菌药物抽样评价_Insert(" & mlng抽样ID & "," & mlng序号 & ",1,'" & str项目编码 & "'," & i & ",1,'" & str项目值 & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                
                str项目编码 = .TextMatrix(i, COL_不合理编码)
                If Val(str项目编码) = 0 Then str项目编码 = "0"
                
                str项目值 = .TextMatrix(i, COL_不合理本院)
                If Not .Cell(flexcpPicture, i, COL_不合理本院) Is Nothing Then '不合理   本院
                    str项目值 = Val(.Cell(flexcpData, i, COL_不合理本院))
                End If
                
                strSQL = "Zl_抗菌药物抽样评价_Insert(" & mlng抽样ID & "," & mlng序号 & ",0,'" & str项目编码 & "'," & i & ",0,'" & str项目值 & "')"
    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
    
                str项目编码 = .TextMatrix(i, COL_不合理编码)
                If Val(str项目编码) = 0 Then str项目编码 = "0"
                
                str项目值 = .TextMatrix(i, COL_不合理中心)
                If Not .Cell(flexcpPicture, i, COL_不合理中心) Is Nothing Then '不合理   中心
                    str项目值 = Val(.Cell(flexcpData, i, COL_不合理中心))
                End If
                
                strSQL = "Zl_抗菌药物抽样评价_Insert(" & mlng抽样ID & "," & mlng序号 & ",1,'" & str项目编码 & "'," & i & ",0,'" & str项目值 & "')"
    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Next
        End With
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mblnChange = False
    mblnInitSaved = False
    SaveData = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get抽样明细SQL() As String
'功能：获取SQL并检查当前输入
    Dim strSQL As String, strTmp As String
    Dim lngTmp As Long
    
    strSQL = "Zl_抗菌药物抽样明细_Update(" & mlng抽样ID & "," & mlng病人ID & "," & mlng主页ID & "," & mlng序号 & ","
    strSQL = strSQL & IIf(mbln手术, 1, 0) & "," & IIf(optInfo(e_D_optInfo_检查_病原学检测_未做_2).Value, 0, 1) & ","
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_病原学检测日期_23).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_病原学检测日期_23).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_病原学检测标本_26).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_病原学检测标本_26).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_病原学检测检出细菌名_25).Text) & "',")
    strSQL = strSQL & IIf(optInfo(e_D_optInfo_检查_药敏试验_未做_6).Value, 0, 1) & ","
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_药敏试验日期_24).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_药敏试验日期_24).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf(optInfo(e_D_optInfo_检查_药敏试验_相符_8).Value, 0, 1) & ","
    
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药前体温_11).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药前体温_11).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药前白细胞计数_13).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药前白细胞计数_13).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药前中性粒细胞_15).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药前中性粒细胞_15).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药前C反应蛋白_17).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药前C反应蛋白_17).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶_19).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶_19).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药前肌酐_21).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药前肌酐_21).Text) & "',")
    
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药后体温_55).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药后体温_55).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药后白细胞计数_53).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药后白细胞计数_53).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药后中性粒细胞_51).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药后中性粒细胞_51).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药后C反应蛋白_49).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药后C反应蛋白_49).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶_45).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶_45).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_检查_用药后肌酐_43).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_检查_用药后肌酐_43).Text) & "',")
 
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药前体温日期_12).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药前体温日期_12).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药前白细胞计数日期_14).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药前白细胞计数日期_14).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药前中性粒细胞日期_16).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药前中性粒细胞日期_16).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药前C反应蛋白日期_18).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药前C反应蛋白日期_18).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶日期_20).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药前丙谷转氨酶日期_20).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药前肌酐日期_22).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药前肌酐日期_22).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药后体温日期_54).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药后体温日期_54).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药后白细胞计数日期_52).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药后白细胞计数日期_52).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药后中性粒细胞日期_50).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药后中性粒细胞日期_50).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药后C反应蛋白日期_48).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药后C反应蛋白日期_48).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶日期_44).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药后丙谷转氨酶日期_44).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_检查_用药后肌酐日期_42).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_检查_用药后肌酐日期_42).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    '影像学诊断
    strTmp = chkInfo(e_D_chkInfo_影像_X线_18).Value & chkInfo(e_D_chkInfo_影像_CT_19).Value & chkInfo(e_D_chkInfo_影像_磁共振_20).Value
    strSQL = strSQL & "'" & strTmp & "',"
    
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_影像_部位_46).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_影像_部位_46).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_影像_结论_47).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_影像_结论_47).Text) & "',")
    
    strSQL = strSQL & Val(txtInfo(e_D_txtInfo_症状_感染症状_35).Tag) & ","
    
    If optInfo(e_D_optInfo_用药目的_未用_17).Value Then
        lngTmp = 0
    ElseIf optInfo(e_D_optInfo_用药目的_预防_18).Value Then
        lngTmp = 1
    ElseIf optInfo(e_D_optInfo_用药目的_治疗_21).Value Then
        lngTmp = 2
    End If
    
    strSQL = strSQL & lngTmp & ","
    strSQL = strSQL & Val(txtInfo(e_D_txtInfo_用药目的_感染诊断_36).Tag) & ","
    
    If optInfo(e_D_optInfo_结果_好转_13).Value Then
        lngTmp = 2
    ElseIf optInfo(e_D_optInfo_结果_治愈_12).Value Then
        lngTmp = 1
    ElseIf optInfo(e_D_optInfo_结果_无效_14).Value Then
        lngTmp = 3
    End If
    strSQL = strSQL & lngTmp & ","
    
    strTmp = ""
    strTmp = "'" & chkInfo(e_D_chkInfo_评价_本院_适应症_5).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_适应症_37).Value & "','" & chkInfo(e_D_chkInfo_评价_本院_药物选择_12).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_药物选择_36).Value & "','" & _
        chkInfo(e_D_chkInfo_评价_本院_单次剂量_13).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_单次剂量_35).Value & "','" & chkInfo(e_D_chkInfo_评价_本院_每日给药频次_14).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_每日给药频次_34).Value & "','" & _
        chkInfo(e_D_chkInfo_评价_本院_溶剂_15).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_溶剂_33).Value & "','" & chkInfo(e_D_chkInfo_评价_本院_给药途径_16).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_给药途径_32).Value & "','" & _
        chkInfo(e_D_chkInfo_评价_本院_用药疗程_17).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_用药疗程_31).Value & "','" & chkInfo(e_D_chkInfo_评价_本院_术前_23).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_术前_28).Value & "','" & _
        chkInfo(e_D_chkInfo_评价_本院_术中_24).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_术中_27).Value & "','" & chkInfo(e_D_chkInfo_评价_本院_术后_25).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_术后_26).Value & "','" & _
        chkInfo(e_D_chkInfo_评价_本院_联合用药_22).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_联合用药_29).Value & "','" & chkInfo(e_D_chkInfo_评价_本院_更换药物_21).Value & "|" & chkInfo(e_D_chkInfo_评价_中心_更换药物_30).Value & "'"
    
    strSQL = strSQL & strTmp & ",'" & txtInfo(e_D_txtInfo_备注_57).Text & "'," & IIf(mbln编辑, 1, 0) & ","
    strSQL = strSQL & IIf(mbln打印, 1, 0) & "," & Val(txtInfo(e_D_txtInfo_用药情况_天数_30).Text) & "," & Val(txtInfo(e_D_txtInfo_用药情况_种数_29).Text) & ","
    strSQL = strSQL & IIf(optInfo(e_D_optInfo_结果_真菌药_有_20).Value, 1, "Null")
    strSQL = strSQL & ")"
    
    Get抽样明细SQL = strSQL
End Function

Private Sub Init用药评价表()
'功能：加用药评价表
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rs项目名称 As ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    Dim intCount As Integer
    Dim intLin As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    
    strSQL = "select 序号,decode(序号,null,decode(名称,'-',999,888),序号) as 排序,编码,名称,是否合理,上级 from 抗菌用药评价项目 where 末级=1 and 是否手术=[1] order by 序号,2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(mbln手术, 1, 0))
    
    strSQL = "select a.编码,a.序号,a.名称 from 抗菌用药评价项目 a,抗菌用药评价项目 b where a.末级=0 and a.上级=b.编码 and b.名称=[1] order by 序号"
    Set rs项目名称 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(mbln手术, "手术", "非手术"))
    
    With vsPJB
        .Clear
        .Redraw = False
        .FixedRows = 2
        .FixedCols = 0
        .Cols = 11
        .Rows = .FixedRows
        .RowHeightMin = 400
        .WordWrap = True
        For i = 0 To COL_不合理中心
            .ColAlignment(.FixedCols + i) = flexAlignCenterCenter
        Next
        .ColWidth(COL_评价项目) = 1200
        .ColWidth(COL_合理内容) = 3400
        .ColWidth(COL_合理本院) = 1000
        .ColWidth(COL_合理中心) = 1000
        
        .ColWidth(COL_不合理内容) = 3400
        .ColWidth(COL_不合理本院) = 1000
        .ColWidth(COL_不合理中心) = 1000
        
        
        .ColHidden(COL_合理编码) = True
        .ColHidden(COL_不合理编码) = True
        .ColHidden(COL_上级编码) = True
        .ColHidden(COL_上级名称) = True
        
        .Cell(flexcpText, 0, 0, 1, 1) = "评价项目"
        
        .Cell(flexcpText, 0, 1, 0, 3) = "合理"
        
        .Cell(flexcpText, 0, 4, 0, 6) = "不合理"

        .TextMatrix(1, COL_合理内容) = "评价内容"
        .TextMatrix(1, COL_不合理内容) = "评价内容"
        
        .TextMatrix(1, COL_合理本院) = "本院评价"
        .TextMatrix(1, COL_不合理本院) = "本院评价"
        
        .TextMatrix(1, COL_合理中心) = "中心或分网评价"
        .TextMatrix(1, COL_不合理中心) = "中心或分网评价"
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFixedOnly
        
        '边框设置
        intLin = 1
        .Select 0, 0, 1, COL_不合理中心
        .CellBorder .GridColor, intLin, intLin, intLin, intLin, 0, 0
        
        .Rows = rsTmp.RecordCount
        lngRow = 1
        For i = 1 To rs项目名称.RecordCount
            
            rsTmp.Filter = "上级='" & rs项目名称!编码 & "' and 是否合理=1"
            intCount = rsTmp.RecordCount
            For j = 1 To rsTmp.RecordCount
                .TextMatrix(lngRow + j, COL_上级编码) = rs项目名称!编码
                .TextMatrix(lngRow + j, COL_上级名称) = rs项目名称!名称
                
                .TextMatrix(lngRow + j, COL_评价项目) = rs项目名称!名称
                
                .TextMatrix(lngRow + j, COL_合理编码) = rsTmp!编码
                
                If Left(rsTmp!名称 & "", 1) = "-" Then
                    strTmp = ""
                Else
                    strTmp = IIf(rsTmp!序号 & "" <> "", rsTmp!序号 & ".", "") & rsTmp!名称
                End If
                
                .TextMatrix(lngRow + j, COL_合理内容) = strTmp
                
                If .TextMatrix(lngRow + j, COL_合理内容) <> "" Then
                    Set .Cell(flexcpPicture, lngRow + j, COL_合理本院) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_合理本院) = 1
                    
                    Set .Cell(flexcpPicture, lngRow + j, COL_合理中心) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_合理中心) = 1
                    
                End If
                rsTmp.MoveNext
            Next
            
            rsTmp.Filter = "上级='" & rs项目名称!编码 & "' and 是否合理=0"
            If intCount < rsTmp.RecordCount Then intCount = rsTmp.RecordCount
            For j = 1 To rsTmp.RecordCount
                .TextMatrix(lngRow + j, COL_上级编码) = rs项目名称!编码
                .TextMatrix(lngRow + j, COL_上级名称) = rs项目名称!名称
                
                .TextMatrix(lngRow + j, COL_评价项目) = rs项目名称!名称
                
                .TextMatrix(lngRow + j, COL_不合理编码) = rsTmp!编码
                
                If Left(rsTmp!名称 & "", 1) = "-" Then
                    strTmp = ""
                Else
                    strTmp = IIf(rsTmp!序号 & "" <> "", rsTmp!序号 & ".", "") & rsTmp!名称
                End If
                
                .TextMatrix(lngRow + j, COL_不合理内容) = strTmp
                
                
                If .TextMatrix(lngRow + j, COL_不合理内容) <> "" Then
                    Set .Cell(flexcpPicture, lngRow + j, COL_不合理本院) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_不合理本院) = 1
                    
                    Set .Cell(flexcpPicture, lngRow + j, COL_不合理中心) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_不合理中心) = 1
                    
                End If
                
                rsTmp.MoveNext
            Next
            
            .Select lngRow + 1, 0, lngRow + intCount, COL_不合理中心
            .CellBorder .GridColor, intLin, intLin, intLin, intLin, 0, 0
            
            rs项目名称.MoveNext
            lngRow = lngRow + intCount
        Next
        .Rows = lngRow + 1
        
        .Cell(flexcpFontBold, 0, COL_评价项目, 1, COL_不合理中心) = True
        .Cell(flexcpAlignment, 2, COL_评价项目, .Rows - 1, COL_评价项目) = flexAlignLeftCenter
        .Cell(flexcpFontBold, 0, COL_评价项目, .Rows - 1, COL_评价项目) = True
        .Cell(flexcpAlignment, 2, COL_合理内容, .Rows - 1, COL_合理内容) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 2, COL_不合理内容, .Rows - 1, COL_不合理内容) = flexAlignLeftCenter
        
        .Cell(flexcpPictureAlignment, 2, COL_合理本院, .Rows - 1, COL_合理中心) = flexPicAlignCenterCenter
        .Cell(flexcpPictureAlignment, 2, COL_不合理本院, .Rows - 1, COL_不合理中心) = flexPicAlignCenterCenter
        .Row = .FixedRows
        .Redraw = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsDrugUse_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'功能：擦除部分边框线
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsDrugUse
        If Col = COL_DRUG_给药频次 Or Col = COL_DRUG_途径 Or Col = COL_DRUG_起止时间 Then
            Call Get一并给药范围(Val(.TextMatrix(Row, COL_DRUG_相关ID)), lngBegin, lngEnd)
            If lngBegin >= lngEnd Then Exit Sub
            
            vRect.Left = Left
            vRect.Right = Right - 1 '保留右边表格线
            
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
            Done = True
        End If
    End With
End Sub

Private Sub vsOperate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_OPE_用药期间 Then vsOperate.Cell(flexcpData, Row, Col) = vsOperate.ComboData
End Sub

Private Sub vsOperate_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsOperate
        If NewCol >= COL_OPE_切口 And NewCol <= COL_OPE_给药情况 Then
            .Editable = flexEDKbdMouse
            If NewCol = COL_OPE_切口 Or NewCol = COL_OPE_给药情况 Or NewCol = COL_OPE_用药期间 Then
                .ComboList = .ColData(NewCol)
            Else
                .ComboList = ""
            End If
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub vsOperate_KeyPress(KeyAscii As Integer)
    With vsOperate
        If KeyAscii = 13 Then
            KeyAscii = 0
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsOperate_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    mblnChange = True
    If KeyAscii = 13 Then
        mblnReturn = True
    ElseIf Col = COL_OPE_结束时间 Or Col = COL_OPE_开始时间 Then
        If InStr("0123456789-: " & Chr(8) & Chr(27) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsOperate_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_OPE_结束时间 Or Col = COL_OPE_开始时间 Then
        vsOperate.Refresh    '如果有弹出提示，不刷新的话，一并给药通过Drawcell被擦除的单元格会再次显示
        If Not AcceptInput(Row, Col) Then
            Cancel = True
        Else
            If mblnReturn Then
                Call vsOperate_KeyPress(13) '定位到一下输入单元
                mblnChange = True
            End If
        End If
    End If
    mblnReturn = False
End Sub

Private Function AcceptInput(ByVal Row As Long, ByVal Col As Long) As Boolean
    
    AcceptInput = False
    With vsOperate
        If .EditText <> "" Then .EditText = zlStr.FullDate(.EditText)
        If .EditText = "" Then AcceptInput = True: Exit Function
        If .EditText = .TextMatrix(Row, Col) Then AcceptInput = True: Exit Function
        
        '检查输入的有效性
        If Not IsDate(.EditText) Then
            MsgBox "请输入一个有效的" & .TextMatrix(0, Col) & " 。", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        
        '必须大于入院时间
        If Format(.EditText, "yyyy-MM-dd HH:mm") <= txtInfo(e_D_txtInfo_基本_入院时间_9).Tag Then
            MsgBox "输入的手术" & IIf(Col = COL_OPE_结束时间, "结束", "开始") & "时间必须大于病人入院时间 " & txtInfo(e_D_txtInfo_基本_入院时间_9).Tag & "。", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        '不能小于出院时间
        If Format(.EditText, "yyyy-MM-dd HH:mm") > txtInfo(e_D_txtInfo_基本_出院时间_10).Tag Then
            MsgBox "输入的手术" & IIf(Col = COL_OPE_结束时间, "结束", "开始") & "时间不应大于病人出院时间 " & txtInfo(e_D_txtInfo_基本_出院时间_10).Tag & "。", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        
        If Col = COL_OPE_结束时间 And .TextMatrix(Row, COL_OPE_开始时间) <> "" Then
            If Format(.EditText, "yyyy-MM-dd HH:mm") < .TextMatrix(Row, COL_OPE_开始时间) Then
                MsgBox "输入的手术结束时间必须大于手术开始时间 " & .TextMatrix(Row, COL_OPE_开始时间) & "。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        ElseIf Col = COL_OPE_开始时间 And .TextMatrix(Row, COL_OPE_结束时间) <> "" Then
            If Format(.EditText, "yyyy-MM-dd HH:mm") > .TextMatrix(Row, COL_OPE_结束时间) Then
                MsgBox "输入的手术开始时间必须小于手术结束时间 " & .TextMatrix(Row, COL_OPE_结束时间) & "。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        End If
        
        .TextMatrix(Row, Col) = .EditText
        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
        mblnChange = True
    End With
    AcceptInput = True
End Function

Private Sub vsPJB_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Visible Then mblnChange = True
End Sub

Private Sub vsPJB_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'功能：擦除部分边框线
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsPJB
        If Col <> COL_评价项目 And .TextMatrix(Row, COL_上级编码) <> "" Then Exit Sub
        
        If Not Get同一上级行(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left + 1 '擦除左边表格线保留
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function Get同一上级行(ByVal lngRow As Long, ByRef lngBegin As Long, ByRef lngEnd As Long) As Boolean
'功能：属于一个处方的行
'参数：lngRow 当前行，lngID 挂号ID
    Dim i As Long
    
    lngBegin = lngRow
    lngEnd = lngRow
    
    With vsPJB
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, COL_上级编码) = "" Then Exit Function
            If .TextMatrix(i, COL_上级编码) <> .TextMatrix(lngRow, COL_上级编码) Then
                Exit For
            Else
                lngBegin = i
            End If
        Next
        
        For i = lngRow + 1 To .Rows - 1
            If .TextMatrix(i, COL_上级编码) = "" Then Exit Function
            If .TextMatrix(i, COL_上级编码) <> .TextMatrix(lngRow, COL_上级编码) Then
                Exit For
            Else
                lngEnd = i
            End If
        Next
    End With
    
    If lngBegin < lngEnd Then Get同一上级行 = True
    
End Function

Private Sub vsPJB_KeyPress(KeyAscii As Integer)
    Dim blnEdit As Boolean
    Dim lngCol As Long
    Dim str项目名称 As String
    Dim str项目明细内容 As String
    
    With vsPJB
        If .Row <= 1 Or .TextMatrix(.Row, COL_上级名称) = "" Then Exit Sub
        
        str项目名称 = .TextMatrix(.Row, COL_上级名称)
        
        If .Col = COL_合理本院 Then
            blnEdit = Edit评价项目(str项目名称, True, True)
            lngCol = COL_合理内容
        ElseIf .Col = COL_合理中心 Then
            blnEdit = Edit评价项目(str项目名称, True, False)
            lngCol = COL_合理内容
        ElseIf .Col = COL_不合理本院 Then
            blnEdit = Edit评价项目(str项目名称, False, True)
            lngCol = COL_不合理内容
        ElseIf .Col = COL_不合理中心 Then
            blnEdit = Edit评价项目(str项目名称, False, False)
            lngCol = COL_不合理内容
        End If
        
        str项目明细内容 = .TextMatrix(.Row, lngCol)
        
        .Editable = flexEDNone
        If blnEdit And str项目明细内容 = "" Then
            If .TextMatrix(.Row - 1, lngCol) <> "" Then .Editable = flexEDKbdMouse
        ElseIf blnEdit Then
            If .Cell(flexcpData, .Row, .Col) = 2 Then '1-不勾选，2－表示勾选
                Set .Cell(flexcpPicture, .Row, .Col) = img16.ListImages("Check").Picture
                .Cell(flexcpData, .Row, .Col) = 1
                mblnChange = True
            ElseIf .Cell(flexcpData, .Row, .Col) = 1 Then
                Set .Cell(flexcpPicture, .Row, .Col) = img16.ListImages("UnCheck").Picture
                .Cell(flexcpData, .Row, .Col) = 2
                mblnChange = True
            End If
        End If
    End With
End Sub

Private Sub vsPJB_DblClick()
    Call vsPJB_KeyPress(32)
End Sub

Private Function Edit评价项目(ByVal str项目 As String, ByVal bln合理 As Boolean, ByVal bln本院 As Boolean) As Boolean
'功能：获取某个评价项目可以进行勾选的范围，如适应证这项目，包含 3 行， 4 5 6
'参数：lngS 开始行，lngE 结束行
 
    If bln本院 And chkInfo(e_D_chkInfo_评价_本院_适应症_5).Value <> 1 And "适应症" <> str项目 Then
        Edit评价项目 = False
        Exit Function
    End If
    
    If Not bln本院 And chkInfo(e_D_chkInfo_评价_中心_适应症_37).Value <> 1 And "适应症" <> str项目 Then
        Edit评价项目 = False
        Exit Function
    End If
    
    Select Case str项目
        Case "适应症"
            If bln本院 Then
                Edit评价项目 = (chkInfo(e_D_chkInfo_评价_本院_适应症_5).Value = 1 And bln合理) Or (chkInfo(e_D_chkInfo_评价_本院_适应症_5).Value <> 1 And Not bln合理)
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_适应症_37).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_适应症_37).Value <> 1 And Not bln合理
            End If
        Case "药物选择"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_药物选择_12).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_药物选择_12).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_药物选择_36).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_药物选择_36).Value <> 1 And Not bln合理
            End If
        Case "单次剂量"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_单次剂量_13).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_单次剂量_13).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_单次剂量_35).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_单次剂量_35).Value <> 1 And Not bln合理
            End If
        Case "每日给药频次"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_每日给药频次_14).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_每日给药频次_14).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_每日给药频次_34).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_每日给药频次_34).Value <> 1 And Not bln合理
            End If
        Case "溶剂"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_溶剂_15).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_溶剂_15).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_溶剂_33).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_溶剂_33).Value <> 1 And Not bln合理
            End If
        Case "给药途径"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_给药途径_16).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_给药途径_16).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_给药途径_32).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_给药途径_32).Value <> 1 And Not bln合理
            End If
        Case "用药疗程"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_用药疗程_17).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_用药疗程_17).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_用药疗程_31).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_用药疗程_31).Value <> 1 And Not bln合理
            End If
        Case "术前用药时间"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_术前_23).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_术前_23).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_术前_28).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_术前_28).Value <> 1 And Not bln合理
            End If
        Case "术中用药"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_术中_24).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_术中_24).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_术中_27).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_术中_27).Value <> 1 And Not bln合理
            End If
        Case "联合用药"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_联合用药_22).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_联合用药_22).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_联合用药_29).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_联合用药_29).Value <> 1 And Not bln合理
            End If
        Case "术后用药"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_术后_25).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_术后_25).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_术后_26).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_术后_26).Value <> 1 And Not bln合理
            End If
        Case "更换药物"
            If bln本院 Then
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_本院_更换药物_21).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_本院_更换药物_21).Value <> 1 And Not bln合理
            Else
                Edit评价项目 = chkInfo(e_D_chkInfo_评价_中心_更换药物_30).Value = 1 And bln合理 Or chkInfo(e_D_chkInfo_评价_中心_更换药物_30).Value <> 1 And Not bln合理
            End If
    End Select
End Function

Private Sub Cls评价明细(ByVal intIndex As Integer)
'功能：当抽样界面的勾选情况发生变化时，调整对应的评价内容
    Select Case intIndex
        Case e_D_chkInfo_评价_本院_适应症_5
            Call Set评价项目值("适应症", True, True)
        Case e_D_chkInfo_评价_本院_药物选择_12
            Call Set评价项目值("药物选择", True, False)
        Case e_D_chkInfo_评价_本院_单次剂量_13
            Call Set评价项目值("单次剂量", True, False)
        Case e_D_chkInfo_评价_本院_每日给药频次_14
            Call Set评价项目值("每日给药频次", True, False)
        Case e_D_chkInfo_评价_本院_溶剂_15
            Call Set评价项目值("溶剂", True, False)
        Case e_D_chkInfo_评价_本院_给药途径_16
            Call Set评价项目值("给药途径", True, False)
        Case e_D_chkInfo_评价_本院_用药疗程_17
            Call Set评价项目值("用药疗程", True, False)
        Case e_D_chkInfo_评价_本院_更换药物_21
            Call Set评价项目值("更换药物", True, False)
        Case e_D_chkInfo_评价_本院_联合用药_22
            Call Set评价项目值("联合用药", True, False)
        Case e_D_chkInfo_评价_本院_术前_23
            Call Set评价项目值("术前用药时间", True, False)
        Case e_D_chkInfo_评价_本院_术中_24
            Call Set评价项目值("术中用药", True, False)
        Case e_D_chkInfo_评价_本院_术后_25
            Call Set评价项目值("术后用药", True, False)
        Case e_D_chkInfo_评价_中心_适应症_37
            Call Set评价项目值("适应症", False, True)
        Case e_D_chkInfo_评价_中心_药物选择_36
            Call Set评价项目值("药物选择", False, False)
        Case e_D_chkInfo_评价_中心_单次剂量_35
            Call Set评价项目值("单次剂量", False, False)
        Case e_D_chkInfo_评价_中心_每日给药频次_34
            Call Set评价项目值("每日给药频次", False, False)
        Case e_D_chkInfo_评价_中心_溶剂_33
            Call Set评价项目值("溶剂", False, False)
        Case e_D_chkInfo_评价_中心_给药途径_32
            Call Set评价项目值("给药途径", False, False)
        Case e_D_chkInfo_评价_中心_用药疗程_31
            Call Set评价项目值("用药疗程", False, False)
        Case e_D_chkInfo_评价_中心_更换药物_30
            Call Set评价项目值("更换药物", False, False)
        Case e_D_chkInfo_评价_中心_联合用药_29
            Call Set评价项目值("联合用药", False, False)
        Case e_D_chkInfo_评价_中心_术前_28
            Call Set评价项目值("术前用药时间", False, False)
        Case e_D_chkInfo_评价_中心_术中_27
            Call Set评价项目值("术中用药", False, False)
        Case e_D_chkInfo_评价_中心_术后_26
            Call Set评价项目值("术后用药", False, False)
    End Select
End Sub

Private Sub Set评价项目值(ByVal str项目 As String, ByVal bln本院 As Boolean, Optional ByVal blnALL As Boolean)
'功能：置设评价项目的值，勾选或不勾选清除录入的文字 处理指定列指定行范围内单元格的值
'参数：blnAll 整列都清空
    Dim i As Long
    
    With vsPJB
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_上级名称) = str项目 And Not blnALL Then
                If bln本院 Then
                    If Not .Cell(flexcpPicture, i, COL_合理本院) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_合理本院) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_合理本院) = 1
                    Else
                        .TextMatrix(i, COL_合理本院) = ""
                    End If
                    
                    If Not .Cell(flexcpPicture, i, COL_不合理本院) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_不合理本院) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_不合理本院) = 1
                    Else
                        .TextMatrix(i, COL_不合理本院) = ""
                    End If
                Else
                    If Not .Cell(flexcpPicture, i, COL_合理中心) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_合理中心) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_合理中心) = 1
                    Else
                        .TextMatrix(i, COL_合理中心) = ""
                    End If
                    
                    If Not .Cell(flexcpPicture, i, COL_不合理中心) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_不合理中心) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_不合理中心) = 1
                    Else
                        .TextMatrix(i, COL_不合理中心) = ""
                    End If
                End If
            End If
            
            If blnALL And bln本院 Then
                If Not .Cell(flexcpPicture, i, COL_合理本院) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_合理本院) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_合理本院) = 1
                Else
                    .TextMatrix(i, COL_合理本院) = ""
                End If
                If Not .Cell(flexcpPicture, i, COL_不合理本院) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_不合理本院) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_不合理本院) = 1
                Else
                    .TextMatrix(i, COL_不合理本院) = ""
                End If
            ElseIf blnALL And Not bln本院 Then
                If Not .Cell(flexcpPicture, i, COL_合理中心) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_合理中心) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_合理中心) = 1
                Else
                    .TextMatrix(i, COL_合理中心) = ""
                End If
                If Not .Cell(flexcpPicture, i, COL_不合理中心) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_不合理中心) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_不合理中心) = 1
                Else
                    .TextMatrix(i, COL_不合理中心) = ""
                End If
            End If
        Next
    End With
End Sub

Private Sub Load用药评价()
'功能：加载病人的用药评价
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lngTmp As String
    Dim i As Long, j As Long
    Dim intCol As Integer
    
    On Error GoTo errH
    strSQL = "select 项目编码,decode(项目值,'1','',项目值) as 项目值,评价类型 from 抗菌药物抽样评价 where 抽样id=[1] and 序号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng抽样ID, mlng序号)
    For i = 1 To rsTmp.RecordCount
        With vsPJB
            For j = .FixedRows To .Rows - 1
                If .TextMatrix(j, COL_合理编码) = rsTmp!项目编码 & "" Then
                    intCol = IIf(Val(rsTmp!评价类型 & "") = 0, COL_合理本院, COL_合理中心)
                    If Val(rsTmp!项目值 & "") = 2 Then
                        Set .Cell(flexcpPicture, j, intCol) = img16.ListImages("UnCheck").Picture
                        .Cell(flexcpData, j, intCol) = 2
                    Else
                        .TextMatrix(j, intCol) = rsTmp!项目值 & ""
                    End If
                ElseIf .TextMatrix(j, COL_不合理编码) = rsTmp!项目编码 & "" Then
                    intCol = IIf(Val(rsTmp!评价类型 & "") = 0, COL_不合理本院, COL_不合理中心)
                    If Val(rsTmp!项目值 & "") = 2 Then
                        Set .Cell(flexcpPicture, j, intCol) = img16.ListImages("UnCheck").Picture
                        .Cell(flexcpData, j, intCol) = 2
                    Else
                        .TextMatrix(j, intCol) = rsTmp!项目值 & ""
                    End If
                End If
            Next
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsPJB_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'自由录入用药评价时不能输入数字
   If InStr("0123456789'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Function UseKssDrug() As Boolean
'功能：病人是否使用了抗菌药，是由医嘱下达产生
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
 
    strSQL = "Select 1" & vbNewLine & _
        "From 病人医嘱记录 A, 药品特性 C" & vbNewLine & _
        "Where a.诊疗类别 = '5' And a.诊疗项目id = c.药名id And Nvl(c.抗生素, 0) <> 0 And Exists" & vbNewLine & _
        " (Select 1 From 病人医嘱发送 B Where a.Id = b.医嘱id) And a.病人id = [1] And a.主页id = [2] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    UseKssDrug = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DateForShow(ByVal strDate As String) As String
'将日期转换为界面显示的日期
    Dim strTmp As String
    If strDate = "" Then DateForShow = "": Exit Function
    If Not IsDate(strDate) Then DateForShow = "": Exit Function
    strTmp = Format(strDate, "yyyy-mm-dd")
    DateForShow = Val(Split(strTmp, "-")(1)) & "/" & Val(Split(strTmp, "-")(2))
End Function

Private Sub InitRS(ByRef rsCtl As ADODB.Recordset)
'功能：控件记录集
    Dim arrFileds() As Variant
    
    Set rsCtl = New ADODB.Recordset
    rsCtl.Fields.Append "信息名", adVarChar, 100
    rsCtl.Fields.Append "信息类型", adInteger '0-日期，1-文本，2-数字
    rsCtl.Fields.Append "控件名", adVarChar, 100
    rsCtl.Fields.Append "控件索引", adBigInt
    rsCtl.Fields.Append "信息长度", adBigInt
    rsCtl.CursorLocation = adUseClient
    rsCtl.LockType = adLockOptimistic
    rsCtl.CursorType = adOpenStatic
    rsCtl.Open
    
    arrFileds = Array("信息名", "信息类型", "控件名", "控件索引", "信息长度")
    
    With rsCtl
        .AddNew arrFileds, Array("用药前体温", Type数字, "txtInfo", e_D_txtInfo_检查_用药前体温_11, 8)
        .AddNew arrFileds, Array("用药后体温", Type数字, "txtInfo", e_D_txtInfo_检查_用药后体温_55, 8)
        .AddNew arrFileds, Array("用药前体温日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药前体温日期_12, 16)
        .AddNew arrFileds, Array("用药后体温日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药后体温日期_54, 16)
        
        .AddNew arrFileds, Array("用药前白细胞计数", Type文本, "txtInfo", e_D_txtInfo_检查_用药前白细胞计数_13, 30)
        .AddNew arrFileds, Array("用药前白细胞计数日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药前白细胞计数日期_14, 16)
        .AddNew arrFileds, Array("用药后白细胞计数", Type文本, "txtInfo", e_D_txtInfo_检查_用药后白细胞计数_53, 30)
        .AddNew arrFileds, Array("用药后白细胞计数日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药后白细胞计数日期_52, 16)
        
        .AddNew arrFileds, Array("用药前中性粒细胞", Type文本, "txtInfo", e_D_txtInfo_检查_用药前中性粒细胞_15, 30)
        .AddNew arrFileds, Array("用药前中性粒细胞日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药前中性粒细胞日期_16, 16)
        .AddNew arrFileds, Array("用药后中性粒细胞", Type文本, "txtInfo", e_D_txtInfo_检查_用药后中性粒细胞_51, 30)
        .AddNew arrFileds, Array("用药后中性粒细胞日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药后中性粒细胞日期_50, 16)
        
        .AddNew arrFileds, Array("用药前C反应蛋白", Type文本, "txtInfo", e_D_txtInfo_检查_用药前C反应蛋白_17, 30)
        .AddNew arrFileds, Array("用药前C反应蛋白日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药前C反应蛋白日期_18, 16)
        .AddNew arrFileds, Array("用药后C反应蛋白", Type文本, "txtInfo", e_D_txtInfo_检查_用药后C反应蛋白_49, 30)
        .AddNew arrFileds, Array("用药后C反应蛋白日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药后C反应蛋白日期_48, 16)
        
        .AddNew arrFileds, Array("用药前谷丙转氨酶", Type文本, "txtInfo", e_D_txtInfo_检查_用药前丙谷转氨酶_19, 30)
        .AddNew arrFileds, Array("用药前谷丙转氨酶日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药前丙谷转氨酶日期_20, 16)
        .AddNew arrFileds, Array("用药后谷丙转氨酶", Type文本, "txtInfo", e_D_txtInfo_检查_用药后丙谷转氨酶_45, 30)
        .AddNew arrFileds, Array("用药后谷丙转氨酶日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药后丙谷转氨酶日期_44, 16)
        
        .AddNew arrFileds, Array("用药前肌酐", Type文本, "txtInfo", e_D_txtInfo_检查_用药前肌酐_21, 30)
        .AddNew arrFileds, Array("用药前肌酐日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药前肌酐日期_22, 16)
        .AddNew arrFileds, Array("用药后肌酐", Type文本, "txtInfo", e_D_txtInfo_检查_用药后肌酐_43, 30)
        .AddNew arrFileds, Array("用药后肌酐日期", Type日期, "txtInfo", e_D_txtInfo_检查_用药后肌酐日期_42, 16)
        
        .AddNew arrFileds, Array("临床微生物检查病源学检测日期", Type日期, "txtInfo", e_D_txtInfo_检查_病原学检测日期_23, 16)
        .AddNew arrFileds, Array("临床微生物检查病源学检测标本", Type文本, "txtInfo", e_D_txtInfo_检查_病原学检测标本_26, 50)
        .AddNew arrFileds, Array("临床微生物检查药敏试验日期", Type文本, "txtInfo", e_D_txtInfo_检查_病原学检测检出细菌名_25, 100)
        
        .AddNew arrFileds, Array("临床微生物检查病源学检测日期", Type日期, "txtInfo", e_D_txtInfo_检查_药敏试验日期_24, 16)
        
        .AddNew arrFileds, Array("影像学诊断部位", Type文本, "txtInfo", e_D_txtInfo_影像_部位_46, 50)
        .AddNew arrFileds, Array("影像学诊断结论", Type文本, "txtInfo", e_D_txtInfo_影像_结论_47, 100)
        
        .AddNew arrFileds, Array("备注", Type文本, "txtInfo", e_D_txtInfo_备注_57, 500)
    End With
    
End Sub

Private Function CheckData() As Boolean
'功能：输入框检查判断
    Dim i As Long
    Dim objCtl As Object
    Dim blnDo As Boolean
    Dim strTmp As String
    Dim strMsg As String
    Dim lngColor As Long
    
    If mrsCtl Is Nothing Then
        Call InitRS(mrsCtl)
    End If
    
    mrsCtl.MoveFirst
    For i = 1 To mrsCtl.RecordCount
        Set objCtl = txtInfo(Val(mrsCtl!控件索引 & ""))
        If objCtl.Enabled And objCtl.Locked = False And objCtl.Text <> "" Then
            lngColor = objCtl.BackColor
            strMsg = ""
            Select Case Val(mrsCtl!信息类型 & "")
            Case 0 '日期
                objCtl.BackColor = &HC0C0FF
                Call txtInfo_Validate(Val(mrsCtl!控件索引 & ""), blnDo)
                objCtl.BackColor = lngColor
                If blnDo Then
                    Exit Function
                End If
            Case 1 '文本
                If Len(objCtl.Text) > Val(mrsCtl!信息长度 & "") Then
                    strMsg = mrsCtl!信息名 & "-内容太长(允许录入" & Val(mrsCtl!信息长度 & "") & "个字符或" & Val(mrsCtl!信息长度 & "") \ 2 & "个汉字)。"
                ElseIf InStr(objCtl.Text, "'") > 0 Then
                    strMsg = mrsCtl!信息名 & "包含特殊字符半角单引号。"
                End If
            Case 2 '数字
                If Not IsNumeric(objCtl.Text) Then
                    strMsg = mrsCtl!信息名 & "要求只能录入数字。"
                ElseIf Len(objCtl.Text) > Val(mrsCtl!信息长度 & "") Then
                    strMsg = mrsCtl!信息名 & "-内容太长(允许录入" & Val(mrsCtl!信息长度 & "") & "个字符。"
                End If
            End Select
            
            If strMsg <> "" Then
                objCtl.BackColor = &HC0C0FF
                MsgBox strMsg, vbInformation, gstrSysName
                objCtl.SetFocus
                objCtl.BackColor = lngColor
                Exit Function
            End If
        End If
        mrsCtl.MoveNext
    Next
    
    CheckData = True
End Function

Private Sub vsPJB_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPJB
        If Len(.EditText) > 200 Then
            MsgBox "自由录入内容过长。", vbInformation, gstrSysName
            Cancel = True
        End If
    End With
End Sub
