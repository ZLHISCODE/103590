VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKssSurveyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˿���ҩ��ʹ����������"
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
   StartUpPosition =   2  '��Ļ����
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
            Caption         =   "����ʱ�䣺               ��"
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
            Name            =   "����"
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
         Caption         =   "������ҩ���������������"
         BeginProperty Font 
            Name            =   "����"
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
                     Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "�������"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "��"
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
               Caption         =   "�Ź���"
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
               Caption         =   "X��"
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
               Caption         =   "3.���ۣ�"
               Height          =   195
               Index           =   106
               Left            =   8085
               TabIndex        =   92
               Top             =   135
               Width           =   795
            End
            Begin VB.Label lblInfo 
               Caption         =   "2.��λ��"
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
               Caption         =   "Ӱ��ѧ���"
               BeginProperty Font 
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "�ٴ�֢״"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "���Ⱦ�йص���Ҫ֢״��"
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
               Caption         =   "��"
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
                  Caption         =   "3.����(��)"
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
                  Caption         =   "2.Ԥ��(��)"
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
                  Caption         =   "1.δ��ҩ"
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
                  Name            =   "����"
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
               Caption         =   "��ҩĿ��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "��Ⱦ���"
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
               Caption         =   "����(Cr)��      (        )"
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
               Caption         =   "�ȱ�ת��ø(ATL)��       (        )"
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
               Caption         =   "C��Ӧ����(CPR)��      (       )"
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
               Caption         =   "������ϸ��(NEUT%)��       (        )"
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
               Caption         =   "��ϸ������(WBC)��       (        )"
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
               Caption         =   "����(t)��    �� ��(        )"
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
               Caption         =   "��"
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
               Caption         =   "��"
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
               Caption         =   "1.δ��"
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
               Caption         =   "2.��"
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
               Caption         =   "1.δ��"
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
               Caption         =   "2.��"
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
               Caption         =   "δ���/"
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
               Caption         =   "���-"
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
               Caption         =   "���/"
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
               Caption         =   "�����"
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
                  Caption         =   "��Ӧ֤(������Ӧ֤�������������¸���)"
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
                  Caption         =   "ҩ��ѡ��"
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
                  Caption         =   "���μ���"
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
                  Caption         =   "ÿ�ո�ҩƵ��"
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
                  Caption         =   "�� ��"
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
                  Caption         =   "��ҩ;��"
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
                  Caption         =   "��ҩ�Ƴ�"
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
                  Caption         =   "����ҩƷ"
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
                  Caption         =   "������ҩ"
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
                  Caption         =   "��ǰ"
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
                  Caption         =   "����"
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
                  Caption         =   "����"
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
                  Caption         =   "Χ��������ҩʱ�䣺"
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
                  Caption         =   "����"
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
                  Caption         =   "����"
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
                  Caption         =   "��ǰ"
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
                  Caption         =   "������ҩ"
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
                  Caption         =   "����ҩƷ"
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
                  Caption         =   "��ҩ�Ƴ�"
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
                  Caption         =   "��ҩ;��"
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
                  Caption         =   "�� ��"
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
                  Caption         =   "ÿ�ո�ҩƵ��"
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
                  Caption         =   "���μ���"
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
                  Caption         =   "ҩ��ѡ��"
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
                  Caption         =   "��Ӧ֤(������Ӧ֤�������������¸���)"
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   120
                  TabIndex        =   159
                  Top             =   30
                  Width           =   3630
               End
               Begin VB.Label lblInfo 
                  Caption         =   "Χ��������ҩʱ�䣺"
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
                  Caption         =   "��"
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
                  Caption         =   "��"
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
                     Name            =   "����"
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
                  Caption         =   "�ۼ�ʹ�ÿ���ҩ        ��        ��"
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
                  Caption         =   "��Ч"
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
                  Caption         =   "��ת"
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
                  Caption         =   "����"
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
                  Caption         =   "��"
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
                  Caption         =   "��"
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
               Caption         =   "˵��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "��ע"
               BeginProperty Font 
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "��ҩ����������"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "����  �� ����"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "��Ժ"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "ʹ�ÿ����ҩ"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "��ҩ��� (ע����ҩ��ͬд���ܼ����Ƽ�����)"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "����(Ԫ)"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "���ƽ��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "סԺ�ܷ��ã�          Ԫ   סԺҩƷ�ܷ��ã�        Ԫ    סԺ����ҩ���ܷ��ã�        Ԫ"
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
               Caption         =   "�̷�(ҽԺ)��Ⱦ"
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
            Caption         =   "�ٴ�΢�����  ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�������˿���ҩ��ʹ����������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����ʱ�䣺"
            Height          =   195
            Index           =   0
            Left            =   4455
            TabIndex        =   137
            Top             =   450
            Width           =   900
         End
         Begin VB.Label lblInfo 
            Caption         =   "               ��"
            Height          =   210
            Index           =   1
            Left            =   5310
            TabIndex        =   136
            Top             =   450
            Width           =   2850
         End
         Begin VB.Label lblInfo 
            Caption         =   "�����������ң�"
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
            Caption         =   "�����ţ�"
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
            Caption         =   "�������˳�Ժ������"
            Height          =   195
            Index           =   4
            Left            =   9720
            TabIndex        =   133
            Top             =   450
            Width           =   1965
         End
         Begin VB.Label lblInfo 
            Caption         =   "��ţ�"
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
               Name            =   "����"
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
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�Ա�"
            Height          =   195
            Index           =   8
            Left            =   1590
            TabIndex        =   129
            Top             =   1455
            Width           =   510
         End
         Begin VB.Label lblInfo 
            Caption         =   "����        "
            Height          =   195
            Index           =   9
            Left            =   3315
            TabIndex        =   128
            Top             =   1455
            Width           =   1365
         End
         Begin VB.Label lblInfo 
            Caption         =   "����        "
            Height          =   195
            Index           =   10
            Left            =   4965
            TabIndex        =   127
            Top             =   1455
            Width           =   1755
         End
         Begin VB.Label lblInfo 
            Caption         =   "��Ժʱ��"
            Height          =   195
            Index           =   11
            Left            =   8100
            TabIndex        =   126
            Top             =   1455
            Width           =   795
         End
         Begin VB.Label lblInfo 
            Caption         =   "��Ժʱ��"
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
               Name            =   "����"
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
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "����ʷ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "������Ʒͨ����"
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
            Caption         =   "����(t)��    �� ��(        )"
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
            Caption         =   "��ϸ������(WBC)��       (        )"
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
            Caption         =   "������ϸ��(NEUT%)��       (        )"
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
            Caption         =   "C��Ӧ����(CPR)��      (       )"
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
            Caption         =   "�ȱ�ת��ø(ATL)��       (        )"
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
            Caption         =   "����(Cr)��      (        )"
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
            Caption         =   "��Դѧ��⣺"
            Height          =   195
            Index           =   30
            Left            =   1590
            TabIndex        =   113
            Top             =   4215
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Caption         =   "ҩ�����飺"
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
            Caption         =   "(                                   ��)"
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
            Caption         =   "�걾"
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
               Name            =   "����"
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
            Caption         =   "ʵ���Ҽ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��ҩǰ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��ҩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Key             =   "������"
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
    
    '����� e_F;e_P��ҩ����;e_D��������ϸ
    e_F_txtInfo_����ʱ��_��_0 = 0
    e_F_txtInfo_����ʱ��_ֹ_1 = 1
    e_F_lblInfo_��Ժ������ǩ_4 = 4
    e_F_txtInfo_��Ժ����_4 = 4
    e_F_txtInfo_��������_2 = 2
    e_F_txtInfo_������_3 = 3
    e_F_txtInfo_���_5 = 5
    
    e_P_txtInfo_����ʱ��_��_28 = 28
    e_P_txtInfo_����ʱ��_ֹ_56 = 56
    
    e_D_lblInfo_����_���_6 = 6
    e_D_txtInfo_����_�Ա�_6 = 6
    e_D_txtInfo_����_����_7 = 7
    e_D_txtInfo_����_����_8 = 8
    e_D_txtInfo_����_��Ժʱ��_9 = 9
    e_D_txtInfo_����_��Ժʱ��_10 = 10
    
    e_D_lblInfo_���_���_13 = 13
    e_D_txtInfo_���_���_27 = 27
    
    e_D_lblInfo_����_���_15 = 15
    e_D_optInfo_����_��_0 = 0
    e_D_optInfo_����_��_1 = 1
    e_D_txtInfo_����_ͨ����_34 = 34
    
    e_D_optInfo_���_��ԭѧ���_��_3 = 3
    e_D_optInfo_���_��ԭѧ���_δ��_2 = 2
    e_D_txtInfo_���_��ԭѧ�������_23 = 23
    e_D_txtInfo_���_��ԭѧ���걾_26 = 26
    e_D_optInfo_���_��ԭѧ���_���_4 = 4
    e_D_optInfo_���_��ԭѧ���_δ���_5 = 5
    e_D_txtInfo_���_��ԭѧ�����ϸ����_25 = 25
    e_D_optInfo_���_ҩ������_δ��_6 = 6
    e_D_optInfo_���_ҩ������_��_7 = 7
    e_D_txtInfo_���_ҩ����������_24 = 24
    e_D_optInfo_���_ҩ������_���_8 = 8
    e_D_optInfo_���_ҩ������_�����_9 = 9
    e_D_txtInfo_���_��ҩǰ����_11 = 11
    e_D_txtInfo_���_��ҩǰ��ϸ������_13 = 13
    e_D_txtInfo_���_��ҩǰ������ϸ��_15 = 15
    e_D_txtInfo_���_��ҩǰC��Ӧ����_17 = 17
    e_D_txtInfo_���_��ҩǰ����ת��ø_19 = 19
    e_D_txtInfo_���_��ҩǰ����_21 = 21
    e_D_txtInfo_���_��ҩ������_55 = 55
    e_D_txtInfo_���_��ҩ���ϸ������_53 = 53
    e_D_txtInfo_���_��ҩ��������ϸ��_51 = 51
    e_D_txtInfo_���_��ҩ��C��Ӧ����_49 = 49
    e_D_txtInfo_���_��ҩ�����ת��ø_45 = 45
    e_D_txtInfo_���_��ҩ����_43 = 43
    e_D_txtInfo_���_��ҩǰ��������_12 = 12
    e_D_txtInfo_���_��ҩǰ��ϸ����������_14 = 14
    e_D_txtInfo_���_��ҩǰ������ϸ������_16 = 16
    e_D_txtInfo_���_��ҩǰC��Ӧ��������_18 = 18
    e_D_txtInfo_���_��ҩǰ����ת��ø����_20 = 20
    e_D_txtInfo_���_��ҩǰ��������_22 = 22
    e_D_txtInfo_���_��ҩ����������_54 = 54
    e_D_txtInfo_���_��ҩ���ϸ����������_52 = 52
    e_D_txtInfo_���_��ҩ��������ϸ������_50 = 50
    e_D_txtInfo_���_��ҩ��C��Ӧ��������_48 = 48
    e_D_txtInfo_���_��ҩ�����ת��ø����_44 = 44
    e_D_txtInfo_���_��ҩ��������_42 = 42
    
    e_D_lblInfo_Ӱ��_���_102 = 102
    e_D_chkInfo_Ӱ��_X��_18 = 18
    e_D_chkInfo_Ӱ��_CT_19 = 19
    e_D_chkInfo_Ӱ��_�Ź���_20 = 20
    e_D_txtInfo_Ӱ��_��λ_46 = 46
    e_D_txtInfo_Ӱ��_����_47 = 47
    
    e_D_lblInfo_֢״_���_101 = 101
    e_D_txtInfo_֢״_��Ⱦ֢״_35 = 35
    
    e_D_picDiff_Ӱ��And֢״����_0 = 0
    e_D_picDiff_�����������_1 = 1
    
    e_D_lblInfo_��ҩĿ��_���_60 = 60
    e_D_lblInfo_��ҩĿ��_��Ⱦ���_68 = 68
    e_D_txtInfo_��ҩĿ��_��Ⱦ���_36 = 36
    e_D_optInfo_��ҩĿ��_δ��_17 = 17
    e_D_optInfo_��ҩĿ��_Ԥ��_18 = 18
    e_D_optInfo_��ҩĿ��_����_21 = 21
    
    e_D_picComm_��ҩĿ������_1 = 1
    e_D_picComm_��ҩ��ϸ����������_0 = 0
    e_D_lblInfo_��ҩ���_���_45 = 45
    e_D_txtInfo_��ҩ���_����_29 = 29
    e_D_txtInfo_��ҩ���_����_30 = 30
    
    e_D_lblInfo_����_���_46 = 46
    e_D_txtInfo_����_�ܷ�_31 = 31
    e_D_txtInfo_����_ҩ��_32 = 32
    e_D_txtInfo_����_��ҩ��_33 = 33
    
    e_D_lblInfo_���_���_47 = 47
    e_D_optInfo_���_����_12 = 12
    e_D_optInfo_���_��ת_13 = 13
    e_D_optInfo_���_��Ч_14 = 14
    e_D_optInfo_���_��Ⱦ_��_15 = 15
    e_D_optInfo_���_��Ⱦ_��_16 = 16
    
    e_D_optInfo_���_���ҩ_��_19 = 19
    e_D_optInfo_���_���ҩ_��_20 = 20
    
    e_D_lblInfo_��Ժ����_���_64 = 64
    e_D_chkInfo_����_��Ժ_��Ӧ֢_5 = 5
    e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12 = 12
    e_D_chkInfo_����_��Ժ_���μ���_13 = 13
    e_D_chkInfo_����_��Ժ_ÿ�ո�ҩƵ��_14 = 14
    e_D_chkInfo_����_��Ժ_�ܼ�_15 = 15
    e_D_chkInfo_����_��Ժ_��ҩ;��_16 = 16
    e_D_chkInfo_����_��Ժ_��ҩ�Ƴ�_17 = 17
    e_D_chkInfo_����_��Ժ_����ҩ��_21 = 21
    e_D_chkInfo_����_��Ժ_������ҩ_22 = 22
    e_D_lblInfo_����_��Ժ_Χ������ǩ_59 = 59
    e_D_chkInfo_����_��Ժ_��ǰ_23 = 23
    e_D_chkInfo_����_��Ժ_����_24 = 24
    e_D_chkInfo_����_��Ժ_����_25 = 25
    
    e_D_lblInfo_��������_���_65 = 65
    e_D_chkInfo_����_����_��Ӧ֢_37 = 37
    e_D_chkInfo_����_����_ҩ��ѡ��_36 = 36
    e_D_chkInfo_����_����_���μ���_35 = 35
    e_D_chkInfo_����_����_ÿ�ո�ҩƵ��_34 = 34
    e_D_chkInfo_����_����_�ܼ�_33 = 33
    e_D_chkInfo_����_����_��ҩ;��_32 = 32
    e_D_chkInfo_����_����_��ҩ�Ƴ�_31 = 31
    e_D_chkInfo_����_����_����ҩ��_30 = 30
    e_D_chkInfo_����_����_������ҩ_29 = 29
    e_D_lblInfo_����_����_Χ������ǩ_70 = 70
    e_D_chkInfo_����_����_��ǰ_28 = 28
    e_D_chkInfo_����_����_����_27 = 27
    e_D_chkInfo_����_����_����_26 = 26
    
    e_D_lblInfo_��ע_���_66 = 66
    e_D_txtInfo_��ע_57 = 57
    
    e_D_lblInfo_˵��_���_69 = 69
    e_D_lblInfo_˵��_����˵���ı�_73 = 73
    e_D_lblInfo_˵��_������˵���ı�_74 = 74
    
    Type���� = 0
    Type�ı� = 1
    Type���� = 2
End Enum

Private Enum COL_VSF '��ҩ�����
    COL_������Ŀ = 0
    COL_��������
    COL_����Ժ
    COL_��������
    COL_����������
    COL_������Ժ
    COL_����������
    COL_�������
    COL_���������
    COL_�ϼ�����
    COL_�ϼ�����
End Enum

Private Enum COL_VSDRUGUSE '����ҩʹ����ϸ���
    COL_DRUG_ͼ�� = 0
    COL_DRUG_ID = 1
    COL_DRUG_���ID
    COL_DRUG_ҩ��ID
    
    COL_DRUG_ҩƷ����
    COL_DRUG_���μ���
    COL_DRUG_��ҩƵ��
    COL_DRUG_;��
    COL_DRUG_������
    COL_DRUG_��ֹʱ��
End Enum

Private Enum COL_VSOPERATE '������������ϸ
    COL_OPE_����ID = 1
    COL_OPE_��������
    COL_OPE_�п�
    COL_OPE_��ʼʱ��
    COL_OPE_����ʱ��
    COL_OPE_��ҩ�ڼ�
    COL_OPE_��ҩ���
End Enum

Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mblnChange As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mlng��ҳID  As Long
Private mlng���  As Long
Private mbln���� As Boolean
Private mblnReturn As Boolean
Private mbln��ҩ���� As Boolean
Private mstr���� As String
Private mlng����  As Long
Private mOldwinproc As Long
Private mbln�༭ As Boolean
Private mbln��ӡ As Boolean
Private mlngYear As Long
Private mblnInitSaved As Boolean '��ʾ�Ƿ������༭���棬һ������֮��ᱣ��һ��������ݣ���ҩ�����������沿�ֻ�����Ϣ
Private mrsCtl As ADODB.Recordset

Public Function ShowMe(frmParent As Object, ByVal lng����id As Long, ByVal lng����ID, ByVal lng��ҳID, ByVal lng��� As Long, ByVal lng���� As Long, ByVal str���� As String, _
            ByVal bln���� As Boolean, ByRef bln�༭ As Boolean, ByRef bln��ӡ As Boolean) As Boolean
'���ܣ�
'������bln���� �Ƿ�����������
    mlng����ID = lng����id
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mbln���� = bln����
    mlng��� = lng���
     
    mstr���� = str����
    mlng���� = lng����
    mbln�༭ = bln�༭
    mbln��ӡ = bln��ӡ
    
    frmKssSurveyEdit.Show 1, frmParent
    
    bln�༭ = mbln�༭
    bln��ӡ = mbln��ӡ
    
    ShowMe = True
End Function

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " ����(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " �����˳�(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&X)"): objControl.BeginGroup = True
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
        Case e_D_chkInfo_����_��Ժ_��Ӧ֢_5
            blnEdit = chkInfo(Index).Value
            For i = e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12 To e_D_chkInfo_����_��Ժ_����_25
                If InStr(",18,19,20,", "," & i & ",") = 0 Then
                    chkInfo(i).Enabled = blnEdit
                    chkInfo(i).Value = 0
                End If
            Next
        Case e_D_chkInfo_����_����_��Ӧ֢_37
            blnEdit = chkInfo(Index).Value
            For i = e_D_chkInfo_����_����_����_26 To e_D_chkInfo_����_����_ҩ��ѡ��_36
                chkInfo(i).Enabled = blnEdit
                chkInfo(i).Value = 0
            Next
    End Select
    If mbln��ҩ���� Then Call Cls������ϸ(Index)
    If Visible Then mblnChange = True
End Sub

Private Sub cmdInfect_Click(Index As Integer)
    '���ܣ�������¼ѡ����---------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnCanle As Boolean
    Dim X As Long, Y As Long
    Dim vRect As RECT
    Dim lngHwnd As Long
    
    If Index = 0 Then
        lngHwnd = txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).hwnd
    Else
        lngHwnd = txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).hwnd
    End If
    GetWindowRect lngHwnd, vRect
    X = vRect.Left * Screen.TwipsPerPixelX
    Y = vRect.Top * Screen.TwipsPerPixelY
    strSQL = "Select id,������� as ��Ⱦ��� From ������ϼ�¼ Where ������� =5 and ����id=[1] and ��ҳid=[2] order by ��ϴ���"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ�������¼", False, "", "", False, False, True, X, Y, 200, blnCanle, False, True, mlng����ID, mlng��ҳID)
    If blnCanle Then Exit Sub
    If rsTmp Is Nothing Then
        MsgBox "�ó�Ժ������ҳ����дδ��д��Ⱦ��ϡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Index = 0 Then
        txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Text = rsTmp!��Ⱦ���
        txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Tag = rsTmp!ID
    Else
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Text = rsTmp!��Ⱦ���
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Tag = rsTmp!ID
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strB As String
    Dim strE As String
    Dim bln�ÿ���ҩ As Boolean
    
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
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "ʹ����������", picUse.hwnd, 0).Tag = "ʹ����������"
        .InsertItem(1, "���������۱�", picReasonable.hwnd, 0).Tag = "���������۱�"
        .Item(0).Selected = True
    End With
    
    Call InitCommandBar
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    strSQL = "select ��Χ��ʼʱ�� as ��ʼʱ��,��Χ����ʱ�� as ����ʱ�� from ����ҩ�������¼ where id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    
    
    txtInfo(e_F_txtInfo_����ʱ��_��_0).Text = Format(rsTmp!��ʼʱ��, "yyyy-mm-dd")
    txtInfo(e_F_txtInfo_����ʱ��_ֹ_1).Text = Format(rsTmp!����ʱ��, "yyyy-mm-dd")
    txtInfo(e_F_txtInfo_��������_2).Text = mstr����
    txtInfo(e_F_txtInfo_��Ժ����_4).Text = mlng����
    txtInfo(e_F_txtInfo_���_5).Text = mlng���
    
    strB = Format(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, "yyyy-06-01")
    strE = Format(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, "yyyy-06-30")
    
    mbln��ҩ���� = False
    
    If Not mbln��ҩ���� Then mbln��ҩ���� = Between(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, strB, strE)
    
    If Not mbln��ҩ���� Then mbln��ҩ���� = Between(txtInfo(e_F_txtInfo_����ʱ��_ֹ_1).Text, strB, strE)
    
    If Not mbln��ҩ���� Then
        mbln��ҩ���� = (strB >= txtInfo(e_F_txtInfo_����ʱ��_��_0).Text And strE <= txtInfo(e_F_txtInfo_����ʱ��_ֹ_1).Text)
    End If
    
    strB = Format(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, "yyyy-12-01")
    strE = Format(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, "yyyy-12-31")
    
    If Not mbln��ҩ���� Then mbln��ҩ���� = Between(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, strB, strE)
    
    If Not mbln��ҩ���� Then mbln��ҩ���� = Between(txtInfo(e_F_txtInfo_����ʱ��_ֹ_1).Text, strB, strE)
    
    If Not mbln��ҩ���� Then
        mbln��ҩ���� = (strB >= txtInfo(e_F_txtInfo_����ʱ��_��_0).Text And strE <= txtInfo(e_F_txtInfo_����ʱ��_ֹ_1).Text)
    End If
    
    tbcSub.Item(1).Visible = mbln��ҩ����
    If mbln��ҩ���� Then
        txtInfo(e_P_txtInfo_����ʱ��_��_28).Text = txtInfo(e_F_txtInfo_����ʱ��_��_0).Text
        txtInfo(e_P_txtInfo_����ʱ��_ֹ_56).Text = txtInfo(e_F_txtInfo_����ʱ��_ֹ_1).Text
        Call Init��ҩ���۱�
    End If
    '��ʼ������ҩ���
    strTmp = "ID;���ID;ҩ��ID;ҩƷ����,3500,1;���μ���,1200,7;��ҩƵ��,1200,1;;��,1200,1;������,800,7;��ֹʱ��(����ʱ��),2530,4"
    Call InitTable(vsDrugUse, strTmp)
    
    Call LoadData
    
    If mblnInitSaved Then
        bln�ÿ���ҩ = UseKssDrug
    Else
        bln�ÿ���ҩ = Not optInfo(e_D_optInfo_��ҩĿ��_δ��_17).Value
    End If
    
    If bln�ÿ���ҩ Then Call LoadDrugUse
 
    If mbln���� Then
        strTmp = "����ID;��������,3000,1;�п�,600,4;������ʼʱ��,1700,4;��������ʱ��,1700,4;��ǰ����Ԥ����ҩʱ��,2500,1;���и�ҩ,1000,1"
        Call InitTable(vsOperate, strTmp)
        Call LoadOperate
    End If
    Call LoadFee
    Call SetCtlProperty
    Call optInfo_Click(e_D_optInfo_���_��ԭѧ���_δ��_2)
    Call optInfo_Click(e_D_optInfo_���_��ԭѧ���_δ���_5)
    Call optInfo_Click(e_D_optInfo_���_ҩ������_δ��_6)
    Call optInfo_Click(e_D_optInfo_��ҩĿ��_δ��_17)
    Call optInfo_Click(e_D_optInfo_��ҩĿ��_Ԥ��_18)
    If mbln��ҩ���� Then Call Load��ҩ����
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
    
    If mbln���� Then
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
    Case conMenu_File_Preview 'Ԥ��
        Call Print�����(1)
    Case conMenu_File_Print '��ӡ
        Call Print�����(2)
    Case conMenu_File_Exit '�˳�
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

Private Sub Print�����(ByVal bytType As Byte)
'���ܣ���ӡ��Ԥ������ bytType=1Ԥ����=2��ӡ
'˵����ӡ֮ǰҪ�ȱ���һЩ���ݡ�
    Dim strRPTName As String
    
    If mbln���� Then
        strRPTName = "ZL1_INSIDE_1269_2"
    Else
        strRPTName = "ZL1_INSIDE_1269_1"
    End If
    
    If mblnChange Then
        mbln�༭ = True
        If Not CheckData Then Exit Sub
        Call SaveData
    ElseIf mblnInitSaved Then
        If Not CheckData Then Exit Sub
        Call SaveData(True)
    End If
    vsPJB.Redraw = flexRDNone
    If bytType = 1 Then
        Call mobjReport.ReportOpen(gcnOracle, 100, strRPTName, Me, "����ID=" & mlng����ID, "���=" & mlng���, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "ReportFormat=" & IIf(tbcSub.Selected.Index = 0, 1, 2), 1)
    ElseIf bytType = 2 Then
        Call mobjReport.ReportOpen(gcnOracle, 100, strRPTName, Me, "����ID=" & mlng����ID, "���=" & mlng���, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "ReportFormat=" & IIf(tbcSub.Selected.Index = 0, 1, 2), 2)
    End If
    vsPJB.Redraw = flexRDDirect
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsc.Value
    lngMin = vsc.Min
    lngMax = vsc.Max
    
    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur - (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMin
        End If
    End If
End Sub

Private Sub Form_Activate()
'������
    Call Form_Resize
    glngPreHWnd = GetWindowLong(picDCB.hwnd, GWL_WNDPROC)
    SetWindowLong picDCB.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
'������
    SetWindowLong picDCB.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_Resize()
    Call picDCB_Resize
    Call picReasonable_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("��ǰ�������������Ѿ������˵�����δ���棬�Ƿ�Ҫ�����˳���", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'�ı��ӡ״̬
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    If mbln��ӡ Then Exit Sub '������Ѿ�����ľͲ����ظ�ִ��
    mbln��ӡ = True
    On Error GoTo errH
    strSQL = Get������ϸSQL
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
    Case e_D_optInfo_���_��ԭѧ���_δ��_2, e_D_optInfo_���_��ԭѧ���_��_3  '��ԭѧ��� �� δ��
        blnTmp = optInfo(e_D_optInfo_���_��ԭѧ���_δ��_2).Value  'δ��
        optInfo(e_D_optInfo_���_��ԭѧ���_���_4).Enabled = Not blnTmp
        optInfo(e_D_optInfo_���_��ԭѧ���_δ���_5).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).Enabled = Not blnTmp 'δ��
        txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).Text = ""
            txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).Text = ""
            txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Text = ""
            txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
            txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
            txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
        Else
            txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
            txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
            txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
        End If
        Call optInfo_Click(4)
    Case e_D_optInfo_���_��ԭѧ���_���_4, e_D_optInfo_���_��ԭѧ���_δ���_5  '���ϸ�� δ��� ���
        blnTmp = optInfo(e_D_optInfo_���_��ԭѧ���_���_4).Value  'δ���
        txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Text = ""
            txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
        Else
            txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
        End If
    Case e_D_optInfo_���_ҩ������_δ��_6, e_D_optInfo_���_ҩ������_��_7  ' ҩ������ �� δ��
        blnTmp = optInfo(e_D_optInfo_���_ҩ������_δ��_6).Value  'δ��
        optInfo(e_D_optInfo_���_ҩ������_���_8).Enabled = Not blnTmp
        optInfo(e_D_optInfo_���_ҩ������_�����_9).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_���_ҩ����������_24).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_���_ҩ����������_24).Text = ""
            txtInfo(e_D_txtInfo_���_ҩ����������_24).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
        Else
            txtInfo(e_D_txtInfo_���_ҩ����������_24).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
        End If
    Case e_D_optInfo_��ҩĿ��_δ��_17, e_D_optInfo_��ҩĿ��_Ԥ��_18, e_D_optInfo_��ҩĿ��_����_21 '��ҩĿ�ģ� δ��ҩ��Ԥ��������
        blnTmp = optInfo(e_D_optInfo_��ҩĿ��_δ��_17).Value  'δ��ҩ
        optInfo(e_D_optInfo_��ҩĿ��_δ��_17).Enabled = blnTmp
        optInfo(e_D_optInfo_��ҩĿ��_Ԥ��_18).Enabled = Not blnTmp
        optInfo(e_D_optInfo_��ҩĿ��_����_21).Enabled = Not blnTmp
        cmdInfect(1).Enabled = Not blnTmp
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Enabled = Not blnTmp
        If blnTmp Then
            txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Text = ""
            txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Tag = ""
            txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
        Else
            txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
        End If
        
        If Index = e_D_optInfo_��ҩĿ��_Ԥ��_18 Or Index = e_D_optInfo_��ҩĿ��_����_21 Then
            blnTmp = optInfo(e_D_optInfo_��ҩĿ��_Ԥ��_18).Value
            cmdInfect(1).Enabled = Not blnTmp
            txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Enabled = Not blnTmp
            If blnTmp Then
                txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Text = ""
                txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Tag = ""
                txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).BackColor = txtInfo(e_F_txtInfo_������_3).BackColor
            Else
                txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).BackColor = txtInfo(e_D_txtInfo_��ע_57).BackColor
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
    
    If mbln���� Then
        picComm(e_D_picComm_��ҩĿ������_1).Move lngL, lngT, lngW, picComm(e_D_picComm_��ҩĿ������_1).Height
        picDiff(e_D_picDiff_�����������_1).Move lngL, lngT + picComm(e_D_picComm_��ҩĿ������_1).Height - 10, lngW, picDiff(e_D_picDiff_�����������_1).Height
        picComm(e_D_picComm_��ҩ��ϸ����������_0).Move lngL, picDiff(e_D_picDiff_�����������_1).Height + picDiff(e_D_picDiff_�����������_1).Top - 10, lngW, picComm(e_D_picComm_��ҩ��ϸ����������_0).Height
    Else
        picDiff(e_D_picDiff_Ӱ��And֢״����_0).Move lngL, lngT, lngW, picDiff(e_D_picDiff_Ӱ��And֢״����_0).Height
        picComm(e_D_picComm_��ҩĿ������_1).Move lngL, lngT + picDiff(e_D_picDiff_Ӱ��And֢״����_0).Height - 10, lngW, picComm(e_D_picComm_��ҩĿ������_1).Height
        picComm(e_D_picComm_��ҩ��ϸ����������_0).Move lngL, picComm(e_D_picComm_��ҩĿ������_1).Height + picComm(e_D_picComm_��ҩĿ������_1).Top - 10, lngW, picComm(e_D_picComm_��ҩ��ϸ����������_0).Height
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
''''��ý���
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
    Case e_D_txtInfo_���_��ԭѧ�������_23, e_D_txtInfo_���_ҩ����������_24, e_D_txtInfo_���_��ҩǰ��������_12, e_D_txtInfo_���_��ҩǰ��ϸ����������_14, e_D_txtInfo_���_��ҩǰ������ϸ������_16, _
    e_D_txtInfo_���_��ҩǰC��Ӧ��������_18, e_D_txtInfo_���_��ҩǰ����ת��ø����_20, e_D_txtInfo_���_��ҩǰ��������_22, e_D_txtInfo_���_��ҩ����������_54, e_D_txtInfo_���_��ҩ���ϸ����������_52, _
    e_D_txtInfo_���_��ҩ��������ϸ������_50, e_D_txtInfo_���_��ҩ��C��Ӧ��������_48, e_D_txtInfo_���_��ҩ�����ת��ø����_44, e_D_txtInfo_���_��ҩ��������_42
        If InStr("0123456789/" & Chr(vbKeyBack) & Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Case e_D_txtInfo_���_��ҩǰ����_11, e_D_txtInfo_���_��ҩ������_55
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
    Case e_D_txtInfo_���_��ԭѧ�������_23, e_D_txtInfo_���_ҩ����������_24, e_D_txtInfo_���_��ҩǰ��������_12, e_D_txtInfo_���_��ҩǰ��ϸ����������_14, e_D_txtInfo_���_��ҩǰ������ϸ������_16, _
    e_D_txtInfo_���_��ҩǰC��Ӧ��������_18, e_D_txtInfo_���_��ҩǰ����ת��ø����_20, e_D_txtInfo_���_��ҩǰ��������_22, e_D_txtInfo_���_��ҩ����������_54, e_D_txtInfo_���_��ҩ���ϸ����������_52, _
    e_D_txtInfo_���_��ҩ��������ϸ������_50, e_D_txtInfo_���_��ҩ��C��Ӧ��������_48, e_D_txtInfo_���_��ҩ�����ת��ø����_44, e_D_txtInfo_���_��ҩ��������_42
        If InStr(strTmp, "/") = 0 Then
            strMsg = "���ڸ�ʽ���ԣ���ȷ��ʽ��12/22"
        Else
            intMM = Val(Split(strTmp, "/")(0))
            intDD = Val(Split(strTmp, "/")(1))
            If intMM = 0 Or intMM > 12 Then
                strMsg = "��д���·ݲ���ȷ��ֻ����1��12"
            Else
                If InStr(",1,3,5,7,8,10,12,", "," & intMM & ",") > 0 Then
                    intTmp = 31
                ElseIf InStr(",4,6,9,11,", "," & intMM & ",") > 0 Then
                    intTmp = 30
                ElseIf intMM = 2 Then
                    strTmp = Split(txtInfo(e_F_txtInfo_����ʱ��_��_0).Text, "-")(0) & "02"
                    intTmp = GetMonthMaxDay(strTmp)
                End If
                
                If intDD > intTmp Or intDD = 0 Then
                    strMsg = "��д�����ں�������ȷ��" & intMM & "�����" & intTmp & "�졣"
                End If
                If strMsg = "" Then '����Ӧ����סԺʱ�䷶Χ��
                    strDate = mlngYear & "/" & strTmp
                    strDate = Format(strDate, "YYYY-MM-DD")
                    If Not Between(strDate, txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Text, txtInfo(e_D_txtInfo_����_��Ժʱ��_10).Text) Then
                        strMsg = "��д������Ӧ�ڲ���סԺʱ�䷶Χ�ڣ�" & txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Text & "��" & txtInfo(e_D_txtInfo_����_��Ժʱ��_10).Text
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
'���ܣ�2�·ݵ�����

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim intMaxDay As Integer
    On Error GoTo errH
    
    strSQL = "Select ��ʼ����, ��ֹ���� From �ڼ�� Where �ڼ� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMonth)
    If Not rsTmp.EOF Then
        strTmp = Format(rsTmp!��ֹ���� & "", "yyyy-mm-dd")
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
'���ܣ����ÿؼ��Ŀɼ��ԣ������
    Dim objCtl As Object
    
    '�����񲼾�
    If mbln���� Then
        picDiff(e_D_picDiff_�����������_1).Visible = True
        picDiff(e_D_picDiff_Ӱ��And֢״����_0).Visible = False
        picComm(e_D_picComm_��ҩĿ������_1).Visible = True
        picComm(e_D_picComm_��ҩ��ϸ����������_0).Visible = True
        
        lblInfo(e_D_lblInfo_˵��_����˵���ı�_73).Visible = True
        lblInfo(e_D_lblInfo_˵��_������˵���ı�_74).Visible = False
        
        lblInfo(e_D_lblInfo_��ҩĿ��_���_60).Caption = "5"
        lblInfo(e_D_lblInfo_��ҩ���_���_45).Caption = "7"
        lblInfo(e_D_lblInfo_����_���_46).Caption = "8"
        lblInfo(e_D_lblInfo_���_���_47).Caption = "9"
        lblInfo(e_D_lblInfo_��Ժ����_���_64).Caption = "10"
        lblInfo(e_D_lblInfo_��������_���_65).Caption = "11"
        lblInfo(e_D_lblInfo_��ע_���_66).Caption = "12"
        lblInfo(e_D_lblInfo_˵��_���_69).Caption = "13"
        vsc.Max = 300
    Else
        picDiff(e_D_picDiff_�����������_1).Visible = False
        picDiff(e_D_picDiff_Ӱ��And֢״����_0).Visible = True
        picComm(e_D_picComm_��ҩĿ������_1).Visible = True
        picComm(e_D_picComm_��ҩ��ϸ����������_0).Visible = True
        
        chkInfo(e_D_chkInfo_����_��Ժ_��ǰ_23).Visible = False
        chkInfo(e_D_chkInfo_����_��Ժ_����_24).Visible = False
        chkInfo(e_D_chkInfo_����_��Ժ_����_25).Visible = False
    
        chkInfo(e_D_chkInfo_����_����_��ǰ_28).Visible = False
        chkInfo(e_D_chkInfo_����_����_����_27).Visible = False
        chkInfo(e_D_chkInfo_����_����_����_26).Visible = False
        
        lblInfo(e_D_lblInfo_����_��Ժ_Χ������ǩ_59).Visible = False
        lblInfo(e_D_lblInfo_����_����_Χ������ǩ_70).Visible = False
        
        lblInfo(e_D_lblInfo_˵��_����˵���ı�_73).Visible = False
        lblInfo(e_D_lblInfo_˵��_������˵���ı�_74).Visible = True
        
        lblN.Caption = "���������˿���ҩ��ʹ����������"
        lblInfo(e_F_lblInfo_��Ժ������ǩ_4).Caption = "���������˳�Ժ������"
        lblYJ.Caption = "��������ҩ���������������"
        vsc.Max = 250
    End If
    
    '�ؼ�����ɫ
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
            Case "OPTIONBUTTON", "CHECKBOX", "LABEL"
                objCtl.BackColor = picDCB.BackColor
        End Select
    Next
    '������ʼ�߶�
    picDCB.Top = 400
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim lng��Ⱦ���ID As Long
    Dim str��Ⱦ��� As String
    Dim strTmp As String
    Dim str��� As String
    Dim lngTmp As Long
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select a.���ƽ��, a.��Ӧ֢, a.ҩ��ѡ��, a.���μ���, a.ÿ�ո�ҩƵ��, a.�ܼ�, a.��ҩ;��, a.��ҩ�Ƴ�, a.��ǰ��ҩʱ��, a.������ҩ, a.������ҩ, a.������ҩ, a.����ҩ��, a.��ע," & vbNewLine & _
        "a.�Ƿ��ӡ, a.�Ƿ�༭, a.��ҩ����, a.����ҩ����, a.�Ƿ�����, a.��ԭѧ���, a.��ԭѧ�������, a.��ԭѧ���걾, a.��ԭѧ�����ϸ����, a.ҩ������, a.ҩ����������," & vbNewLine & _
        "a.ҩ�������Ƿ����, a.��ҩǰ����, a.��ҩǰ��ϸ������, a.��ҩǰ������ϸ��, a.��ҩǰc��Ӧ����, a.��ҩǰ����ת��ø, a.��ҩǰ����, a.��ҩ������, a.��ҩ���ϸ������, a.��ҩ��������ϸ��," & vbNewLine & _
        "a.��ҩ��c��Ӧ����, a.��ҩ�����ת��ø, a.��ҩ����, a.��ҩǰ��������, a.��ҩǰ��ϸ����������, a.��ҩǰ������ϸ������, a.��ҩǰc��Ӧ��������, a.��ҩǰ����ת��ø����, a.��ҩǰ��������," & vbNewLine & _
        "a.��ҩ����������, a.��ҩ���ϸ����������, a.��ҩ��������ϸ������, a.��ҩ��c��Ӧ��������, a.��ҩ�����ת��ø����, a.��ҩ��������, a.Ӱ��ѧ���, a.Ӱ��ѧ��ϲ�λ, a.Ӱ��ѧ��Ͻ���," & vbNewLine & _
        "a.�ٴ�֢״, a.��ҩĿ��, a.��Ⱦ���,a.�Ƿ��ÿ����ҩ,b.�Ա�,b.����,b.סԺ�� as ������,b.��Ժ���� as ��Ժʱ��,b.��Ժ���� as ��Ժʱ��,b.����" & vbNewLine & _
        "From ����ҩ�������ϸ A,������ҳ b" & vbNewLine & _
        "Where a.����id=b.����id and a.��ҳid=b.��ҳid and a.����id = [1] And a.��� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng���)
    
    '�����һ��֮�󣬸��ֶ�Ҫô��1��0�������δ���������ֶ�ΪNull
    mblnInitSaved = rsTmp!�Ƿ����� & "" = ""
    
    txtInfo(e_F_txtInfo_������_3).Text = rsTmp!������ & ""
    
    txtInfo(e_D_txtInfo_����_�Ա�_6).Text = rsTmp!�Ա� & ""
    txtInfo(e_D_txtInfo_����_����_7).Text = rsTmp!���� & ""
    
    '������Ϣ
    If Val(rsTmp!���� & "") > 0 Then
        txtInfo(e_D_txtInfo_����_����_8).Text = Val(rsTmp!���� & "") & "Kg"
    Else
        strSQL = "Select b.��Ŀ��λ as ��λ, b.��Ŀ���� as ��Ϣ��, b.��¼���� as ��Ϣֵ" & _
            " From ���˻����¼ A, ���˻������� B Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2] and b.��Ŀ����='����'"
        Set rsOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsOther.EOF Then txtInfo(e_D_txtInfo_����_����_8).Text = rsOther!��Ϣֵ & rsOther!��λ & ""
        Set rsOther = Nothing
    End If
    
    txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Text = Format(rsTmp!��Ժʱ�� & "", "yyyy-MM-dd")
    txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Tag = Format(rsTmp!��Ժʱ�� & "", "yyyy-MM-dd HH:mm")
    txtInfo(e_D_txtInfo_����_��Ժʱ��_10).Text = Format(rsTmp!��Ժʱ�� & "", "yyyy-MM-dd")
    txtInfo(e_D_txtInfo_����_��Ժʱ��_10).Tag = Format(rsTmp!��Ժʱ�� & "", "yyyy-MM-dd HH:mm")
    
    mlngYear = Val(Split(txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Text, "-")(0))
    
    optInfo(e_D_optInfo_���_��ԭѧ���_δ��_2).Value = Val(rsTmp!��ԭѧ��� & "") = 0
    optInfo(e_D_optInfo_���_��ԭѧ���_δ��_2 + 1).Value = Val(rsTmp!��ԭѧ��� & "") = 1
    
    If rsTmp!��ԭѧ������� & "" <> "" Then
        strTmp = Format(rsTmp!��ԭѧ�������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).Text = rsTmp!��ԭѧ���걾 & ""
    
    optInfo(e_D_optInfo_���_��ԭѧ���_���_4).Value = True
    optInfo(e_D_optInfo_���_��ԭѧ���_���_4 + 1).Value = False
    If rsTmp!��ԭѧ�����ϸ���� & "" <> "" Then
        optInfo(e_D_optInfo_���_��ԭѧ���_���_4).Value = False
        optInfo(e_D_optInfo_���_��ԭѧ���_���_4 + 1).Value = True
        txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Text = rsTmp!��ԭѧ�����ϸ���� & ""
    End If
    optInfo(e_D_optInfo_���_ҩ������_δ��_6).Value = Val(rsTmp!ҩ������ & "") = 0
    optInfo(e_D_optInfo_���_ҩ������_δ��_6 + 1).Value = Val(rsTmp!ҩ������ & "") = 1
    
    If rsTmp!ҩ���������� & "" <> "" Then
        strTmp = Format(rsTmp!ҩ����������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_ҩ����������_24).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_ҩ����������_24).Tag = strTmp
    End If
    
    optInfo(e_D_optInfo_���_ҩ������_���_8).Value = Val(rsTmp!ҩ�������Ƿ���� & "") = 0
    optInfo(e_D_optInfo_���_ҩ������_���_8 + 1).Value = Val(rsTmp!ҩ�������Ƿ���� & "") = 1
    
    txtInfo(e_D_txtInfo_���_��ҩǰ����_11).Text = rsTmp!��ҩǰ���� & ""
    If rsTmp!��ҩǰ�������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩǰ��������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩǰ��������_12).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩǰ��������_12).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ������_13).Text = rsTmp!��ҩǰ��ϸ������ & ""
    If rsTmp!��ҩǰ��ϸ���������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩǰ��ϸ����������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ����������_14).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ����������_14).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ��_15).Text = rsTmp!��ҩǰ������ϸ�� & ""
    If rsTmp!��ҩǰ������ϸ������ & "" <> "" Then
        strTmp = Format(rsTmp!��ҩǰ������ϸ������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ������_16).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ������_16).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ����_17).Text = rsTmp!��ҩǰc��Ӧ���� & ""
    If rsTmp!��ҩǰc��Ӧ�������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩǰc��Ӧ��������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ��������_18).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ��������_18).Tag = strTmp
    End If
     
    txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø_19).Text = rsTmp!��ҩǰ����ת��ø & ""
    If rsTmp!��ҩǰ����ת��ø���� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩǰ����ת��ø����, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø����_20).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø����_20).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩǰ����_21).Text = rsTmp!��ҩǰ���� & ""
    If rsTmp!��ҩǰ�������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩǰ��������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩǰ��������_22).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩǰ��������_22).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩ������_55).Text = rsTmp!��ҩ������ & ""
    If rsTmp!��ҩ���������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩ����������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩ����������_54).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩ����������_54).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩ���ϸ������_53).Text = rsTmp!��ҩ���ϸ������ & ""
    If rsTmp!��ҩ���ϸ���������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩ���ϸ����������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩ���ϸ����������_52).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩ���ϸ����������_52).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩ��������ϸ��_51).Text = rsTmp!��ҩ��������ϸ�� & ""
    If rsTmp!��ҩ��������ϸ������ & "" <> "" Then
        strTmp = Format(rsTmp!��ҩ��������ϸ������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩ��������ϸ������_50).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩ��������ϸ������_50).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ����_49).Text = rsTmp!��ҩ��C��Ӧ���� & ""
    If rsTmp!��ҩ��C��Ӧ�������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩ��C��Ӧ��������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ��������_48).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ��������_48).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø_45).Text = rsTmp!��ҩ�����ת��ø & ""
    If rsTmp!��ҩ�����ת��ø���� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩ�����ת��ø����, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø����_44).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø����_44).Tag = strTmp
    End If
    
    txtInfo(e_D_txtInfo_���_��ҩ����_43).Text = rsTmp!��ҩ���� & ""
    If rsTmp!��ҩ�������� & "" <> "" Then
        strTmp = Format(rsTmp!��ҩ��������, "yyyy-MM-dd")
        txtInfo(e_D_txtInfo_���_��ҩ��������_42).Text = DateForShow(strTmp)
        txtInfo(e_D_txtInfo_���_��ҩ��������_42).Tag = strTmp
    End If
    
    strTmp = rsTmp!Ӱ��ѧ��� & ""
    If Len(strTmp) = 3 Then
        chkInfo(e_D_chkInfo_Ӱ��_X��_18).Value = Mid(strTmp, 1, 1)
        chkInfo(e_D_chkInfo_Ӱ��_CT_19).Value = Mid(strTmp, 2, 1)
        chkInfo(e_D_chkInfo_Ӱ��_�Ź���_20).Value = Mid(strTmp, 3, 1)
    End If
    
    txtInfo(e_D_txtInfo_Ӱ��_��λ_46).Text = rsTmp!Ӱ��ѧ��ϲ�λ & ""
    txtInfo(e_D_txtInfo_Ӱ��_����_47).Text = rsTmp!Ӱ��ѧ��Ͻ��� & ""
    
    '��ҩĿ��
    lngTmp = Val(rsTmp!��ҩĿ�� & "")
    optInfo(e_D_optInfo_��ҩĿ��_δ��_17 + lngTmp).Value = True
    
    '��ҩ����
    txtInfo(e_D_txtInfo_��ҩ���_����_30).Text = Val(rsTmp!��ҩ���� & "")
    
    '����ҩ����
    txtInfo(e_D_txtInfo_��ҩ���_����_29).Text = Val(rsTmp!����ҩ���� & "")
    
    optInfo(e_D_optInfo_���_���ҩ_��_20).Value = Val(rsTmp!�Ƿ��ÿ����ҩ & "") <> 0
    
    '��ҩ����
    strTmp = rsTmp!��Ӧ֢ & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_��Ӧ֢_5).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_��Ӧ֢_37).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!ҩ��ѡ�� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_ҩ��ѡ��_36).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!���μ��� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_���μ���_13).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_���μ���_35).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!ÿ�ո�ҩƵ�� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_ÿ�ո�ҩƵ��_14).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_ÿ�ո�ҩƵ��_34).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!�ܼ� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_�ܼ�_15).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_�ܼ�_33).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!��ҩ;�� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_��ҩ;��_16).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_��ҩ;��_32).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!��ҩ�Ƴ� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_��ҩ�Ƴ�_17).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_��ҩ�Ƴ�_31).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!��ǰ��ҩʱ�� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_��ǰ_23).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_��ǰ_28).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!������ҩ & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_����_24).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_����_27).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!������ҩ & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_����_25).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_����_26).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!������ҩ & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_������ҩ_22).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_������ҩ_29).Value = Val(Split(strTmp, "|")(1))
    End If
    
    strTmp = rsTmp!����ҩ�� & ""
    If strTmp <> "" Then
        chkInfo(e_D_chkInfo_����_��Ժ_����ҩ��_21).Value = Val(Split(strTmp, "|")(0))
        chkInfo(e_D_chkInfo_����_����_����ҩ��_30).Value = Val(Split(strTmp, "|")(1))
    End If
    
    txtInfo(e_D_txtInfo_��ע_57).Text = rsTmp!��ע & ""
    
    '����ʷ
    strSQL = "Select ҩ���� From ���˹�����¼ Where ����id = [1] And ��ҳid = [2]"
    Set rsOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    optInfo(e_D_optInfo_����_��_0).Value = rsOther.EOF
    optInfo(e_D_optInfo_����_��_0 + 1).Value = Not rsOther.EOF
    
    If Not rsOther.EOF Then
        For i = 1 To rsOther.RecordCount
            strTmp = strTmp & "," & rsOther!ҩ����
            rsOther.MoveNext
        Next
        txtInfo(e_D_txtInfo_����_ͨ����_34).Text = Mid(strTmp, 2): strTmp = ""
    End If
    
'----------------------------------------------------------------------------------------------------------------------------------------------
    
    '��� ��ȡ��������� ȡ��ҳ�����е����
    strSQL = "Select ID,��¼��Դ,�������,�������,��Ժ��� From ������ϼ�¼ Where ����id = [1] and ��ҳid =[2]  And NVL(�������,1) = 1 order by ��¼��Դ,��ϴ���"
    Set rs��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    '1.���  ��ҳ�е����г�Ժ���
    rs���.Filter = "(�������=3 and ��¼��Դ=3) or (�������=13 and ��¼��Դ=3)"
    For i = 1 To rs���.RecordCount
        str��� = str��� & "," & i & "." & rs���!�������
        If InStr("," & strTmp & ",", "," & rs���!��Ժ��� & ",") = 0 Then
            strTmp = strTmp & "," & rs���!��Ժ���
        End If
        rs���.MoveNext
    Next
    '���
    txtInfo(e_D_txtInfo_���_���_27).Text = IIf("" = str���, "��", Mid(str���, 2))
    '���ƽ��
    optInfo(e_D_optInfo_���_��ת_13).Value = False
    optInfo(e_D_optInfo_���_����_12).Value = False
    optInfo(e_D_optInfo_���_��Ч_14).Value = False
    If mbln�༭ Then
        If InStr(",1,2,3,", "," & rsTmp!���ƽ�� & ",") > 0 Then
            lngTmp = 11 + Val(rsTmp!���ƽ�� & "")
            optInfo(lngTmp).Value = True
        End If
    Else
        If InStr(strTmp, "��ת") > 0 Then
            optInfo(e_D_optInfo_���_��ת_13).Value = True
        ElseIf InStr(strTmp, "����") > 0 Then
            optInfo(e_D_optInfo_���_����_12).Value = True
        ElseIf InStr(strTmp, "����") > 0 Then
            optInfo(e_D_optInfo_���_��Ч_14).Value = True
        End If
    End If
    strTmp = ""
    
    '��Ⱦ���   ����
    rs���.Filter = "�������=5"
    optInfo(e_D_optInfo_���_��Ⱦ_��_16).Value = rs���.EOF
    optInfo(e_D_optInfo_���_��Ⱦ_��_15).Value = Not rs���.EOF
    If Not rs���.EOF Then
        str��Ⱦ��� = rs���!������� & ""
        lng��Ⱦ���ID = Val(rs���!ID & "")
    End If
    
    '��Ĭ�����ã��ٸ����Ƿ�༭��������
    '�ٴ�֢״-���Ⱦ�й�'��ҩĿ�ģ���Ⱦ���
    If mbln�༭ Then
        lng��Ⱦ���ID = 0: str��Ⱦ��� = ""
        lng��Ⱦ���ID = Val(rsTmp!�ٴ�֢״ & "")
        rs���.Filter = "id=" & lng��Ⱦ���ID
        If Not rs���.EOF Then
            str��Ⱦ��� = rs���!������� & ""
            lng��Ⱦ���ID = Val(rs���!ID & "")
        End If
        txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Tag = lng��Ⱦ���ID
        txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Text = str��Ⱦ���
        lng��Ⱦ���ID = 0: str��Ⱦ��� = ""
        
        lng��Ⱦ���ID = Val(rsTmp!��Ⱦ��� & "")
        rs���.Filter = "id=" & lng��Ⱦ���ID
        If Not rs���.EOF Then
            str��Ⱦ��� = rs���!������� & ""
            lng��Ⱦ���ID = Val(rs���!ID & "")
        End If
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Tag = lng��Ⱦ���ID
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Text = str��Ⱦ���
        If lng��Ⱦ���ID <> 0 Then lblInfo(e_D_lblInfo_��ҩĿ��_��Ⱦ���_68).Tag = lng��Ⱦ���ID & "," & str��Ⱦ���
    Else
        'δ�༭����Ĭ��ָ��һ��
        txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Tag = lng��Ⱦ���ID
        txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Text = str��Ⱦ���
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Tag = lng��Ⱦ���ID
        txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Text = str��Ⱦ���
        If lng��Ⱦ���ID <> 0 Then lblInfo(e_D_lblInfo_��ҩĿ��_��Ⱦ���_68).Tag = lng��Ⱦ���ID & "," & str��Ⱦ���
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
    strSQL = "Select Sum(Decode(Nvl(e.������, 0), 0, 0, a.���ʽ��)) As ����ҩ��," & vbNewLine & _
        "Sum(Decode(a.�շ����, '5', a.���ʽ��, '6', a.���ʽ��, '7', a.���ʽ��, 0)) As ��ҩ��, Sum(a.���ʽ��) As סԺ����" & vbNewLine & _
        "From סԺ���ü�¼ A, ҩƷ��� D, ҩƷ���� E" & vbNewLine & _
        "Where a.����id = [1] And a.��ҳid = [2] and a.��¼״̬<>0 And a.�շ�ϸĿid = d.ҩƷid(+) And d.ҩ��id = e.ҩ��id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        txtInfo(e_D_txtInfo_����_�ܷ�_31).Text = Format(Val(rsTmp!סԺ���� & ""), "0.00")
        txtInfo(e_D_txtInfo_����_ҩ��_32).Text = Format(Val(rsTmp!��ҩ�� & ""), "0.00")
        txtInfo(e_D_txtInfo_����_��ҩ��_33).Text = Format(Val(rsTmp!����ҩ�� & ""), "0.00")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDrugUse()
'���ܣ�������ҩ����ͷ��ã��ȴ����������ж�һ�Σ����û�������ٴ�ҽ����¼��ȡ
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim strҽ��IDs As String
    Dim arrTmp As Variant
    Dim str����ҩ��IDs As String
    Dim str���� As String
    Dim str��ҩĿ�� As String
    Dim lngRow As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim dblTmp As Double
    Dim dbl�״����� As Double
    Dim dbl�������� As Double
    Dim dbl���� As Double
    Dim lng���ID As Long
    Dim lng���� As Long
    Dim int��ҩĿ�� As Integer
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = "select ҩ��id,ͼ��,ҩ��,����,Ƶ��,;��,����,��ֹ���� from ����ҩ�������ҩ where ����id=[1] and ���=[2] order by ҩƷ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng���)
    
    With vsDrugUse
        .Redraw = False
        .Rows = .FixedRows
        lng���ID = 1
        For i = 1 To rsTmp.RecordCount
            .AddItem "": lngRow = .Rows - 1
            
            If rsTmp!ͼ�� & "" <> "" Then
                If rsTmp!ͼ�� & "" = "��" Then lng���ID = lng���ID + 1
            Else
                lng���ID = lng���ID + 1
            End If
            
            .TextMatrix(lngRow, COL_DRUG_���ID) = lng���ID
            .TextMatrix(lngRow, COL_DRUG_ҩ��ID) = Val(rsTmp!ҩ��id & "")
            .TextMatrix(lngRow, COL_DRUG_ͼ��) = rsTmp!ͼ�� & ""
            .TextMatrix(lngRow, COL_DRUG_ҩƷ����) = rsTmp!ҩ�� & ""
            .TextMatrix(lngRow, COL_DRUG_���μ���) = rsTmp!���� & ""
            .TextMatrix(lngRow, COL_DRUG_��ҩƵ��) = rsTmp!Ƶ�� & ""
            .TextMatrix(lngRow, COL_DRUG_;��) = rsTmp!;�� & ""
            .TextMatrix(lngRow, COL_DRUG_������) = rsTmp!���� & ""
            .TextMatrix(lngRow, COL_DRUG_��ֹʱ��) = rsTmp!��ֹ���� & ""
            rsTmp.MoveNext
        Next
        .Redraw = True
    End With
    
    If rsTmp.RecordCount > 0 Then Exit Sub
    Set rsTmp = Nothing
    
    strSQL = "Select d.Id,d.���id,e.Id As ҩ��id,Decode(e.Id,b.ҩ��id,b.������,0) As ������,d.ҽ������ As ҩƷ����,c.ҽ������ As ;��," & vbNewLine & _
        "d.�״�����, d.��������,d.ִ��Ƶ�� As ��ҩƵ��, e.���㵥λ As ��λ, d.��ҩĿ��,to_char(d.��ʼִ��ʱ��, 'MM-DD HH24:MI')||' - '||" & vbNewLine & _
        "to_char(Nvl(Nvl(d.�ϴ�ִ��ʱ��,d.ִ����ֹʱ��),d.ͣ��ʱ��),'MM-DD HH24:MI') As ��ֹʱ��" & vbNewLine & _
        "From ����ҽ����¼ A, ҩƷ���� B, ����ҽ����¼ C, ����ҽ����¼ D, ������ĿĿ¼ E" & vbNewLine & _
        "Where a.������Ŀid = b.ҩ��id And a.���id = c.Id And a.������� = '5' And Nvl(b.������, 0) <> 0 And c.Id = d.���id And d.������Ŀid = e.Id And" & vbNewLine & _
        "d.������� In ('5','6') And a.ҽ��״̬ in (8,9) And a.����id =[1] And a.��ҳid =[2] Order By d.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.EOF Then Exit Sub
    For i = 1 To rsTmp.RecordCount
        If InStr("," & strҽ��IDs & ",", "," & rsTmp!���ID & ",") = 0 Then
            strҽ��IDs = strҽ��IDs & "," & rsTmp!���ID
        End If
        rsTmp.MoveNext
    Next
    strҽ��IDs = Mid(strҽ��IDs, 2)
    rsTmp.MoveFirst
    
    strSQL = "select id as ��ҽ��id,Zl_Adviceexetimes(Id,��ʼִ��ʱ��,Nvl(Nvl(�ϴ�ִ��ʱ��,ִ����ֹʱ��),ͣ��ʱ��)," & _
        "ִ��ʱ�䷽��,��ʼִ��ʱ��,��ʼִ��ʱ��-1,Ƶ�ʼ��,�����λ,ҽ����Ч) as �ֽ�ʱ�� From ����ҽ����¼ where Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    'rs���� �����¼��������ҽ������������ҩ����
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҽ��IDs)
    
    With vsDrugUse
        .Redraw = False
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .AddItem "": lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_DRUG_ID) = Val(rsTmp!ID & "")
            .TextMatrix(lngRow, COL_DRUG_���ID) = Val(rsTmp!���ID & "")
            .TextMatrix(lngRow, COL_DRUG_ҩ��ID) = Val(rsTmp!ҩ��id & "")
            .TextMatrix(lngRow, COL_DRUG_ҩƷ����) = rsTmp!ҩƷ���� & ""
            .TextMatrix(lngRow, COL_DRUG_��ҩƵ��) = rsTmp!��ҩƵ�� & ""
            .TextMatrix(lngRow, COL_DRUG_;��) = rsTmp!;�� & ""
            .TextMatrix(lngRow, COL_DRUG_��ֹʱ��) = rsTmp!��ֹʱ�� & ""
            strTmp = ""
            dblTmp = Val(rsTmp!�״����� & "")
            dbl�״����� = dblTmp
            If dblTmp > 0 Then
                If Mid(dblTmp, 1, 1) = "." Then
                    strTmp = "0" & dblTmp
                Else
                    strTmp = dblTmp
                End If
            End If
            
            dblTmp = Val(rsTmp!�������� & "")
            dbl�������� = dblTmp
            If dblTmp > 0 Then
                If Mid(dblTmp, 1, 1) = "." Then
                    strTmp = IIf(strTmp = "", "", strTmp & ":") & "0" & dblTmp
                Else
                    strTmp = IIf(strTmp = "", "", strTmp & ":") & dblTmp
                End If
            End If
            
            .TextMatrix(lngRow, COL_DRUG_���μ���) = strTmp & rsTmp!��λ: strTmp = "": dblTmp = 0
            
            If lng���ID <> Val(rsTmp!���ID & "") Then
                lng���ID = Val(rsTmp!���ID & "")
                '�������㣬ȷ������������
                rs����.Filter = "��ҽ��id=" & Val(rsTmp!���ID & "")
                If Not rs����.EOF Then
                    strTmp = rs����!�ֽ�ʱ�� & ""
                    lng���� = 0
                    If strTmp <> "" Then
                        If InStr(strTmp, ",") = 0 Then
                            strTmp = Format(rs����!�ֽ�ʱ��, "YYYY-MM-DD HH:MM:SS")
                        End If
                        arrTmp = Split(strTmp, ",")
                        
                        lng���� = UBound(arrTmp) + 1
                        
                        For j = 0 To UBound(arrTmp)
                            strTmp = Format(arrTmp(j), "YYYY-MM-DD")
                            If InStr("," & str���� & ",", "," & strTmp & ",") = 0 Then
                                str���� = str���� & "," & strTmp
                            End If
                        Next
                    End If
                End If
            End If
            
            strTmp = "0"
            If lng���� > 0 Then
                If dbl�״����� <> 0 Then
                    dbl���� = dbl�״����� + dbl�������� * (lng���� - 1)
                Else
                    dbl���� = dbl�������� * lng����
                End If
                
                If Mid(dbl����, 1, 1) = "." Then
                    strTmp = "0" & dbl����
                Else
                    strTmp = dbl����
                End If
            End If
            
            .TextMatrix(lngRow, COL_DRUG_������) = strTmp & rsTmp!��λ: strTmp = "": dblTmp = 0
             
             
            If InStr("," & str����ҩ��IDs & ",", "," & rsTmp!ҩ��id & ",") = 0 And Val(rsTmp!������ & "") <> 0 Then
                str����ҩ��IDs = str����ҩ��IDs & "," & rsTmp!ҩ��id
            End If
            
            If InStr("," & str��ҩĿ�� & ",", "," & Val(rsTmp!��ҩĿ�� & "") & ",") = 0 And Val(rsTmp!��ҩĿ�� & "") <> 0 Then
                str��ҩĿ�� = str��ҩĿ�� & "," & Val(rsTmp!��ҩĿ�� & "")
            End If
            rsTmp.MoveNext
        Next
        
        '��ҩ����
        str���� = Mid(str����, 2)
        If str���� <> "" Then txtInfo(e_D_txtInfo_��ҩ���_����_30).Text = UBound(Split(str����, ",")) + 1
        
        '����ҩ����
        str����ҩ��IDs = Mid(str����ҩ��IDs, 2)
        If str����ҩ��IDs <> "" Then txtInfo(e_D_txtInfo_��ҩ���_����_29).Text = UBound(Split(str����ҩ��IDs, ",")) + 1
        
        '��ҩĿ��
        str��ҩĿ�� = Mid(str��ҩĿ��, 2)
        If str��ҩĿ�� <> "" Then
            If InStr("," & str��ҩĿ�� & ",", ",1,") > 0 And InStr("," & str��ҩĿ�� & ",", ",2,") = 0 Then
                int��ҩĿ�� = 1 'Ԥ��
            ElseIf InStr("," & str��ҩĿ�� & ",", ",1,") = 0 And InStr("," & str��ҩĿ�� & ",", ",2,") > 0 Then
                int��ҩĿ�� = 2 '����
            End If
        End If
        
        optInfo(e_D_optInfo_��ҩĿ��_δ��_17 + 1).Value = True
        If int��ҩĿ�� <> 0 Then optInfo(e_D_optInfo_��ҩĿ��_δ��_17 + int��ҩĿ��).Value = True
        '��ͼ��
        For i = 1 To .Rows - 1
            Call Getһ����ҩ��Χ(Val(.TextMatrix(i, COL_DRUG_���ID)), lngBegin, lngEnd)
            For j = lngBegin To lngEnd
                If lngBegin <> lngEnd Then
                    If j = lngBegin Then
                        .TextMatrix(j, 0) = "��"
                    ElseIf j = lngEnd Then
                        .TextMatrix(j, 0) = "��"
                    Else
                        .TextMatrix(j, 0) = "��"
                    End If
                End If
            Next
        Next
        .Redraw = True
    End With
    Call Insert��ҩ
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Insert��ҩ()
'���ܣ�������ҩ���
    Dim blnTrans As Boolean
    Dim blnDo As Boolean
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    arrSQL = Array()
    strSQL = "Zl_����ҩ�������ҩ_Delete(" & mlng����ID & "," & mlng��� & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    With vsDrugUse
        j = 1
        For i = .FixedRows To .Rows - 1
            blnDo = True
            strSQL = "Zl_����ҩ�������ҩ_Insert(" & mlng����ID & "," & mlng��� & "," & j & "," & Val(.TextMatrix(i, COL_DRUG_ҩ��ID)) & "," & _
                IIf(.TextMatrix(i, COL_DRUG_ͼ��) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_ͼ��) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_ҩƷ����) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_ҩƷ����) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_���μ���) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_���μ���) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_��ҩƵ��) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_��ҩƵ��) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_;��) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_;��) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_������) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_������) & "'") & "," & _
                IIf(.TextMatrix(i, COL_DRUG_��ֹʱ��) = "", "NULL", "'" & .TextMatrix(i, COL_DRUG_��ֹʱ��) & "'") & ")"
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
'���ܣ������������ �ȴ����������ж�һ�Σ����û�������ٴ� ���������¼ ��ȡ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim str��ҩ�ڼ� As String
    Dim lngRow As Long
    Dim i As Long
    
    On Error GoTo errH

    strSQL = "select ����,���� from ����Ԥ����ҩ�ڼ� order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        strTmp = ""
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "|#" & rsTmp!���� & ";" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    str��ҩ�ڼ� = Mid(strTmp, 2)
    
    With vsOperate
        .Rows = .FixedRows
        .ColData(COL_OPE_�п�) = "��|��|��"
        .ColData(COL_OPE_��ҩ���) = "δ׷��|��׷��"
        .ColData(COL_OPE_��ҩ�ڼ�) = str��ҩ�ڼ�
    End With
    
    strSQL = "Select a.����id,a.��������,a.�п�,b.���� as ��ҩ�ڼ�,a.��ʼʱ��," & vbNewLine & _
        " a.����ʱ��,decode(nvl(a.��ҩ���,0),0,'δ׷��',1,'��׷��',null) as ��ҩ���,b.����" & vbNewLine & _
        " From ����ҩ��������� A,����Ԥ����ҩ�ڼ� b" & vbNewLine & _
        " Where a.Ԥ����ҩ�ڼ� = b.����(+) And a.����id = [1] And a.��� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng���)
    
    If rsTmp.EOF Then
        Set rsTmp = Nothing
        strSQL = "Select a.Id As ����id, a.�������� As ��������, a.�п�,Null As ��ҩ�ڼ�, a.������ʼʱ�� As ��ʼʱ��,a.��������ʱ�� As ����ʱ��,'δ׷��' As ��ҩ���,null as ����" & vbNewLine & _
            "From ���������¼ A Where a.����id =[1] And a.��ҳid =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    End If
    
    With vsOperate
        For i = 1 To rsTmp.RecordCount
            .AddItem "": lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_OPE_����ID) = Val(rsTmp!����id & "")
            .TextMatrix(lngRow, COL_OPE_��������) = rsTmp!�������� & ""
            .TextMatrix(lngRow, COL_OPE_�п�) = rsTmp!�п� & ""
            If rsTmp!��ʼʱ�� & "" <> "" Then
                .TextMatrix(lngRow, COL_OPE_��ʼʱ��) = Format(rsTmp!��ʼʱ��, "YYYY-MM-DD HH:MM")
            End If
            If rsTmp!����ʱ�� & "" <> "" Then
                .TextMatrix(lngRow, COL_OPE_����ʱ��) = Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:MM")
            End If
            .TextMatrix(lngRow, COL_OPE_��ҩ�ڼ�) = rsTmp!��ҩ�ڼ� & ""
            .Cell(flexcpData, lngRow, COL_OPE_��ҩ�ڼ�) = Val(rsTmp!���� & "")
            .TextMatrix(lngRow, COL_OPE_��ҩ���) = rsTmp!��ҩ��� & ""
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

Private Sub Getһ����ҩ��Χ(ByVal lng���ID As Long, lngBegin As Long, lngEnd As Long)
'���ܣ�������صĸ�ҩ;��ҽ��ID,ȷ��һ����ҩ��һ��ҩƷ����ֹ�к�
    Dim i As Long
    lngBegin = vsDrugUse.FindRow(CStr(lng���ID), , COL_DRUG_���ID)
    lngEnd = lngBegin
    If lngBegin = -1 Then Exit Sub
    For i = lngBegin To vsDrugUse.Rows - 1
        If Not vsDrugUse.RowHidden(i) Then
            If Val(vsDrugUse.TextMatrix(i, COL_DRUG_���ID)) = lng���ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Function SaveData(Optional ByVal blnFirst As Boolean) As Boolean
'������blnFirst ��true ��δ�����޸����ٱ���һ�Σ�����һ���ݳ�ʼֵ =false�������޸ĵı���
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim strTmp As String
    Dim str��Ŀ���� As String
    Dim str��Ŀֵ As String
    Dim lngTmp As String
    Dim strResult As String
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    Dim i As Integer, j As Integer
    Dim blnInit As Boolean
    
    blnInit = mblnInitSaved And blnFirst
    
    On Error GoTo errH
    If Not blnFirst Then mbln�༭ = True      '�޸ĺ󱻱�������Ϊ���Ѿ��༭����
    strSQL = Get������ϸSQL
    arrSQL = Array()
    If mblnChange Or blnInit Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    
    '�������
    If mbln���� And (mblnChange Or blnInit) Then
        strSQL = "Zl_����ҩ���������_Delete(" & mlng����ID & "," & mlng��� & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        With vsOperate
            For i = .FixedRows To .Rows - 1
                lngTmp = Val(.Cell(flexcpData, i, COL_OPE_��ҩ�ڼ�))
                strSQL = "Zl_����ҩ���������_Insert(" & mlng����ID & "," & mlng��� & "," & Val(.TextMatrix(i, COL_OPE_����ID)) & ",'" & .TextMatrix(i, COL_OPE_��������) & "'," & _
                   IIf(.TextMatrix(i, COL_OPE_�п�) = "", "NULL", "'" & .TextMatrix(i, COL_OPE_�п�) & "'") & "," & _
                   IIf(.TextMatrix(i, COL_OPE_��ʼʱ��) = "", "NULL", "to_date('" & .TextMatrix(i, COL_OPE_��ʼʱ��) & "','YYYY-MM-DD HH24:MI')") & "," & _
                   IIf(.TextMatrix(i, COL_OPE_����ʱ��) = "", "NULL", "to_date('" & .TextMatrix(i, COL_OPE_����ʱ��) & "','YYYY-MM-DD HH24:MI')") & "," & _
                   IIf(lngTmp = 0, "NULL", lngTmp) & "," & IIf(.TextMatrix(i, COL_OPE_��ҩ���) = "��׷��", 1, 0) & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Next
        End With
    End If
    
    '��ҩ����
    If mbln��ҩ���� And (mblnChange Or blnInit) Then
        strSQL = "Zl_����ҩ���������_Delete(" & mlng����ID & "," & mlng��� & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        With vsPJB
            For i = .FixedRows To .Rows - 1
                
                str��Ŀ���� = .TextMatrix(i, COL_�������)
                If Val(str��Ŀ����) = 0 Then str��Ŀ���� = "0"
                
                str��Ŀֵ = .TextMatrix(i, COL_����Ժ)
                If Not .Cell(flexcpPicture, i, COL_����Ժ) Is Nothing Then '����   ��Ժ
                    str��Ŀֵ = Val(.Cell(flexcpData, i, COL_����Ժ))
                End If
                
                strSQL = "Zl_����ҩ���������_Insert(" & mlng����ID & "," & mlng��� & ",0,'" & str��Ŀ���� & "'," & i & ",1,'" & str��Ŀֵ & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
    
                str��Ŀ���� = .TextMatrix(i, COL_�������)
                If Val(str��Ŀ����) = 0 Then str��Ŀ���� = "0"
                
                str��Ŀֵ = .TextMatrix(i, COL_��������)
                If Not .Cell(flexcpPicture, i, COL_��������) Is Nothing Then '����   ����
                    str��Ŀֵ = Val(.Cell(flexcpData, i, COL_��������))
                End If
                
                strSQL = "Zl_����ҩ���������_Insert(" & mlng����ID & "," & mlng��� & ",1,'" & str��Ŀ���� & "'," & i & ",1,'" & str��Ŀֵ & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                
                str��Ŀ���� = .TextMatrix(i, COL_���������)
                If Val(str��Ŀ����) = 0 Then str��Ŀ���� = "0"
                
                str��Ŀֵ = .TextMatrix(i, COL_������Ժ)
                If Not .Cell(flexcpPicture, i, COL_������Ժ) Is Nothing Then '������   ��Ժ
                    str��Ŀֵ = Val(.Cell(flexcpData, i, COL_������Ժ))
                End If
                
                strSQL = "Zl_����ҩ���������_Insert(" & mlng����ID & "," & mlng��� & ",0,'" & str��Ŀ���� & "'," & i & ",0,'" & str��Ŀֵ & "')"
    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
    
                str��Ŀ���� = .TextMatrix(i, COL_���������)
                If Val(str��Ŀ����) = 0 Then str��Ŀ���� = "0"
                
                str��Ŀֵ = .TextMatrix(i, COL_����������)
                If Not .Cell(flexcpPicture, i, COL_����������) Is Nothing Then '������   ����
                    str��Ŀֵ = Val(.Cell(flexcpData, i, COL_����������))
                End If
                
                strSQL = "Zl_����ҩ���������_Insert(" & mlng����ID & "," & mlng��� & ",1,'" & str��Ŀ���� & "'," & i & ",0,'" & str��Ŀֵ & "')"
    
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

Private Function Get������ϸSQL() As String
'���ܣ���ȡSQL����鵱ǰ����
    Dim strSQL As String, strTmp As String
    Dim lngTmp As Long
    
    strSQL = "Zl_����ҩ�������ϸ_Update(" & mlng����ID & "," & mlng����ID & "," & mlng��ҳID & "," & mlng��� & ","
    strSQL = strSQL & IIf(mbln����, 1, 0) & "," & IIf(optInfo(e_D_optInfo_���_��ԭѧ���_δ��_2).Value, 0, 1) & ","
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ԭѧ�������_23).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ԭѧ���걾_26).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ԭѧ�����ϸ����_25).Text) & "',")
    strSQL = strSQL & IIf(optInfo(e_D_optInfo_���_ҩ������_δ��_6).Value, 0, 1) & ","
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_ҩ����������_24).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_ҩ����������_24).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf(optInfo(e_D_optInfo_���_ҩ������_���_8).Value, 0, 1) & ","
    
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩǰ����_11).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩǰ����_11).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ������_13).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ������_13).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ��_15).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ��_15).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ����_17).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ����_17).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø_19).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø_19).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩǰ����_21).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩǰ����_21).Text) & "',")
    
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩ������_55).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩ������_55).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩ���ϸ������_53).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩ���ϸ������_53).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩ��������ϸ��_51).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩ��������ϸ��_51).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ����_49).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ����_49).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø_45).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø_45).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_���_��ҩ����_43).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_���_��ҩ����_43).Text) & "',")
 
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩǰ��������_12).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩǰ��������_12).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ����������_14).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩǰ��ϸ����������_14).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ������_16).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩǰ������ϸ������_16).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ��������_18).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩǰC��Ӧ��������_18).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø����_20).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩǰ����ת��ø����_20).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩǰ��������_22).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩǰ��������_22).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩ����������_54).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩ����������_54).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩ���ϸ����������_52).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩ���ϸ����������_52).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩ��������ϸ������_50).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩ��������ϸ������_50).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ��������_48).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩ��C��Ӧ��������_48).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø����_44).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩ�����ת��ø����_44).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    strSQL = strSQL & IIf("" = txtInfo(e_D_txtInfo_���_��ҩ��������_42).Tag, "NULL,", "to_date('" & txtInfo(e_D_txtInfo_���_��ҩ��������_42).Tag & "','YYYY-MM-DD HH24:MI:SS'),")
    'Ӱ��ѧ���
    strTmp = chkInfo(e_D_chkInfo_Ӱ��_X��_18).Value & chkInfo(e_D_chkInfo_Ӱ��_CT_19).Value & chkInfo(e_D_chkInfo_Ӱ��_�Ź���_20).Value
    strSQL = strSQL & "'" & strTmp & "',"
    
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_Ӱ��_��λ_46).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_Ӱ��_��λ_46).Text) & "',")
    strSQL = strSQL & IIf("" = Trim(txtInfo(e_D_txtInfo_Ӱ��_����_47).Text), "NULL,", "'" & Trim(txtInfo(e_D_txtInfo_Ӱ��_����_47).Text) & "',")
    
    strSQL = strSQL & Val(txtInfo(e_D_txtInfo_֢״_��Ⱦ֢״_35).Tag) & ","
    
    If optInfo(e_D_optInfo_��ҩĿ��_δ��_17).Value Then
        lngTmp = 0
    ElseIf optInfo(e_D_optInfo_��ҩĿ��_Ԥ��_18).Value Then
        lngTmp = 1
    ElseIf optInfo(e_D_optInfo_��ҩĿ��_����_21).Value Then
        lngTmp = 2
    End If
    
    strSQL = strSQL & lngTmp & ","
    strSQL = strSQL & Val(txtInfo(e_D_txtInfo_��ҩĿ��_��Ⱦ���_36).Tag) & ","
    
    If optInfo(e_D_optInfo_���_��ת_13).Value Then
        lngTmp = 2
    ElseIf optInfo(e_D_optInfo_���_����_12).Value Then
        lngTmp = 1
    ElseIf optInfo(e_D_optInfo_���_��Ч_14).Value Then
        lngTmp = 3
    End If
    strSQL = strSQL & lngTmp & ","
    
    strTmp = ""
    strTmp = "'" & chkInfo(e_D_chkInfo_����_��Ժ_��Ӧ֢_5).Value & "|" & chkInfo(e_D_chkInfo_����_����_��Ӧ֢_37).Value & "','" & chkInfo(e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12).Value & "|" & chkInfo(e_D_chkInfo_����_����_ҩ��ѡ��_36).Value & "','" & _
        chkInfo(e_D_chkInfo_����_��Ժ_���μ���_13).Value & "|" & chkInfo(e_D_chkInfo_����_����_���μ���_35).Value & "','" & chkInfo(e_D_chkInfo_����_��Ժ_ÿ�ո�ҩƵ��_14).Value & "|" & chkInfo(e_D_chkInfo_����_����_ÿ�ո�ҩƵ��_34).Value & "','" & _
        chkInfo(e_D_chkInfo_����_��Ժ_�ܼ�_15).Value & "|" & chkInfo(e_D_chkInfo_����_����_�ܼ�_33).Value & "','" & chkInfo(e_D_chkInfo_����_��Ժ_��ҩ;��_16).Value & "|" & chkInfo(e_D_chkInfo_����_����_��ҩ;��_32).Value & "','" & _
        chkInfo(e_D_chkInfo_����_��Ժ_��ҩ�Ƴ�_17).Value & "|" & chkInfo(e_D_chkInfo_����_����_��ҩ�Ƴ�_31).Value & "','" & chkInfo(e_D_chkInfo_����_��Ժ_��ǰ_23).Value & "|" & chkInfo(e_D_chkInfo_����_����_��ǰ_28).Value & "','" & _
        chkInfo(e_D_chkInfo_����_��Ժ_����_24).Value & "|" & chkInfo(e_D_chkInfo_����_����_����_27).Value & "','" & chkInfo(e_D_chkInfo_����_��Ժ_����_25).Value & "|" & chkInfo(e_D_chkInfo_����_����_����_26).Value & "','" & _
        chkInfo(e_D_chkInfo_����_��Ժ_������ҩ_22).Value & "|" & chkInfo(e_D_chkInfo_����_����_������ҩ_29).Value & "','" & chkInfo(e_D_chkInfo_����_��Ժ_����ҩ��_21).Value & "|" & chkInfo(e_D_chkInfo_����_����_����ҩ��_30).Value & "'"
    
    strSQL = strSQL & strTmp & ",'" & txtInfo(e_D_txtInfo_��ע_57).Text & "'," & IIf(mbln�༭, 1, 0) & ","
    strSQL = strSQL & IIf(mbln��ӡ, 1, 0) & "," & Val(txtInfo(e_D_txtInfo_��ҩ���_����_30).Text) & "," & Val(txtInfo(e_D_txtInfo_��ҩ���_����_29).Text) & ","
    strSQL = strSQL & IIf(optInfo(e_D_optInfo_���_���ҩ_��_20).Value, 1, "Null")
    strSQL = strSQL & ")"
    
    Get������ϸSQL = strSQL
End Function

Private Sub Init��ҩ���۱�()
'���ܣ�����ҩ���۱�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rs��Ŀ���� As ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    Dim intCount As Integer
    Dim intLin As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    
    strSQL = "select ���,decode(���,null,decode(����,'-',999,888),���) as ����,����,����,�Ƿ����,�ϼ� from ������ҩ������Ŀ where ĩ��=1 and �Ƿ�����=[1] order by ���,2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(mbln����, 1, 0))
    
    strSQL = "select a.����,a.���,a.���� from ������ҩ������Ŀ a,������ҩ������Ŀ b where a.ĩ��=0 and a.�ϼ�=b.���� and b.����=[1] order by ���"
    Set rs��Ŀ���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(mbln����, "����", "������"))
    
    With vsPJB
        .Clear
        .Redraw = False
        .FixedRows = 2
        .FixedCols = 0
        .Cols = 11
        .Rows = .FixedRows
        .RowHeightMin = 400
        .WordWrap = True
        For i = 0 To COL_����������
            .ColAlignment(.FixedCols + i) = flexAlignCenterCenter
        Next
        .ColWidth(COL_������Ŀ) = 1200
        .ColWidth(COL_��������) = 3400
        .ColWidth(COL_����Ժ) = 1000
        .ColWidth(COL_��������) = 1000
        
        .ColWidth(COL_����������) = 3400
        .ColWidth(COL_������Ժ) = 1000
        .ColWidth(COL_����������) = 1000
        
        
        .ColHidden(COL_�������) = True
        .ColHidden(COL_���������) = True
        .ColHidden(COL_�ϼ�����) = True
        .ColHidden(COL_�ϼ�����) = True
        
        .Cell(flexcpText, 0, 0, 1, 1) = "������Ŀ"
        
        .Cell(flexcpText, 0, 1, 0, 3) = "����"
        
        .Cell(flexcpText, 0, 4, 0, 6) = "������"

        .TextMatrix(1, COL_��������) = "��������"
        .TextMatrix(1, COL_����������) = "��������"
        
        .TextMatrix(1, COL_����Ժ) = "��Ժ����"
        .TextMatrix(1, COL_������Ժ) = "��Ժ����"
        
        .TextMatrix(1, COL_��������) = "���Ļ��������"
        .TextMatrix(1, COL_����������) = "���Ļ��������"
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFixedOnly
        
        '�߿�����
        intLin = 1
        .Select 0, 0, 1, COL_����������
        .CellBorder .GridColor, intLin, intLin, intLin, intLin, 0, 0
        
        .Rows = rsTmp.RecordCount
        lngRow = 1
        For i = 1 To rs��Ŀ����.RecordCount
            
            rsTmp.Filter = "�ϼ�='" & rs��Ŀ����!���� & "' and �Ƿ����=1"
            intCount = rsTmp.RecordCount
            For j = 1 To rsTmp.RecordCount
                .TextMatrix(lngRow + j, COL_�ϼ�����) = rs��Ŀ����!����
                .TextMatrix(lngRow + j, COL_�ϼ�����) = rs��Ŀ����!����
                
                .TextMatrix(lngRow + j, COL_������Ŀ) = rs��Ŀ����!����
                
                .TextMatrix(lngRow + j, COL_�������) = rsTmp!����
                
                If Left(rsTmp!���� & "", 1) = "-" Then
                    strTmp = ""
                Else
                    strTmp = IIf(rsTmp!��� & "" <> "", rsTmp!��� & ".", "") & rsTmp!����
                End If
                
                .TextMatrix(lngRow + j, COL_��������) = strTmp
                
                If .TextMatrix(lngRow + j, COL_��������) <> "" Then
                    Set .Cell(flexcpPicture, lngRow + j, COL_����Ժ) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_����Ժ) = 1
                    
                    Set .Cell(flexcpPicture, lngRow + j, COL_��������) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_��������) = 1
                    
                End If
                rsTmp.MoveNext
            Next
            
            rsTmp.Filter = "�ϼ�='" & rs��Ŀ����!���� & "' and �Ƿ����=0"
            If intCount < rsTmp.RecordCount Then intCount = rsTmp.RecordCount
            For j = 1 To rsTmp.RecordCount
                .TextMatrix(lngRow + j, COL_�ϼ�����) = rs��Ŀ����!����
                .TextMatrix(lngRow + j, COL_�ϼ�����) = rs��Ŀ����!����
                
                .TextMatrix(lngRow + j, COL_������Ŀ) = rs��Ŀ����!����
                
                .TextMatrix(lngRow + j, COL_���������) = rsTmp!����
                
                If Left(rsTmp!���� & "", 1) = "-" Then
                    strTmp = ""
                Else
                    strTmp = IIf(rsTmp!��� & "" <> "", rsTmp!��� & ".", "") & rsTmp!����
                End If
                
                .TextMatrix(lngRow + j, COL_����������) = strTmp
                
                
                If .TextMatrix(lngRow + j, COL_����������) <> "" Then
                    Set .Cell(flexcpPicture, lngRow + j, COL_������Ժ) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_������Ժ) = 1
                    
                    Set .Cell(flexcpPicture, lngRow + j, COL_����������) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, lngRow + j, COL_����������) = 1
                    
                End If
                
                rsTmp.MoveNext
            Next
            
            .Select lngRow + 1, 0, lngRow + intCount, COL_����������
            .CellBorder .GridColor, intLin, intLin, intLin, intLin, 0, 0
            
            rs��Ŀ����.MoveNext
            lngRow = lngRow + intCount
        Next
        .Rows = lngRow + 1
        
        .Cell(flexcpFontBold, 0, COL_������Ŀ, 1, COL_����������) = True
        .Cell(flexcpAlignment, 2, COL_������Ŀ, .Rows - 1, COL_������Ŀ) = flexAlignLeftCenter
        .Cell(flexcpFontBold, 0, COL_������Ŀ, .Rows - 1, COL_������Ŀ) = True
        .Cell(flexcpAlignment, 2, COL_��������, .Rows - 1, COL_��������) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 2, COL_����������, .Rows - 1, COL_����������) = flexAlignLeftCenter
        
        .Cell(flexcpPictureAlignment, 2, COL_����Ժ, .Rows - 1, COL_��������) = flexPicAlignCenterCenter
        .Cell(flexcpPictureAlignment, 2, COL_������Ժ, .Rows - 1, COL_����������) = flexPicAlignCenterCenter
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
'���ܣ��������ֱ߿���
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsDrugUse
        If Col = COL_DRUG_��ҩƵ�� Or Col = COL_DRUG_;�� Or Col = COL_DRUG_��ֹʱ�� Then
            Call Getһ����ҩ��Χ(Val(.TextMatrix(Row, COL_DRUG_���ID)), lngBegin, lngEnd)
            If lngBegin >= lngEnd Then Exit Sub
            
            vRect.Left = Left
            vRect.Right = Right - 1 '�����ұ߱����
            
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
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
    If Col = COL_OPE_��ҩ�ڼ� Then vsOperate.Cell(flexcpData, Row, Col) = vsOperate.ComboData
End Sub

Private Sub vsOperate_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsOperate
        If NewCol >= COL_OPE_�п� And NewCol <= COL_OPE_��ҩ��� Then
            .Editable = flexEDKbdMouse
            If NewCol = COL_OPE_�п� Or NewCol = COL_OPE_��ҩ��� Or NewCol = COL_OPE_��ҩ�ڼ� Then
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
    ElseIf Col = COL_OPE_����ʱ�� Or Col = COL_OPE_��ʼʱ�� Then
        If InStr("0123456789-: " & Chr(8) & Chr(27) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsOperate_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_OPE_����ʱ�� Or Col = COL_OPE_��ʼʱ�� Then
        vsOperate.Refresh    '����е�����ʾ����ˢ�µĻ���һ����ҩͨ��Drawcell�������ĵ�Ԫ����ٴ���ʾ
        If Not AcceptInput(Row, Col) Then
            Cancel = True
        Else
            If mblnReturn Then
                Call vsOperate_KeyPress(13) '��λ��һ�����뵥Ԫ
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
        
        '����������Ч��
        If Not IsDate(.EditText) Then
            MsgBox "������һ����Ч��" & .TextMatrix(0, Col) & " ��", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        
        '���������Ժʱ��
        If Format(.EditText, "yyyy-MM-dd HH:mm") <= txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Tag Then
            MsgBox "���������" & IIf(Col = COL_OPE_����ʱ��, "����", "��ʼ") & "ʱ�������ڲ�����Ժʱ�� " & txtInfo(e_D_txtInfo_����_��Ժʱ��_9).Tag & "��", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        '����С�ڳ�Ժʱ��
        If Format(.EditText, "yyyy-MM-dd HH:mm") > txtInfo(e_D_txtInfo_����_��Ժʱ��_10).Tag Then
            MsgBox "���������" & IIf(Col = COL_OPE_����ʱ��, "����", "��ʼ") & "ʱ�䲻Ӧ���ڲ��˳�Ժʱ�� " & txtInfo(e_D_txtInfo_����_��Ժʱ��_10).Tag & "��", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        
        If Col = COL_OPE_����ʱ�� And .TextMatrix(Row, COL_OPE_��ʼʱ��) <> "" Then
            If Format(.EditText, "yyyy-MM-dd HH:mm") < .TextMatrix(Row, COL_OPE_��ʼʱ��) Then
                MsgBox "�������������ʱ��������������ʼʱ�� " & .TextMatrix(Row, COL_OPE_��ʼʱ��) & "��", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        ElseIf Col = COL_OPE_��ʼʱ�� And .TextMatrix(Row, COL_OPE_����ʱ��) <> "" Then
            If Format(.EditText, "yyyy-MM-dd HH:mm") > .TextMatrix(Row, COL_OPE_����ʱ��) Then
                MsgBox "�����������ʼʱ�����С����������ʱ�� " & .TextMatrix(Row, COL_OPE_����ʱ��) & "��", vbInformation, gstrSysName
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
'���ܣ��������ֱ߿���
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsPJB
        If Col <> COL_������Ŀ And .TextMatrix(Row, COL_�ϼ�����) <> "" Then Exit Sub
        
        If Not Getͬһ�ϼ���(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left + 1 '������߱���߱���
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
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

Private Function Getͬһ�ϼ���(ByVal lngRow As Long, ByRef lngBegin As Long, ByRef lngEnd As Long) As Boolean
'���ܣ�����һ����������
'������lngRow ��ǰ�У�lngID �Һ�ID
    Dim i As Long
    
    lngBegin = lngRow
    lngEnd = lngRow
    
    With vsPJB
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, COL_�ϼ�����) = "" Then Exit Function
            If .TextMatrix(i, COL_�ϼ�����) <> .TextMatrix(lngRow, COL_�ϼ�����) Then
                Exit For
            Else
                lngBegin = i
            End If
        Next
        
        For i = lngRow + 1 To .Rows - 1
            If .TextMatrix(i, COL_�ϼ�����) = "" Then Exit Function
            If .TextMatrix(i, COL_�ϼ�����) <> .TextMatrix(lngRow, COL_�ϼ�����) Then
                Exit For
            Else
                lngEnd = i
            End If
        Next
    End With
    
    If lngBegin < lngEnd Then Getͬһ�ϼ��� = True
    
End Function

Private Sub vsPJB_KeyPress(KeyAscii As Integer)
    Dim blnEdit As Boolean
    Dim lngCol As Long
    Dim str��Ŀ���� As String
    Dim str��Ŀ��ϸ���� As String
    
    With vsPJB
        If .Row <= 1 Or .TextMatrix(.Row, COL_�ϼ�����) = "" Then Exit Sub
        
        str��Ŀ���� = .TextMatrix(.Row, COL_�ϼ�����)
        
        If .Col = COL_����Ժ Then
            blnEdit = Edit������Ŀ(str��Ŀ����, True, True)
            lngCol = COL_��������
        ElseIf .Col = COL_�������� Then
            blnEdit = Edit������Ŀ(str��Ŀ����, True, False)
            lngCol = COL_��������
        ElseIf .Col = COL_������Ժ Then
            blnEdit = Edit������Ŀ(str��Ŀ����, False, True)
            lngCol = COL_����������
        ElseIf .Col = COL_���������� Then
            blnEdit = Edit������Ŀ(str��Ŀ����, False, False)
            lngCol = COL_����������
        End If
        
        str��Ŀ��ϸ���� = .TextMatrix(.Row, lngCol)
        
        .Editable = flexEDNone
        If blnEdit And str��Ŀ��ϸ���� = "" Then
            If .TextMatrix(.Row - 1, lngCol) <> "" Then .Editable = flexEDKbdMouse
        ElseIf blnEdit Then
            If .Cell(flexcpData, .Row, .Col) = 2 Then '1-����ѡ��2����ʾ��ѡ
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

Private Function Edit������Ŀ(ByVal str��Ŀ As String, ByVal bln���� As Boolean, ByVal bln��Ժ As Boolean) As Boolean
'���ܣ���ȡĳ��������Ŀ���Խ��й�ѡ�ķ�Χ������Ӧ֤����Ŀ������ 3 �У� 4 5 6
'������lngS ��ʼ�У�lngE ������
 
    If bln��Ժ And chkInfo(e_D_chkInfo_����_��Ժ_��Ӧ֢_5).Value <> 1 And "��Ӧ֢" <> str��Ŀ Then
        Edit������Ŀ = False
        Exit Function
    End If
    
    If Not bln��Ժ And chkInfo(e_D_chkInfo_����_����_��Ӧ֢_37).Value <> 1 And "��Ӧ֢" <> str��Ŀ Then
        Edit������Ŀ = False
        Exit Function
    End If
    
    Select Case str��Ŀ
        Case "��Ӧ֢"
            If bln��Ժ Then
                Edit������Ŀ = (chkInfo(e_D_chkInfo_����_��Ժ_��Ӧ֢_5).Value = 1 And bln����) Or (chkInfo(e_D_chkInfo_����_��Ժ_��Ӧ֢_5).Value <> 1 And Not bln����)
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_��Ӧ֢_37).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_��Ӧ֢_37).Value <> 1 And Not bln����
            End If
        Case "ҩ��ѡ��"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_ҩ��ѡ��_36).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_ҩ��ѡ��_36).Value <> 1 And Not bln����
            End If
        Case "���μ���"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_���μ���_13).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_���μ���_13).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_���μ���_35).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_���μ���_35).Value <> 1 And Not bln����
            End If
        Case "ÿ�ո�ҩƵ��"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_ÿ�ո�ҩƵ��_14).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_ÿ�ո�ҩƵ��_14).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_ÿ�ո�ҩƵ��_34).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_ÿ�ո�ҩƵ��_34).Value <> 1 And Not bln����
            End If
        Case "�ܼ�"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_�ܼ�_15).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_�ܼ�_15).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_�ܼ�_33).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_�ܼ�_33).Value <> 1 And Not bln����
            End If
        Case "��ҩ;��"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_��ҩ;��_16).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_��ҩ;��_16).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_��ҩ;��_32).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_��ҩ;��_32).Value <> 1 And Not bln����
            End If
        Case "��ҩ�Ƴ�"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_��ҩ�Ƴ�_17).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_��ҩ�Ƴ�_17).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_��ҩ�Ƴ�_31).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_��ҩ�Ƴ�_31).Value <> 1 And Not bln����
            End If
        Case "��ǰ��ҩʱ��"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_��ǰ_23).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_��ǰ_23).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_��ǰ_28).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_��ǰ_28).Value <> 1 And Not bln����
            End If
        Case "������ҩ"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_����_24).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_����_24).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_����_27).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_����_27).Value <> 1 And Not bln����
            End If
        Case "������ҩ"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_������ҩ_22).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_������ҩ_22).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_������ҩ_29).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_������ҩ_29).Value <> 1 And Not bln����
            End If
        Case "������ҩ"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_����_25).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_����_25).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_����_26).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_����_26).Value <> 1 And Not bln����
            End If
        Case "����ҩ��"
            If bln��Ժ Then
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_��Ժ_����ҩ��_21).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_��Ժ_����ҩ��_21).Value <> 1 And Not bln����
            Else
                Edit������Ŀ = chkInfo(e_D_chkInfo_����_����_����ҩ��_30).Value = 1 And bln���� Or chkInfo(e_D_chkInfo_����_����_����ҩ��_30).Value <> 1 And Not bln����
            End If
    End Select
End Function

Private Sub Cls������ϸ(ByVal intIndex As Integer)
'���ܣ�����������Ĺ�ѡ��������仯ʱ��������Ӧ����������
    Select Case intIndex
        Case e_D_chkInfo_����_��Ժ_��Ӧ֢_5
            Call Set������Ŀֵ("��Ӧ֢", True, True)
        Case e_D_chkInfo_����_��Ժ_ҩ��ѡ��_12
            Call Set������Ŀֵ("ҩ��ѡ��", True, False)
        Case e_D_chkInfo_����_��Ժ_���μ���_13
            Call Set������Ŀֵ("���μ���", True, False)
        Case e_D_chkInfo_����_��Ժ_ÿ�ո�ҩƵ��_14
            Call Set������Ŀֵ("ÿ�ո�ҩƵ��", True, False)
        Case e_D_chkInfo_����_��Ժ_�ܼ�_15
            Call Set������Ŀֵ("�ܼ�", True, False)
        Case e_D_chkInfo_����_��Ժ_��ҩ;��_16
            Call Set������Ŀֵ("��ҩ;��", True, False)
        Case e_D_chkInfo_����_��Ժ_��ҩ�Ƴ�_17
            Call Set������Ŀֵ("��ҩ�Ƴ�", True, False)
        Case e_D_chkInfo_����_��Ժ_����ҩ��_21
            Call Set������Ŀֵ("����ҩ��", True, False)
        Case e_D_chkInfo_����_��Ժ_������ҩ_22
            Call Set������Ŀֵ("������ҩ", True, False)
        Case e_D_chkInfo_����_��Ժ_��ǰ_23
            Call Set������Ŀֵ("��ǰ��ҩʱ��", True, False)
        Case e_D_chkInfo_����_��Ժ_����_24
            Call Set������Ŀֵ("������ҩ", True, False)
        Case e_D_chkInfo_����_��Ժ_����_25
            Call Set������Ŀֵ("������ҩ", True, False)
        Case e_D_chkInfo_����_����_��Ӧ֢_37
            Call Set������Ŀֵ("��Ӧ֢", False, True)
        Case e_D_chkInfo_����_����_ҩ��ѡ��_36
            Call Set������Ŀֵ("ҩ��ѡ��", False, False)
        Case e_D_chkInfo_����_����_���μ���_35
            Call Set������Ŀֵ("���μ���", False, False)
        Case e_D_chkInfo_����_����_ÿ�ո�ҩƵ��_34
            Call Set������Ŀֵ("ÿ�ո�ҩƵ��", False, False)
        Case e_D_chkInfo_����_����_�ܼ�_33
            Call Set������Ŀֵ("�ܼ�", False, False)
        Case e_D_chkInfo_����_����_��ҩ;��_32
            Call Set������Ŀֵ("��ҩ;��", False, False)
        Case e_D_chkInfo_����_����_��ҩ�Ƴ�_31
            Call Set������Ŀֵ("��ҩ�Ƴ�", False, False)
        Case e_D_chkInfo_����_����_����ҩ��_30
            Call Set������Ŀֵ("����ҩ��", False, False)
        Case e_D_chkInfo_����_����_������ҩ_29
            Call Set������Ŀֵ("������ҩ", False, False)
        Case e_D_chkInfo_����_����_��ǰ_28
            Call Set������Ŀֵ("��ǰ��ҩʱ��", False, False)
        Case e_D_chkInfo_����_����_����_27
            Call Set������Ŀֵ("������ҩ", False, False)
        Case e_D_chkInfo_����_����_����_26
            Call Set������Ŀֵ("������ҩ", False, False)
    End Select
End Sub

Private Sub Set������Ŀֵ(ByVal str��Ŀ As String, ByVal bln��Ժ As Boolean, Optional ByVal blnALL As Boolean)
'���ܣ�����������Ŀ��ֵ����ѡ�򲻹�ѡ���¼������� ����ָ����ָ���з�Χ�ڵ�Ԫ���ֵ
'������blnAll ���ж����
    Dim i As Long
    
    With vsPJB
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_�ϼ�����) = str��Ŀ And Not blnALL Then
                If bln��Ժ Then
                    If Not .Cell(flexcpPicture, i, COL_����Ժ) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_����Ժ) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_����Ժ) = 1
                    Else
                        .TextMatrix(i, COL_����Ժ) = ""
                    End If
                    
                    If Not .Cell(flexcpPicture, i, COL_������Ժ) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_������Ժ) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_������Ժ) = 1
                    Else
                        .TextMatrix(i, COL_������Ժ) = ""
                    End If
                Else
                    If Not .Cell(flexcpPicture, i, COL_��������) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_��������) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_��������) = 1
                    Else
                        .TextMatrix(i, COL_��������) = ""
                    End If
                    
                    If Not .Cell(flexcpPicture, i, COL_����������) Is Nothing Then
                        Set .Cell(flexcpPicture, i, COL_����������) = img16.ListImages("Check").Picture
                        .Cell(flexcpData, i, COL_����������) = 1
                    Else
                        .TextMatrix(i, COL_����������) = ""
                    End If
                End If
            End If
            
            If blnALL And bln��Ժ Then
                If Not .Cell(flexcpPicture, i, COL_����Ժ) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_����Ժ) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_����Ժ) = 1
                Else
                    .TextMatrix(i, COL_����Ժ) = ""
                End If
                If Not .Cell(flexcpPicture, i, COL_������Ժ) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_������Ժ) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_������Ժ) = 1
                Else
                    .TextMatrix(i, COL_������Ժ) = ""
                End If
            ElseIf blnALL And Not bln��Ժ Then
                If Not .Cell(flexcpPicture, i, COL_��������) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_��������) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_��������) = 1
                Else
                    .TextMatrix(i, COL_��������) = ""
                End If
                If Not .Cell(flexcpPicture, i, COL_����������) Is Nothing Then
                    Set .Cell(flexcpPicture, i, COL_����������) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, COL_����������) = 1
                Else
                    .TextMatrix(i, COL_����������) = ""
                End If
            End If
        Next
    End With
End Sub

Private Sub Load��ҩ����()
'���ܣ����ز��˵���ҩ����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lngTmp As String
    Dim i As Long, j As Long
    Dim intCol As Integer
    
    On Error GoTo errH
    strSQL = "select ��Ŀ����,decode(��Ŀֵ,'1','',��Ŀֵ) as ��Ŀֵ,�������� from ����ҩ��������� where ����id=[1] and ���=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng���)
    For i = 1 To rsTmp.RecordCount
        With vsPJB
            For j = .FixedRows To .Rows - 1
                If .TextMatrix(j, COL_�������) = rsTmp!��Ŀ���� & "" Then
                    intCol = IIf(Val(rsTmp!�������� & "") = 0, COL_����Ժ, COL_��������)
                    If Val(rsTmp!��Ŀֵ & "") = 2 Then
                        Set .Cell(flexcpPicture, j, intCol) = img16.ListImages("UnCheck").Picture
                        .Cell(flexcpData, j, intCol) = 2
                    Else
                        .TextMatrix(j, intCol) = rsTmp!��Ŀֵ & ""
                    End If
                ElseIf .TextMatrix(j, COL_���������) = rsTmp!��Ŀ���� & "" Then
                    intCol = IIf(Val(rsTmp!�������� & "") = 0, COL_������Ժ, COL_����������)
                    If Val(rsTmp!��Ŀֵ & "") = 2 Then
                        Set .Cell(flexcpPicture, j, intCol) = img16.ListImages("UnCheck").Picture
                        .Cell(flexcpData, j, intCol) = 2
                    Else
                        .TextMatrix(j, intCol) = rsTmp!��Ŀֵ & ""
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
'����¼����ҩ����ʱ������������
   If InStr("0123456789'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Function UseKssDrug() As Boolean
'���ܣ������Ƿ�ʹ���˿���ҩ������ҽ���´����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
 
    strSQL = "Select 1" & vbNewLine & _
        "From ����ҽ����¼ A, ҩƷ���� C" & vbNewLine & _
        "Where a.������� = '5' And a.������Ŀid = c.ҩ��id And Nvl(c.������, 0) <> 0 And Exists" & vbNewLine & _
        " (Select 1 From ����ҽ������ B Where a.Id = b.ҽ��id) And a.����id = [1] And a.��ҳid = [2] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    UseKssDrug = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DateForShow(ByVal strDate As String) As String
'������ת��Ϊ������ʾ������
    Dim strTmp As String
    If strDate = "" Then DateForShow = "": Exit Function
    If Not IsDate(strDate) Then DateForShow = "": Exit Function
    strTmp = Format(strDate, "yyyy-mm-dd")
    DateForShow = Val(Split(strTmp, "-")(1)) & "/" & Val(Split(strTmp, "-")(2))
End Function

Private Sub InitRS(ByRef rsCtl As ADODB.Recordset)
'���ܣ��ؼ���¼��
    Dim arrFileds() As Variant
    
    Set rsCtl = New ADODB.Recordset
    rsCtl.Fields.Append "��Ϣ��", adVarChar, 100
    rsCtl.Fields.Append "��Ϣ����", adInteger '0-���ڣ�1-�ı���2-����
    rsCtl.Fields.Append "�ؼ���", adVarChar, 100
    rsCtl.Fields.Append "�ؼ�����", adBigInt
    rsCtl.Fields.Append "��Ϣ����", adBigInt
    rsCtl.CursorLocation = adUseClient
    rsCtl.LockType = adLockOptimistic
    rsCtl.CursorType = adOpenStatic
    rsCtl.Open
    
    arrFileds = Array("��Ϣ��", "��Ϣ����", "�ؼ���", "�ؼ�����", "��Ϣ����")
    
    With rsCtl
        .AddNew arrFileds, Array("��ҩǰ����", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰ����_11, 8)
        .AddNew arrFileds, Array("��ҩ������", Type����, "txtInfo", e_D_txtInfo_���_��ҩ������_55, 8)
        .AddNew arrFileds, Array("��ҩǰ��������", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰ��������_12, 16)
        .AddNew arrFileds, Array("��ҩ����������", Type����, "txtInfo", e_D_txtInfo_���_��ҩ����������_54, 16)
        
        .AddNew arrFileds, Array("��ҩǰ��ϸ������", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩǰ��ϸ������_13, 30)
        .AddNew arrFileds, Array("��ҩǰ��ϸ����������", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰ��ϸ����������_14, 16)
        .AddNew arrFileds, Array("��ҩ���ϸ������", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩ���ϸ������_53, 30)
        .AddNew arrFileds, Array("��ҩ���ϸ����������", Type����, "txtInfo", e_D_txtInfo_���_��ҩ���ϸ����������_52, 16)
        
        .AddNew arrFileds, Array("��ҩǰ������ϸ��", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩǰ������ϸ��_15, 30)
        .AddNew arrFileds, Array("��ҩǰ������ϸ������", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰ������ϸ������_16, 16)
        .AddNew arrFileds, Array("��ҩ��������ϸ��", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩ��������ϸ��_51, 30)
        .AddNew arrFileds, Array("��ҩ��������ϸ������", Type����, "txtInfo", e_D_txtInfo_���_��ҩ��������ϸ������_50, 16)
        
        .AddNew arrFileds, Array("��ҩǰC��Ӧ����", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩǰC��Ӧ����_17, 30)
        .AddNew arrFileds, Array("��ҩǰC��Ӧ��������", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰC��Ӧ��������_18, 16)
        .AddNew arrFileds, Array("��ҩ��C��Ӧ����", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩ��C��Ӧ����_49, 30)
        .AddNew arrFileds, Array("��ҩ��C��Ӧ��������", Type����, "txtInfo", e_D_txtInfo_���_��ҩ��C��Ӧ��������_48, 16)
        
        .AddNew arrFileds, Array("��ҩǰ�ȱ�ת��ø", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩǰ����ת��ø_19, 30)
        .AddNew arrFileds, Array("��ҩǰ�ȱ�ת��ø����", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰ����ת��ø����_20, 16)
        .AddNew arrFileds, Array("��ҩ��ȱ�ת��ø", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩ�����ת��ø_45, 30)
        .AddNew arrFileds, Array("��ҩ��ȱ�ת��ø����", Type����, "txtInfo", e_D_txtInfo_���_��ҩ�����ת��ø����_44, 16)
        
        .AddNew arrFileds, Array("��ҩǰ����", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩǰ����_21, 30)
        .AddNew arrFileds, Array("��ҩǰ��������", Type����, "txtInfo", e_D_txtInfo_���_��ҩǰ��������_22, 16)
        .AddNew arrFileds, Array("��ҩ����", Type�ı�, "txtInfo", e_D_txtInfo_���_��ҩ����_43, 30)
        .AddNew arrFileds, Array("��ҩ��������", Type����, "txtInfo", e_D_txtInfo_���_��ҩ��������_42, 16)
        
        .AddNew arrFileds, Array("�ٴ�΢�����鲡Դѧ�������", Type����, "txtInfo", e_D_txtInfo_���_��ԭѧ�������_23, 16)
        .AddNew arrFileds, Array("�ٴ�΢�����鲡Դѧ���걾", Type�ı�, "txtInfo", e_D_txtInfo_���_��ԭѧ���걾_26, 50)
        .AddNew arrFileds, Array("�ٴ�΢������ҩ����������", Type�ı�, "txtInfo", e_D_txtInfo_���_��ԭѧ�����ϸ����_25, 100)
        
        .AddNew arrFileds, Array("�ٴ�΢�����鲡Դѧ�������", Type����, "txtInfo", e_D_txtInfo_���_ҩ����������_24, 16)
        
        .AddNew arrFileds, Array("Ӱ��ѧ��ϲ�λ", Type�ı�, "txtInfo", e_D_txtInfo_Ӱ��_��λ_46, 50)
        .AddNew arrFileds, Array("Ӱ��ѧ��Ͻ���", Type�ı�, "txtInfo", e_D_txtInfo_Ӱ��_����_47, 100)
        
        .AddNew arrFileds, Array("��ע", Type�ı�, "txtInfo", e_D_txtInfo_��ע_57, 500)
    End With
    
End Sub

Private Function CheckData() As Boolean
'���ܣ���������ж�
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
        Set objCtl = txtInfo(Val(mrsCtl!�ؼ����� & ""))
        If objCtl.Enabled And objCtl.Locked = False And objCtl.Text <> "" Then
            lngColor = objCtl.BackColor
            strMsg = ""
            Select Case Val(mrsCtl!��Ϣ���� & "")
            Case 0 '����
                objCtl.BackColor = &HC0C0FF
                Call txtInfo_Validate(Val(mrsCtl!�ؼ����� & ""), blnDo)
                objCtl.BackColor = lngColor
                If blnDo Then
                    Exit Function
                End If
            Case 1 '�ı�
                If Len(objCtl.Text) > Val(mrsCtl!��Ϣ���� & "") Then
                    strMsg = mrsCtl!��Ϣ�� & "-����̫��(����¼��" & Val(mrsCtl!��Ϣ���� & "") & "���ַ���" & Val(mrsCtl!��Ϣ���� & "") \ 2 & "������)��"
                ElseIf InStr(objCtl.Text, "'") > 0 Then
                    strMsg = mrsCtl!��Ϣ�� & "���������ַ���ǵ����š�"
                End If
            Case 2 '����
                If Not IsNumeric(objCtl.Text) Then
                    strMsg = mrsCtl!��Ϣ�� & "Ҫ��ֻ��¼�����֡�"
                ElseIf Len(objCtl.Text) > Val(mrsCtl!��Ϣ���� & "") Then
                    strMsg = mrsCtl!��Ϣ�� & "-����̫��(����¼��" & Val(mrsCtl!��Ϣ���� & "") & "���ַ���"
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
            MsgBox "����¼�����ݹ�����", vbInformation, gstrSysName
            Cancel = True
        End If
    End With
End Sub
