VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPacsApplication 
   Caption         =   "检查申请"
   ClientHeight    =   10860
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPacsApplication.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   612
      Left            =   6600
      TabIndex        =   52
      Top             =   120
      Width           =   5052
      _Version        =   589884
      _ExtentX        =   8911
      _ExtentY        =   1080
      _StockProps     =   64
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12450
      Top             =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   11820
      Top             =   660
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMethod 
      Height          =   1815
      Left            =   7020
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6855
      Visible         =   0   'False
      Width           =   2055
      _cx             =   1993543209
      _cy             =   1993542785
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   12648447
      BackColorAlternate=   12648447
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPacsApplication.frx":1422
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin XtremeSuiteControls.TabControl tbcRequest 
      Height          =   336
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   6420
      _Version        =   589884
      _ExtentX        =   11324
      _ExtentY        =   593
      _StockProps     =   64
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   11310
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   15
      ScaleHeight     =   795
      ScaleWidth      =   6285
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   30
      Width           =   6285
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   9900
      Left            =   -60
      ScaleHeight     =   9900
      ScaleWidth      =   13950
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   13950
      Begin zlCISKernel.ucScrollPanel uspRequestPage 
         Bindings        =   "frmPacsApplication.frx":145E
         Height          =   9915
         Left            =   75
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   13785
         _ExtentX        =   24289
         _ExtentY        =   17515
         UCBackColor     =   14737632
         Begin VB.PictureBox picPart 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3900
            Left            =   0
            ScaleHeight     =   3900
            ScaleWidth      =   13575
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "5330-105"
            Top             =   4425
            Visible         =   0   'False
            Width           =   13575
            Begin VB.CheckBox chkPriority 
               BackColor       =   &H00FFFFFF&
               Caption         =   "紧急医嘱"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   4575
               TabIndex        =   7
               Top             =   270
               Width           =   1185
            End
            Begin MSComCtl2.DTPicker dtpExeTime 
               Height          =   360
               Left            =   6855
               TabIndex        =   10
               Top             =   240
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   635
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   251330563
               CurrentDate     =   41348.5555555556
            End
            Begin VB.ComboBox cbxExeRoom 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   10425
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   240
               Width           =   3015
            End
            Begin MSComctlLib.ImageList img16 
               Left            =   5415
               Top             =   1125
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
                     Picture         =   "frmPacsApplication.frx":1472
                     Key             =   "c0"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmPacsApplication.frx":1A0C
                     Key             =   "c1"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmPacsApplication.frx":1FA6
                     Key             =   "o0"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmPacsApplication.frx":2540
                     Key             =   "o1"
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtFind 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1140
               TabIndex        =   6
               Top             =   240
               Width           =   3300
            End
            Begin VSFlex8Ctl.VSFlexGrid vfgList 
               Height          =   3135
               Left            =   5865
               TabIndex        =   12
               Top             =   675
               Width           =   7590
               _cx             =   13388
               _cy             =   5530
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
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
               BackColorSel    =   16769985
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
               Rows            =   7
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmPacsApplication.frx":2ADA
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
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
               Begin VB.CommandButton cmd 
                  Caption         =   "…"
                  Height          =   240
                  Left            =   3435
                  TabIndex        =   8
                  TabStop         =   0   'False
                  ToolTipText     =   "选择项目(*)"
                  Top             =   1035
                  Visible         =   0   'False
                  Width           =   270
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vfgRequestProject 
               Height          =   3150
               Left            =   135
               TabIndex        =   9
               Top             =   675
               Width           =   5640
               _cx             =   9948
               _cy             =   5556
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
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
               BackColorSel    =   16769985
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
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   ""
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
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
               FrozenRows      =   1
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
            Begin VB.Label labExeRoom 
               BackStyle       =   0  'Transparent
               Caption         =   "执行科室："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   9435
               TabIndex        =   44
               Top             =   300
               Width           =   1530
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "执行时间："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   5850
               TabIndex        =   42
               Top             =   300
               Width           =   1575
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "项目定位："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   165
               TabIndex        =   39
               Top             =   300
               Width           =   1350
            End
            Begin VB.Label labPartArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检查部位区域"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   4950
               TabIndex        =   22
               Top             =   975
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.PictureBox picBaseInf 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   0
            ScaleHeight     =   1185
            ScaleWidth      =   13635
            TabIndex        =   3
            TabStop         =   0   'False
            Tag             =   "90-0"
            Top             =   90
            Visible         =   0   'False
            Width           =   13635
            Begin VB.PictureBox picAuditing 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               FillColor       =   &H80000008&
               ForeColor       =   &H80000008&
               Height          =   525
               Left            =   150
               ScaleHeight     =   495
               ScaleWidth      =   1125
               TabIndex        =   48
               Top             =   105
               Visible         =   0   'False
               Width           =   1155
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "已校验"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000C000&
                  Height          =   270
                  Left            =   90
                  TabIndex        =   49
                  Top             =   120
                  Width           =   915
               End
            End
            Begin VB.Label labConditionTag 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "急!!"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   12090
               TabIndex        =   47
               Top             =   120
               Width           =   675
            End
            Begin VB.Shape shpBase 
               BorderStyle     =   0  'Transparent
               FillStyle       =   0  'Solid
               Height          =   60
               Left            =   0
               Top             =   1095
               Width           =   13590
            End
            Begin VB.Label labInPatientValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "123456789"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   11490
               TabIndex        =   35
               Top             =   735
               Width           =   1410
            End
            Begin VB.Label labInPatient 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10410
               TabIndex        =   34
               Top             =   735
               Width           =   1200
            End
            Begin VB.Label labOutPatientValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "123456789"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8190
               TabIndex        =   33
               Top             =   735
               Width           =   1410
            End
            Begin VB.Label labAgeValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "30岁"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5715
               TabIndex        =   32
               Top             =   735
               Width           =   585
            End
            Begin VB.Label labSexValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "女"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3735
               TabIndex        =   31
               Top             =   735
               Width           =   285
            End
            Begin VB.Label labNameValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "张三丰"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1005
               TabIndex        =   30
               Top             =   735
               Width           =   855
            End
            Begin VB.Label labOutPatientNo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7110
               TabIndex        =   29
               Top             =   735
               Width           =   1200
            End
            Begin VB.Label labAge 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4890
               TabIndex        =   28
               Top             =   735
               Width           =   855
            End
            Begin VB.Label labSex 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2985
               TabIndex        =   27
               Top             =   735
               Width           =   870
            End
            Begin VB.Label labName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   165
               TabIndex        =   26
               Top             =   735
               Width           =   900
            End
            Begin VB.Label labTitle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CT检查申请单"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   4845
               TabIndex        =   25
               Top             =   105
               Width           =   2730
            End
            Begin VB.Label labBaseArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "基本信息区域"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   2520
               TabIndex        =   24
               Top             =   405
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.PictureBox picInput 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3150
            Left            =   15
            ScaleHeight     =   3150
            ScaleWidth      =   13575
            TabIndex        =   2
            TabStop         =   0   'False
            Tag             =   "1865-90"
            Top             =   1275
            Visible         =   0   'False
            Width           =   13575
            Begin VB.CommandButton cmdInput 
               Appearance      =   0  'Flat
               Caption         =   "…"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   12990
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   45
               Width           =   375
            End
            Begin VB.TextBox rtbInputPro 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   1680
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               TabStop         =   0   'False
               Text            =   "frmPacsApplication.frx":2BE6
               Top             =   15
               Visible         =   0   'False
               Width           =   11280
            End
            Begin VB.Label labMustPro 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   585
               TabIndex        =   41
               Top             =   345
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape shpBackLine 
               BorderStyle     =   0  'Transparent
               FillStyle       =   0  'Solid
               Height          =   30
               Index           =   0
               Left            =   1680
               Top             =   405
               Visible         =   0   'False
               Width           =   11715
            End
            Begin VB.Shape shpInputPro 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00E0E0E0&
               FillStyle       =   0  'Solid
               Height          =   45
               Index           =   0
               Left            =   60
               Top             =   1395
               Visible         =   0   'False
               Width           =   13380
            End
            Begin VB.Label labInputPro 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "录入项目："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   45
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label labInputArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "录入项目区域"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   4680
               TabIndex        =   23
               Top             =   1800
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.PictureBox picRequestInf 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   0
            ScaleHeight     =   1485
            ScaleWidth      =   13545
            TabIndex        =   13
            TabStop         =   0   'False
            Tag             =   "9405-135"
            Top             =   8340
            Visible         =   0   'False
            Width           =   13545
            Begin VB.TextBox txtCurStudyProject 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   15
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   90
               Width           =   13470
            End
            Begin VB.Label lblInsureInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保信息："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   51
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label labPrice 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "￥：---"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   315
               Left            =   30
               TabIndex        =   46
               Top             =   585
               Width           =   1200
            End
            Begin VB.Shape shpRequest 
               BorderStyle     =   0  'Transparent
               FillStyle       =   0  'Solid
               Height          =   60
               Left            =   15
               Top             =   435
               Width           =   13470
            End
            Begin VB.Label labRequestInfArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "申请信息区域"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   4860
               TabIndex        =   20
               Top             =   75
               Visible         =   0   'False
               Width           =   1800
            End
            Begin VB.Label labRequestDoct 
               BackStyle       =   0  'Transparent
               Caption         =   "申请人："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4080
               TabIndex        =   19
               Top             =   570
               Width           =   1245
            End
            Begin VB.Label labRequestDoctValue 
               BackStyle       =   0  'Transparent
               Caption         =   "才吃饭"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5130
               TabIndex        =   18
               Top             =   555
               Width           =   1365
            End
            Begin VB.Label labRequestRoom 
               BackStyle       =   0  'Transparent
               Caption         =   "申请科室："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6465
               TabIndex        =   17
               Top             =   570
               Width           =   1530
            End
            Begin VB.Label labRequestRoomValue 
               BackStyle       =   0  'Transparent
               Caption         =   "耳鼻喉科室-"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7800
               TabIndex        =   16
               Top             =   555
               Width           =   1755
            End
            Begin VB.Label labRequestTime 
               BackStyle       =   0  'Transparent
               Caption         =   "申请时间："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   9525
               TabIndex        =   15
               Top             =   570
               Width           =   1575
            End
            Begin VB.Label labRequestTimeValue 
               BackStyle       =   0  'Transparent
               Caption         =   "2013-03-06 17:50"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10950
               TabIndex        =   14
               Top             =   555
               Width           =   2565
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   50
      Top             =   10500
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2381
            MinWidth        =   882
            Picture         =   "frmPacsApplication.frx":2BEF
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13600
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPacsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_FIXEDELEMENT_DIAGNOSE As String = "最后诊断"

'菜单id定义
Private Enum TApplicationMenu
    amAppFile = 0
    amAppPreview = 1
    amAppPrint = 2
    
    amAppEdit = 3
    amAppSave = 4
    amAppDel = 5
    amAppExit = 6
    
    amAppKind = 7
    amAppSend = 8
    
    amAppType = 9
End Enum


'项目列名定义
Private Enum TProjectCol
    pcId = 0                'Id列
    pcName = 1              '项目名称列
    pcRoomType = 2          '科室类型列
    pcMethod = 3            '部分方法组合列
    pcExeRoom = 4           '执行科室列
    pcNormalCol = 5            '常规列
    pcBedCol = 6               '床旁列
    pcOperCol = 7              '书中列
End Enum


'病人来源类型
Private Enum TPatientFrom
    pfOutPatient = 1    '门诊病人
    pfInPatient = 2     '住院病人
End Enum


Private mrsRequestPart As ADODB.Recordset       '保存该申请单的所有检查部位，避免选择项目时，从数据库中读取
          

'病人信息
Private Type TPatientInf
    lngID As Long                       '当前病人ID：由外部传入
    lngFrom As Long                     '病人来源：由外部传入 ,1表示门诊，2表示住院
    lngInsure As Long                   '病人险类，用于对医保项目进行检查时传入
        
    lngPageId As Long                   '主页Id：住院病人才有主页Id
    lngRegId  As Long                   '挂号id：门诊病人才有挂号Id
    
    strConditionTag As String           '病情标记 如“危！！ 急！！”

    strName As String                   '病人姓名
    strSex As String                    '病人性别
    strRegNo As String                  '挂号单据
    strRegDate As String                '登记时间
    lngRoomId As Long                   '病人科室ID
    strAge As String                    '年龄
    strInNO As String                   '住院号
    strOutNo As String                  '门诊号
    strInHospitalDate As String         '入院时间
End Type


'医生信息
Private Type TDoctorInf
    lngID As Long
    str用户名 As String
    str姓名 As String
    lng部门ID As Long
End Type


Private Type TRECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Type TPoint
    X As Long
    Y As Long
End Type

Private mCurPatientInf As TPatientInf
Private mCurDoctorInf As TDoctorInf

Public mobjAppDatas As New Scripting.Dictionary         '保存申请页面的数据

Private mblnPageUpdateState As Boolean
Private mblnIsLoadRequestPage As Boolean        '是否加载申请页
Public mblnIsSaveRequestPage As Boolean        '是否保存了申请页

Private mintBabyID As Integer                   '婴儿序号
Private mlngCurDeptId As Long                   '当前科室Id
Private mlngUpdateAppNo As Long                 '需要更新的医嘱申请序号
Private mstrDoctorName As String
Private mlngProjectId As Long                   '需要自动定位的诊疗项目的项目ID
Private mlngRequestPageCount As Long
Private mblnIsRestoreTab As Boolean

Private mfrmPacsApplyWord As frmPacsApplyWord

Private mstrRequestAffixConfig As String        '申请附项配置，附项名1:必填,排列,要素Id|附项名2:必填,排列,要素Id|附项名n:必填,排列,要素Id

Private mblnShowWord  As Boolean        '按钮是否显示词句界面，True-词句界面，False-模板界面

Private mobjLastControl As Object
Private mobjOwner As Object
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mobjEmrInterface As Object           '新版病历申请附项读取部件

'Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As TRECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As TPoint)
Private Declare Sub OffsetRect Lib "user32" (lpRect As TRECT, ByVal X As Long, ByVal Y As Long)


Public Sub InitComponents(ByVal lngDeptID As Long, objOwner As Object)
    mlngCurDeptId = lngDeptID
    Set mobjOwner = objOwner
    
    Call GetUserInfo
End Sub


Public Sub GetUserInfo()
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    If Not rsTmp.EOF Then
        mCurDoctorInf.lngID = rsTmp!ID
        mCurDoctorInf.lng部门ID = IIF(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        mCurDoctorInf.str姓名 = IIF(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        mCurDoctorInf.str用户名 = IIF(IsNull(rsTmp!用户名), "", rsTmp!用户名)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function ShowApplicationForm(ByVal lngPatientID As Long, _
                                    ByVal lngCallFrom As Long, _
                                    ByVal lngPatientRegId As Long, _
                                    ByVal lngPatientPageId As Long, _
                                    ByVal lngUpdateAppNoOrAdvId As Long, _
                                    ByRef objAppPages() As clsApplicationData, _
                                    Optional ByVal intBabyID As Integer = 0, _
                                    Optional ByVal blnEdit As Boolean = True, _
                                    Optional ByVal lngProjectId As Long = 0) As Boolean
    Dim curItem
    Dim i As Long
    
    ShowApplicationForm = False
    
    If lngPatientID <= 0 Then
        MsgBox "当前病人ID无效，ID值为 [" & lngPatientID & "]，不能下达检查申请。", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If mlngCurDeptId <= 0 Then
        MsgBox "当前科室ID无效，ID值为 [" & mlngCurDeptId & "]，不能下达检查申请。", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    '释放申请单数据
    Set mobjAppDatas = New Scripting.Dictionary
    
    'objAppPages有值时，为修改申请，此时lngUpdateAppNoOrAdvId>0
    If SafeArrayGetDim(objAppPages) > 0 Then
        For i = 0 To UBound(objAppPages)
            mobjAppDatas.Add "_" & objAppPages(i).lngApplicationPageId & "_" & i, objAppPages(i)
        Next
    End If

    mblnPageUpdateState = lngUpdateAppNoOrAdvId = 1
    mblnIsLoadRequestPage = True
    mblnIsSaveRequestPage = False
    mlngProjectId = lngProjectId
    
    mCurPatientInf.lngFrom = lngCallFrom
    mCurPatientInf.lngID = lngPatientID
    mCurPatientInf.lngRegId = lngPatientRegId
    mCurPatientInf.lngPageId = lngPatientPageId
    mintBabyID = intBabyID
    
    mlngUpdateAppNo = lngUpdateAppNoOrAdvId
    
    picBaseInf.Enabled = blnEdit = True
    picInput.Enabled = blnEdit = True
    picPart.Enabled = blnEdit = True
    txtCurStudyProject.Enabled = blnEdit = True
    
    Me.Show 1, mobjOwner
    
    If Me.mobjAppDatas.Count > 0 Then
        ReDim objAppPages(Me.mobjAppDatas.Count - 1)
        
        i = 0
        For Each curItem In Me.mobjAppDatas.Items
            If Not curItem Is Nothing Then
                Set objAppPages(i) = curItem
                i = i + 1
            End If
        Next
    End If
    
    '释放数据
    For i = mobjAppDatas.Count - 1 To 0 Step -1
        Call mobjAppDatas.Remove(mobjAppDatas.Keys(i))
        Set mobjAppDatas.Item(i) = Nothing
    Next i
    
    ShowApplicationForm = Me.mblnIsSaveRequestPage
    
End Function

Private Function GetCurRequestPageFormat(ByVal lngTabIndex As Long) As clsApplicationData
'获取当前申请页面的保存格式

    Dim objAppData As New clsApplicationData
    Dim strAffix As String
    Dim lngExeType As Long
    Dim objText As TextBox
    Dim strExeRoomInf As String
    Dim lngProjectRowIndex As Long
    Dim strProName As String
    Dim strElement As String
    Dim lngRequestPageIndex As Long
        
    Set GetCurRequestPageFormat = Nothing
    
    strAffix = ""
    
    For Each objText In rtbInputPro
        If objText.Index > 0 And objText.Visible Then
            strProName = GetInputProName(objText.Tag)
            strElement = GetInputProElement(objText.Tag)
            
            If strAffix <> "" Then strAffix = strAffix & "|"
            strAffix = strAffix & strProName & ":" & objText.Text
            
            If strElement = M_STR_FIXEDELEMENT_DIAGNOSE Then
                '保存临床诊断的诊断Id
                objAppData.strDiagnoseId = objText.ToolTipText
            End If
        End If
    Next
    
    lngRequestPageIndex = tbcPage.Selected.Index
    lngProjectRowIndex = Val(vfgRequestProject.Tag)
    
    lngExeType = 0    '执行类型
    If Val(vfgRequestProject.Cell(flexcpData, lngProjectRowIndex, TProjectCol.pcBedCol)) = 1 Then
        lngExeType = 1
    ElseIf Val(vfgRequestProject.Cell(flexcpData, lngProjectRowIndex, TProjectCol.pcOperCol)) = 1 Then
        lngExeType = 2
    End If

    strExeRoomInf = vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom)
    
    objAppData.strApplicationPageName = tbcPage.Item(lngRequestPageIndex).Caption                '申请单名称
    objAppData.lngApplicationPageId = GetRequestId(tbcPage.Item(lngRequestPageIndex).Tag)  '申请单Id
    objAppData.strRequestTime = zlDatabase.Currentdate                                      '申请时间
    objAppData.strRequestAffixCfg = mstrRequestAffixConfig
    objAppData.blnIsModify = True
    
    objAppData.blnIsPriority = IIF(chkPriority.value <> 0, True, False)                 '是否紧急
    objAppData.lngProjectId = Val(vfgRequestProject.Cell(flexcpData, lngProjectRowIndex, TProjectCol.pcId))  '诊疗项目Id
    objAppData.lngExeType = lngExeType                                                  '执行类型
    objAppData.strStartExeTime = dtpExeTime.value                                       '执行时间
    objAppData.lngExeRoomId = Val(strExeRoomInf)                                        '执行科室Id
    objAppData.strExeRoomName = Replace(strExeRoomInf, Val(strExeRoomInf) & "-", "") '执行科室名称
    objAppData.lngExeRoomType = Val(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcRoomType))
    objAppData.strPartMethod = Trim(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod))   '部位方法
    objAppData.strRequestAffix = strAffix                                               '申请附项(10)
    objAppData.lngRequestRoomId = mlngCurDeptId                                         '申请科室(11)
    objAppData.strRequestDoctor = UserInfo.姓名
    
    If mCurPatientInf.lngFrom = TPatientFrom.pfInPatient Then                           '只有住院医嘱才存在补录情况
        objAppData.blnIsAdditionalRec = IIF(DateDiff("n", dtpExeTime.value, zlDatabase.Currentdate) >= gint补录间隔, True, False)  '是否补录医嘱
    Else
        objAppData.blnIsAdditionalRec = False
    End If
    
    Set GetCurRequestPageFormat = objAppData
End Function

Private Function VerificationDataIsRight() As String
'验证数据是否正确
    Dim objRich As TextBox
    Dim i As Long
    Dim aryPart() As String
    Dim lngProjectRowIndex As Long
    Dim strMethod As String
    
    VerificationDataIsRight = ""
    
    lngProjectRowIndex = Val(vfgRequestProject.Tag)
    
    '判断必录字段正确性
    For Each objRich In rtbInputPro
        If Not objRich Is Nothing Then
            '判断该数据是否为必录，如果为必录，则必须录入数据后才能申请
            If labMustPro(objRich.Index).Visible Then
                If Trim(objRich.Text) = "" Then
                    VerificationDataIsRight = "【" & objRich.Tag & "】不能为空。"
                    objRich.SetFocus
                    
                    Exit Function
                End If
            End If
        End If
    Next
    
    
    '判断检查部位正确性
    If lngProjectRowIndex < 0 Then
        VerificationDataIsRight = "请选择所需检查的项目。"
        vfgRequestProject.SetFocus
        
        Exit Function
    End If
    
    '判断部位方法正确性
    If vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod) = "" Then
        VerificationDataIsRight = "请设置检查项目【" & vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcName) & "】对应的检查部位和方法。"
        If vfgRequestProject.Row <> lngProjectRowIndex Then vfgRequestProject.Row = lngProjectRowIndex
        vfgList.SetFocus
        
        Exit Function
    End If
    
    '验证所选部位是否设置了检查方法
    strMethod = vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod)
    
    '当选择了检查部位后，则需判断是否设置对应的检查方法，如果未选择，则不允许进行保存
    If Trim(strMethod) <> "" Then
        aryPart = Split(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod), "|")
        For i = 0 To UBound(aryPart)
            If Split(aryPart(i), ";")(1) = "" Then
                VerificationDataIsRight = "请对【" & vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcName) & "】项目的" & _
                                            vfgList.TextMatrix(0, 1) & " 【" & Split(aryPart(i), ";")(0) & "】对应的" & vfgList.TextMatrix(0, 2) & "设置完整。"
                If vfgRequestProject.Row <> lngProjectRowIndex Then vfgRequestProject.Row = lngProjectRowIndex
                Call vfgList.SetFocus
                
                Exit Function
            End If
        Next i
    End If
    
    '判断是否设置执行科室
    If vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom) = "" Then
        VerificationDataIsRight = "请设置检查项目【" & vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcName) & "】对应的检查科室。"
        If vfgRequestProject.Row <> lngProjectRowIndex Then vfgRequestProject.Row = lngProjectRowIndex
        
        cbxExeRoom.SetFocus
        
        Exit Function
    End If
    
End Function


Private Sub ReadRequestPatientInf()
'载入申请单的基本信息如病人姓名性别,申请科室,申请人等
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngDate As Long

    If mCurPatientInf.lngRegId <> 0 Then '查询门诊
        strSQL = "select a.姓名,a.性别,a.年龄,a.险类,a.门诊号,a.住院号," & _
                " Nvl(Nvl(b.续诊科室ID,Decode(b.转诊状态,1,b.转诊科室ID,NULL)),b.执行部门ID) as 病人科室ID,b.Id as 挂号Id, b.No as 挂号单,b.登记时间,b.急诊 " & _
                " from 病人信息 a, 病人挂号记录 b " & _
                " where a.病人Id=b.病人ID and a.病人Id=[1] and  b.id=[2]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询申请单基本信息", mCurPatientInf.lngID, mCurPatientInf.lngRegId)
    Else    '查询住院
        strSQL = "select a.姓名,a.性别,a.年龄,a.险类,a.门诊号,a.住院号,a.入院时间, a.当前科室ID as 病人科室Id,b.主页Id, b.病人性质,b.当前病况 from 病人信息 a, 病案主页 b " & _
                " where a.病人Id=b.病人ID and a.病人Id=[1] and b.主页Id=[2] "
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询申请单基本信息", mCurPatientInf.lngID, mCurPatientInf.lngPageId)
    End If
    
    Call ClearRequestPatientInf
    
    If rsData.RecordCount > 0 Then
    
        mCurPatientInf.lngRoomId = Val(NVL(rsData!病人科室id))
        mCurPatientInf.strName = NVL(rsData!姓名)
        mCurPatientInf.strSex = NVL(rsData!性别)
        mCurPatientInf.strAge = NVL(rsData!年龄)
        mCurPatientInf.lngInsure = Val(NVL(rsData!险类))
        mCurPatientInf.strInNO = NVL(rsData!住院号)
        mCurPatientInf.strOutNo = NVL(rsData!门诊号)
        
        If mCurPatientInf.lngInsure <= 0 Then lblInsureInfo.Caption = "医保信息：非医保病人"
        
        If mCurPatientInf.lngRegId <> 0 Then
            mCurPatientInf.strRegNo = NVL(rsData!挂号单)
            mCurPatientInf.strConditionTag = IIF(Val(NVL(rsData!急诊)) <> 0, "急!!", "")
            mCurPatientInf.strRegDate = NVL(rsData!登记时间)
        ElseIf mCurPatientInf.lngPageId <> 0 Then
            '如果是住院病人，则根据病人性质从新判断病人来源
            mCurPatientInf.lngFrom = IIF(Val(NVL(rsData!病人性质)) = 1, 1, 2)
            mCurPatientInf.strConditionTag = Decode(Val(NVL(rsData!当前病况)), 9, "重!!", 10, "危!!", "")
            mCurPatientInf.strInHospitalDate = NVL(rsData!入院时间)
        Else
            '...
        End If
    End If
    
    If mintBabyID > 0 Then
        strSQL = "select 婴儿姓名,婴儿性别,出生时间,Round(Decode(死亡时间,NULL,SysDate,死亡时间)-出生时间) || '天' As 婴儿年龄 from 病人新生儿记录 where 病人id=[1] and 主页id=[2] and 序号=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询申请单基本信息", mCurPatientInf.lngID, mCurPatientInf.lngPageId, mintBabyID)
         
        If rsData.RecordCount > 0 Then
            mCurPatientInf.strName = NVL(rsData!婴儿姓名)
            mCurPatientInf.strSex = NVL(rsData!婴儿性别)
            mCurPatientInf.strAge = NVL(rsData!婴儿年龄)
        End If
    End If
End Sub


Private Sub LoadRequestBaseInf()
'载入申请单共有的基本信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    picBaseInf.Visible = True
    
    labNameValue.Caption = mCurPatientInf.strName
    labSexValue.Caption = mCurPatientInf.strSex
    labAgeValue.Caption = mCurPatientInf.strAge
    labInPatientValue.Caption = mCurPatientInf.strInNO
    labOutPatientValue.Caption = mCurPatientInf.strOutNo
    labConditionTag.Caption = mCurPatientInf.strConditionTag
    
    If labRequestDoctValue.Caption = "" Then labRequestDoctValue.Caption = mCurDoctorInf.str姓名
    
    picRequestInf.Visible = True
    
    '读取申请相关信息
    strSQL = "select 名称 from 部门表 where id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询申请科室信息", mlngCurDeptId)
    
    If rsData.RecordCount > 0 Then
        labRequestRoomValue.Caption = NVL(rsData!名称)
    Else
        labRequestRoomValue.Caption = ""
    End If
    
    
End Sub

Private Function GetOrderInspectInfo(ByVal lng病人ID As Long, ByVal strCondition As String, ByVal intType As Integer, ByVal lng就诊ID As Long) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等
    Dim strText As String
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, lng病人ID, lng就诊ID, strCondition)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(lng病人ID, strCondition)
        End If
    End If
    GetOrderInspectInfo = strText
End Function

Public Function GetAppendItemValue(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal lng婴儿 As Long, ByVal str中文名 As String, ByVal str项目 As String, ByVal lng要素ID As Long) As String
'功能：获取指定的申请附项值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim intType As Integer
    Dim lng就诊ID As Long
     
    On Error GoTo errH
    
    If str挂号单 <> "" Then
        '从病人之前已保存的医嘱中提取,以最后填写的为准
        strSQL = " Select 内容 From (" & _
            " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
            " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
            IIF(str挂号单 <> "", " And A.挂号单=[2]", " And Nvl(A.主页ID,0)=[3]") & _
            " And B.项目=[5] And B.内容 is Not Null and nvl(a.医嘱状态,0)<>4" & _
            " Order by A.开嘱时间 Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplication", lng病人ID, str挂号单, lng主页ID, lng婴儿, str项目)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
    End If
    
    '未读取到值则如果有对应要素，从要素提取函数读取
    If lng要素ID <> 0 And strText = "" Then
        '先老版，再新版
        If str挂号单 <> "" Then '门诊
            strSQL = "Select Zl_Replace_Element_Value(B.中文名,[1],A.ID,1) as 内容" & _
                " From 病人挂号记录 A,诊治所见项目 B Where A.NO=[2] And B.ID=[3] And a.记录性质=1 And a.记录状态=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, str挂号单, lng要素ID)
        Else
            strSQL = "Select Zl_Replace_Element_Value(中文名,[1],[2],2) as 内容 From 诊治所见项目 Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, lng要素ID)
        End If
        If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
        If strText = "" Then
            If str挂号单 <> "" Then
                strSQL = "select a.id From 病人挂号记录 A Where A.NO=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str挂号单)
                lng就诊ID = Val(rsTmp!ID & "")
                intType = 1
            Else
                lng就诊ID = lng主页ID
                intType = 2
            End If
            strText = GetOrderInspectInfo(lng病人ID, str中文名, intType, lng就诊ID)
        End If
    End If
    
    If str挂号单 = "" And strText = "" Then
        '从病人之前已保存的医嘱中提取,以最后填写的为准
        strSQL = " Select 内容 From (" & _
            " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
            " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
            IIF(str挂号单 <> "", " And A.挂号单=[2]", " And Nvl(A.主页ID,0)=[3]") & _
            " And B.项目=[5] And B.内容 is Not Null and nvl(a.医嘱状态,0)<>4" & _
            " Order by A.开嘱时间 Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplication", lng病人ID, str挂号单, lng主页ID, lng婴儿, str项目)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
    End If
    
    GetAppendItemValue = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'执行菜单事件
On Error GoTo errHandle
    '转移焦点，使其dtpicker的change事件能立即触发
    If tbcPage.Visible Then tbcPage.SetFocus
    
    Select Case Control.ID
        Case TApplicationMenu.amAppSave
            If Not SaveRequest(tbcRequest.Selected.Index) Then
                Exit Sub
            End If
            
        Case TApplicationMenu.amAppDel
            Call CancelRequest(tbcRequest.Selected.Index)
            
        Case TApplicationMenu.amAppType * 100 To TApplicationMenu.amAppType * 100 + 99
            Call SwitchRequestPageType(Control.Category, IIF(Len(Control.Parent.Title) > 0, Control.Parent.Title & "-", "") & Control.Caption)
            
        Case TApplicationMenu.amAppEdit
            Unload Me
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetRequestItemCount(ByVal lngRequestId As Long) As Long
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 0 To tbcRequest.ItemCount - 1
        If GetRequestId(tbcRequest.Item(i).Tag) = lngRequestId Then
            lngCount = lngCount + 1
        End If
    Next i
    
    GetRequestItemCount = lngCount
End Function

Private Sub CancelRequest(ByVal lngRequestTabIndex As Long)
'撤销申请
    Dim blnIsSelect As Boolean
    Dim i As Long
    
    If MsgBox("撤销本次申请后，数据将不能恢复，是否继续？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '从mobjAppDatas中移除保存的数据
    If mobjAppDatas.Exists("_" & CStr(tbcRequest.Item(lngRequestTabIndex).Tag)) Then
        Call mobjAppDatas.Remove("_" & tbcRequest.Item(lngRequestTabIndex).Tag)
    End If
    
    mblnPageUpdateState = False
    
    If GetRequestItemCount(GetRequestId(tbcRequest.Item(lngRequestTabIndex).Tag)) > 1 Then
        If lngRequestTabIndex > 0 Then
            '向前移动申请页面
            tbcRequest.Item(lngRequestTabIndex).Tag = ""    '避免撤销时提示保存
            Call tbcRequest.RemoveItem(lngRequestTabIndex)
        ElseIf lngRequestTabIndex < tbcRequest.ItemCount - 1 Then
            '向后移动申请页面
            tbcRequest.Item(lngRequestTabIndex).Tag = ""    '避免撤销时提示保存
            Call tbcRequest.RemoveItem(lngRequestTabIndex)
        End If
        
        blnIsSelect = False
        For i = lngRequestTabIndex - 1 To 0 Step -1
            If tbcRequest.Item(i).Visible = True Then
                tbcRequest.Item(i).Selected = True
                blnIsSelect = True
                Exit For
            End If
        Next i
        
        If blnIsSelect = False Then
            For i = lngRequestTabIndex + 1 To tbcRequest.ItemCount - 1
                If tbcRequest.Item(i).Visible = True Then
                    tbcRequest.Item(i).Selected = True
                    Exit For
                End If
            Next i
        End If
        
        tbcRequest.Tag = tbcRequest.Selected.Index & "-" & tbcRequest.Selected.Caption
        
    Else
        tbcRequest.Item(lngRequestTabIndex).Caption = "新项目"
        Call LoadRequestPage(GetRequestId(tbcRequest.Item(lngRequestTabIndex).Tag), tbcRequest.Item(lngRequestTabIndex).Tag)
    End If
End Sub


Private Function SaveRequest(ByVal lngRequestTabIndex As Long) As Boolean
    Dim strVerification As String
    Dim lngCurAdviceId As Long
    
    Dim strDataKey As String
    Dim objOldAppData As clsApplicationData
    Dim objCurAppData As clsApplicationData
    
    Dim strMsg As String
    Dim strCheckContext As String   '需要对码的检查项目
    Dim lngMsgResult As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strCurProName As String
    
    SaveRequest = False
    
    If lngRequestTabIndex < 0 Then Exit Function
    
    strVerification = VerificationDataIsRight()
    If strVerification <> "" Then
        MsgBox strVerification, vbOKOnly, Me.Caption
        Exit Function
    End If

    If mCurPatientInf.lngFrom = TPatientFrom.pfOutPatient Then  '门诊病人
        
        If Trim(mCurPatientInf.strRegDate) <> "" Then
            If dtpExeTime.value < CDate(mCurPatientInf.strRegDate) Then
                Call MsgBox("执行时间不允许早于病人的挂号时间。", vbInformation + vbOKOnly, Me.Caption)
                dtpExeTime.SetFocus
                
                Exit Function
            End If
        End If
                
    Else    '住院病人
        
        If Trim(mCurPatientInf.strInHospitalDate) <> "" Then
            If dtpExeTime.value < CDate(mCurPatientInf.strInHospitalDate) Then
                Call MsgBox("执行时间不允许早于病人的入院时间。", vbInformation + vbOKOnly, Me.Caption)
                dtpExeTime.SetFocus
                
                Exit Function
            End If
        End If
        
        '判断是否为补录医嘱
        If DateDiff("n", dtpExeTime.value, zlDatabase.Currentdate) >= gint补录间隔 And gint补录间隔 > 0 Then
            If MsgBox("检查的开始执行时间早于当前时间，将自动按照补录方式进行处理，是否继续？", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                dtpExeTime.SetFocus
                Exit Function
            End If
            
            If chkPriority.value <> 0 Then
                Call MsgBox("作为补录医嘱进行处理时，不允许勾选“紧急医嘱”。", vbInformation + vbOKOnly, Me.Caption)
                chkPriority.SetFocus
                
                Exit Function
            End If
        End If
    End If
    
    strCurProName = vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcName)
    
    '判断检查项目是否已经存在
    If tbcRequest.ItemCount > 1 Then
        For i = 0 To tbcRequest.ItemCount - 1
            If tbcRequest.Item(i).Caption = strCurProName And i <> lngRequestTabIndex Then
                Call MsgBox("检查项目 [" & strCurProName & "] 已在本次申请中存在，不能重复申请。", vbInformation + vbOKOnly, Me.Caption)
                Exit Function
            End If
        Next i
    End If
    
    
    strDataKey = tbcRequest.Item(lngRequestTabIndex).Tag
    
    Set objCurAppData = GetCurRequestPageFormat(lngRequestTabIndex)
    
    '医保项目对码
    strCheckContext = objCurAppData.lngProjectId & ":" & objCurAppData.lngExeRoomId
    strMsg = CheckAdviceInsure(mCurPatientInf.lngInsure, True, mCurPatientInf.lngID, mCurPatientInf.lngFrom, _
                                "", strCheckContext, "当前项目")
                                
    If strMsg <> "" Then
        If gint医保对码 = 1 Then
            lngMsgResult = MsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", vbYesNo, gstrSysName)
            If lngMsgResult = vbNo Then Exit Function
        ElseIf gint医保对码 = 2 Then
            Call MsgBox(strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName)
            Exit Function
        End If
    End If
                                        
    
    If mobjAppDatas.Exists("_" & strDataKey) Then
        Set objOldAppData = mobjAppDatas.Item("_" & strDataKey)
        
        objCurAppData.lngUpdateAdviceId = objOldAppData.lngUpdateAdviceId
        objCurAppData.lngUpdateAppNo = objOldAppData.lngUpdateAppNo
        objCurAppData.blnAllowUpdate = objOldAppData.blnAllowUpdate
        
        '如果已经在mobjAppDatas保存了申请数据，则执行删除后重新保存
        Set mobjAppDatas.Item("_" & strDataKey) = Nothing
        Call mobjAppDatas.Remove("_" & strDataKey)
    End If
    
    Call mobjAppDatas.Add("_" & strDataKey, objCurAppData)
    
    
    Call SetRequestPageState(False)
    
    tbcRequest.Item(lngRequestTabIndex).Caption = strCurProName
    
    mblnIsSaveRequestPage = True
    
    If mCurPatientInf.lngInsure > 0 Then
        strSQL = "Select b.收费项目id From 诊疗项目目录 a, 诊疗收费关系 b Where a.id = b.诊疗项目id  And a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcId))
        
        If rsTemp.RecordCount > 0 Then
            lblInsureInfo.Caption = "医保信息：" & gclsInsure.GetItemInfo(mCurPatientInf.lngInsure, mCurPatientInf.lngID, NVL(rsTemp!收费项目ID), "", 0, "", TProjectCol.pcId & "||" & mCurPatientInf.lngFrom)
        End If
    End If
    
    SaveRequest = True
End Function


Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Select Case Control.ID
        Case TApplicationMenu.amAppSave
            Control.Enabled = mblnPageUpdateState
            
        Case TApplicationMenu.amAppType * 100 + 1 To TApplicationMenu.amAppType * 100 + 99
            If Not tbcPage.Visible Or tbcPage.ItemCount <= 0 Then Exit Sub
            
            '如果对应申请单已经保存，则设置图标
            Control.IconId = IIF(mobjAppDatas.Exists("_" & Control.Category), 3558, 1)
            
        Case TApplicationMenu.amAppDel
            If Not tbcPage.Visible Or tbcPage.ItemCount <= 0 Then
                Control.Enabled = False
                Exit Sub
            End If
            
            Control.Enabled = mobjAppDatas.Exists("_" & CStr(tbcRequest.Selected.Tag)) And mlngUpdateAppNo <= 0
    End Select
Exit Sub
errHandle:
End Sub

Private Sub cbxExeRoom_Click()
On Error GoTo errHandle
    Dim lngProjectRowIndex As Long
    Dim lngExeRoomId As Long
    
    '当mblnIsLoadRequestPage为true时，表示正在加载执行科室设置
    If cbxExeRoom.ListIndex < 0 Or mblnIsLoadRequestPage Then Exit Sub
    
    lngProjectRowIndex = Val(vfgRequestProject.Tag)
    
    If lngProjectRowIndex < 0 And vfgRequestProject.RowSel < 0 Then Exit Sub
    
    If lngProjectRowIndex <> vfgRequestProject.RowSel Then
        Call SelectRequestProject(vfgRequestProject.RowSel)
        lngProjectRowIndex = vfgRequestProject.RowSel
    End If
    
    lngExeRoomId = Val(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom))
    
    If lngExeRoomId <> cbxExeRoom.ItemData(cbxExeRoom.ListIndex) Then
        vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom) = cbxExeRoom.ItemData(cbxExeRoom.ListIndex) & "-" & Replace(cbxExeRoom.Text, Val(cbxExeRoom.Text) & "-", "")
        
        Call SetRequestPageState(True)
    End If
    
    Call ShowArrangeState
    Call ShowStudyProjectToTxt
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SetRequestPageState(ByVal blnModifyState As Boolean)
'设置申请界面的修改状态
    mblnPageUpdateState = blnModifyState
End Sub

Private Sub cbxExeRoom_DropDown()
On Error GoTo errHandle
    Call SendMessage(cbxExeRoom.hwnd, &H160, 200, 0)
errHandle:
End Sub

Private Sub SwitchRequestPageType(ByVal lngRequestPageId As Long, ByVal strRequestPageName As String)
On Error GoTo errHandle
    Dim i As Long
    Dim objControlItem As XtremeSuiteControls.TabControlItem
    Dim objPage As XtremeSuiteControls.TabControlItem
    
    Dim lngFindIndex As Long
    Dim lngHintResult As Long
    Dim lngItemAdviceTag As String
    Dim blnRefresh As Boolean
    Dim strName As String
    Dim strRequestName As String
    Dim blnIsNew As Boolean
    Dim lngPageIndex As Long
    Dim blnIsChangeRequestType As Boolean
    
 
    '去掉菜单后的快捷键
    strName = Mid(strRequestPageName, 1, InStr(strRequestPageName, "(&") - 1)

    If mblnPageUpdateState Then
        strRequestName = Replace(tbcRequest.Tag, Val(tbcRequest.Tag) & "-", "")
        '对已经改变的申请单进行保存提示
        lngHintResult = MsgBox("【" & strRequestName & "】内容已经改变，是否保存？", vbYesNoCancel, Me.Caption)
        
        If lngHintResult = vbYes Then
            '保存申请单
            If Not SaveRequest(Val(tbcRequest.Tag)) Then
                '如果保存失败，则退出页面切换
                Exit Sub
            End If
            
            Call SetRequestPageState(False)
            
        ElseIf lngHintResult = vbNo Then
            Call SetRequestPageState(False)
            
            '清除诊断编辑器中的缓存数据
            If Not mclsDiagEdit Is Nothing Then
                Call mclsDiagEdit.DeleteApplyDiag(Val(tbcRequest.Item(Val(tbcRequest.Tag)).Tag))
            End If
        Else
            Exit Sub
        End If
    End If
    
    lngPageIndex = -1
    lngFindIndex = -1
    
    blnIsNew = False
    blnRefresh = False
    
    If tbcRequest.Selected Is Nothing Then
        blnIsChangeRequestType = True
    Else
        blnIsChangeRequestType = GetRequestId(tbcRequest.Selected.Tag) <> lngRequestPageId
    End If
    
    '计算申请页的索引
    For i = 0 To tbcPage.ItemCount - 1
        If tbcPage.Item(i).Tag = lngRequestPageId Then
            lngPageIndex = i
            Exit For
        End If
    Next i
    
        '配置tab页显示
    If lngPageIndex >= 0 Then
       '切换到申请类型页面
        If Not tbcPage.Selected Is Nothing Then
            If tbcPage.Selected.Index <> lngPageIndex Then
                tbcPage.Item(lngPageIndex).Selected = True
                Call tbcPage_SelectedChanged(tbcPage.Selected)
            End If
        End If
        
        
        For i = 0 To tbcRequest.ItemCount - 1
            If GetRequestId(tbcRequest.Item(i).Tag) = lngRequestPageId Then
                If mobjAppDatas.Exists("_" & CStr(tbcRequest.Item(i).Tag)) = False Then
                    '如果有没有保存的申请，则直接进行切换
                    blnIsNew = False
                    
                    If i <> tbcRequest.Selected.Index Then blnRefresh = True
                    
                    tbcRequest.Item(i).Selected = True
                    lngFindIndex = i
                    
                    Exit For
                End If
            End If
        Next i
    Else
        Set objPage = tbcPage.InsertItem(tbcPage.ItemCount, strName, picTab.hwnd, 0)
        
        objPage.Tag = CStr(lngRequestPageId)
        objPage.Selected = True
        
        Call tbcPage_SelectedChanged(objPage)
    End If
    
    '如果申请页面类型改变，则需要清除诊断
    If blnIsChangeRequestType Then
        '清除诊断编辑器中的缓存数据
        If Not mclsDiagEdit Is Nothing Then
            Call mclsDiagEdit.DeleteApplyDiag(GetRequestId(tbcRequest.Item(lngFindIndex).Tag))
        End If
    End If


    '如果lngFindIndex 大于-1，说明已经加载了该页面
    If lngFindIndex < 0 Then
        '创建新的tab页面
        Set objControlItem = tbcRequest.InsertItem(tbcRequest.ItemCount, "新项目", picTab.hwnd, 0)
        
        '允许重复申请，则tag值为“申请ID_申请时间”
        objControlItem.Tag = CStr(lngRequestPageId) & "_" & Format(Now, "mmddhhmmss")
        objControlItem.Selected = True
        
        blnRefresh = True
    End If
    
    '当tab数量为一或者申请页面类型改变时，由这里载入申请单相关数据
    If tbcRequest.ItemCount = 1 Or blnRefresh Then
        Call tbcRequest_SelectedChanged(tbcRequest.Selected)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub chkPriority_Click()
On Error GoTo errHandle
    If mblnIsLoadRequestPage Then Exit Sub
    
    Call SetRequestPageState(True)
    
'    chkPriority.FontBold = IIF(chkPriority.value <> 0, True, False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdDiagnose_Click()
'打开诊断编码
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim strPro As String
    Dim strCode As String
    Dim i As Long
    
    Call CloseDiagnoseCodeInput
    
    strPro = rtbInputPro(Val(cmdInput.Tag)).Text
    strCode = GetProCode(strPro, "1")
    
    '诊断编码
    Set rsTemp = zlDatabase.ShowILLSelect(Me, "1", mCurPatientInf.lngRoomId, mCurPatientInf.strSex, True, False, strCode)
    
    If rsTemp Is Nothing Then
        Exit Sub
    End If
    
    If rsTemp.RecordCount <= 0 Then
        rtbInputPro(Val(cmdInput.Tag)).Text = ""
    Else
        i = 1
        strPro = ""
        
        While Not rsTemp.EOF
            strPro = strPro & i & "、" & NVL(rsTemp!名称) & "  "
            i = i + 1
            
            Call rsTemp.MoveNext
        Wend
        
        rtbInputPro(Val(cmdInput.Tag)).Text = strPro
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExitCode_Click()
On Error GoTo errHandle
    Call CloseDiagnoseCodeInput
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function GetProCode(ByVal strProContext As String, ByVal strType As String) As String
'strProContext:项目对应的录入内容
'strType:选择器的类型，D:表示疾病编码器，1:表示诊断编码器
'返回项目的对应编码
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim strPro As String
    Dim aryPro() As String
    Dim strCode As String
    
    strCode = ""
    strPro = strProContext
    
    '获取已经选择的疾病编码组合“,J02.901,M03.756,”
    If strPro <> "" Then
        For i = 1 To 9
            strPro = Replace(strPro, i & "、", "<#>")
        Next i
        
        aryPro() = Split(strPro, "<#>")
        
        strPro = ""
        For i = LBound(aryPro) To UBound(aryPro)
            If Trim(aryPro(i)) <> "" Then
                strPro = strPro & Trim(aryPro(i)) & ","
            End If
        Next i
        
        strSQL = "select 编码 from " & IIF(strType = "D", "疾病编码目录", "疾病诊断目录") & " a, " & _
                " (select column_value from Table(Cast(f_Str2List([1]) As zlTools.t_StrList))) b " & _
                " where a.名称=b.column_value and 类别=[2]"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询编码", strPro, strType)
        
        While Not rsTemp.EOF
            strCode = strCode & "," & NVL(rsTemp!编码)
            Call rsTemp.MoveNext
        Wend
        
        If strCode <> "" Then strCode = strCode & ","
    End If
    
    GetProCode = strCode
End Function


Private Sub cmdIcd10_Click()
'打开疾病编码
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim strPro As String
    Dim strCode As String
    Dim i As Long
    
    Call CloseDiagnoseCodeInput
    
    strPro = rtbInputPro(Val(cmdInput.Tag)).Text
    strCode = GetProCode(strPro, "D")
    
    'D-ICD-10疾病编码
    Set rsTemp = zlDatabase.ShowILLSelect(Me, "D", mCurPatientInf.lngRoomId, mCurPatientInf.strSex, True, True, strCode)
    
    If rsTemp Is Nothing Then
        Exit Sub
    End If
    
    If rsTemp.RecordCount <= 0 Then
        rtbInputPro(Val(cmdInput.Tag)).Text = ""
    Else
        i = 1
        strPro = ""
        
        While Not rsTemp.EOF
            strPro = strPro & i & "、" & NVL(rsTemp!名称) & "  "
            i = i + 1
            
            Call rsTemp.MoveNext
        Wend
        
        rtbInputPro(Val(cmdInput.Tag)).Text = strPro
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetRequestId(ByVal strTag As String) As Long
'获取申请ID
    GetRequestId = Val(Split(strTag & "_")(0))
End Function

Private Function OpenDiagnoseEdit(ByRef str诊断Id As String, ByRef str诊断内容 As String) As Boolean
'打开诊断编辑器
    
    Dim lngApplicationPageId As Long
    Dim lngAdviceID As Long
    Dim objAppData As clsApplicationData
    
    OpenDiagnoseEdit = False
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mCurPatientInf.lngFrom = TPatientFrom.pfOutPatient, 1260, 1261))
    End If
    
    lngApplicationPageId = GetRequestId(tbcRequest.Selected.Tag)
    
    '判断当前申请单是否在mobjAppDatas中存在
    If mobjAppDatas.Exists("_" & tbcRequest.Selected.Tag) Then
        Set objAppData = mobjAppDatas("_" & tbcRequest.Selected.Tag)
        
        str诊断Id = objAppData.strDiagnoseId
        lngAdviceID = objAppData.lngUpdateAdviceId
    Else
        lngAdviceID = 0
    End If
    
    '显示诊断编辑器界面
    OpenDiagnoseEdit = mclsDiagEdit.ShowDiagEdit(Me, _
                                    lngApplicationPageId, _
                                    mCurPatientInf.lngID, _
                                    IIF(mCurPatientInf.lngFrom = 2, mCurPatientInf.lngPageId, mCurPatientInf.lngRegId), _
                                    mCurPatientInf.lngFrom, _
                                    mCurPatientInf.lngRoomId, _
                                    str诊断Id, _
                                    str诊断内容, _
                                    1, _
                                    lngAdviceID)
                                    
End Function

Private Sub cmdInput_Click()
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strPro As String
    Dim strElement As String
    Dim rct As TRECT
    Dim pointOffset As TPoint
    Dim str诊断IDs As String
    Dim str诊断内容 As String

    If mblnShowWord Then
        If mfrmPacsApplyWord Is Nothing Then
            Set mfrmPacsApplyWord = New frmPacsApplyWord
        End If
        rtbInputPro(Val(cmdInput.Tag)).SelText = rtbInputPro(Val(cmdInput.Tag)).SelText & mfrmPacsApplyWord.ShowPacsApplyWord(mlngCurDeptId, mCurDoctorInf.lngID, labInputPro(Val(cmdInput.Tag)).Caption, Me)
    Else
        strPro = GetInputProName(rtbInputPro(Val(cmdInput.Tag)).Tag)
        strElement = GetInputProElement(rtbInputPro(Val(cmdInput.Tag)).Tag)
        
        If strElement = M_STR_FIXEDELEMENT_DIAGNOSE Then
            '打开诊断录入器
            If OpenDiagnoseEdit(str诊断IDs, str诊断内容) Then
                rtbInputPro(Val(cmdInput.Tag)).Text = str诊断内容
                rtbInputPro(Val(cmdInput.Tag)).ToolTipText = str诊断IDs
            End If
    
        Else
            strSQL = "select id, 模板标题,模板内容,使用次数 from 病历附项模板 where 病历文件Id=" & GetRequestId(tbcRequest.Selected.Tag) & " and 单据附项='" & strPro & "' order by 使用次数 desc "
            Set rsData = zlDatabase.ShowSelect(Me, strSQL, 0, strPro)
            
            '如果没有被选择，则直接退出
            If rsData Is Nothing Then Exit Sub
            
            rtbInputPro(Val(cmdInput.Tag)).SelText = NVL(rsData!模板内容)
            
            strSQL = "zl_病历附项模板_Update(" & rsData!ID & ",'" & NVL(rsData!模板标题) & "','" & NVL(rsData!模板内容) & "'," & Val(NVL(rsData!使用次数)) + 1 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "更新模板使用次数")
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
 

Private Sub dtpExeTime_Change()
    Dim lngProjectRowIndex As Long
    
    If mblnIsLoadRequestPage Then Exit Sub

'    lngProjectRowIndex = Val(vfgRequestProject.Tag)
'    If lngProjectRowIndex < 0 Then Exit Sub
'
'    If lngProjectRowIndex <> vfgRequestProject.RowSel Then Call SelectRequestProject(vfgRequestProject.RowSel)
    
    Call ShowArrangeState
    Call SetRequestPageState(True)
    
'    Debug.Print "dtpExeTime_Change:" & Time
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandle
    Dim i As Long
    Dim objCon As Object
    
    If Not uspRequestPage.UCEnabled Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        
        Case vbKeyS
            If Shift <> 2 Then Exit Sub
            
            If mblnPageUpdateState And Not tbcRequest.Selected Is Nothing Then
                '转移焦点，使其dtpicker的change事件能立即触发
                If tbcRequest.Visible Then tbcRequest.SetFocus
                
                '保存申请单
                Call SaveRequest(tbcRequest.Selected.Index)
            End If
        
        Case vbKeyLeft
            If Shift <> 2 Then Exit Sub
            If tbcRequest.Selected Is Nothing Then Exit Sub
            
            '切换tab页面
            If tbcRequest.Selected.Index > 0 Then tbcRequest.Item(tbcRequest.Selected.Index - 1).Selected = True
        
        Case vbKeyRight
            If Shift <> 2 Then Exit Sub
            If tbcRequest.Selected Is Nothing Then Exit Sub
            
            '切换tab页面
            If tbcRequest.Selected.Index < tbcRequest.ItemCount - 1 Then tbcRequest.Item(tbcRequest.Selected.Index + 1).Selected = True
            
        Case vbKeyReturn
'            zlCommFun.PressKey vbKeyTab
            If Shift <> 2 Then Exit Sub
            
            'ctrl+enter 执行焦点切换操作
            Select Case UCase(TypeName(Screen.ActiveControl))
                Case "TEXTBOX"
                    If UCase(Screen.ActiveControl.Name) = "TXTFIND" Then
                        chkPriority.SetFocus
                    Else
                        For Each objCon In rtbInputPro
                            If objCon.Index >= Screen.ActiveControl.Index + 1 And objCon.Visible And objCon.TabStop Then
                                objCon.SetFocus
                                Exit Sub
                            End If
                        Next
                        
                        txtFind.SetFocus
                    End If
                    
                Case "VSFLEXGRID"
                    If UCase(Screen.ActiveControl.Name) = "VFGREQUESTPROJECT" Then
                        dtpExeTime.SetFocus
                    Else
                        If rtbInputPro.Count > 1 And picInput.Visible Then
                            rtbInputPro(1).SetFocus
                        Else
                            txtFind.SetFocus
                        End If
                    End If
                Case "DTPICKER"
                    cbxExeRoom.SetFocus
                    
                Case "COMBOBOX"
                    If UCase(Screen.ActiveControl.Name) = "CBXEXEROOM" Then
                        vfgList.SetFocus
                    Else
                        If rtbInputPro.Count > 1 And picInput.Visible Then
                            rtbInputPro(1).SetFocus
                        Else
                            txtFind.SetFocus
                        End If
                    End If
                    
                Case "CHECKBOX"
                    vfgRequestProject.SetFocus
                    
                Case Else
                    If rtbInputPro.Count > 1 And picInput.Visible Then
                        rtbInputPro(1).SetFocus
                    Else
                        txtFind.SetFocus
                    End If
            End Select
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim i As Long

    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFaceElement
    Call FreeRequestInputControl
    
    Call ReadRequestPatientInf

    Call InitRequestTab
    Call InitCommandBars
    
    labRequestDoctValue.Caption = ""
    If mlngUpdateAppNo > 0 Then
        '配置医嘱更新页面
        Call CfgUpdatePage(mlngUpdateAppNo)
        
        '载入申请单共有的基本信息
        Call LoadRequestBaseInf
        
        Me.Caption = "检查更新"
    Else
        '查找申请类型菜单
        If mlngProjectId <> 0 Then Set cbrControl = GetDefaultPage(mlngProjectId)
        If cbrControl Is Nothing Then
            Set cbrControl = cbrMain.FindControl(, TApplicationMenu.amAppType * 100 + 1, False, True)
        End If
        
        If cbrControl Is Nothing Then
            uspRequestPage.UCEnabled = False
        Else
            Call cbrMain_Execute(cbrControl)
            
            '载入申请单共有的基本信息
            Call LoadRequestBaseInf
        End If
        
        Me.Caption = "检查申请"
    End If
    
    If mCurPatientInf.lngID <= 0 Then
        uspRequestPage.UCEnabled = False
    End If
End Sub

Private Function GetDefaultPage(ByVal lngProjectId As Long) As Object
'功能：获取省页面
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngTmp As Long, i As Long
    Dim cbrControl As CommandBarControl
    
    strSQL = "Select Max(病历文件id) As PageId From 病历单据应用 Where 诊疗项目id =[1] And 应用场合=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngProjectId, mCurPatientInf.lngFrom)
    
    lngTmp = Val(rsData!PageId & "")
 
    If lngTmp <> 0 Then
        For i = 1 To mlngRequestPageCount
            Set cbrControl = cbrMain.FindControl(, TApplicationMenu.amAppType * 100 + i, False, True)
            If Not cbrControl Is Nothing Then
                If Val(cbrControl.Category) = lngTmp Then
                    Exit For
                Else
                    Set cbrControl = Nothing
                End If
            End If
        Next
    End If
    Set GetDefaultPage = cbrControl
End Function

Private Function GetProjectName(ByVal lngProId As Long) As String
'获取项目名称
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetProjectName = ""
    
    strSQL = "Select 名称 From 诊疗项目目录 Where Id =[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询项目名称", lngProId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetProjectName = NVL(rsData!名称)
End Function


Private Sub CfgUpdatePage(ByVal lngUpdateAppNo As Long)
'配置更新页面
    Dim i As Integer
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strTbcTag As String
    Dim strInsertTags As String
    
    Dim objAppData As clsApplicationData
    Dim curTab As XtremeSuiteControls.TabControlItem
    Dim objCurPage As XtremeSuiteControls.TabControlItem
    
    Dim blnOldUpdateState As Boolean
    
    blnOldUpdateState = mblnPageUpdateState
    
    If mobjAppDatas.Count <= 0 Then Exit Sub
    
    mblnPageUpdateState = False
    
    strInsertTags = ""
    For i = 0 To mobjAppDatas.Count - 1
        Set objAppData = mobjAppDatas(mobjAppDatas.Keys(i))
        If objAppData Is Nothing Then Exit Sub
        
        labRequestTimeValue.Caption = Format(IIF(Trim(objAppData.strRequestTime) = "", Now, objAppData.strRequestTime), "yyyy-mm-dd hh:mm")
        Set curTab = tbcRequest.InsertItem(0, GetProjectName(objAppData.lngProjectId), picTab.hwnd, 0)
        
        curTab.Tag = CStr(objAppData.lngApplicationPageId) & "_" & i
        
        If InStr(strInsertTags, objAppData.lngApplicationPageId) <= 0 Then
            strInsertTags = strInsertTags & "," & objAppData.lngApplicationPageId
            
            Set objCurPage = tbcPage.InsertItem(0, objAppData.strApplicationPageName, picTab.hwnd, 0)
            objCurPage.Tag = CStr(objAppData.lngApplicationPageId)
        End If
        
        If strTbcTag = "" Then strTbcTag = curTab.Index & "-" & objAppData.strApplicationPageName
    Next
        
    If tbcPage.ItemCount > 0 Then
        tbcPage.Item(0).Selected = True
        Call tbcPage_SelectedChanged(tbcPage.Selected)
        
        tbcRequest.Item(0).Selected = True
        Call tbcRequest_SelectedChanged(tbcRequest.Selected)
    End If
    
    mblnPageUpdateState = blnOldUpdateState
End Sub


Private Sub InitFaceElement()
'初始化申请单界面相关的元素
    Dim dtNow As Date
    
    dtNow = zlDatabase.Currentdate
    
    vfgRequestProject.Tag = -1
    vfgRequestProject.Rows = 0
    vfgRequestProject.Cols = 8
    
    vfgRequestProject.ColHidden(TProjectCol.pcRoomType) = True
    vfgRequestProject.ColHidden(TProjectCol.pcMethod) = True
    vfgRequestProject.ColHidden(TProjectCol.pcExeRoom) = True
    
    labRequestTimeValue.Caption = Format(dtNow, "yyyy-mm-dd hh:mm")
    dtpExeTime.value = dtNow
    
    tbcRequest.Tag = -1     '保存当前选中的页面，以便在tab页切换的事件中，能够判断上一次的选择页面，默认不选择任何页面
    
    picInput.Tag = 0        '为0时，表示不会显示申请附项的录入
End Sub


Private Sub InitRequestTab()
    Dim objfont As New StdFont
    
    objfont.Name = "宋体"
    objfont.Size = 12
    
    With tbcPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPageFlat ' xtpTabAppearancePropertyPageSelected ' xtpTabAppearancePropertyPage2003 'xtpTabAppearancePropertyPageFlat
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .PaintManager.HotTracking = True
        .PaintManager.ButtonMargin.Top = 11
        .PaintManager.ButtonMargin.Bottom = 11
        
        Set .PaintManager.Font = objfont
    End With
    
    With tbcRequest
        .PaintManager.Appearance = xtpTabAppearancePropertyPageFlat ' xtpTabAppearancePropertyPageSelected ' xtpTabAppearancePropertyPage2003 'xtpTabAppearancePropertyPageFlat
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = False
        .PaintManager.HotTracking = True
        .PaintManager.ButtonMargin.Top = 4
        .PaintManager.ButtonMargin.Bottom = 2
        
        Set .PaintManager.Font = objfont
    End With
End Sub


Private Sub LoadRequestPageKind(mnuParent As Object)
'载入检查申请单类别
    Dim objParentMenu As CommandBarControl
    Dim objMenuBar As CommandBarControl
    Dim objControl As CommandBarControl
    
    Dim blnExist As Boolean
    Dim lngCount As Long
    Dim lngID As Long
    Dim arrName() As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim objControlItem As TabControlItem
    Dim i As Long
    Dim j As Long
    
    Set objParentMenu = mnuParent
    
    strSQL = "Select a.Id, a.名称 From 病历文件列表 A, 病历单据应用 B Where a.Id = b.病历文件id And b.应用场合 = [1] And a.种类 = 7 And a.子类 = '检查' Group By a.Id, a.名称, a.编号 Order By a.编号"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查申请单", mCurPatientInf.lngFrom)

    
    If rsData.RecordCount <= 0 Then
        mnuParent.Enabled = False
        Exit Sub
    End If
    mlngRequestPageCount = rsData.RecordCount
    i = 1
    lngID = 1
    
    With objParentMenu.CommandBar
        While Not rsData.EOF
            blnExist = False
            lngCount = 1
            '如果名称中包含‘-’,则‘-’前面部分为申请分类，后面部分为申请单名称
            If InStr(NVL(rsData!名称), "-") = 0 Then
                Set objMenuBar = .Controls.Add(xtpControlButton, TApplicationMenu.amAppType * 100 + lngID, _
                                                NVL(rsData!名称) & "申请(&" & Chr(IIF(48 + i > 57, 56 + i, 48 + i)) & ")", "")
                
                objMenuBar.Category = NVL(rsData!ID)
                objMenuBar.IconId = 1
                lngID = lngID + 1
                i = i + 1
            Else
                arrName = Split(NVL(rsData!名称), "-")
                '检查申请分类是否已存在
                For j = 1 To .Controls.Count
                    If NVL(arrName(0)) = Mid(.Controls(j).Caption, 1, InStr(.Controls(j).Caption, "(&") - 1) Then
                        Set objMenuBar = .Controls(j)
                        lngCount = objMenuBar.CommandBar.Controls.Count + 1
                        blnExist = True
                        Exit For
                    End If
                Next
                
                '申请分类不存在时，增加分类
                '为了保证申请单的ID保持一致（xtpControlButton, TApplicationMenu.amAppType * 100），将分类的ID设为xtpControlButtonPopup, TApplicationMenu.amAppType * 1000
                If Not blnExist Then
                    Set objMenuBar = .Controls.Add(xtpControlButtonPopup, TApplicationMenu.amAppType * 1000 + i, NVL(arrName(0)) & "(&" & Chr(IIF(48 + i > 57, 56 + i, 48 + i)) & ")", "")
                    objMenuBar.IconId = 1
                    i = i + 1
                End If
                
                '显示申请
                Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppType * 100 + lngID, NVL(arrName(1)) & "申请(&" & Chr(IIF(48 + lngCount > 57, 56 + lngCount, 48 + lngCount)) & ")", "")
                lngID = lngID + 1
                '如果新增了分类，记录分类名称，用于标题显示
                If Not blnExist Then
                    objControl.Parent.Title = NVL(arrName(0))
                End If
                objControl.Category = NVL(rsData!ID)
                objControl.IconId = 1
            End If
            
            
            Call rsData.MoveNext
        Wend
    End With

End Sub



Private Sub LoadRequestAffixInputCfg(ByVal lngRequestPageId As Long, Optional ByVal strAffixs As String = "")
'载入申请单附项录入
    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strProName As String
    Dim lngProOrder As Long
    Dim rsRequestContext As ADODB.Recordset
    Dim strInputValue As String
    Dim arrAffix() As String
    Dim strElement As String
    Dim i As Long
    
    labTitle.Caption = tbcPage.Selected.Caption
    
    strSQL = "select a.要素id,a.文件ID,a.项目,a.必填,a.排列,要素Id,b.中文名 as 要素名, a.内容,a.只读,b.中文名  " & _
            " from 病历单据附项 a, 诊治所见项目 b  " & _
            " where a.要素id=b.id(+) and a.文件Id=[1] order by 排列 "
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询录入项目", lngRequestPageId)
        
    If rsData.RecordCount <= 0 Then
        picInput.Visible = False
        picInput.Tag = 0
        
        Exit Sub
    End If

    arrAffix() = Split(strAffixs, "|")
    mstrRequestAffixConfig = ""
    
    While Not rsData.EOF
        '加载录入项目到界面中
        strProName = NVL(rsData!项目)
        lngProOrder = Val(NVL(rsData!排列))
        
        strInputValue = ""
        
        If mstrRequestAffixConfig <> "" Then mstrRequestAffixConfig = mstrRequestAffixConfig & "|"
        mstrRequestAffixConfig = mstrRequestAffixConfig & strProName & ":" & Val(NVL(rsData!必填)) & "," & lngProOrder & "," & NVL(rsData!要素ID) & ","
        
        
        If strAffixs <> "" Then
            For i = 0 To UBound(arrAffix)
                If Split(arrAffix(i), ":")(0) = strProName Then
                    strSQL = arrAffix(i)
                    strInputValue = Replace(strSQL, Mid(strSQL, 1, InStr(strSQL, ":")), "")
                    Exit For
                End If
            Next i
        End If
        
        If Trim(strInputValue) = "" Then
            strInputValue = GetAppendItemValue(mCurPatientInf.lngID, mCurPatientInf.lngPageId, mCurPatientInf.strRegNo, mintBabyID, NVL(rsData!中文名), strProName, NVL(rsData!要素ID, 0))
            
            If strInputValue = "" Then
                strInputValue = NVL(rsData!内容)
            End If
        End If
 
        '获取要素名称
        strElement = NVL(rsData!要素名)
        
        Call SetInputControl(strProName, strInputValue, lngProOrder, Val(NVL(rsData!必填)), _
                            strElement, IIF(rsData.RecordCount <= 4, True, False), IIF(Val(NVL(rsData!只读)) = 1, True, False))

        Call rsData.MoveNext
    Wend
    
    picInput.Height = mobjLastControl.Top + mobjLastControl.Height + 120
    picInput.Visible = True
    picInput.Tag = 1
    
    If rtbInputPro.Count > 1 Then rtbInputPro(1).TabIndex = 0
End Sub



Private Sub LoadRequestPageProject(ByVal lngRequestPageId As Long, _
    Optional ByVal lngProjectId As Long = 0, _
    Optional ByVal lngExeRoomId As Long = 0, _
    Optional ByVal strExeRoomName As String = "", _
    Optional ByVal lngExeType As Long = 0, _
    Optional ByVal strPartMethod As String = "")
    
'载入申请单对应的检查项目
'需要对项目范围进行过滤，如病人为住院病人时，只过滤属于住院的检查项目
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim strUseDeptIds As String
    Dim blnShowNormal As Boolean
    Dim lngSelIndex As Long
    
    picPart.Visible = True
    
    strUseDeptIds = ",||" & mCurPatientInf.lngRoomId & "||,||" & mlngCurDeptId & "||,"
    
    strSQL = "select distinct a.Id,编码,a.名称, a.执行科室,a.执行标记 from 诊疗项目目录 a, 病历单据应用 b " & _
            " where a.Id=b.诊疗项目Id " & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) " & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " & _
                    " And A.服务对象 IN(" & IIF(mCurPatientInf.lngFrom = 3, "1,2,4", mCurPatientInf.lngFrom) & ",3) " & _
                    " and b.应用场合=[3] " & _
                    " And Nvl(A.单独应用,0)=1" & _
                    " And Nvl(A.适用性别,0) IN (" & IIF(mCurPatientInf.strSex Like "*男*", "1,0)", "2,0)") & _
                    " And Nvl(A.执行频率,0) IN(0,1) " & _
                    " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And Instr([2],',||'||科室ID||'||,')>0)" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                    " And 病历文件Id=[1] " & _
            " order by 编码"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询诊疗单据对应检查项目", lngRequestPageId, strUseDeptIds, mCurPatientInf.lngFrom)
        
    vfgRequestProject.Rows = 0
    vfgRequestProject.Tag = -1
    txtCurStudyProject.Text = ""
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    vfgRequestProject.Rows = rsData.RecordCount
  
    lngSelIndex = -1
    i = 0
    blnShowNormal = False
    
    While Not rsData.EOF
        '新开检查时，如果诊疗项目ID=mlngProjectId,则自动定位到此项目；修改检查时，如果诊疗项目ID=lngProjectId,则自动定位到此项目；
        If Val(NVL(rsData!ID)) = lngProjectId Or Val(NVL(rsData!ID)) = mlngProjectId Then
            lngSelIndex = i
        End If
        
        Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcId) = img16.ListImages("o0").Picture
        
        
        vfgRequestProject.ColWidth(TProjectCol.pcId) = 260
        
        vfgRequestProject.Cell(flexcpData, i, TProjectCol.pcId) = NVL(rsData!ID)
        vfgRequestProject.Cell(flexcpData, i, TProjectCol.pcName) = NVL(rsData!名称)
        
        vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcName) = NVL(rsData!编码) & "-" & NVL(rsData!名称)
        vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcRoomType) = Val(NVL(rsData!执行科室))
        
        If Val(NVL(rsData!执行标记)) <> 0 Then
            '增加常规，床旁，术中选择列
            blnShowNormal = True
            
            vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcNormalCol) = "常"
            Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcNormalCol) = img16.ListImages(IIF(lngExeType = 0, "o1", "o0")).Picture
            
            vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcBedCol) = "床"
            Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcBedCol) = img16.ListImages(IIF(lngExeType = 1, "o1", "o0")).Picture
            
            vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcOperCol) = "术"
            Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcOperCol) = img16.ListImages(IIF(lngExeType = 2, "o1", "o0")).Picture
        End If
    
        i = i + 1
        
        Call rsData.MoveNext
    Wend
    
    If blnShowNormal Then
        vfgRequestProject.ColWidth(TProjectCol.pcName) = 3400
        
        vfgRequestProject.ColWidth(TProjectCol.pcNormalCol) = 550
        vfgRequestProject.ColWidth(TProjectCol.pcBedCol) = 550
        vfgRequestProject.ColWidth(TProjectCol.pcOperCol) = 550
        
        vfgRequestProject.ColHidden(TProjectCol.pcNormalCol) = False
        vfgRequestProject.ColHidden(TProjectCol.pcBedCol) = False
        vfgRequestProject.ColHidden(TProjectCol.pcOperCol) = False
    Else
        vfgRequestProject.ColHidden(TProjectCol.pcNormalCol) = True
        vfgRequestProject.ColHidden(TProjectCol.pcBedCol) = True
        vfgRequestProject.ColHidden(TProjectCol.pcOperCol) = True
    End If
    
    If lngSelIndex >= 0 Then
        Call SelectRequestProject(lngSelIndex, True)
        
        vfgRequestProject.Cell(flexcpText, lngSelIndex, TProjectCol.pcMethod) = strPartMethod
        vfgRequestProject.Cell(flexcpText, lngSelIndex, TProjectCol.pcExeRoom) = lngExeRoomId & "-" & strExeRoomName
        
        Call vfgRequestProject.ShowCell(lngSelIndex, 0)
        Call vfgRequestProject.Select(lngSelIndex, 0)
    Else
        If vfgRequestProject.Rows > 0 Then
            Call vfgRequestProject.ShowCell(0, 0)
            Call vfgRequestProject.Select(0, 0)
        End If
    End If
    
    If vfgRequestProject.RowSel > 0 Then vfgRequestProject.TopRow = vfgRequestProject.RowSel
End Sub


Public Function Get操作员部门ID(ByVal int服务对象 As Integer) As Long
'功能：取操作员所属服务对指定对象的部门，缺省部门优先
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        strSQL = "Select Distinct B.部门ID,Nvl(B.缺省,0) as 缺省,C.服务对象 From 部门人员 B,部门性质说明 C" & _
            " Where B.人员ID = [1] And B.部门ID=C.部门ID" & _
            " Order by 缺省 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", mCurDoctorInf.lngID)
    End If
    rsTmp.Filter = "服务对象 = 3 or 服务对象 = " & int服务对象
    
    If Not rsTmp.EOF Then
        Get操作员部门ID = rsTmp!部门ID
    Else
        Get操作员部门ID = mCurDoctorInf.lng部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadRequestProjectExeRoom(ByVal lngProjectId As Long, ByVal lngProjectRoomType As Long, _
    ByVal lngCurExeDeptId As Long)
'[1]：病人来源，[2]：病人科室Id，[3]：诊疗项目Id，[4]：当前执行科室Id即默认执行科室ID
'[5]：操作员科室ID， [6]：病人Id，[7]：病人主页ID，[8]：开嘱科室ID
    Dim strSQL As String
    Dim lng操作员科室ID As Long
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim lngDefExeRoomId As Long
    Dim bytDay As Byte
    Dim bln上班安排 As Boolean
    
'    strDefExeRoomId = Get诊疗执行科室ID(mCurPatientInf.lngID, mCurPatientInf.lngPageID, "D", , lngProjectRoomType, mCurPatientInf.lngRoomId, mlngCurDeptId, 0)
    
    
    cbxExeRoom.Clear

    lngDefExeRoomId = 0
    
    Select Case lngProjectRoomType
        Case 0 '0-无执行的叮嘱
            Exit Sub
        Case 5  '5-院外执行
            '使用API快速加入,不然可能有点慢
            AddComboItem cbxExeRoom.hwnd, CB_ADDSTRING, 0, "-"
            SetComboData cbxExeRoom.hwnd, CB_SETITEMDATA, i - 1, -1
            
            Call Cbo.SetIndex(cbxExeRoom.hwnd, 0)

            Exit Sub
        Case 1 '1-病人所在科室
            strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([2],[4]) Order by 编码"
        Case 2 '2-病人所在病区
            If mCurPatientInf.lngFrom = 1 Then
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([2],[4]) Order by 编码"
            Else
                strSQL = _
                    " Select A.ID,A.编码,A.简码,A.名称" & _
                    " From 部门表 A,病案主页 B" & _
                    " Where A.ID=B.当前病区ID And B.病人ID=[6] And B.主页ID=[7]" & _
                    " Union " & _
                    " Select ID,编码,简码,名称 From 部门表 Where ID=[4]" & _
                    " Order by 编码"
            End If
        Case 3 '3-操作员所在科室
            strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([5],[4]) Order by 编码"
            lng操作员科室ID = Get操作员部门ID(mCurPatientInf.lngFrom)
        Case 4 '4-指定科室
'            strSql = _
'                " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
'                " From 部门表 A,诊疗执行科室 B,部门性质说明 C" & _
'                " Where A.ID=B.执行科室ID And B.诊疗项目ID=[3] And A.ID=C.部门ID" & _
'                " And C.服务对象 IN([1],3) And (B.病人来源 is NULL Or B.病人来源=[1])" & _
'                " And (B.开单科室ID is NULL Or B.开单科室ID=[2])" & _
'                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
'                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
'                " Union Select ID,编码,简码,名称 From 部门表 Where ID=[4]" & _
'                " Order by 编码"

                If lngProjectId = 0 Then
                    If mCurPatientInf.lngFrom = 1 Then  '1表示门诊
                        lngDefExeRoomId = mCurPatientInf.lngRoomId
                        strSQL = "select ID,编码,简码,名称 From 部门表 where id=[10]"
                        
                    ElseIf mCurPatientInf.lngFrom = 2 Then  '2表示住院
                        lngDefExeRoomId = GetPatiUnitID(mCurPatientInf.lngID, mCurPatientInf.lngPageId)
                        strSQL = "select ID,编码,简码,名称 From 部门表 where id=[10]"
                    End If
                    
                Else
                    'pacs检查项目都属于临嘱，因此需要对临嘱判断上班时间
                    bln上班安排 = Check上班安排(False)
                    
                    If Not bln上班安排 Then
                        strSQL = "Select Distinct c.Id, c.编码,c.简码,c.名称, Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                            " From 诊疗执行科室 A,部门性质说明 B,部门表 C" & _
                            " Where A.执行科室ID=B.部门ID And A.诊疗项目ID=[3]" & _
                            " And B.服务对象 IN([1],3) And (A.病人来源 is NULL Or A.病人来源=[1])" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And A.执行科室ID=C.ID And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " Order by 排序" '默认科室优先
                    Else
                        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
                        strSQL = _
                            " Select Distinct d.Id, d.编码,d.简码,d.名称, Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                            " From 诊疗执行科室 A,部门安排 B,部门性质说明 C,部门表 D" & _
                            " Where A.执行科室ID+0=B.部门ID And B.星期=[9]" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                            " And A.执行科室ID=C.部门ID And C.服务对象 IN([1],3) And (A.病人来源 is NULL Or A.病人来源=[1])" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And A.执行科室ID=D.ID And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & _
                            " And (D.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or D.撤档时间 is NULL)" & _
                            " And A.诊疗项目ID=[3]" & _
                            " Order by 排序"
                    End If
                End If
        Case 6 '6-开单人所在科室
            strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([8],[4]) Order by 编码"
    End Select
    
        

    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询执行科室", mCurPatientInf.lngFrom, mCurPatientInf.lngRoomId, _
                                        lngProjectId, 0, mCurDoctorInf.lng部门ID, _
                                        mCurPatientInf.lngID, mCurPatientInf.lngPageId, mlngCurDeptId, bytDay, lngDefExeRoomId)
                                        
    If rsData.RecordCount > 0 Then lngDefExeRoomId = rsData!ID
    
    '4表示指定执行科室
    If lngProjectRoomType = 4 Then
        If Not rsData.EOF Then
            lngDefExeRoomId = rsData!执行科室ID
            rsData.Filter = "开单科室ID=" & mCurPatientInf.lngRoomId
            
            If rsData.EOF Then rsData.Filter = "执行科室ID=" & mCurPatientInf.lngRoomId
            If rsData.EOF And mCurPatientInf.lngFrom = 2 Then rsData.Filter = "执行科室ID=" & GetPatiUnitID(mCurPatientInf.lngID, mCurPatientInf.lngPageId)
            
            If Not rsData.EOF Then lngDefExeRoomId = rsData!执行科室ID
            
        ElseIf gbln指定医嘱在其他科室执行 Then
            If mCurPatientInf.lngFrom = 1 Then         '1表示门诊
                lngDefExeRoomId = mCurPatientInf.lngRoomId
                strSQL = "select ID,编码,简码,名称 From 部门表 where id=[10]"
                
            ElseIf mCurPatientInf.lngFrom = 2 Then     '2表示住院
                lngDefExeRoomId = GetPatiUnitID(mCurPatientInf.lngID, mCurPatientInf.lngPageId)
                strSQL = "select ID,编码,简码,名称 From 部门表 where id=[10]"
                
            End If
            
            '重新获取科室信息
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询执行科室", mCurPatientInf.lngFrom, mCurPatientInf.lngRoomId, _
                                                lngProjectId, 0, mCurDoctorInf.lng部门ID, _
                                                mCurPatientInf.lngID, mCurPatientInf.lngPageId, mlngCurDeptId, bytDay, lngDefExeRoomId)
            If rsData.RecordCount > 0 Then lngDefExeRoomId = rsData!ID
            
        End If
    End If
    
    rsData.Filter = ""

    For i = 1 To rsData.RecordCount
        '使用API快速加入,不然可能有点慢
        AddComboItem cbxExeRoom.hwnd, CB_ADDSTRING, 0, rsData!编码 & "-" & rsData!名称
        SetComboData cbxExeRoom.hwnd, CB_SETITEMDATA, i - 1, CLng(rsData!ID)
        
        '设置默认的检查科室
        If lngDefExeRoomId = rsData!ID Then
            Call Cbo.SetIndex(cbxExeRoom.hwnd, i - 1)
        End If

        '设置当前已经选择的检查科室
        If lngCurExeDeptId = rsData!ID Then
            Call Cbo.SetIndex(cbxExeRoom.hwnd, i - 1)
        End If

        rsData.MoveNext
    Next
    
    '只有执行科室为1时，才设置默认科室，否则需要由用户手动选择
    If cbxExeRoom.ListCount = 1 And cbxExeRoom.ListIndex < 0 Then cbxExeRoom.ListIndex = 0
End Sub

Private Sub ClearRequestPatientInf()
'清除申请单相关的病人部分信息
    mCurPatientInf.lngRoomId = 0
    mCurPatientInf.strName = ""
    mCurPatientInf.strSex = ""
    mCurPatientInf.strRegNo = ""
    mCurPatientInf.strAge = ""
    mCurPatientInf.strInNO = ""
    mCurPatientInf.strOutNo = ""
    mCurPatientInf.strConditionTag = ""
End Sub


Private Sub RestoreRequestPageCfg()
'恢复和医嘱相关的申请单界面配置
    Dim objControl As Object
    
    Set mobjLastControl = Nothing
    
    '清除申请单基本信息
    txtFind.Text = ""
    txtCurStudyProject.Text = ""
    
'    cmdInput.Visible = False
    
    labPrice.Caption = "￥：---"
    
    chkPriority.value = 0
    chkPriority.FontBold = False
    
    dtpExeTime.value = zlDatabase.Currentdate
    
    '隐藏已经加载的录入组件
    For Each objControl In labInputPro
        objControl.Visible = False
    Next
    
    For Each objControl In labMustPro
        objControl.Visible = False
    Next
    
    For Each objControl In rtbInputPro
        objControl.Visible = False
    Next
    
    For Each objControl In shpBackLine
        objControl.Visible = False
    Next
    
    For Each objControl In shpInputPro
        objControl.Visible = False
    Next
    
    '删除部位选择
    vfgList.Rows = 1
    
    '删除检查项目
    vfgRequestProject.Rows = 0

    
    '删除科室选择
    cbxExeRoom.Clear
    
    picInput.Visible = False
    picInput.Tag = 0
End Sub


Private Sub SetInputControl(ByVal strProName As String, ByVal strDefaultContext As String, ByVal lngProOrder As Long, _
    Optional ByVal blnIsMustInput As Boolean, Optional ByVal strElementName As String = "", Optional blnIsBigDistance As Boolean = False, _
    Optional ByVal blnIsReadOnly As Boolean = False)
    
'设置申请单界面的录入项目
'strProName：录入项名称
'strDefaultContext：默认值
'lngProOrder：录入顺序
'blnIsMustInput：是否必录项
'strElementName：要素名称
'blnIsBigDistance：录入项目之间是否使用较大的间隔距离
'blnIsReadOnly：项目是否允许进行编辑

    Dim objLab As Label
    Dim objMust As Label
    Dim objRchEdit As TextBox
    Dim objShp As Shape
    
    
    If InStr(strProName, "-") > 0 Then
        '载入分割线
        If Not HasControl("shpInputPro", lngProOrder) Then
            Load shpInputPro(lngProOrder)
        End If
        
        Set objShp = shpInputPro(lngProOrder)
        objShp.Visible = True
        
        If mobjLastControl Is Nothing Then
            objShp.Left = 120
            objShp.Top = 120
            
        Else
            objShp.Left = 120
            objShp.Top = mobjLastControl.Top + mobjLastControl.Height + 360
            
        End If
        
        Set mobjLastControl = objShp
    Else
        '载入录入标签
        If Not HasControl("labInputPro", lngProOrder) Then
            Load labInputPro(lngProOrder)
        End If
        
        Set objLab = labInputPro(lngProOrder)
        
        objLab.Caption = strProName
        objLab.Visible = True
        
        If mobjLastControl Is Nothing Then
            objLab.Left = 120
            objLab.Top = IIF(blnIsBigDistance, 360, 180)
            
        Else
            objLab.Left = 120
            objLab.Top = mobjLastControl.Top + mobjLastControl.Height + 120 + IIF(blnIsBigDistance, 360, 180)
            
        End If
        
        '如果为必填项目，则载入必填的标记标签

        If Not HasControl("labMustPro", lngProOrder) Then
            Load labMustPro(lngProOrder)
        End If
        
        Set objMust = labMustPro(lngProOrder)
        objMust.Left = objLab.Left + objLab.Width + 60   ' Fix(objLab.Width / 2) + objLab.Left - 60
        objMust.Top = objLab.Top ' objLab.Top + objLab.Height - 10
        objMust.Visible = blnIsMustInput
        
        
        
        '载入录入框
        If Not HasControl("rtbInputPro", lngProOrder) Then
            Load rtbInputPro(lngProOrder)
        End If
        
        Set objRchEdit = rtbInputPro(lngProOrder)
        
        objRchEdit.Left = objLab.Left + objLab.Width + objMust.Width + 120
        objRchEdit.Top = objLab.Top - 30 ' objLab.Height + 60
        objRchEdit.Width = picInput.ScaleWidth - objRchEdit.Left - 60 - cmdInput.Width  '减去cmdInput的宽度，是考虑到可能存在使用附项模板输入的情况
        objRchEdit.Text = strDefaultContext
        objRchEdit.Tag = strProName & IIF(strElementName <> "", ">[" & strElementName & "]", "")
        objRchEdit.Visible = True
        objRchEdit.TabStop = True
        objRchEdit.TabIndex = rtbInputPro.Count - 1
        
        objRchEdit.TabStop = Not (blnIsReadOnly Or strElementName = M_STR_FIXEDELEMENT_DIAGNOSE)
        
        If blnIsReadOnly Or strElementName = M_STR_FIXEDELEMENT_DIAGNOSE Then
            '只读状态的设置
            objRchEdit.Locked = True
            objRchEdit.BackColor = &HC0FFFF
        Else
            objRchEdit.Locked = False
            objRchEdit.BackColor = rtbInputPro(0).BackColor
        End If
        
        Call SetInputControlHight(objRchEdit)
        
        '载入背景线
        If Not HasControl("shpBackLine", lngProOrder) Then
            Load shpBackLine(lngProOrder)
        End If
        
        Set objShp = shpBackLine(lngProOrder)
        
        objShp.Left = objRchEdit.Left
        objShp.Top = objRchEdit.Top + objRchEdit.Height
        objShp.Width = objRchEdit.Width
        objShp.Visible = Not (blnIsReadOnly Or strElementName = M_STR_FIXEDELEMENT_DIAGNOSE)
        
        Set mobjLastControl = objRchEdit
    End If
End Sub


Private Function GetInputProName(ByVal strTag As String) As String
'获取录入项目的项目名称
    Dim lngProIndex As Long
    
    GetInputProName = strTag
    
    lngProIndex = InStrRev(strTag, ">[")
    If lngProIndex <= 0 Then Exit Function
    
    GetInputProName = Mid(strTag, 1, lngProIndex - 1)
End Function

Private Function GetInputProElement(ByVal strTag As String) As String
'获取录入项目的要素名称,strTag示例: 病人诊断>[[临床诊断]]
    Dim lngProIndex As Long
    Dim strResult As String
    
    GetInputProElement = ""
    
    lngProIndex = InStr(strTag, ">[")
    If lngProIndex <= 0 Then Exit Function
    
    '去掉字符">["
    strResult = Mid(strTag, lngProIndex + 2, 255)
    
    '去掉字符"]"
    strResult = Mid(strResult, 1, Len(strResult) - 1)
    
    GetInputProElement = strResult
End Function


Private Function HasControl(ByVal strControlName As String, ByVal lngIndex As Long) As Boolean
On Error GoTo errHandle
    Dim objControl As Object
    Select Case strControlName
        Case "labInputPro"
            For Each objControl In labInputPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
        Case "labMustPro"
            For Each objControl In labMustPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
        Case "shpInputPro"
            For Each objControl In shpInputPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
        Case "rtbInputPro"
            For Each objControl In rtbInputPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
        Case "shpBackLine"
            For Each objControl In shpBackLine
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
    End Select
    
    HasControl = False
Exit Function
errHandle:
    HasControl = False
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo errHandle
    Dim strRequestName As String
    Dim lngHintResult As Long
    
    If Not mblnPageUpdateState Then Exit Sub
    
    strRequestName = Replace(tbcRequest.Tag, Val(tbcRequest.Tag) & "-", "")
    '对已经改变的申请单进行保存提示
    lngHintResult = MsgBox("【" & strRequestName & "】内容已经改变，是否保存？", vbYesNoCancel, Me.Caption)
    
    If lngHintResult = vbYes Then
        '保存申请单
        If Not SaveRequest(Val(tbcRequest.Tag)) Then
            '如果保存失败，则退出页面切换
            Cancel = 1
            Exit Sub
        End If
        
        Call SetRequestPageState(False)
        
    ElseIf lngHintResult = vbNo Then
        Call SetRequestPageState(False)
        
    Else
        Cancel = 1
        Exit Sub
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
    picMenu.Left = 0
    picMenu.Top = 0
    picMenu.Width = ScaleWidth
    
    tbcPage.Left = 3520 '2580
    tbcPage.Top = 60
    tbcPage.Width = ScaleWidth - 3540 ' 2600
    
    tbcRequest.Left = 0
    tbcRequest.Top = picMenu.Height - 100
    tbcRequest.Width = ScaleWidth
    
    picBack.Left = 0
    picBack.Top = tbcRequest.Top + tbcRequest.Height
    picBack.Width = ScaleWidth
    picBack.Height = ScaleHeight - picMenu.Height - stbThis.Height - tbcRequest.Height + 100

Exit Sub
errHandle:
End Sub


Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸

        
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '设置控件显示风格
        .EnableCustomization False                             '是否允许自定义设置
        .ActiveMenuBar.Visible = False
        
        
        Set .Icons = zlCommFun.GetPubIcons                     '设置关联的图标控件
    End With

    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)

    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppSave, "保存", "保存检查申请 (Ctrl+S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppDel, "删除", "删除检查申请"): cbrControl.IconId = 4114: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButtonPopup, TApplicationMenu.amAppType, "新增申请▼", "选择检查申请单的类型"): cbrControl.IconId = 807
    
    '载入申请单类型菜单
    Call LoadRequestPageKind(cbrControl)
    If mlngUpdateAppNo > 0 Or mCurPatientInf.lngID <= 0 Then cbrControl.Enabled = False

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppEdit, "退出", "退出检查申请"): cbrControl.IconId = 2613
    cbrControl.BeginGroup = True
        
        
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
End Sub


Private Sub Form_Terminate()
    Set mobjAppDatas = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
'    Dim i As Long
    
    Call FreeRequestInputControl
    
'    '不能在这里释放数据，因为窗口关闭后，还需要访问此属性返回申请单内容
'    For i = mobjAppDatas.Count - 1 To 0 Step -1
'        Call mobjAppDatas.Remove(mobjAppDatas.Keys(i))
'        Set mobjAppDatas.Item(i) = Nothing
'    Next i
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mfrmPacsApplyWord = Nothing
Exit Sub
errHandle:
End Sub

Private Sub SetNullMethod(ByVal lngProjectRowIndex As Long)
'设置检查项目的方法为空方法，当为空方法时，保存申请则不需要确认选择部位和检查方法
    Dim lngRow As Long
    
    lngRow = lngProjectRowIndex
    If lngRow < 0 Then
        lngRow = vfgRequestProject.RowSel
    End If
    
    If lngRow >= 0 Then vfgRequestProject.Cell(flexcpText, lngRow, TProjectCol.pcMethod) = " "
End Sub

Private Sub LoadRequestPagePart(ByVal lngProjectId As Long, ByVal strSex As String, ByVal strExtData As String)
'功能：初始化检查部位表格格式及数据
'参数：mstrExtData=包含检查部位的信息,为空时表示新输入检查组合项目
On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long, i As Integer
    Dim str类型 As String, str名称 As String
    Dim arrData As Variant, strNoneRegion As String
    Dim blnNone As Boolean
    Dim Y As Long, str方法 As String
    Dim lngRowHideCount As Long
    Dim lngProjectRowIndex As Long
    
    vfgList.Rows = 1
    
    With vfgList
        .WordWrap = True
        .FixedRows = 1
        .FixedCols = 0
        .Rows = .FixedRows + 1
        .Cols = 4
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        
        If str类型 = "病理" Then
            .TextMatrix(0, 0) = "标本名称"
            .TextMatrix(0, 1) = "标本名称"
            .TextMatrix(0, 2) = "材料类别"
        Else
            .TextMatrix(0, 0) = "检查部位"
            .TextMatrix(0, 1) = "检查部位"
            .TextMatrix(0, 2) = "检查方法"
        End If
        
        .TextMatrix(0, 3) = "备注"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1: .ColWidth(i) = 1600
        Next
        
    End With
    
    '读取检查项目基本信息
    strSQL = "Select 名称,操作类型,执行标记 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngProjectId)
    
    If rsTmp.RecordCount <= 0 Then
        lngProjectRowIndex = Val(vfgRequestProject.RowSel)
        
        Call SetNullMethod(lngProjectRowIndex)
        
        Exit Sub
    End If
        
    str类型 = rsTmp!操作类型
    str名称 = rsTmp!名称
        
    '过滤检查部位信息
    Set rsTmp = mrsRequestPart
    rsTmp.Filter = "诊疗项目Id=" & lngProjectId & " and 类型='" & str类型 & "'"
    rsTmp.Sort = "分组"
    
    blnNone = rsTmp.EOF
    
    If rsTmp.RecordCount <= 0 Then
        lngProjectRowIndex = Val(vfgRequestProject.RowSel)
        
        Call SetNullMethod(lngProjectRowIndex)
        
        Exit Sub
    End If
    
    If rsTmp.RecordCount > 0 Then
        lngRowHideCount = 50
        vfgList.Rows = vfgList.Rows + lngRowHideCount
        
        '设置固定数量的隐藏行，便于将选中的部位移动到首要位置
        For i = 1 To lngRowHideCount
            vfgList.RowHidden(i) = True
        Next i
    End If


    With vfgList
        '显示基准的部位及默认方法
        If blnNone Then
            .Editable = flexEDNone
            .TabStop = False
        Else
            .Editable = flexEDKbdMouse
        End If
        
        If str类型 = "病理" Then
            .TextMatrix(0, 0) = "标本名称"
            .TextMatrix(0, 1) = "标本名称"
            .TextMatrix(0, 2) = "材料类别"
        Else
            .TextMatrix(0, 0) = "检查部位"
            .TextMatrix(0, 1) = "检查部位"
            .TextMatrix(0, 2) = "检查方法"
        End If
        
        .TextMatrix(0, 3) = "备注"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        
        Do While Not rsTmp.EOF
            If .TextMatrix(.Rows - 1, 1) <> rsTmp!部位 Then
            
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                End If
            
                .TextMatrix(.Rows - 1, 0) = zlCommFun.GetNeedName("" & rsTmp!分组)
                .TextMatrix(.Rows - 1, 1) = rsTmp!部位
                Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                .Cell(flexcpData, .Rows - 1, 2) = CStr(NVL(rsTmp!检查方法))  '供方法选择器使用
                .TextMatrix(.Rows - 1, 3) = NVL(rsTmp!备注)
            End If
            
            If NVL(rsTmp!默认, 0) = 1 Then '以"方法名1,方法名2,..."的方式显示部位检查方法
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & NVL(rsTmp!方法)
                If Left(.TextMatrix(.Rows - 1, 2), 1) = "," Then
                    .TextMatrix(.Rows - 1, 2) = Mid(.TextMatrix(.Rows - 1, 2), 2)
                End If
            End If
            
            rsTmp.MoveNext
        Loop
        
        '修改时套入已有的内容
        '  如果为空，也可能是以前的单部位检查项目，这时要以新增的方式重新选择部位
        '  或者对于以前的单部位项目，强行传入以前的部位(没有方法)，现还可能有同名部位
        If strExtData <> "" Then
            arrData = Split(Split(strExtData, vbTab)(0), "|")
            
            For i = 0 To UBound(arrData)
                lngIdx = .FindRow(CStr(Split(arrData(i), ";")(0)), 1, 1, , True)
                str方法 = ""
                
                If lngIdx <> -1 Then
                
                    '检查方法有没有不存在的
                    For Y = 0 To UBound(Split(Split(arrData(i), ";")(1), ","))
                        If InStr(.Cell(flexcpData, lngIdx, 2), CStr(Split(Split(arrData(i), ";")(1), ",")(Y))) = 0 Then
                            strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0) & "(" & Split(Split(arrData(i), ";")(1), ",")(Y) & ")"
                        Else
                            str方法 = str方法 & "," & Split(Split(arrData(i), ";")(1), ",")(Y)
                        End If
                    Next
                    
                    '该部位的方法:可能以前的数据只有部位没有方法
                    If UBound(Split(arrData(i), ";")) >= 1 Then
                        .TextMatrix(lngIdx, 2) = Mid(str方法, 2)
                    Else
                        .TextMatrix(lngIdx, 2) = ""
                    End If
                    
                    .Cell(flexcpData, lngIdx, 1) = 1 '表明该部位已选择
                    .Cell(flexcpFontBold, lngIdx, 1, lngIdx, 3) = True
                    .Cell(flexcpBackColor, lngIdx, 1, lngIdx, 3) = &HC0E0FF
                    
                    Set .Cell(flexcpPicture, lngIdx, 1) = img16.ListImages("c1").Picture
                    
                    '将被选择行移动到行首
                    Call MoveToFirst(lngIdx)
                Else
                    '该部位设置已不存在
                    strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0)
                End If
            Next
        End If
        
        .Row = 1: .Col = 1
        .ShowCell .Row, .Col
        
        '确定表格尺寸
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) < 500 Then .ColWidth(0) = 500
        If .ColWidth(0) > 850 Then .ColWidth(0) = 850
        If .ColWidth(1) < 800 Then .ColWidth(1) = 800
        If .ColWidth(1) > 1600 Then .ColWidth(1) = 1600
        If .ColWidth(2) < 2500 Then .ColWidth(2) = 2500
        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        If .ColWidth(3) < 800 Then .ColWidth(3) = 800
        If .ColWidth(3) > 2000 Then .ColWidth(3) = 2000
        
        lngIdx = 0
        For i = 0 To .Cols - 1
            lngIdx = lngIdx + .ColWidth(i) + 15
        Next
        
    End With
    
    '删除隐藏的数据行
    For i = vfgList.Rows - 1 To 1 Step -1
        If vfgList.RowHidden(i) Then Call vfgList.RemoveItem(i)
    Next i
        
    '已不存在的部位提示
    If strNoneRegion <> "" Then
        If str类型 = "病理" Then
            MsgBox "以下病理标本在项目设置中已不存在：" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        Else
            MsgBox "以下检查部位或方法在项目设置中已不存在或不适用：" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
            
            '如果存在不适用的方法，则允许手动编辑方法
'            vfgList.Editable = flexEDKbdMouse
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub MoveToFirst(ByVal lngCurRow As Long)
'移动到第一行
    Dim strRowData As Variant
    Dim strRowText As Variant
    Dim varRowPic  As Variant
    Dim blnFontBold As Boolean
    Dim blnRowHidden As Boolean
    Dim lngBakColor As OLE_COLOR
    
    Dim i As Long
    Dim lngUpRow As Long
    
    If vfgList.Rows = 2 Then Exit Sub

    '查询首次没有被勾选的行
    For i = 1 To vfgList.Rows - 1
        If vfgList.Cell(flexcpData, i, 1) <> 1 Then
            lngUpRow = i
            Exit For
        End If
    Next i
    
    blnRowHidden = vfgList.RowHidden(lngUpRow)
    
    For i = 0 To vfgList.Cols - 1
        
        strRowText = vfgList.TextMatrix(lngUpRow, i)
        strRowData = vfgList.Cell(flexcpData, lngUpRow, i)
        blnFontBold = vfgList.Cell(flexcpFontBold, lngUpRow, i)
        lngBakColor = vfgList.Cell(flexcpBackColor, lngUpRow, i)
        
        Set varRowPic = vfgList.Cell(flexcpPicture, lngUpRow, i)
        
        vfgList.TextMatrix(lngUpRow, i) = vfgList.TextMatrix(lngCurRow, i)
        vfgList.Cell(flexcpData, lngUpRow, i) = vfgList.Cell(flexcpData, lngCurRow, i)
        vfgList.Cell(flexcpPicture, lngUpRow, i) = vfgList.Cell(flexcpPicture, lngCurRow, i)
        vfgList.Cell(flexcpFontBold, lngUpRow, i) = vfgList.Cell(flexcpFontBold, lngCurRow, i)
        vfgList.Cell(flexcpBackColor, lngUpRow, i) = vfgList.Cell(flexcpBackColor, lngCurRow, i)
        
        vfgList.TextMatrix(lngCurRow, i) = strRowText
        vfgList.Cell(flexcpData, lngCurRow, i) = strRowData
        vfgList.Cell(flexcpPicture, lngCurRow, i) = varRowPic
        vfgList.Cell(flexcpFontBold, lngCurRow, i) = blnFontBold
        vfgList.Cell(flexcpBackColor, lngCurRow, i) = lngBakColor
    Next i
    
    vfgList.RowHidden(lngUpRow) = vfgList.RowHidden(lngCurRow)
    vfgList.RowHidden(lngCurRow) = blnRowHidden
End Sub


Private Sub picBack_Resize()
On Error Resume Next
    Dim lngFreeHeight As Long
    
    picPart.Height = 3900
    
    Call AdjustRequestFace
    
    uspRequestPage.Left = 0
    uspRequestPage.Top = 0
    uspRequestPage.Width = picBack.ScaleWidth
    uspRequestPage.Height = picBack.ScaleHeight
    
    '重新调整项目选择高度，使其能够自动适应页面大小
    If picRequestInf.Top + picRequestInf.Height < picBack.ScaleHeight - 240 Then
        lngFreeHeight = picBack.Height - picRequestInf.Top - picRequestInf.Height - 240
        picPart.Height = 3900 + lngFreeHeight
    Else
        picPart.Height = 3900
    End If

    '计算申请页面的其他组件的大小和位置
    Call AdjustRequestFace

End Sub


Private Sub CloseDiagnoseCodeInput()
'    Call ClipCursor(0)
    
    picBack.Enabled = True
    
    cmdInput.SetFocus
End Sub


Private Sub picInput_Click()
On Error GoTo errHandle
    Dim curPos As PointAPI
    Dim objShp As Shape
    
 
    '获取当前指针对应在控件中的位置
    Call GetCursorPos(curPos)
    
    Call ScreenToClient(picInput.hwnd, curPos)
    
    curPos.X = picInput.ScaleX(curPos.X, vbPixels, vbTwips)
    curPos.Y = picInput.ScaleY(curPos.Y, vbPixels, vbTwips)
     
    For Each objShp In shpBackLine
        If curPos.X > objShp.Left And curPos.X < objShp.Left + objShp.Width And curPos.Y < objShp.Top And objShp.Visible Then
            rtbInputPro(objShp.Index).SetFocus
            Exit Sub
        End If
    Next
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub picPart_Resize()
On Error Resume Next
    '调整项目选择框的大小，使其能够自动适应
    vfgRequestProject.Height = picPart.Height - vfgRequestProject.Top - 120
    vfgList.Height = picPart.Height - vfgList.Top - 120
err.Clear
End Sub


Private Sub AutoSizeInputFace()
'自动适应录入窗口界面
    Dim objLastControl As Object
    Dim objCurControl As Object
    
    Dim objLab As Object
    Dim objMust As Object
    Dim objText As Object
    Dim objShp As Object
    
    Dim lngControlCount As Long
    Dim blnInputObj As Boolean
    Dim blnIsBigDistance As Boolean
    Dim i As Long
    Dim blnExeCase As Boolean
    
    blnIsBigDistance = IIF(rtbInputPro.Count <= 4, True, False)
    Set objLastControl = Nothing
    
    lngControlCount = (shpBackLine.Count - 1) + (shpInputPro.Count - 1)
    
    '标签，必录标志，文本框，下划线
    For i = 1 To lngControlCount

        blnExeCase = Not HasControl("shpInputPro", i)
        
        If HasControl("shpInputPro", i) Then
            blnExeCase = Not shpInputPro(i).Visible
        End If
        
        If blnExeCase Then
            If HasControl("labInputPro", i) Then
            
                Set objCurControl = labInputPro(i)
    
                If objCurControl.Visible Then
                
                    If Not (objLastControl Is Nothing) Then
                        objCurControl.Top = objLastControl.Top + objLastControl.Height + 120 + IIF(blnIsBigDistance, 360, 180)
                    End If
                    
                    Set objLab = objCurControl
                    
                    If HasControl("labMustPro", i) Then
                        Set objMust = labMustPro(i)
                        objMust.Top = objLab.Top
                    End If
                    
                    If HasControl("rtbInputPro", i) Then
                        Set objText = rtbInputPro(i)
                        objText.Top = objLab.Top - 30
                    End If
                                
                    Set objShp = shpBackLine(i)
                    objShp.Top = objText.Top + objText.Height
                    
                    Set objLastControl = objShp
                End If
            End If
            
        Else
            If HasControl("shpInputPro", i) Then
                Set objCurControl = shpInputPro(i)
                
                If objCurControl.Visible Then
                    If Not (objLastControl Is Nothing) Then
                        objCurControl.Top = objLastControl.Top + objLastControl.Height + 360
                    End If
                    
                    Set objLastControl = objCurControl
                End If
            End If
        End If
        
    Next i
    
    If cmdInput.Visible Then
        cmdInput.Top = rtbInputPro(Val(cmdInput.Tag)).Top - 20
        cmdInput.Height = rtbInputPro(Val(cmdInput.Tag)).Height + 40
    End If
    
    picInput.Height = objLastControl.Top + objLastControl.Height + 120
    
End Sub

Private Function SetInputControlHight(objText As TextBox) As Boolean
'设置输入框控件的高度
    Dim lngTxtHeight As Long
    Dim lngBase As Long
    Dim lngOldHeight As Long
    Dim strTestText As String
    
    lngBase = Me.TextHeight("啊")
    SetInputControlHight = False
    
    With objText
        strTestText = .Text & "啊啊啊啊啊啊啊啊啊啊啊啊啊啊啊"
        
        lngTxtHeight = Me.TextHeight(strTestText) * IIF(Me.TextWidth(strTestText) > .Width, Fix(Me.TextWidth(strTestText) / .Width) + 1, 1)
        
        If lngTxtHeight = .Height Then Exit Function
        
        If lngTxtHeight < 375 Then
            If .Height = 375 Then Exit Function
            
            .Height = 375
        ElseIf lngTxtHeight > 1635 Then
            If .Height = 1635 Then Exit Function
            
            .Height = 1635
        Else
            If lngTxtHeight + Fix(lngTxtHeight / lngBase) * 60 > 1635 Then
                .Height = 1635
            Else
                If lngTxtHeight + Fix(lngTxtHeight / lngBase) * 60 = .Height Then Exit Function
                .Height = lngTxtHeight + Fix(lngTxtHeight / lngBase) * 60
            End If
        End If
    End With
    
    SetInputControlHight = True
End Function


Private Sub AdjustRtbInputControlHeight(ByVal Index As Integer)
'根据录入框文本内容，调整申请附项录入控件高度
    Dim blnChangeHeight As Boolean
    Dim objCurText As TextBox
    
    Set objCurText = rtbInputPro(Index)
    blnChangeHeight = SetInputControlHight(objCurText)
        
    If blnChangeHeight Then
        Call AutoSizeInputFace
        Call picBack_Resize
        
        Call uspRequestPage.CalcScroll
    End If
End Sub


Private Sub rtbInputPro_Change(Index As Integer)
On Error GoTo errHandle
    
    If rtbInputPro(Index).Visible Then
        Call SetRequestPageState(True)
        
        Call AdjustRtbInputControlHeight(Index)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Resume
End Sub


Private Function GetTabAdviceId(ByVal strSelTabTag As String) As Long
'获取tab也对应的医嘱ID
    Dim lngTemp As Long
    
    lngTemp = Val(strSelTabTag)
    
    GetTabAdviceId = Val(Replace(strSelTabTag, lngTemp & IIF(InStr(strSelTabTag, "-") > 0, "-", ""), ""))
End Function

Private Function IsEqualTabPage(ByVal tabItem As XtremeSuiteControls.ITabControlItem) As Boolean
    IsEqualTabPage = IIF(tabItem.Index & "-" & tabItem.Caption = tbcRequest.Tag, True, False)
End Function

Private Sub rtbInputPro_GotFocus(Index As Integer)
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strPro As String
    Dim strElement As String
    
    strElement = GetInputProElement(rtbInputPro(Index).Tag)
    
    cmdInput.Left = rtbInputPro(Index).Left + rtbInputPro(Index).Width
    cmdInput.Top = rtbInputPro(Index).Top - 20 '+ rtbInputPro(Index).Height - cmdInput.Height
    cmdInput.Height = rtbInputPro(Index).Height + 40
    
    cmdInput.Tag = Index '存储录入框的对象索引
        
    If Not rtbInputPro(Index).Visible Or (rtbInputPro(Index).Locked And strElement <> M_STR_FIXEDELEMENT_DIAGNOSE) Then
'        cmdInput.Visible = False
        mblnShowWord = True
        Exit Sub
    End If
    
    strPro = GetInputProName(rtbInputPro(Index).Tag)
    
    strSQL = "select id from 病历附项模板 where 病历文件id=[1] and 单据附项=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询附项模板", GetRequestId(tbcRequest.Selected.Tag), strPro)
    
    If rsData.RecordCount > 0 Or strElement = M_STR_FIXEDELEMENT_DIAGNOSE Then
        '如果存在附项内容模板，则显示模板弹出按钮
'        rtbInputPro(Index).Width = shpBackLine(Index).Width - cmdInput.Width
        
        mblnShowWord = False
'        cmdInput.Visible = True
    Else
'        cmdInput.Visible = False
        mblnShowWord = True
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub rtbInputPro_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHandle
'    If KeyAscii <> 13 Then Exit Sub
'    If Not cmdInput.Visible Then Exit Sub
'
'    Call cmdInput_Click
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    Dim lngRequestId As Long
    Dim i As Long
    Dim lngStart As Long
    Dim objCurItem As XtremeSuiteControls.ITabControlItem
    
    If Item.Tag = "" Then Exit Sub  '如果没有tag值，则可能正在执行删除操作
    lngRequestId = Val(Item.Tag)
    
    lngStart = -1
    For i = 0 To tbcRequest.ItemCount - 1
        If GetRequestId(tbcRequest.Item(i).Tag) <> lngRequestId Then
            tbcRequest.Item(i).Visible = False
        Else
            If lngStart = -1 Then lngStart = i
            If tbcRequest.Item(i).Caption = "新项目" Then lngStart = i
            
            tbcRequest.Item(i).Visible = True
        End If
    Next i
    
    If tbcRequest.ItemCount <= 0 Or lngStart < 0 Then Exit Sub
    
    If mblnIsRestoreTab Then    '恢复页面操作
        Exit Sub
    End If

    tbcRequest.Item(lngStart).Selected = True
    
    If lngStart <> tbcRequest.Selected.Index Then    '恢复页面操作
        Exit Sub
    End If
    
    tbcPage.Tag = Item.Index
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tbcRequest_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    Dim lngHintResult As Long
    Dim strRequestName As String
    Dim objCurAppData As clsApplicationData
    
    mblnIsRestoreTab = False
    
    If mblnIsRestoreTab Then Exit Sub
    
    '如果切换后的页面相同或者item无对应数据，则退出处理
    If IsEqualTabPage(Item) Or Item.Tag = "" Then Exit Sub
    
    mblnIsLoadRequestPage = True
    
    If mblnPageUpdateState Then
        strRequestName = Replace(tbcRequest.Tag, Val(tbcRequest.Tag) & "-", "")
        lngHintResult = MsgBox("【" & strRequestName & "】内容已经改变，是否保存？", vbYesNoCancel, Me.Caption)
        
        If lngHintResult = vbYes Then
            '保存申请单
            If Not SaveRequest(Val(tbcRequest.Tag)) Then
                mblnIsRestoreTab = True
                
                '如果保存失败，则恢复页面切换
                tbcPage.Item(Val(tbcPage.Tag)).Selected = True
                tbcRequest.Item(Val(tbcRequest.Tag)).Selected = True
                Exit Sub
            End If
            
            Call SetRequestPageState(False)
            
        ElseIf lngHintResult = vbNo Then
            Call SetRequestPageState(False)
            
            '清除诊断编辑器中的缓存数据
            If Not mclsDiagEdit Is Nothing Then
                Call mclsDiagEdit.DeleteApplyDiag(Val(tbcRequest.Item(Val(tbcRequest.Tag)).Tag))
            End If
        Else
            tbcRequest.Item(Val(tbcRequest.Tag)).Selected = True
            Exit Sub
            
        End If
        
    End If
    
    tbcRequest.Tag = Item.Index & "-" & Item.Caption
    
    Call LoadRequestPage(GetRequestId(Item.Tag), Item.Tag)

    
    mblnIsLoadRequestPage = False
Exit Sub
errHandle:
    mblnIsLoadRequestPage = False
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadRequestPage(ByVal lngApplicationPageId As Long, ByVal strApplicationPageTag As String)
'载入申请页面
On Error GoTo errHandle
    Dim objCurAppData As clsApplicationData
    Dim strDataKey As String
    Dim lngFreeHeight As Long
    
    Set objCurAppData = New clsApplicationData '如果没有设置对应属性，属性则使用默认值
    
    strDataKey = strApplicationPageTag
    If mobjAppDatas.Exists("_" & strDataKey) Then
        Set objCurAppData = Nothing
        Set objCurAppData = mobjAppDatas.Item("_" & strDataKey)
    End If
    
    '先恢复申请页面配置
    Call RestoreRequestPageCfg
    
    rtbInputPro(0).BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)

    
    '载入检查申请对应的所有检查项目及部位到临时数据集
    Call LoadRequestPartDataSet(lngApplicationPageId, mCurPatientInf.strSex)
    
    '载入申请录入配置
    Call LoadRequestAffixInputCfg(lngApplicationPageId, objCurAppData.strRequestAffix)
    
    '载入申请项目配置
    Call LoadRequestPageProject(lngApplicationPageId, objCurAppData.lngProjectId, _
                                objCurAppData.lngExeRoomId, objCurAppData.strExeRoomName, _
                                objCurAppData.lngExeType, objCurAppData.strPartMethod)
    
    '如果默认选择了一个项目，则读取默认项目对应的检查部位方法的组合
    If vfgRequestProject.RowSel < 0 Then
        Call LoadRequestPagePart(0, "", "")
    Else
        Call LoadRequestPagePart(vfgRequestProject.Cell(flexcpData, vfgRequestProject.RowSel, TProjectCol.pcId), _
                                mCurPatientInf.strSex, _
                                vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcMethod))
    End If
    
    '更新医嘱的执行时间
    If objCurAppData.strStartExeTime <> "" Then
        dtpExeTime.value = objCurAppData.strStartExeTime
    End If
    
    chkPriority.value = IIF(objCurAppData.blnIsPriority, 1, 0)
'    chkPriority.FontBold = IIF(objCurAppData.blnIsPriority, True, False)
    

    '调整申请页界面布局
    picPart.Height = 3900
    
    Call AdjustRequestFace
    
    '根据页面加载后的情况，从新计算picPart的高度，使其自动适应
    If picRequestInf.Top + picRequestInf.Height < picBack.Height - 240 Then
        lngFreeHeight = picBack.Height - picRequestInf.Top - picRequestInf.Height - 240
        picPart.Height = 3900 + lngFreeHeight
    Else
        picPart.Height = 3900
    End If

    '调整申请页中其他组件的大小和位置
    Call AdjustRequestFace
    
    uspRequestPage.UCScrollState = False
    Call uspRequestPage.CalcScroll
        
    Call ShowArrangeState
    Call ShowStudyProjectToTxt
    
    labRequestDoctValue.Caption = IIF(objCurAppData.strRequestDoctor <> "", objCurAppData.strRequestDoctor, UserInfo.姓名)
     
    '如果医嘱不允许修改，则控件可编辑属性设置为enable
    picAuditing.Visible = Not objCurAppData.blnAllowUpdate
    
    txtFind.Enabled = objCurAppData.blnAllowUpdate
    txtFind.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    
    dtpExeTime.Enabled = objCurAppData.blnAllowUpdate
    cbxExeRoom.Enabled = objCurAppData.blnAllowUpdate
    
    
    chkPriority.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    chkPriority.Enabled = objCurAppData.blnAllowUpdate

    vfgRequestProject.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgRequestProject.BackColorBkg = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgRequestProject.Editable = IIF(objCurAppData.blnAllowUpdate, flexEDKbdMouse, flexEDNone)
    
    vfgList.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgList.BackColorBkg = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgList.Editable = IIF(objCurAppData.blnAllowUpdate, flexEDKbdMouse, flexEDNone)
    
    picBaseInf.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picInput.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picPart.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picRequestInf.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    
    uspRequestPage.UCBackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picInput.Enabled = objCurAppData.blnAllowUpdate
'    picPart.Enabled = objCurAppData.blnAllowUpdate
    
    tmrFocus.Enabled = True
    
    mblnIsLoadRequestPage = False
Exit Sub
errHandle:
    mblnIsLoadRequestPage = False
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadRequestPartDataSet(ByVal lngRequestId As Long, ByVal strSex As String)
'载入检查申请单的所有检查部位信息
    Dim strSQL As String

    strSQL = "select distinct a.Id,a.类型,a.默认, c.诊疗项目ID, b.适用性别, b.分组, a.部位, a.方法, b.方法 as 检查方法, b.编码,b.备注 " & _
            " from 诊疗项目部位 a, 诊疗检查部位 b, 病历单据应用 c " & _
            " where a.类型=b.类型 and a.部位=b.名称 and a.项目Id = c.诊疗项目Id " & _
                    " And (b.适用性别 = [2] or Nvl(b.适用性别,0)=0) and c.病历文件Id=[1] order by 分组,编码"

    Set mrsRequestPart = zlDatabase.OpenSQLRecord(strSQL, "查询申请单检查部位", lngRequestId, _
                                                    IIF(strSex = "男", 1, IIF(strSex = "女", 2, 0)))
End Sub



Private Sub AdjustRequestFace()
    picBaseInf.Left = IIF(uspRequestPage.Width > picBaseInf.Width, Fix((uspRequestPage.Width - picBaseInf.Width) / 2), 0)
    picBaseInf.Top = 0
    
    picInput.Left = picBaseInf.Left
    picInput.Top = picBaseInf.Top + picBaseInf.Height
    picInput.Width = picBaseInf.Width
    
    picPart.Left = picBaseInf.Left
    picPart.Top = picBaseInf.Top + picBaseInf.Height + IIF(Val(picInput.Tag) = 1, picInput.Height, 0)
    picPart.Width = picBaseInf.Width
    
    picRequestInf.Left = picBaseInf.Left
    picRequestInf.Top = picPart.Top + picPart.Height
    picRequestInf.Width = picBaseInf.Width
End Sub

Private Sub FreeRequestInputControl()
    Dim objFree As Object
    
    '删除申请附项中的录入配置......
    For Each objFree In labInputPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then
                Unload objFree
            End If
        End If
    Next
    
    For Each objFree In labMustPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then
                Unload objFree
            End If
        End If
    Next
    
    For Each objFree In rtbInputPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    
    For Each objFree In shpInputPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    
    For Each objFree In shpBackLine
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    Set mobjLastControl = Nothing
End Sub

Private Sub Timer1_Timer()
On Error GoTo errHandle
    If mlngUpdateAppNo = 0 Then
        labRequestTimeValue.Caption = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm")
    End If
Exit Sub
errHandle:
End Sub

Private Sub tmrFocus_Timer()
On Error GoTo errHandle
    tmrFocus.Enabled = False
    If rtbInputPro.Count > 1 Then rtbInputPro(1).SetFocus
Exit Sub
errHandle:
End Sub

Private Sub txtFind_Change()
On Error GoTo errHandle
    Call FindRequestProject(txtFind.Text)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtFind_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii = 13 And Trim(txtFind.Text) <> "" Then Call FindRequestProject(txtFind.Text, True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub FindRequestProject(ByVal strFind As String, Optional ByVal blnIsEnter As Boolean = False)
'查找检查申请项目
    Dim i As Long
    Dim lngFindIndex As Long
    Dim strPY As String

    lngFindIndex = -1


    '从被选择的数据后面重新开始查找
    For i = IIF(vfgRequestProject.Tag > -1, Val(vfgRequestProject.Tag), 0) To vfgRequestProject.Rows - 1
        strPY = GetPYCode(vfgRequestProject.Cell(flexcpText, i, 1))

        If UCase(strPY) Like "*" & UCase(strFind) & "*" Then
            If lngFindIndex <= -1 And Val(vfgRequestProject.Tag) <> i Then lngFindIndex = i
        End If
    Next i

    '如果没有找到，则从数据开始位置重新查找
    If lngFindIndex <= -1 Then
        For i = 0 To Val(vfgRequestProject.Tag)
            strPY = GetPYCode(vfgRequestProject.Cell(flexcpText, i, 1))
            If UCase(strPY) Like "*" & UCase(strFind) & "*" Then
                If lngFindIndex <= -1 Then lngFindIndex = i
            End If
        Next i
    End If

    If lngFindIndex >= 0 Then
        Call vfgRequestProject.Select(lngFindIndex, 0)
        Call vfgRequestProject.ShowCell(lngFindIndex, 0)
        
        Call SelectRequestProject(lngFindIndex)
        Call ShowStudyProjectToTxt
    End If
End Sub


Private Function GetPYCode(ByVal strChinese As String) As String
    Dim i As Long
    
    GetPYCode = ""
    
    For i = 1 To Len(strChinese)
        GetPYCode = GetPYCode & GetWordChar1(Mid(strChinese, i, 1))
    Next i

End Function


Private Function GetWordChar1(ByVal strWord As String) As String
'获得汉字的拼音简码
On Error Resume Next
    If Asc(strWord) < 0 Then
        If Asc(Left(strWord, 1)) < Asc("啊") Then
            GetWordChar1 = "0":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("啊") And Asc(Left(strWord, 1)) < Asc("芭") Then
            GetWordChar1 = "A":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("芭") And Asc(Left(strWord, 1)) < Asc("擦") Then
            GetWordChar1 = "B":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("擦") And Asc(Left(strWord, 1)) < Asc("搭") Then
            GetWordChar1 = "C":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("搭") And Asc(Left(strWord, 1)) < Asc("蛾") Then
            GetWordChar1 = "D":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("蛾") And Asc(Left(strWord, 1)) < Asc("发") Then
            GetWordChar1 = "E":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("发") And Asc(Left(strWord, 1)) < Asc("噶") Then
            GetWordChar1 = "F":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("噶") And Asc(Left(strWord, 1)) < Asc("哈") Then
            GetWordChar1 = "G":    Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("哈") And Asc(Left(strWord, 1)) < Asc("击") Then
            GetWordChar1 = "H":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("击") And Asc(Left(strWord, 1)) < Asc("喀") Then
            GetWordChar1 = "J":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("喀") And Asc(Left(strWord, 1)) < Asc("垃") Then
            GetWordChar1 = "K":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("垃") And Asc(Left(strWord, 1)) < Asc("妈") Then
            GetWordChar1 = "L":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("妈") And Asc(Left(strWord, 1)) < Asc("拿") Then
            GetWordChar1 = "M":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("拿") And Asc(Left(strWord, 1)) < Asc("哦") Then
            GetWordChar1 = "N":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("哦") And Asc(Left(strWord, 1)) < Asc("啪") Then
            GetWordChar1 = "O":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("啪") And Asc(Left(strWord, 1)) < Asc("期") Then
            GetWordChar1 = "P":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("期") And Asc(Left(strWord, 1)) < Asc("然") Then
            GetWordChar1 = "Q":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("然") And Asc(Left(strWord, 1)) < Asc("撒") Then
            GetWordChar1 = "R":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("撒") And Asc(Left(strWord, 1)) < Asc("塌") Then
            GetWordChar1 = "S":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("塌") And Asc(Left(strWord, 1)) < Asc("挖") Then
            GetWordChar1 = "T":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("挖") And Asc(Left(strWord, 1)) < Asc("昔") Then
            GetWordChar1 = "W":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("昔") And Asc(Left(strWord, 1)) < Asc("压") Then
            GetWordChar1 = "X":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("压") And Asc(Left(strWord, 1)) < Asc("匝") Then
            GetWordChar1 = "Y":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("匝") Then
            GetWordChar1 = "Z":        Exit Function
        End If
    Else
        If UCase(strWord) <= "Z" And UCase(strWord) >= "A" Then
            GetWordChar1 = UCase(Left(strWord, 1))
        Else
            GetWordChar1 = strWord
        End If
    End If
End Function

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strMethod As String, i As Long, j As Long
    Dim arrMethod As Variant, arrSub As Variant
    Dim k As Long
    Dim blnDo As Boolean
    
    If vfgList.Editable = flexEDNone Then Exit Sub
    
    strMethod = vfgList.Cell(flexcpData, Row, Col)
    If strMethod = "" Then
        MsgBox "该检查部位没有设置可供选择的检查方法。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMethod
        .Rows = 0
        
        arrMethod = Split(Replace(strMethod, vbTab, ";" & vbTab), ";")
        
        For i = 0 To UBound(arrMethod)
            arrSub = Split(arrMethod(i), ",")
            
            For j = 0 To UBound(arrSub)
                .Rows = .Rows + 1
                If j = 0 Then
                    If InStr(1, arrMethod(i), vbTab) > 0 Then
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 2 '表明是共选项
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 3) '第一位是造影剂标志
                        If InStr("," & vfgList.TextMatrix(vfgList.Row, 2) & ",", "," & Mid(arrSub(j), 3) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c1").Picture
                            
                            .Cell(flexcpFontBold, .Rows - 1, 0) = True
                            .Cell(flexcpData, .Rows - 1, 0) = 1
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c0").Picture
                            
                            .Cell(flexcpFontBold, .Rows - 1, 0) = False
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbWhite
                        End If
                    Else
                        '排斥项
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 1 '表明是排斥项
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 2) '第一位是造影剂标志
                        If InStr("," & vfgList.TextMatrix(vfgList.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o1").Picture
                           
                            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 1) = True
                            .Cell(flexcpData, .Rows - 1, 0) = 1 '1为选中
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o0").Picture
                            
                            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 1) = False
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbWhite
                        End If
                    End If
                Else
                    '共选子项
                    .RowData(.Rows - 1) = 3 '表明是共选子项
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Mid(arrSub(j), 2)

                    If InStr("," & vfgList.TextMatrix(vfgList.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                        blnDo = True
                        '主项没有选择时,子项不能选择
                        For k = .Rows - 2 To 0 Step -1
                            If .RowData(k) <> 3 Then
                                If .Cell(flexcpData, k, 0) = 0 Then blnDo = False
                                Exit For
                            End If
                        Next
                    Else
                        blnDo = False
                    End If
                    
                    If blnDo Then
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c1").Picture
                        
                        .Cell(flexcpFontBold, .Rows - 1, 1) = True
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                        
                        .Cell(flexcpFontBold, .Rows - 1, 1) = False
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                    End If
                End If
            Next
        Next
        
        .Row = 0: .Col = 1
        
        .Height = .Rows * (.RowHeightMin + 15) + 30
        If .Height > Me.ScaleHeight - 100 Then .Height = Me.ScaleHeight - 100
        If .Height < 3 * (.RowHeightMin + 15) + 30 Then .Height = 3 * (.RowHeightMin + 15) + 30
        
        .Width = (vfgList.Width - 30) - (vfgList.CellLeft + 15)
        
        .Left = picPart.Left + vfgList.Left + vfgList.CellLeft + 15
        
        .Top = picBack.Top + uspRequestPage.Top + picPart.Top + vfgList.Top + vfgList.CellTop + vfgList.CellHeight '+ .Height
        If .Top + .Height > Me.ScaleHeight Then
            .Top = .Top - .Height - vfgList.CellHeight
        End If
        
        .ZOrder
        If .Tag = "AutoPopup" Then
            .Visible = .Rows > 1
        Else
            .Visible = True
        End If
        If .Visible Then .SetFocus
    End With
End Sub


Private Sub vfgList_DblClick()
On Error GoTo errHandle
    If vfgList.Editable <> flexEDNone And vfgList.MouseCol = 1 And vfgList.MouseRow >= vfgList.FixedRows Then
        Call vfgList_KeyPress(vbKeySpace)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
'功能：非编辑状态时，自动移动单元格
    Dim strPartGroup As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
 
            If vfgList.Col <= 1 Then
                vfgList.Col = vfgList.Col + 1
            ElseIf vfgList.Col = 2 And vfgList.Row <= vfgList.Rows - 2 Then
                vfgList.Row = vfgList.Row + 1
                vfgList.Col = 1
            ElseIf vfgList.Col = 2 And vfgList.Row = vfgList.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        
    ElseIf KeyAscii = Asc("*") Then
 
            If vfgList.Col = 2 Then
                Call vfgList_CellButtonClick(vfgList.Row, vfgList.Col)
            End If

    ElseIf KeyAscii = vbKeySpace Then
        If vfgList.Editable <> flexEDNone Then
            If vfgList.Col = 1 Then
                If vfgList.Cell(flexcpData, vfgList.Row, vfgList.Col) = 1 Then
                
                    
                    Set vfgList.Cell(flexcpPicture, vfgList.Row, vfgList.Col) = img16.ListImages("c0").Picture
                    
                    vfgList.Cell(flexcpData, vfgList.Row, vfgList.Col) = 0
                    vfgList.Cell(flexcpFontBold, vfgList.Row, 1, vfgList.Row, 3) = False
                    vfgList.Cell(flexcpBackColor, vfgList.Row, 1, vfgList.Row, vfgList.Cols - 1) = vbWhite
                    
                Else
                    
                    Set vfgList.Cell(flexcpPicture, vfgList.Row, vfgList.Col) = img16.ListImages("c1").Picture
                    
                    vfgList.Cell(flexcpData, vfgList.Row, vfgList.Col) = 1
                    vfgList.Cell(flexcpFontBold, vfgList.Row, 1, vfgList.Row, 3) = True
                    vfgList.Cell(flexcpBackColor, vfgList.Row, 1, vfgList.Row, vfgList.Cols - 1) = &HC0E0FF
                    
                    '自动弹出方法选择器
                    vfgList.Col = 2
                    vsMethod.Tag = "AutoPopup"
                    Call vfgList_CellButtonClick(vfgList.Row, vfgList.Col)
                    vsMethod.Tag = ""
                End If
                            
                '将当前部位配置保存到对应项目中
                strPartGroup = GetPartMethod
                vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, 3) = strPartGroup
                
                If strPartGroup <> "" Then Call SelectRequestProject(vfgRequestProject.RowSel)
                
                Call ShowStudyProjectToTxt
                
                Call SetRequestPageState(True)
            ElseIf vfgList.Col = 2 Then
                Call vfgList_CellButtonClick(vfgList.Row, vfgList.Col)
            End If
        End If
    End If
End Sub


Private Function GetPartMethod() As String
'收集部位及方法的情况
    Dim i As Long
    
    With vfgList
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, 1) = 1 Then
                
                GetPartMethod = GetPartMethod & "|" & .TextMatrix(i, 1) & ";" & .TextMatrix(i, 2)
            End If
        Next
        
        If GetPartMethod <> "" Then
            GetPartMethod = Mid(GetPartMethod, 2)
        End If
    End With
End Function



Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errHandle
    If vfgList.Col = 1 And vfgList.MouseCol = 1 Then
        If X <= vfgList.CellLeft + 250 Then
            Call vfgList_KeyPress(vbKeySpace)
        End If
    End If
 
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SelectRequestProject(ByVal lngRowSel As Long, Optional ByVal blnIsLoad As Boolean = False)
    Dim oldSelectIndex As Long
    
    If lngRowSel <= -1 Then Exit Sub
    
    oldSelectIndex = Val(vfgRequestProject.Tag)
    
    '恢复之前被选中的行图标
    If oldSelectIndex >= 0 Then
        If oldSelectIndex <> lngRowSel Then
            Set vfgRequestProject.Cell(flexcpPicture, oldSelectIndex, TProjectCol.pcId) = img16.ListImages("o0").Picture
            vfgRequestProject.Cell(flexcpFontBold, oldSelectIndex, 0, oldSelectIndex, 2) = False
            vfgRequestProject.Cell(flexcpBackColor, oldSelectIndex, 0, oldSelectIndex, vfgRequestProject.Cols - 1) = vbWhite
        
            vfgRequestProject.Cell(flexcpText, oldSelectIndex, TProjectCol.pcMethod) = ""
        End If
    End If
    
    Set vfgRequestProject.Cell(flexcpPicture, lngRowSel, TProjectCol.pcId) = img16.ListImages("o1").Picture
    vfgRequestProject.Cell(flexcpFontBold, lngRowSel, 0, lngRowSel, 2) = True
    vfgRequestProject.Cell(flexcpBackColor, lngRowSel, 0, lngRowSel, vfgRequestProject.Cols - 1) = &HC0E0FF
    
    '设置执行科室
    If cbxExeRoom.ListIndex >= 0 Then
        vfgRequestProject.Cell(flexcpText, lngRowSel, TProjectCol.pcExeRoom) = cbxExeRoom.ItemData(cbxExeRoom.ListIndex) & "-" & Replace(cbxExeRoom.Text, Val(cbxExeRoom.Text) & "-", "")
    End If
    
    vfgRequestProject.Tag = lngRowSel
    
    If Not blnIsLoad Then Call SetRequestPageState(True)
    
    Call ShowArrangeState
End Sub

Public Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Sub ShowArrangeState()
'显示执行项目的当前安排状态
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strHint As String
    Dim lngProjectIndex As Long
    Dim lngRoomId As Long
    Dim lngProjectId As Long
    Dim strDateRang As String
    
    
    lngRoomId = 0
    strHint = ""
    
    strDateRang = "To_Date([1],'YYYY-MM-DD HH24:MI:SS') and To_Date([2],'YYYY-MM-DD HH24:MI:SS')"
    
    If cbxExeRoom.ListIndex >= 0 Then
        lngRoomId = cbxExeRoom.ItemData(cbxExeRoom.ListIndex)
        
        '查询科室等待情况
        strSQL = "SELECT count(*) as 数量 FROM 病人医嘱记录 a, 病人医嘱发送 b where a.id=b.医嘱id and a.相关id is null and a.诊疗类别='D' " & _
                    " and 开始执行时间 between " & strDateRang & " and nvl(b.执行过程,0)<3 and nvl(b.执行状态,0)<> 2 and b.执行部门id+0=[3]"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询科室等待情况", _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 00:00:00"), _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 23:59:59"), _
                                            lngRoomId)
                                            
        strHint = "所选科室共有【" & Val(NVL(rsTemp!数量)) & "】人等待检查    "
    End If
    
    lngProjectIndex = Val(vfgRequestProject.Tag)
    
    lngProjectId = 0
    
    If lngProjectIndex >= 0 Then
        lngProjectId = Val(vfgRequestProject.Cell(flexcpData, lngProjectIndex, TProjectCol.pcId))
        
        '查询项目等待情况
        strSQL = "SELECT count(*) as 数量 FROM 病人医嘱记录 a, 病人医嘱发送 b where a.id=b.医嘱id and a.相关id is null and a.诊疗类别='D' " & _
                    " and 开始执行时间 between " & strDateRang & " and nvl(b.执行过程,0)<3 and nvl(b.执行状态,0)<> 2 and a.诊疗项目id+0=[3]"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询项目等待情况", _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 00:00:00"), _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 23:59:59"), _
                                            lngProjectId)
                                            
        strHint = strHint & "所选项目共有【" & Val(NVL(rsTemp!数量)) & "】人等待检查    "
    End If
    
'    '查询当前科室下的项目等待情况
'    If lngRoomId <> 0 And lngProjectId <> 0 Then
'        strSQL = "SELECT count(*) as 数量 FROM 病人医嘱记录 a, 病人医嘱发送 b where a.id=b.医嘱id and a.相关id is null and a.诊疗类别='D' " & _
'                    " and 开始执行时间 between " & strDateRang & " and nvl(b.执行过程,0)<3 and nvl(b.执行状态,0)<> 2 and a.诊疗项目id=[3] and b.执行部门Id=[4]"
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询项目等待情况", _
'                                            Format(dtpExeTime.value, "yyyy-mm-dd 00:00:00"), _
'                                            Format(dtpExeTime.value, "yyyy-mm-dd 23:59:59"), _
'                                            lngProjectId, lngRoomId)
'
'        strHint = strHint & "当前选择的项目有【" & Val(Nvl(rsTemp!数量)) & "】人需到所选科室检查"
'    End If
    
    stbThis.Panels(2).Text = strHint
    
End Sub


Private Sub ShowStudyProjectToTxt()
'将选择的检查项目和方法显示到txt只读文本中
    Dim lngProjectIndex As Long
    Dim dblPrice As Double
    Dim lngExeType As Long
    Dim strExeRoomName As String
    Dim lngExeRoomId As Long
    
    lngProjectIndex = Val(vfgRequestProject.Tag)
    If lngProjectIndex < 0 Then
        txtCurStudyProject.Text = ""
        Exit Sub
    End If
    
    strExeRoomName = vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcExeRoom)
    strExeRoomName = Replace(strExeRoomName, Val(strExeRoomName) & "-", "")
    
    txtCurStudyProject.Text = Replace("【" & strExeRoomName & "：" & _
                                        vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcName) & _
                                        "】" & vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcMethod), "|", "■")
                            
    lngExeType = 0    '执行类型
    If Val(vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcBedCol)) = 1 Then
        lngExeType = 1
    ElseIf Val(vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcOperCol)) = 1 Then
        lngExeType = 2
    End If
    
    If cbxExeRoom.ListIndex < 0 Then
        lngExeRoomId = 0
    Else
        lngExeRoomId = Val(cbxExeRoom.ItemData(cbxExeRoom.ListIndex))
    End If
                        
    '计算检查费用
    dblPrice = GetPrice(mCurPatientInf.lngID, mCurPatientInf.lngPageId, vfgRequestProject.Cell(flexcpData, lngProjectIndex, TProjectCol.pcId), _
                        Trim(vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcMethod)), _
                        lngExeType, _
                        mCurPatientInf.lngFrom, _
                        lngExeRoomId)
                        
    labPrice = "￥：" & Format(dblPrice, "0.00")
End Sub


Private Sub vfgList_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vfgList.EditSelStart = 0
    vfgList.EditSelLength = zlCommFun.ActualLen(vfgList.EditText)
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo errHandle
    If Col <> 2 Then Cancel = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_DblClick()
On Error GoTo errHandle
    If vfgRequestProject.RowSel <= -1 Then Exit Sub
    If vfgRequestProject.Editable = flexEDNone Then Exit Sub
    
    Call SelectRequestProject(vfgRequestProject.RowSel)
    Call ShowStudyProjectToTxt
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_GotFocus()
On Error GoTo errHandle
    If vfgRequestProject.RowSel < 0 And vfgRequestProject.Rows > 0 Then
        vfgRequestProject.Row = 0
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii <> vbKeySpace Then Exit Sub
    If vfgRequestProject.Editable = flexEDNone Then Exit Sub
    
    Call SelectRequestProject(vfgRequestProject.RowSel)

    If vfgRequestProject.Col >= TProjectCol.pcNormalCol Then
        Call SelectProjectExeType(vfgRequestProject.RowSel, vfgRequestProject.Col)
    End If

    
    Call ShowStudyProjectToTxt
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SelectProjectExeType(ByVal lngRow As Long, ByVal lngCol As Long)
    '如果没有内容，则不需要设置选择状态
    If vfgRequestProject.Cell(flexcpText, lngRow, lngCol) = "" Then Exit Sub
    
    Set vfgRequestProject.Cell(flexcpPicture, lngRow, lngCol) = img16.ListImages("o1").Picture
    vfgRequestProject.Cell(flexcpData, lngRow, lngCol) = 1
    
    Select Case lngCol
        Case TProjectCol.pcNormalCol
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcBedCol) = img16.ListImages("o0").Picture
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcOperCol) = img16.ListImages("o0").Picture
             
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcBedCol) = 0
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcOperCol) = 0
            
        Case TProjectCol.pcBedCol
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcNormalCol) = img16.ListImages("o0").Picture
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcOperCol) = img16.ListImages("o0").Picture
            
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcNormalCol) = 0
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcOperCol) = 0
            
        Case TProjectCol.pcOperCol
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcNormalCol) = img16.ListImages("o0").Picture
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcBedCol) = img16.ListImages("o0").Picture
            
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcNormalCol) = 0
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcBedCol) = 0
    End Select
End Sub


Private Sub vfgRequestProject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errHandle
    '只有选择第一列时，才能进行项目选择
    If vfgRequestProject.RowSel < 0 Then Exit Sub
    If vfgRequestProject.Editable = flexEDNone Then Exit Sub
    
    If vfgRequestProject.ColSel = 0 Then
        Call SelectRequestProject(vfgRequestProject.RowSel)
        
        Call ShowStudyProjectToTxt
        
    ElseIf vfgRequestProject.ColSel >= TProjectCol.pcNormalCol _
        And vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, vfgRequestProject.Col) <> "" Then
        
        Call SelectRequestProject(vfgRequestProject.RowSel)
        
        Call SelectProjectExeType(vfgRequestProject.RowSel, vfgRequestProject.Col)
        
        Call ShowStudyProjectToTxt
        
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_SelChange()
On Error GoTo errHandle
    Dim blnIsLoading As Boolean
    
    If vfgRequestProject.RowSel < 0 Then
        Call LoadRequestPagePart(0, "", "")
        Exit Sub
    End If
    
    blnIsLoading = False
    
    '判断数据原本是否处于载入中
    If mblnIsLoadRequestPage Then blnIsLoading = True
    
    mblnIsLoadRequestPage = True

    '载入项目对应的部位
    Call LoadRequestPagePart(vfgRequestProject.Cell(flexcpData, vfgRequestProject.RowSel, TProjectCol.pcId), _
                            mCurPatientInf.strSex, _
                            vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcMethod))
    
    '载入项目的执行科室
    Call LoadRequestProjectExeRoom(vfgRequestProject.Cell(flexcpData, vfgRequestProject.RowSel, TProjectCol.pcId), _
                                    vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcRoomType), _
                                    Val(vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcExeRoom)))
                                    
                                    
    '检查列表编辑属性要与项目的可编辑属性相同
    vfgList.Editable = vfgRequestProject.Editable
                                    
    If Not blnIsLoading Then mblnIsLoadRequestPage = False
                                            
Exit Sub
errHandle:
    If Not blnIsLoading Then mblnIsLoadRequestPage = False
    
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    Cancel = True
err.Clear
End Sub

Private Sub vsMethod_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 And NewRow <> -1 Then
        If vsMethod.TextMatrix(NewRow, 0) = "" Then
            Cancel = True
            vsMethod.Col = 1
        End If
    End If
End Sub

Private Sub vsMethod_DblClick()
    Call vsMethod_KeyPress(13)
End Sub

Private Sub vsMethod_KeyPress(KeyAscii As Integer)
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMethod
        If KeyAscii = 13 Then
            Call ConfirmMethod
            vsMethod.Visible = False
            vfgList.SetFocus
        ElseIf KeyAscii = vbKeySpace Then
            '检查方法的选择与取消
            If .Cell(flexcpData, .Row, 0) = 1 Then
                '单选项目前也允许取消选择
                .Cell(flexcpData, .Row, 0) = 0
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = False
                
                Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o0", "c0")).Picture
                '同时取消该单选项的子项
                If .RowData(.Row) = 1 Then
                    For i = .Row + 1 To .Rows - 1
                        If .RowData(i) = 3 Then
                            If .Cell(flexcpData, i, 0) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                                
                                Set .Cell(flexcpPicture, i, 1) = img16.ListImages("c0").Picture
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            Else
                blnDo = True
                If .RowData(.Row) = 3 Then
                    '主项没有选择时,子项不能选择
                    For i = .Row - 1 To 0 Step -1
                        If .RowData(i) <> 3 Then
                            If .Cell(flexcpData, i, 0) = 0 Then blnDo = False
                            Exit For
                        End If
                    Next
                End If
                
                If blnDo Then
                    .Cell(flexcpData, .Row, 0) = 1
                    .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                    
                    Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o1", "c1")).Picture
                    
                    If .RowData(.Row) = 1 Then '单选项选中时，取消其他单选项
                        For i = 0 To .Rows - 1
                            If i <> .Row And .RowData(i) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                                
                                Set .Cell(flexcpPicture, i, 0, i, 1) = img16.ListImages("o0").Picture
                                For j = i + 1 To .Rows - 1 '同时取消该单选项的子项
                                    If .RowData(j) = 3 Then
                                        If .Cell(flexcpData, j, 0) = 1 Then
                                            .Cell(flexcpData, j, 0) = 0
                                            .Cell(flexcpFontBold, j, 0, j, .Cols - 1) = False
                                            
                                            Set .Cell(flexcpPicture, j, 1) = img16.ListImages("c0").Picture
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            
            Call ConfirmMethod
        End If
    End With
End Sub


Private Sub ConfirmMethod()
'功能：检查方法的确认
    Dim strMethod As String, i As Long
    Dim strPartGroup As String
        
    With vsMethod
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strMethod = strMethod & "," & .TextMatrix(i, 1)
            End If
        Next
        
        vfgList.TextMatrix(vfgList.Row, 2) = Mid(strMethod, 2)
        
        '方法设置后，自动选中该部位
        If vfgList.TextMatrix(vfgList.Row, 2) <> "" Then
            vfgList.Cell(flexcpData, vfgList.Row, 1) = 1
            vfgList.Cell(flexcpFontBold, vfgList.Row, 1, vfgList.Row, 3) = True
            
            Set vfgList.Cell(flexcpPicture, vfgList.Row, 1) = img16.ListImages("c1").Picture
        End If
    End With
    
    strPartGroup = GetPartMethod
    vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, 3) = strPartGroup
    
    If strPartGroup <> "" Then
        Call SelectRequestProject(vfgRequestProject.RowSel)
    End If
    
    Call ShowStudyProjectToTxt
End Sub

Private Sub vsMethod_LostFocus()
    vsMethod.Visible = False
End Sub

Private Sub vsMethod_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = vsMethod.MouseCol
    If vsMethod.Col = lngCol And vsMethod.Text <> "" Then
        If X <= vsMethod.CellLeft + 250 Then
            Call vsMethod_KeyPress(vbKeySpace)
        End If
    End If
End Sub
