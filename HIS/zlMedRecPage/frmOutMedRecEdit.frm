VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOutMedRecEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊首页"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmOutMedRecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8205
      TabIndex        =   98
      Top             =   7650
      Width           =   8205
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改病人基本信息"
         Height          =   350
         Left            =   360
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   60
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6840
         TabIndex        =   99
         Top             =   60
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确认(&O)"
         Height          =   350
         Left            =   5640
         TabIndex        =   97
         ToolTipText     =   "热键：F2"
         Top             =   60
         Width           =   1100
      End
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   7575
      Left            =   120
      TabIndex        =   100
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本信息"
      TabPicture(0)   =   "frmOutMedRecEdit.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraMain(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "就诊信息"
      TabPicture(1)   =   "frmOutMedRecEdit.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraMain(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraMain 
         BorderStyle     =   0  'None
         Height          =   7000
         Index           =   1
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   7515
         Begin VB.Frame fraDocSum 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   0
            TabIndex        =   105
            Top             =   1560
            Width           =   7335
            Begin VB.TextBox txtInfo 
               Height          =   555
               Index           =   12
               Left            =   120
               MaxLength       =   1000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   72
               Top             =   320
               Width           =   7125
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   " 就诊摘要 "
               Height          =   180
               Index           =   20
               Left            =   360
               TabIndex        =   71
               Top             =   90
               Width           =   900
            End
            Begin VB.Line linDocSum 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   120
               X2              =   7200
               Y1              =   180
               Y2              =   180
            End
            Begin VB.Line linDocSum 
               BorderColor     =   &H80000014&
               Index           =   1
               X1              =   120
               X2              =   7200
               Y1              =   195
               Y2              =   195
            End
         End
         Begin VB.Frame fraOtherInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1845
            Left            =   120
            TabIndex        =   104
            Top             =   5160
            Width           =   7335
            Begin zlMedRecPage.UCPatiVitalSigns UCPatiVitalSigns 
               Height          =   750
               Left            =   360
               TabIndex        =   84
               Top             =   375
               Width           =   5745
               _extentx        =   10134
               _extenty        =   1323
               textbackcolor   =   -2147483643
               font            =   "frmOutMedRecEdit.frx":0044
               forecolor       =   0
               xdis            =   300
               ydis            =   80
               labtotxt        =   30
            End
            Begin VB.ComboBox cboBaseInfo 
               Height          =   300
               Index           =   9
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   30
               Width           =   2160
            End
            Begin VB.OptionButton optState 
               Caption         =   "复诊"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   80
               Top             =   60
               Width           =   855
            End
            Begin VB.CommandButton cmdInfo 
               Caption         =   "…"
               Height          =   255
               Index           =   17
               Left            =   3360
               TabIndex        =   94
               TabStop         =   0   'False
               ToolTipText     =   "选择(*)"
               Top             =   1560
               Width           =   285
            End
            Begin VB.TextBox txtInfo 
               Height          =   300
               Index           =   18
               Left            =   5160
               MaxLength       =   100
               TabIndex        =   96
               Top             =   1560
               Width           =   1935
            End
            Begin VB.TextBox txtInfo 
               Height          =   300
               Index           =   17
               Left            =   870
               MaxLength       =   100
               TabIndex        =   93
               Top             =   1560
               Width           =   2760
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Caption         =   "传染病上传(&U)"
               Height          =   195
               Index           =   16
               Left            =   2040
               TabIndex        =   81
               Top             =   90
               Width           =   1470
            End
            Begin VB.OptionButton optState 
               Caption         =   "初诊"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   79
               Top             =   60
               Value           =   -1  'True
               Width           =   855
            End
            Begin MSMask.MaskEdBox mskDateInfo 
               Height          =   300
               Index           =   6
               Left            =   5160
               TabIndex        =   88
               Tag             =   "####-##-##"
               Top             =   1185
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "####-##-##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDateInfo 
               Height          =   300
               Index           =   7
               Left            =   6195
               TabIndex        =   89
               Tag             =   "##:##"
               Top             =   1185
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   529
               _Version        =   393216
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   5
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtInfo 
               Height          =   300
               Index           =   16
               Left            =   870
               MaxLength       =   200
               TabIndex        =   86
               Top             =   1215
               Width           =   2760
            End
            Begin VB.TextBox txtDateInfo 
               Height          =   300
               Index           =   7
               Left            =   6195
               MaxLength       =   30
               TabIndex        =   91
               Top             =   1185
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.TextBox txtDateInfo 
               Height          =   300
               Index           =   6
               Left            =   5160
               MaxLength       =   30
               TabIndex        =   90
               Top             =   1185
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblInfo 
               Caption         =   "其他医学警示"
               Height          =   180
               Index           =   18
               Left            =   3960
               TabIndex        =   95
               Top             =   1560
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Caption         =   "医学警示"
               Height          =   180
               Index           =   17
               Left            =   120
               TabIndex        =   92
               Top             =   1620
               Width           =   720
            End
            Begin VB.Label lblBaseInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "去向"
               Height          =   180
               Index           =   9
               Left            =   4440
               TabIndex        =   82
               Top             =   90
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病地址"
               Height          =   180
               Index           =   16
               Left            =   120
               TabIndex        =   85
               Top             =   1245
               Width           =   720
            End
            Begin VB.Label lblDateInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病时间"
               Height          =   180
               Index           =   6
               Left            =   4320
               TabIndex        =   87
               Top             =   1200
               Width           =   720
            End
         End
         Begin VB.Frame fraAller 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1580
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   7335
            Begin VB.CheckBox chkInfo 
               Caption         =   "无过敏记录"
               Height          =   195
               Index           =   30
               Left            =   1440
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   83
               Width           =   1290
            End
            Begin VB.OptionButton optAller 
               Caption         =   "根据药品目录输入(&1)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   0
               Left            =   2880
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   90
               Value           =   -1  'True
               Width           =   2130
            End
            Begin VB.OptionButton optAller 
               Caption         =   "根据过敏源输入(&2)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   1
               Left            =   5070
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   90
               Width           =   1890
            End
            Begin VSFlex8Ctl.VSFlexGrid vsAller 
               Height          =   1260
               Left            =   120
               TabIndex        =   70
               Top             =   315
               Width           =   7125
               _cx             =   12568
               _cy             =   2222
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
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmOutMedRecEdit.frx":0068
               ScrollTrack     =   -1  'True
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
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   " 过敏记录 "
               Height          =   180
               Index           =   21
               Left            =   360
               TabIndex        =   66
               Top             =   90
               Width           =   900
            End
            Begin VB.Line linAller 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   120
               X2              =   7200
               Y1              =   180
               Y2              =   180
            End
            Begin VB.Line linAller 
               BorderColor     =   &H80000014&
               Index           =   1
               X1              =   120
               X2              =   7200
               Y1              =   195
               Y2              =   195
            End
         End
         Begin VB.Frame fraDiag 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2715
            Left            =   0
            TabIndex        =   102
            Top             =   2500
            Width           =   7335
            Begin VB.OptionButton optDiag 
               Caption         =   "根据诊断标准输入(&3)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   0
               Left            =   2820
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   90
               Value           =   -1  'True
               Width           =   2010
            End
            Begin VB.OptionButton optDiag 
               Caption         =   "根据疾病编码输入(&4)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   1
               Left            =   4890
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   90
               Width           =   2010
            End
            Begin VB.CommandButton cmdMakeLog 
               Height          =   255
               Left            =   1560
               Picture         =   "frmOutMedRecEdit.frx":011B
               Style           =   1  'Graphical
               TabIndex        =   74
               TabStop         =   0   'False
               ToolTipText     =   "根据诊断生成就诊摘要(F12)"
               Top             =   53
               Width           =   345
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   1260
               Left            =   120
               TabIndex        =   77
               Top             =   360
               Width           =   7125
               _cx             =   12568
               _cy             =   2222
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
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmOutMedRecEdit.frx":06A5
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   115
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
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   960
               Left            =   120
               TabIndex        =   78
               Top             =   1700
               Width           =   7125
               _cx             =   12568
               _cy             =   1693
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
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   24
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmOutMedRecEdit.frx":0961
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
               ScrollTips      =   0   'False
               MergeCells      =   115
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
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   " 诊断记录 "
               Height          =   180
               Index           =   19
               Left            =   360
               TabIndex        =   73
               Top             =   90
               Width           =   900
            End
            Begin VB.Line linDiag 
               BorderColor     =   &H80000014&
               Index           =   1
               X1              =   120
               X2              =   7215
               Y1              =   195
               Y2              =   195
            End
            Begin VB.Line linDiag 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   120
               X2              =   7215
               Y1              =   180
               Y2              =   180
            End
         End
      End
      Begin VB.Frame fraMain 
         BorderStyle     =   0  'None
         Height          =   6975
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   7425
         Begin VB.CommandButton cmdPicClear 
            Caption         =   "清除"
            Height          =   350
            Left            =   6600
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   5640
            Width           =   600
         End
         Begin VB.CommandButton cmdPicCollect 
            Caption         =   "采集"
            Height          =   350
            Left            =   6600
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   5280
            Width           =   600
         End
         Begin VB.CommandButton cmdPicFile 
            Caption         =   "文件"
            Height          =   350
            Left            =   6600
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   4920
            Width           =   585
         End
         Begin VB.PictureBox picPatient 
            Height          =   2025
            Left            =   4080
            ScaleHeight     =   1965
            ScaleWidth      =   2475
            TabIndex        =   106
            Top             =   4920
            Width           =   2535
            Begin VB.Image imgPatient 
               Height          =   1950
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2460
            End
         End
         Begin VB.CommandButton cmdAdressInfo 
            Caption         =   "…"
            Height          =   255
            Index           =   2
            Left            =   6930
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3705
            Width           =   285
         End
         Begin VB.CommandButton cmdAdressInfo 
            Caption         =   "…"
            Height          =   255
            Index           =   6
            Left            =   6930
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2985
            Width           =   285
         End
         Begin VB.CommandButton cmdAdressInfo 
            Caption         =   "…"
            Height          =   255
            Index           =   0
            Left            =   6930
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2625
            Width           =   285
         End
         Begin VB.CommandButton cmdAdressInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   1
            Left            =   6945
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2265
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Caption         =   "…"
            Height          =   255
            Index           =   3
            Left            =   6930
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   4413
            Width           =   285
         End
         Begin VB.CommandButton cmdAdressInfo 
            Caption         =   "…"
            Height          =   240
            Index           =   5
            Left            =   6945
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1905
            Width           =   270
         End
         Begin VB.TextBox txtAdressInfo 
            Height          =   300
            Index           =   1
            Left            =   4020
            MaxLength       =   100
            TabIndex        =   31
            ToolTipText     =   "按*键显示区域列表"
            Top             =   2235
            Width           =   3195
         End
         Begin VB.TextBox txtAdressInfo 
            Height          =   300
            Index           =   3
            Left            =   900
            MaxLength       =   100
            TabIndex        =   51
            ToolTipText     =   "按*键显示地区列表"
            Top             =   4390
            Width           =   6315
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   135
            Width           =   1200
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   3
            Left            =   900
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   135
            Width           =   1635
         End
         Begin VB.TextBox txtSpecificInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   15
            Left            =   3510
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   495
            Width           =   675
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            HelpContextID   =   1000
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   3510
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   135
            Width           =   1305
         End
         Begin VB.ComboBox cboSpecificInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            Left            =   4200
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   495
            Width           =   615
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   4
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1005
            Width           =   2355
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   5
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1005
            Width           =   3195
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   2
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1365
            Width           =   2355
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   3
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1365
            Width           =   3195
         End
         Begin VB.TextBox txtAdressInfo 
            Height          =   300
            Index           =   6
            Left            =   900
            MaxLength       =   100
            TabIndex        =   37
            ToolTipText     =   "按*键显示合约单位列表"
            Top             =   2955
            Width           =   6315
         End
         Begin VB.TextBox txtSpecificInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   900
            MaxLength       =   20
            TabIndex        =   40
            Top             =   3315
            Width           =   3090
         End
         Begin VB.TextBox txtSpecificInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   4905
            MaxLength       =   6
            TabIndex        =   42
            Top             =   3315
            Width           =   2310
         End
         Begin VB.TextBox txtAdressInfo 
            Height          =   300
            Index           =   2
            Left            =   900
            MaxLength       =   100
            TabIndex        =   44
            ToolTipText     =   "按*键显示地区列表"
            Top             =   3675
            Width           =   6315
         End
         Begin VB.TextBox txtSpecificInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   900
            MaxLength       =   20
            TabIndex        =   47
            Top             =   4035
            Width           =   3090
         End
         Begin VB.TextBox txtSpecificInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   4905
            MaxLength       =   6
            TabIndex        =   49
            Top             =   4035
            Width           =   2310
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   0
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   495
            Width           =   1215
         End
         Begin VB.TextBox txtAdressInfo 
            Height          =   300
            Index           =   0
            Left            =   900
            MaxLength       =   100
            TabIndex        =   34
            ToolTipText     =   "按*键显示地区列表"
            Top             =   2595
            Width           =   6315
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   15
            Left            =   900
            MaxLength       =   64
            TabIndex        =   56
            Top             =   5146
            Width           =   2595
         End
         Begin VB.TextBox txtAdressInfo 
            Height          =   300
            Index           =   5
            Left            =   4020
            MaxLength       =   30
            TabIndex        =   26
            ToolTipText     =   "按*键显示区域列表"
            Top             =   1875
            Width           =   3195
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   4
            Left            =   900
            MaxLength       =   20
            TabIndex        =   29
            Top             =   2235
            Width           =   2340
         End
         Begin VB.TextBox txtSpecificInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   900
            MaxLength       =   6
            TabIndex        =   54
            Top             =   4768
            Width           =   2595
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   8
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   5524
            Width           =   2595
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   43
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   5902
            Width           =   2595
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   38
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   6660
            Width           =   2595
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   36
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   6280
            Width           =   2595
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   300
            Index           =   0
            Left            =   900
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##"
            Top             =   495
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   16
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   61
            ItemData        =   "frmOutMedRecEdit.frx":0B53
            Left            =   900
            List            =   "frmOutMedRecEdit.frx":0B55
            TabIndex        =   24
            Top             =   1875
            Width           =   2340
         End
         Begin MSComDlg.CommonDialog cmdialog 
            Left            =   6600
            Top             =   6240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtDateInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   0
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   495
            Width           =   1635
         End
         Begin VB.Line linInfo 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   -15
            X2              =   7335
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊号"
            Height          =   180
            Index           =   14
            Left            =   5400
            TabIndex        =   5
            Top             =   195
            Width           =   540
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            Height          =   180
            Index           =   3
            Left            =   480
            TabIndex        =   1
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            Height          =   180
            Index           =   1
            Left            =   3090
            TabIndex        =   3
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            Height          =   180
            Index           =   15
            Left            =   3090
            TabIndex        =   10
            Top             =   555
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻状况"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "职业"
            Height          =   180
            Index           =   3
            Left            =   3555
            TabIndex        =   21
            Top             =   1425
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "民族"
            Height          =   180
            Index           =   5
            Left            =   3555
            TabIndex        =   17
            Top             =   1065
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "国籍"
            Height          =   180
            Index           =   4
            Left            =   480
            TabIndex        =   14
            Top             =   1065
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证号"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label lblAdressInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位名称"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   36
            Top             =   3015
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位电话"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   3375
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位邮编"
            Height          =   180
            Index           =   2
            Left            =   4095
            TabIndex        =   41
            Top             =   3375
            Width           =   720
         End
         Begin VB.Label lblAdressInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭地址"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   3735
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭电话"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   4095
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭邮编"
            Height          =   180
            Index           =   4
            Left            =   4095
            TabIndex        =   48
            Top             =   4095
            Width           =   720
         End
         Begin VB.Line linInfo 
            BorderColor     =   &H80000014&
            Index           =   0
            X1              =   -60
            X2              =   7290
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Line linInfo 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   -150
            X2              =   7200
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Line linInfo 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   -105
            X2              =   7245
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付费方式"
            Height          =   180
            Index           =   0
            Left            =   5220
            TabIndex        =   13
            Top             =   555
            Width           =   720
         End
         Begin VB.Label lblAdressInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生地点"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   2655
            Width           =   720
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   555
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "监护人"
            Height          =   180
            Index           =   15
            Left            =   300
            TabIndex        =   55
            Top             =   5175
            Width           =   540
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "区域"
            Height          =   180
            Index           =   5
            Left            =   3555
            TabIndex        =   25
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "其他证件"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   2295
            Width           =   720
         End
         Begin VB.Label lblAdressInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口地址"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   50
            Top             =   4455
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口邮编"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   53
            Top             =   4810
            Width           =   720
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "籍贯"
            Height          =   180
            Index           =   1
            Left            =   3555
            TabIndex        =   30
            Top             =   2295
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "文化程度"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Top             =   5584
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "生育状况"
            Height          =   180
            Index           =   43
            Left            =   120
            TabIndex        =   59
            Top             =   5962
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Caption         =   "血型"
            Height          =   180
            Index           =   36
            Left            =   480
            TabIndex        =   61
            Top             =   6340
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Caption         =   "Rh"
            Height          =   180
            Index           =   38
            Left            =   645
            TabIndex        =   63
            Top             =   6720
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   -64800
         TabIndex        =   101
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   6060
         Width           =   270
      End
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   7920
      Picture         =   "frmOutMedRecEdit.frx":0B57
      Top             =   6480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   7920
      Picture         =   "frmOutMedRecEdit.frx":10E1
      Top             =   6840
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmOutMedRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowMe() As Boolean
    Me.Show , gclsPros.MainForm
    ShowMe = True
End Function

Private Sub cboBaseInfo_Click(Index As Integer)
    Call CboBaseInfoClick(Index)
End Sub

Private Sub cboBaseInfo_GotFocus(Index As Integer)
    Call CboBaseInfoGotFocus(Index)
End Sub

Private Sub cboBaseInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CboBaseInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cboBaseInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyUp(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_Validate(Index As Integer, Cancel As Boolean)
    Call cboBaseInfoValidate(Index, Cancel)
End Sub

Private Sub cboBaseInfo_Change(Index As Integer)
    Call CboBaseInfoChange(Index)
End Sub

Private Sub CboSpecificInfo_Click(Index As Integer)
    Call CboSpecificInfoClick(Index)
End Sub

Private Sub cboSpecificInfo_GotFocus(Index As Integer)
    Call CboSpecificInfoGotFocus(Index)
End Sub

Private Sub cboSpecificInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboSpecificInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub cboSpecificInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call cboSpecificInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Call chkInfoClick(Index)
End Sub

Private Sub chkInfo_GotFocus(Index As Integer)
    Call ChkInfoGotFocus(Index)
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call ChkInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cmdAdressInfo_Click(Index As Integer)
    Call CmdAdressInfoClick(Index)
End Sub

Private Sub cmdCancel_Click()
    Call CmdCancelClick
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Call CmdInfoClick(Index)
End Sub

Private Sub cmdMakeLog_Click()
    Call cmdMakeLogClick
End Sub

Private Sub cmdModify_Click()
    Call cmdModifyClick
End Sub

Private Sub cmdOK_Click()
    Call CmdOKClick
End Sub

'暂时不共享，主要时只有该窗体有该功能
Private Sub cmdPicClear_Click()
    '问题号:74421
    imgPatient.Picture = Nothing
    picPatient.Tag = ""
End Sub
'暂时不共享，主要时只有该窗体有该功能
Private Sub cmdPicCollect_Click()
    Dim strPictureFile As String
    If gobjPatient Is Nothing Then
        On Error Resume Next
        Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
        Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
    End If
    If gobjPatient Is Nothing Then
        MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, gclsPros.CurrentForm.Caption
        Exit Sub
    End If
    If gobjPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
    Set imgPatient.Picture = LoadPicture(strPictureFile)
    picPatient.Tag = strPictureFile
End Sub
'暂时不共享，主要时只有该窗体有该功能
Private Sub cmdPicFile_Click()
    '问题号:74421
    Dim strFileDir As String
    On Error GoTo Errhand:
    With cmdialog
        .CancelError = False
        .Flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .FilterIndex = 2
        .ShowOpen
        strFileDir = .FileName
        If strFileDir = "" Then Exit Sub
        imgPatient.Picture = LoadPicture(strFileDir)
        picPatient.Tag = strFileDir
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FormKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call FormKeyPress(KeyAscii)
End Sub

Private Sub Form_Load()
    If Not FormLoad Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FormUnLoad(Cancel)
End Sub

Private Sub mskDateInfo_Change(Index As Integer)
    Call DateInfoChange(Index)
End Sub

Private Sub mskDateInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call DateInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub mskDateInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call DateInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub mskDateInfo_Validate(Index As Integer, Cancel As Boolean)
    Call DateInfoValidate(Index, Cancel)
End Sub

Private Sub optAller_Click(Index As Integer)
    Call OptAllerClick(Index)
End Sub

Private Sub optAller_KeyPress(Index As Integer, KeyAscii As Integer)
    Call OptAllerKeyPress(Index, KeyAscii)
End Sub

Private Sub optDiag_Click(Index As Integer)
    Call optDiagClick(Index)
End Sub

Private Sub optDiag_KeyPress(Index As Integer, KeyAscii As Integer)
    Call optDiagKeyPress(Index, KeyAscii)
End Sub

Private Sub optState_Click(Index As Integer)
    Call optStateClick(Index)
End Sub

Private Sub txtAdressInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call txtAdressInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtAdressInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call txtAdressInfoMouseDown(Index, Button, Shift, x, Y)
End Sub

Private Sub txtAdressInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call txtAdressInfoMouseUp(Index, Button, Shift, x, Y)
End Sub

Private Sub txtInfo_Change(Index As Integer)
    Call TxtInfoChange(Index)
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call TxtInfoGotFocus(Index)
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call TxtInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call TxtInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call TxtInfoMouseDown(Index, Button, Shift, x, Y)
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call TxtInfoMouseUp(Index, Button, Shift, x, Y)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Call TxtInfoValidate(Index, Cancel)
End Sub

Private Sub txtSpecificInfo_Change(Index As Integer)
    Call SpecificInfoChange(Index)
End Sub

Private Sub txtSpecificInfo_GotFocus(Index As Integer)
    Call SpecificInfoGotFocus(Index)
End Sub

Private Sub txtSpecificInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call SpecificInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtSpecificInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SpecificInfoMouseDown(Index, Button, Shift, x, Y)
End Sub

Private Sub txtSpecificInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SpecificInfoMouseUp(Index, Button, Shift, x, Y)
End Sub

Private Sub txtSpecificInfo_Validate(Index As Integer, Cancel As Boolean)
    Call SpecificInfoValidate(Index, Cancel)
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call AllerAfterEdit(vsAller, Row, Col)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call AllerAfterRowColChange(vsAller, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsAller_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerBeforeEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call AllerCellButtonClick(vsAller, Row, Col)
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Call AllerKeyDown(vsAller, KeyCode, Shift)
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    Call AllerKeyPress(vsAller, KeyAscii)
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call AllerKeyPressEdit(vsAller, Row, Col, KeyAscii)
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call AllerSetupEditWindow(vsAller, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerStartEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerValidateEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagXY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagXY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterUserResize(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_GotFocus()
    Call DiagGotFocus(vsDiagXY)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagXY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagXY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagZY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterUserResize(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_GotFocus()
    Call DiagGotFocus(vsDiagZY)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagZY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagZY, Row, Col, Cancel)
End Sub

