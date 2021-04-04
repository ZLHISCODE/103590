VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmOutMedRecEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊首页"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frmOutMedRecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   7785
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
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   8160
      Width           =   7785
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6480
         TabIndex        =   90
         Top             =   60
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确认(&O)"
         Height          =   350
         Left            =   5280
         TabIndex        =   89
         ToolTipText     =   "热键：F2"
         Top             =   60
         Width           =   1100
      End
   End
   Begin VB.Timer timThis 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   8205
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   8160
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   14393
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本信息"
      TabPicture(0)   =   "frmOutMedRecEdit.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "就诊信息"
      TabPicture(1)   =   "frmOutMedRecEdit.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraInfo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ucPatiVitalSigns"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin zl9CISJob.UCPatiVitalSigns ucPatiVitalSigns 
         Height          =   750
         Left            =   -73890
         TabIndex        =   76
         Top             =   5865
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   1323
         TextBackColor   =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         XDis            =   300
         YDis            =   80
         LabToTxt        =   85
      End
      Begin VB.Frame fraInfo 
         BorderStyle     =   0  'None
         Height          =   7650
         Index           =   1
         Left            =   -74760
         TabIndex        =   93
         Top             =   360
         Width           =   7425
         Begin VB.Frame fraDocSum 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   0
            TabIndex        =   98
            Top             =   1560
            Width           =   7335
            Begin VB.TextBox txtEdit 
               Height          =   555
               Index           =   12
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   67
               Top             =   320
               Width           =   7125
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   " 就诊摘要 "
               Height          =   180
               Index           =   20
               Left            =   360
               TabIndex        =   66
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
            Height          =   2370
            Left            =   120
            TabIndex        =   97
            Top             =   5235
            Width           =   7335
            Begin VB.ComboBox cboEdit 
               Height          =   300
               Index           =   11
               Left            =   1200
               TabIndex        =   83
               Text            =   "cboEdit"
               Top             =   1365
               Width           =   2760
            End
            Begin VB.OptionButton optState 
               Caption         =   "复诊"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   74
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "…"
               Height          =   255
               Index           =   23
               Left            =   6705
               TabIndex        =   86
               TabStop         =   0   'False
               ToolTipText     =   "选择(*)"
               Top             =   1065
               Width           =   285
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   24
               Left            =   1200
               MaxLength       =   100
               TabIndex        =   88
               Top             =   2025
               Width           =   5895
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   23
               Left            =   1200
               MaxLength       =   100
               TabIndex        =   85
               Top             =   1695
               Width           =   5895
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   20
               Left            =   4260
               MaxLength       =   100
               TabIndex        =   81
               Top             =   1035
               Width           =   2760
            End
            Begin VB.CheckBox chkEdit 
               Alignment       =   1  'Right Justify
               Caption         =   "传染病上传(&U)"
               Height          =   195
               Index           =   1
               Left            =   3600
               TabIndex        =   75
               Top             =   30
               Width           =   1470
            End
            Begin VB.OptionButton optState 
               Caption         =   "初诊"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   73
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin MSMask.MaskEdBox txt发病日期 
               Height          =   300
               Left            =   1200
               TabIndex        =   78
               Top             =   1035
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Format          =   "yyyy-MM-dd"
               Mask            =   "####-##-##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txt发病时间 
               Height          =   300
               Left            =   2295
               TabIndex        =   79
               Top             =   1035
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   529
               _Version        =   393216
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "HH:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label lblEdit 
               Caption         =   "其他医学警示"
               Height          =   180
               Index           =   33
               Left            =   45
               TabIndex        =   87
               Top             =   2100
               Width           =   1080
            End
            Begin VB.Label lblEdit 
               Caption         =   "医学警示"
               Height          =   180
               Index           =   34
               Left            =   390
               TabIndex        =   84
               Top             =   1770
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "去向"
               Height          =   180
               Index           =   32
               Left            =   750
               TabIndex        =   82
               Top             =   1425
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病地址"
               Height          =   180
               Index           =   27
               Left            =   3450
               TabIndex        =   80
               Top             =   1065
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病时间"
               Height          =   180
               Index           =   21
               Left            =   390
               TabIndex        =   77
               Top             =   1080
               Width           =   720
            End
         End
         Begin VB.Frame fraAller 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1580
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   7335
            Begin VB.OptionButton optAller 
               Caption         =   "根据药品目录输入(&1)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   0
               Left            =   2880
               TabIndex        =   100
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
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   90
               Width           =   1890
            End
            Begin VSFlex8Ctl.VSFlexGrid vsAller 
               Height          =   1260
               Left            =   120
               TabIndex        =   65
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmOutMedRecEdit.frx":0044
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
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   " 过敏记录 "
               Height          =   180
               Index           =   18
               Left            =   360
               TabIndex        =   64
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
            TabIndex        =   94
            Top             =   2500
            Width           =   7335
            Begin VB.OptionButton optInput 
               Caption         =   "根据诊断标准输入(&3)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   0
               Left            =   2820
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   90
               Value           =   -1  'True
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "根据疾病编码输入(&4)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   1
               Left            =   4890
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   90
               Width           =   2010
            End
            Begin VB.CommandButton cmdMakeLog 
               Height          =   255
               Left            =   1560
               Picture         =   "frmOutMedRecEdit.frx":00DB
               Style           =   1  'Graphical
               TabIndex        =   95
               TabStop         =   0   'False
               ToolTipText     =   "根据诊断生成就诊摘要(F12)"
               Top             =   53
               Width           =   345
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   1260
               Left            =   120
               TabIndex        =   71
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
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmOutMedRecEdit.frx":0665
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
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   960
               Left            =   120
               TabIndex        =   72
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
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmOutMedRecEdit.frx":078E
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
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   " 诊断记录 "
               Height          =   180
               Index           =   19
               Left            =   360
               TabIndex        =   68
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
      Begin VB.Frame fraInfo 
         BorderStyle     =   0  'None
         Height          =   7725
         Index           =   0
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   7425
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   9
            Left            =   4620
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   5108
            Width           =   2595
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   7
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   5108
            Width           =   2595
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   4905
            MaxLength       =   64
            TabIndex        =   101
            Top             =   4750
            Width           =   2310
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   255
            Index           =   9
            Left            =   6930
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3705
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
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
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   255
            Index           =   5
            Left            =   6930
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2625
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   240
            Index           =   19
            Left            =   6945
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2265
            Width           =   270
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   255
            Index           =   17
            Left            =   6930
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   4413
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   240
            Index           =   13
            Left            =   6945
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1905
            Width           =   270
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   19
            Left            =   4020
            MaxLength       =   30
            TabIndex        =   31
            Top             =   2235
            Width           =   3195
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   17
            Left            =   840
            MaxLength       =   100
            TabIndex        =   51
            Top             =   4390
            Width           =   6315
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   135
            Width           =   1200
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   0
            Left            =   900
            MaxLength       =   64
            TabIndex        =   2
            Top             =   135
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   3
            Left            =   3510
            MaxLength       =   10
            TabIndex        =   11
            Top             =   495
            Width           =   675
         End
         Begin VB.ComboBox cboEdit 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   3510
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   135
            Width           =   1305
         End
         Begin VB.ComboBox cboEdit 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   495
            Width           =   615
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   3
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1005
            Width           =   2355
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   4
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1005
            Width           =   3195
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   5
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1365
            Width           =   2355
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   6
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1365
            Width           =   3195
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   900
            MaxLength       =   18
            TabIndex        =   24
            Top             =   1875
            Width           =   2340
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   900
            MaxLength       =   100
            TabIndex        =   37
            Top             =   2955
            Width           =   6315
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   900
            MaxLength       =   20
            TabIndex        =   40
            Top             =   3315
            Width           =   3090
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   4905
            MaxLength       =   6
            TabIndex        =   42
            Top             =   3315
            Width           =   2310
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   9
            Left            =   900
            MaxLength       =   100
            TabIndex        =   44
            Top             =   3675
            Width           =   6315
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   10
            Left            =   900
            MaxLength       =   20
            TabIndex        =   47
            Top             =   4035
            Width           =   3090
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   4905
            MaxLength       =   6
            TabIndex        =   49
            Top             =   4035
            Width           =   2310
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   2
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   495
            Width           =   1215
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   900
            MaxLength       =   100
            TabIndex        =   34
            Top             =   2595
            Width           =   6315
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   13
            Left            =   4020
            MaxLength       =   30
            TabIndex        =   26
            Top             =   1875
            Width           =   3195
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   16
            Left            =   900
            MaxLength       =   20
            TabIndex        =   29
            Top             =   2235
            Width           =   2340
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            Left            =   900
            MaxLength       =   6
            TabIndex        =   54
            Top             =   4750
            Width           =   2595
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   8
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   5466
            Width           =   2595
         End
         Begin VB.ComboBox cboEdit 
            Height          =   300
            Index           =   10
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   5824
            Width           =   2595
         End
         Begin MSMask.MaskEdBox txt出生时间 
            Height          =   300
            Left            =   1950
            TabIndex        =   9
            Top             =   495
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   300
            Left            =   900
            TabIndex        =   8
            Top             =   495
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   -15
            X2              =   7335
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊号"
            Height          =   180
            Index           =   3
            Left            =   5400
            TabIndex        =   5
            Top             =   195
            Width           =   540
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   1
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblEdit 
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
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            Height          =   180
            Index           =   2
            Left            =   3090
            TabIndex        =   10
            Top             =   555
            Width           =   360
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻状况"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   19
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "职业"
            Height          =   180
            Index           =   9
            Left            =   3555
            TabIndex        =   21
            Top             =   1425
            Width           =   360
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "民族"
            Height          =   180
            Index           =   7
            Left            =   3555
            TabIndex        =   17
            Top             =   1065
            Width           =   360
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "国籍"
            Height          =   180
            Index           =   6
            Left            =   480
            TabIndex        =   15
            Top             =   1065
            Width           =   360
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证号"
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   23
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位名称"
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   36
            Top             =   3015
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位电话"
            Height          =   180
            Index           =   13
            Left            =   120
            TabIndex        =   39
            Top             =   3375
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位邮编"
            Height          =   180
            Index           =   14
            Left            =   4095
            TabIndex        =   41
            Top             =   3375
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭地址"
            Height          =   180
            Index           =   15
            Left            =   120
            TabIndex        =   43
            Top             =   3735
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭电话"
            Height          =   180
            Index           =   16
            Left            =   120
            TabIndex        =   46
            Top             =   4095
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭邮编"
            Height          =   180
            Index           =   17
            Left            =   4095
            TabIndex        =   48
            Top             =   4095
            Width           =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            X1              =   -60
            X2              =   7290
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            X1              =   -150
            X2              =   7200
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            X1              =   -105
            X2              =   7245
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款方式"
            Height          =   180
            Index           =   5
            Left            =   5220
            TabIndex        =   13
            Top             =   555
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生地点"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   33
            Top             =   2655
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   7
            Top             =   555
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "监护人"
            Height          =   180
            Index           =   22
            Left            =   4275
            TabIndex        =   55
            Top             =   4815
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "区域"
            Height          =   180
            Index           =   11
            Left            =   3555
            TabIndex        =   25
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl其它证件 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "其他证件"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   2295
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口地址"
            Height          =   180
            Index           =   25
            Left            =   120
            TabIndex        =   50
            Top             =   4455
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口邮编"
            Height          =   180
            Index           =   26
            Left            =   120
            TabIndex        =   53
            Top             =   4810
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "籍贯"
            Height          =   180
            Index           =   19
            Left            =   3555
            TabIndex        =   30
            Top             =   2295
            Width           =   360
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "文化程度"
            Height          =   180
            Index           =   28
            Left            =   120
            TabIndex        =   56
            Top             =   5168
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "生育状况"
            Height          =   180
            Index           =   29
            Left            =   120
            TabIndex        =   57
            Top             =   5526
            Width           =   720
         End
         Begin VB.Label lblEdit 
            Caption         =   "血型"
            Height          =   180
            Index           =   35
            Left            =   480
            TabIndex        =   59
            Top             =   5884
            Width           =   360
         End
         Begin VB.Label lblEdit 
            Caption         =   "Rh"
            Height          =   180
            Index           =   36
            Left            =   4335
            TabIndex        =   62
            Top             =   5190
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   -64800
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   6060
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmOutMedRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReadOnly As Boolean
Private mblnDiagnose As Boolean
Private mstrPrivs As String
Private mlng病人ID As Long
Private mstr挂号单 As String
Private mlng挂号ID As Long
Private mlng科室ID As Long
Private mint险类 As Integer
Private mint社区 As Integer
Private mstr社区号 As String
Private mbln中医 As Boolean
Private mstr疾病ID As String   '用于保存疾病ID,在FormClosed事件中传递给父窗体
Private mstr诊断ID As String   '用于保存诊断ID,在FormClosed事件中传递给父窗体

Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mbln过敏药物Edit As Boolean

Private mrsXYDiag  As ADODB.Recordset '西医诊断记录集
Private mrsZYDiag  As ADODB.Recordset '中医诊断记录集
Private mblnUseTYT As Boolean '使用太元通接口
Private mint过敏输入来源 As Integer '医生站的过敏输入来源

Private Enum TXT_ENUM
    txt姓名 = 0
    txt门诊号 = 1
    txt监护人 = 2
    txt年龄 = 3
    txt身份证号 = 4
    txt出生地点 = 5
    txt单位名称 = 6
    txt单位电话 = 7
    txt单位邮编 = 8
    txt家庭地址 = 9
    txt家庭电话 = 10
    txt家庭邮编 = 11
    txt就诊摘要 = 12
    txt区域 = 13
'    txt身高 = 14
'    txt体重 = 15
    txt其他证件 = 16
    txt户口地址 = 17
    txt户口地址邮编 = 18
    txt籍贯 = 19
    txt发病地址 = 20
'    txt收缩压 = 21
'    txt舒张压 = 22
    txt医学警示 = 23
    txt其他医学警示 = 24
'    txt体温 = 25
End Enum

Private Enum CBO_ENUM
    cbo性别 = 0
    cbo年龄 = 1
    cbo付款 = 2
    cbo国籍 = 3
    cbo民族 = 4
    cbo婚姻 = 5
    cbo职业 = 6
    cbo文化程度 = 7
    cbo生育状况 = 8
    cboRh = 9
    cbo血型 = 10
    cbo去向 = 11
End Enum

Private Enum CHK_ENUM
    chk传染病上传 = 1
End Enum

Private Enum OPT_ENUM
    opt初诊 = 0
    opt复诊 = 1
End Enum

Private Enum COL_ENUM
    col类型 = 0
    col编码 = 1
    col诊断 = 2
    col中医证候 = 3
    col发病时间 = 4
    col疑诊 = 5
    col诊断ID = 6
    col疾病ID = 7
    col证候ID = 8
    col医嘱ID = 9
End Enum

Private Enum AllerColS
    AC_过敏时间 = 0
    AC_过敏药物 = 1
    AC_过敏反应 = 2
    AC_过敏源编码 = 3
End Enum

Private mlngNum As Long
Private mlngSelNum As Long
Private mlngNumBack As Long
Private mstrEmail As String
Private mstrQQ As String

Public Function ShowMe(frmParent As Object, ByVal str挂号单 As String, ByVal strPrivs As String, Optional blnDiagnose As Boolean, Optional ByVal blnReadOnly As Boolean, _
Optional ByRef str疾病ID As String, Optional ByRef str诊断ID As String) As Boolean

'参数：blnDiagnose=是否调用用于填写诊断
'返回：blnDiagnose=是否填写了病人的诊断
    mblnReadOnly = blnReadOnly
    mblnDiagnose = blnDiagnose
    mstr挂号单 = str挂号单
    mstrPrivs = strPrivs
    
    mstr疾病ID = ""
    mstr诊断ID = ""

    On Error Resume Next
    Me.Show 1, frmParent
    str疾病ID = mstr疾病ID
    str诊断ID = mstr诊断ID

    On Error GoTo 0
    
    blnDiagnose = mblnDiagnose
    ShowMe = mblnOk
End Function

Private Sub SetFaceEditable(ByVal blnReadOnly As Boolean)
'功能：根据当前是否只读，设置界面的可编辑属性
    Dim objControl As Object

    For Each objControl In Me.Controls
        If InStr("TextBox;MaskEdBox;ComboBox;CheckBox;VSFlexGrid", TypeName(objControl)) > 0 Then
            'TabStop=False表示当前确实不可编辑的
            If objControl.Container.Name = "fraInfo" And objControl.TabStop = True Then
                If TypeName(objControl) = "TextBox" And objControl.Enabled Then
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                    objControl.Locked = blnReadOnly
                ElseIf TypeName(objControl) = "MaskEdBox" Then
                    '没有Locked属性,用Enabled实现
                    objControl.Enabled = Not blnReadOnly
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "ComboBox" And objControl.Enabled Then
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                    objControl.Locked = blnReadOnly
                ElseIf TypeName(objControl) = "CheckBox" Then
                    '没有Locked属性,用Enabled实现
                    objControl.Enabled = Not blnReadOnly
                ElseIf TypeName(objControl) = "VSFlexGrid" Then
                    '同时注意要在键盘鼠标事件中进行一些控制
                    objControl.Editable = IIf(blnReadOnly, flexEDNone, flexEDKbdMouse)
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                    objControl.BackColorBkg = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                End If
            End If
        End If
    Next
End Sub

Private Function InitMedData() As Boolean
'功能：初始化编辑环境和必要的数据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    Call zlControl.CboSetHeight(cboEdit(cbo民族), cboEdit(cbo民族).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cbo国籍), cboEdit(cbo国籍).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cbo职业), cboEdit(cbo职业).Height * 16)
    vsDiagXY.MergeCol(0) = True
    vsDiagZY.MergeCol(0) = True
    
    Call SetCboFromList(Array("岁", "月", "天"), cboEdit(cbo年龄), 0)
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 性别 Order by 编码", Array(cbo性别))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 医疗付款方式 Order by 编码", Array(cbo付款))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 民族 Order by 编码", Array(cbo民族))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 国籍 Order by 编码", Array(cbo国籍))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 婚姻状况 Order by 编码", Array(cbo婚姻))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 职业 Order by 编码", Array(cbo职业))
    
    strSQL = "Select 名称, 编码 From 病人去向"
    cboEdit(cbo去向).Clear
    cboEdit(cbo去向).AddItem ("")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cbo去向), rsTmp, False)
    
    Call SetCboFromList(Array("", "9-文盲和半文盲", "8-小学（包括村学）", "7-初中", "6-高中", "4-中专", "3-大专", "2-大学", "1-研究生及以上"), cboEdit(cbo文化程度), 0)
    Call SetCboFromList(Array("", "0-未生育", "1-生育1胎", "2-生育2胎及以上", "4-不详"), cboEdit(cbo生育状况), 0)
    Call SetCboFromList(Array("", "A型", "B型", "O型", "AB型", "不详"), cboEdit(cbo血型), 0)
    Call SetCboFromList(Array("", "阴", "阳", "不详", "未查"), cboEdit(cboRh), 0)
    optInput(0).TabStop = False: optInput(1).TabStop = False '要强行代码执行一次
    
    InitMedData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadMedRec() As Boolean
'功能：读取门诊首页的各种信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long
    
    On Error GoTo errH
    
    '基本信息
    strSQL = "Select A.病人ID,B.ID as 挂号ID,B.摘要,B.复诊,a.籍贯," & _
        " Nvl(Nvl(B.续诊科室ID,Decode(B.转诊状态,1,B.转诊科室ID,NULL)),B.执行部门ID) as 科室ID," & _
        " B.传染病上传,B.发病时间,B.发病地址,A.险类,A.门诊号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,A.出生日期,A.医疗付款方式," & _
        " A.国籍,A.民族,A.婚姻状况,A.职业,A.身份证号,A.出生地点,A.监护人,A.家庭地址,A.家庭电话," & _
        " A.区域,A.家庭地址邮编,A.工作单位,A.合同单位id,A.单位电话,A.单位邮编,B.社区,C.社区号,A.其他证件,A.户口地址,a.户口地址邮编,a.qq,a.email" & _
        " From 病人信息 A,病人挂号记录 B,病人社区信息 C" & _
        " Where A.病人ID=B.病人ID And B.病人ID=C.病人ID(+) And B.社区=C.社区(+) And B.NO=[1] And B.记录性质=1 And B.记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    If rsTmp.EOF Then Exit Function
    
    mlng病人ID = rsTmp!病人ID
    mlng挂号ID = rsTmp!挂号ID
    mlng科室ID = rsTmp!科室ID
    mint险类 = Nvl(rsTmp!险类, 0)
    mint社区 = Nvl(rsTmp!社区, 0)
    mstr社区号 = Nvl(rsTmp!社区号)
    mbln中医 = Have部门性质(rsTmp!科室ID, "中医科")
    mstrEmail = Nvl(rsTmp!Email)
    mstrQQ = Nvl(rsTmp!QQ)
        
    txtEdit(txt姓名).Text = rsTmp!姓名
    txtEdit(txt姓名).Tag = rsTmp!姓名 '记录原始的姓名
    Call GetCboIndex(cboEdit(cbo性别), Nvl(rsTmp!性别))
    txtEdit(txt门诊号).Text = Nvl(rsTmp!门诊号)
    
    If Not IsNull(rsTmp!出生日期) Then
        txt出生日期.Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
        If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
            txt出生时间.Text = Format(rsTmp!出生日期, "HH:mm")
        End If
    End If
        
    Call LoadOldData("" & rsTmp!年龄, txtEdit(txt年龄), cboEdit(cbo年龄))
    

    Call GetCboIndex(cboEdit(cbo付款), Nvl(rsTmp!医疗付款方式))
    Call GetCboIndex(cboEdit(cbo国籍), Nvl(rsTmp!国籍))
    Call GetCboIndex(cboEdit(cbo民族), Nvl(rsTmp!民族))
    Call GetCboIndex(cboEdit(cbo婚姻), Nvl(rsTmp!婚姻状况))
    Call GetCboIndex(cboEdit(cbo职业), Nvl(rsTmp!职业))
    txtEdit(txt区域).Text = Nvl(rsTmp!区域)
    txtEdit(txt籍贯).Text = Nvl(rsTmp!籍贯)
    txtEdit(txt监护人).Text = Nvl(rsTmp!监护人)
    txtEdit(txt身份证号).Text = Nvl(rsTmp!身份证号)
    txtEdit(txt其他证件).Text = Nvl(rsTmp!其他证件)
    txtEdit(txt出生地点).Text = Nvl(rsTmp!出生地点)
    txtEdit(txt单位名称).Text = Nvl(rsTmp!工作单位)
    txtEdit(txt单位名称).Tag = Val("" & rsTmp!合同单位id)
    If InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") = 0 And Not IsNull(rsTmp!合同单位id) Then
        txtEdit(txt单位名称).Enabled = False
        cmdEdit(txt单位名称).Enabled = False
    End If
    
    txtEdit(txt单位电话).Text = Nvl(rsTmp!单位电话)
    txtEdit(txt单位邮编).Text = Nvl(rsTmp!单位邮编)
    txtEdit(txt家庭地址).Text = Nvl(rsTmp!家庭地址)
    txtEdit(txt家庭电话).Text = Nvl(rsTmp!家庭电话)
    txtEdit(txt家庭邮编).Text = Nvl(rsTmp!家庭地址邮编)
    txtEdit(txt户口地址).Text = Nvl(rsTmp!户口地址)
    txtEdit(txt户口地址邮编).Text = Nvl(rsTmp!户口地址邮编)
    txtEdit(txt就诊摘要).Text = Nvl(rsTmp!摘要)
    If Nvl(rsTmp!复诊, 0) = 1 Then
        optState(opt复诊).Value = True
    End If
    chkEdit(chk传染病上传).Value = Nvl(rsTmp!传染病上传, 0)
    If Not IsNull(rsTmp!发病时间) Then
        txt发病日期.Text = Format(rsTmp!发病时间, "yyyy-MM-dd")
        txt发病时间.Text = Format(rsTmp!发病时间, "HH:mm")
        If txt发病时间.Text = "00:00" Then txt发病时间.Text = "__:__"
    End If
    txtEdit(txt发病地址).Text = Nvl(rsTmp!发病地址)
    
    '附加信息
    Call ucPatiVitalSigns.LoadPatiVitalSigns(mlng病人ID, mlng挂号ID)
    strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID=[1] And (就诊ID=[2] Or 就诊ID is Null) Order by Nvl(就诊ID,999999999)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    rsTmp.Filter = "信息名='文化程度'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cbo文化程度), Nvl(rsTmp!信息值))
    rsTmp.Filter = "信息名='生育状况'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cbo生育状况), Nvl(rsTmp!信息值))
    rsTmp.Filter = "信息名='去向'"
    If Not rsTmp.EOF Then cboEdit(cbo去向).Text = Nvl(rsTmp!信息值)
    rsTmp.Filter = "信息名='血型'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cbo血型), Nvl(rsTmp!信息值))
    rsTmp.Filter = "信息名='RH'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cboRh), Nvl(rsTmp!信息值))
    rsTmp.Filter = "信息名='医学警示'"
    If Not rsTmp.EOF Then txtEdit(txt医学警示).Text = Nvl(rsTmp!信息值)
    rsTmp.Filter = "信息名='其他医学警示'"
    If Not rsTmp.EOF Then txtEdit(txt其他医学警示).Text = Nvl(rsTmp!信息值)
    '过敏信息:本次挂号的,过敏的
    strSQL = "Select 记录来源,NVL(过敏时间,记录时间) as 过敏时间,药物ID,药物名,过敏反应,过敏源编码 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by NVL(过敏时间,记录时间),药物名"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 2 '固定行+新行
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , AC_过敏药物)
                End If
                If lngRow = -1 Then
                    .TextMatrix(i, AC_过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, AC_过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")  '用于保存
                    .TextMatrix(i, AC_过敏药物) = Nvl(rsTmp!药物名)
                    .Cell(flexcpData, i, AC_过敏药物) = .TextMatrix(i, AC_过敏药物) '用于输入恢复
                    .TextMatrix(i, AC_过敏反应) = Nvl(rsTmp!过敏反应)
                    .Cell(flexcpData, i, AC_过敏反应) = .TextMatrix(i, AC_过敏反应)   '用于输入恢复
                    .TextMatrix(i, AC_过敏源编码) = Nvl(rsTmp!过敏源编码)
                    .RowData(i) = Val(Nvl(rsTmp!药物ID, 0))
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    vsAller.Row = 1: vsAller.Col = AC_过敏药物
    vsAller.Tag = "未修改"
    
    '读取诊断信息
    Call LoadPatiDiag(False)
    
    If Not mbln中医 Then
        vsDiagZY.Visible = False
        vsDiagXY.Height = vsDiagZY.Top + vsDiagZY.Height - vsDiagXY.Top
        vsDiagXY.ColHidden(0) = True
        vsDiagXY.ColWidth(1) = vsDiagXY.ColWidth(1) + vsDiagXY.ColWidth(0)
    End If
    LoadMedRec = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPatiDiag(ByVal blnLast As Boolean) As Boolean
'功能：读取并显示病人诊断
'参数：blnLast=是否不读取本次就诊的诊断，而读取最后一次就诊的诊断
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    If blnLast Then
        strSQL = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1 " & _
                " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
        strSQL = "Select a.ID,a.记录来源,a.诊断类型,a.疾病ID,a.诊断ID,a.证候ID,a.诊断描述,a.是否疑诊, a.记录日期, a.记录人,b.编码 as 疾病编码,c.编码 as 诊断编码,d.编码 as 证候编码,A.发病时间,A.诊断次序, " & _
                " (Select f_List2str(Cast(Collect(c.医嘱id || '') As t_Strlist)) 医嘱id From 病人诊断医嘱 C,病人医嘱记录 D where c.医嘱id=d.id  and c.诊断id=a.id And d.医嘱状态<>-1 And D.医嘱状态<>4) as 医嘱ID  From 病人诊断医嘱 C where c.诊断ID=A.ID ) 医嘱ID  " & _
                "From 病人诊断记录  A, 疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
                " Where  a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.证候ID=d.ID(+) And a.记录来源 IN(1,3) And a.诊断类型 IN(1,11)" & _
                " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=(" & strSQL & ")" & _
                " Order by a.诊断类型,a.诊断次序,a.编码序号"
    Else
        strSQL = "Select a.ID,a.记录来源,a.诊断类型,a.疾病ID,a.诊断ID,a.证候ID,a.诊断描述,a.是否疑诊, a.记录日期, a.记录人,b.编码 as 疾病编码,c.编码 as 诊断编码,d.编码 as 证候编码,A.发病时间,A.诊断次序, " & _
                " (Select f_List2str(Cast(Collect(c.医嘱id || '') As t_Strlist)) 医嘱id From 病人诊断医嘱 C,病人医嘱记录 D where c.医嘱id=d.id and c.诊断id=a.id And d.医嘱状态<>-1 And D.医嘱状态<>4) as 医嘱ID " & _
                " From 病人诊断记录  A, 疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
                " Where  a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.证候ID=d.ID(+) And a.记录来源 IN(1,3) And a.诊断类型 IN(1,11)" & _
                " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=[2]" & _
                " Order by a.诊断类型,a.诊断次序,a.编码序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    
    If Not rsTmp.EOF Then
        '西医诊断
        rsTmp.Filter = "诊断类型=1 And 记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "诊断类型=1 And 记录来源<>3": '其它来源的作为缺省显示
        With vsDiagXY
            Set mrsXYDiag = zlDatabase.CopyNewRec(rsTmp)
            .Rows = rsTmp.RecordCount + 2
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!诊断描述) Then
                    .TextMatrix(i, col编码) = ""
                    .TextMatrix(i, col诊断) = ""
                Else
                    If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                        '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                        If Val(rsTmp!疾病id & "") <> 0 Then
                            .TextMatrix(i, col编码) = Nvl(rsTmp!疾病编码)
                        ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                            .TextMatrix(i, col编码) = Nvl(rsTmp!诊断编码)
                        Else
                            .TextMatrix(i, col编码) = ""
                        End If
                        .TextMatrix(i, col诊断) = rsTmp!诊断描述
                    Else
                        .TextMatrix(i, col编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                        .TextMatrix(i, col诊断) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                    End If
                End If
                If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                    .Cell(flexcpData, i, col诊断) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                Else
                    .Cell(flexcpData, i, col诊断) = .TextMatrix(i, col诊断)
                End If
                .TextMatrix(i, col疑诊) = IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                .Cell(flexcpData, i, col疑诊) = Val(rsTmp!ID & "")
                .TextMatrix(i, col诊断ID) = Nvl(rsTmp!诊断id, 0)
                .TextMatrix(i, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                .TextMatrix(i, col医嘱ID) = rsTmp!医嘱ID & ""
                .TextMatrix(i, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                If Val(rsTmp!诊断次序 & "") = 1 And .TextMatrix(i, col发病时间) <> "" Then
                    '如果填写了发病时间，则下面的发病时间则不允许填写了
                    txt发病日期.BackColor = vbButtonFace
                    txt发病日期.Enabled = False
                    txt发病时间.BackColor = vbButtonFace
                    txt发病时间.Enabled = False
                End If
                rsTmp.MoveNext
            Next
            .Cell(flexcpText, .FixedRows, col类型, .Rows - 1, col类型) = "西医"
            .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
        End With
        '中医诊断
        If mbln中医 Then
            rsTmp.Filter = "诊断类型=11 And 记录来源=3"
            If rsTmp.EOF Then rsTmp.Filter = "诊断类型=11 And 记录来源<>3"
            With vsDiagZY
                Set mrsZYDiag = zlDatabase.CopyNewRec(rsTmp)
                .Rows = rsTmp.RecordCount + 1
                For i = 0 To rsTmp.RecordCount - 1
                    If IsNull(rsTmp!诊断描述) Then
                        .TextMatrix(i, col编码) = ""
                        .TextMatrix(i, col诊断) = ""
                    Else
                        If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                            '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                            If Val(rsTmp!疾病id & "") <> 0 Then
                                .TextMatrix(i, col编码) = Nvl(rsTmp!疾病编码)
                            ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                                .TextMatrix(i, col编码) = Nvl(rsTmp!诊断编码)
                            Else
                                .TextMatrix(i, col编码) = ""
                            End If
                            .TextMatrix(i, col诊断) = rsTmp!诊断描述
                        Else
                            .TextMatrix(i, col编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                            .TextMatrix(i, col诊断) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                        End If
                    End If

                    .TextMatrix(i, col疑诊) = IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                    .Cell(flexcpData, i, col疑诊) = Val(rsTmp!ID & "")
                    .TextMatrix(i, col诊断ID) = Nvl(rsTmp!诊断id, 0)
                    .TextMatrix(i, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                    .TextMatrix(i, col证候ID) = Nvl(rsTmp!证候id, 0)
                    .TextMatrix(i, col医嘱ID) = rsTmp!医嘱ID & ""
                    .TextMatrix(i, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                    If Val(rsTmp!诊断次序 & "") = 1 And .TextMatrix(i, col发病时间) <> "" Then
                        '如果填写了发病时间，则下面的发病时间则不允许填写了
                        txt发病日期.BackColor = vbButtonFace
                        txt发病日期.Enabled = False
                        txt发病时间.BackColor = vbButtonFace
                        txt发病时间.Enabled = False
                    End If
                    '取证候名称
                    If InStr(.TextMatrix(i, col诊断), "(") > 0 And InStr(.TextMatrix(i, col诊断), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(i, col诊断), InStrRev(.TextMatrix(i, col诊断), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '先取证候
                        .TextMatrix(i, col中医证候) = strTmp
                        '去掉诊断描述的证候
                        .TextMatrix(i, col诊断) = Mid(.TextMatrix(i, col诊断), 1, InStrRev(.TextMatrix(i, col诊断), "(") - 1)
                    Else
                       .TextMatrix(i, col中医证候) = ""
                    End If
                    
                    '自由录入诊断的诊断描述，需要去掉证候，因此此句代码后移
                    If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                        .Cell(flexcpData, i, col诊断) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                    Else
                        .Cell(flexcpData, i, col诊断) = .TextMatrix(i, col诊断)
                    End If
                    rsTmp.MoveNext
                Next
                .Cell(flexcpText, .FixedRows, col类型, .Rows - 1, col类型) = "中医"
                .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
            End With
        End If
    End If
    vsDiagXY.Row = vsDiagXY.FixedRows: vsDiagXY.Col = 0: vsDiagXY.Col = col诊断
    vsDiagZY.Row = vsDiagZY.FixedRows: vsDiagZY.Col = 0: vsDiagZY.Col = col诊断
    vsDiagXY.Tag = "未修改"
    vsDiagZY.Tag = "未修改"
    
    LoadPatiDiag = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMedRec(Optional blnDiagnose As Boolean) As Boolean
'功能：检查首页输入数据合法性
'返回：blnDiagnose=是否填写了诊断
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str身份证 As String, str出生日期 As String, lng性别 As Long
    Dim str年龄 As String, i As Long, j As Long, k As Long
    Dim str疾病IDs As String, str诊断IDs As String
    
    blnDiagnose = False
    curDate = zlDatabase.Currentdate
    
    '必须要输入的内容检查
    '-----------------------------------------------------------------------------------------
    If InStr(mstrPrivs, "修改基本信息") > 0 Then
        arrInfo = Array(cbo付款)
        arrName = Array("付款方式")
        For i = 0 To UBound(arrInfo)
            If cboEdit(arrInfo(i)).Enabled And Not cboEdit(arrInfo(i)).Locked And cboEdit(arrInfo(i)).ListIndex = -1 Then
                Call ShowMessage(cboEdit(arrInfo(i)), "必须输入病人的" & arrName(i) & "。")
                Exit Function
            End If
        Next
    End If
        
    '项目输入的长度检查
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "输入内容过长，请检查。(该项目最多允许 " & objTmp.MaxLength & " 个字符或 " & objTmp.MaxLength \ 2 & " 个汉字)")
                Exit Function
            End If
        End If
    Next

    '输入内容的有效性检查
    '-----------------------------------------------------------------------------------------
    '15岁以下应为未婚
    If Not (cboEdit(cbo婚姻).Text = "" Or cboEdit(cbo婚姻).ListIndex = -1) And IsDate(txt出生日期.Text) Then
        If DateDiff("yyyy", CDate(txt出生日期.Text), curDate) < 15 Then
            If InStr(cboEdit(cbo婚姻).Text, "已婚") > 0 _
                Or InStr(cboEdit(cbo婚姻).Text, "丧偶") > 0 Or InStr(cboEdit(cbo婚姻).Text, "离婚") > 0 Then
                Call ShowMessage(cboEdit(cbo婚姻), "该病人年龄太小，当前填写的婚姻状况信息不适合。")
                Exit Function
            End If
        End If
    End If
            
    '身份证号码检查
    '对身份证号进行验证
    str身份证 = txtEdit(txt身份证号).Text
    If str身份证 <> "" Then
        If Len(str身份证) <> 15 And Len(str身份证) <> 18 Then
            Call ShowMessage(txtEdit(txt身份证号), "身份证号码的长度不正确，应为15位或18位。")
            Exit Function
        End If

        If Len(str身份证) = 15 Then
            str出生日期 = Mid(str身份证, 7, 6)
            str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
            lng性别 = Val(Right(str身份证, 1))
        Else
            str出生日期 = Mid(str身份证, 7, 8)
            str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
            lng性别 = Val(Mid(str身份证, 17, 1))
        End If
        If Not IsDate(str出生日期) Then
            If ShowMessage(txtEdit(txt身份证号), "身份证号码中的出生日期信息不正确，是否继续？", True) = vbNo Then Exit Function
        ElseIf IsDate(txt出生日期.Text) Then
            If Format(str出生日期, "yyyy-MM-dd") <> Format(txt出生日期.Text, "yyyy-MM-dd") Then
                If ShowMessage(txtEdit(txt身份证号), "身份证号码中的出生日期信息与病人的出生日期不符，是否继续？", True) = vbNo Then Exit Function
            End If
        End If
        If (lng性别 Mod 2 = 1 And InStr(cboEdit(cbo性别).Text, "女") > 0) Or (lng性别 Mod 2 = 0 And InStr(cboEdit(cbo性别).Text, "男") > 0) Then
            If ShowMessage(txtEdit(txt身份证号), "身份证号码中的性别信息与病人的性别不符，是否继续？", True) = vbNo Then Exit Function
        End If
    End If
    
    '诊断表格的检查
    '-----------------------------------------------------------------------------------------
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col诊断)) <> "" Then
                If mint险类 = 920 Then '北京医保无理要求
                    If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 82 Then
                        .Row = i: .Col = col诊断
                        Call ShowMessage(vsDiagXY, "诊断内容太长，只允许82个字符或41个汉字。")
                        Exit Function
                    End If
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 200 Then
                    .Row = i: .Col = col诊断
                    Call ShowMessage(vsDiagXY, "诊断内容太长，只允许200个字符或100个汉字。")
                    Exit Function
                End If
                If .TextMatrix(i, col发病时间) <> "" Then
                    If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col发病时间), "YYYY-MM-DD HH:mm") Then
                         .Row = i: .Col = col发病时间
                        Call ShowMessage(vsDiagXY, "发病时间应该早于当前时间。")
                        Exit Function
                    End If
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, col诊断)) <> "" Then
                        If .TextMatrix(j, col诊断) = .TextMatrix(i, col诊断) Then
                            .Row = i: .Col = col诊断
                            Call ShowMessage(vsDiagXY, "发现存在两行相同的诊断信息。")
                            Exit Function
                        ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                            If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                .Row = i: .Col = col诊断
                                Call ShowMessage(vsDiagXY, "发现存在两行相同的诊断信息。")
                                Exit Function
                            End If
                        ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 Then
                            If Val(.TextMatrix(j, col诊断ID)) = Val(.TextMatrix(i, col诊断ID)) Then
                                .Row = i: .Col = col诊断
                                Call ShowMessage(vsDiagXY, "发现存在两行相同的诊断信息。")
                                Exit Function
                            End If
                        End If
                    End If
                Next
                
                If Val(.TextMatrix(i, col疾病ID)) <> 0 Then str疾病IDs = str疾病IDs & "," & Val(.TextMatrix(i, col疾病ID))
                If Val(.TextMatrix(i, col诊断ID)) <> 0 Then str诊断IDs = str诊断IDs & "," & Val(.TextMatrix(i, col诊断ID))
                
                blnDiagnose = True
            End If
        Next
    End With
        
    If mbln中医 Then
        With vsDiagZY
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col诊断)) <> "" Then
                    If mint险类 = 920 Then '北京医保无理要求
                        If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 82 Then
                            .Row = i: .Col = col诊断
                            Call ShowMessage(vsDiagZY, "诊断内容太长，只允许82个字符或41个汉字。")
                            Exit Function
                        End If
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 200 Then
                        .Row = i: .Col = col诊断
                        Call ShowMessage(vsDiagZY, "诊断内容太长，只允许200个字符或100个汉字。")
                        Exit Function
                    End If
                    If .TextMatrix(i, col发病时间) <> "" Then
                        If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col发病时间), "YYYY-MM-DD HH:mm") Then
                             .Row = i: .Col = col发病时间
                            Call ShowMessage(vsDiagZY, "发病时间应该早于当前时间。")
                            Exit Function
                        End If
                    End If
                    For j = i + 1 To .Rows - 1
                        If Trim(.TextMatrix(j, col诊断)) <> "" Then
                            If .TextMatrix(j, col诊断) = .TextMatrix(i, col诊断) Then
                                .Row = i: .Col = col诊断
                                Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
                                Exit Function
                            ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                                If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                    .Row = i: .Col = col诊断
                                    Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
                                    Exit Function
                                End If
                            ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 Then
                                '因中医诊断带证候,可能无对应证候ID,诊断ID又相同
'                                If Val(.TextMatrix(j, col诊断ID)) & "," & Val(.TextMatrix(j, col证候ID)) _
'                                    = Val(.TextMatrix(i, col诊断ID)) & "," & Val(.TextMatrix(i, col证候ID)) Then
'                                    .Row = i: .Col = col诊断
'                                    Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
'                                    Exit Function
'                                End If
                            End If
                        End If
                    Next
                     '中医诊断和西医诊断的自由录入医嘱不能存在相同的
                    If .TextMatrix(i, col编码) = "" Then
                        For k = vsDiagXY.FixedRows To vsDiagXY.Rows - 1
                            If Trim(vsDiagXY.TextMatrix(k, col诊断)) <> "" And vsDiagXY.TextMatrix(k, col编码) = "" Then
                                If vsDiagXY.TextMatrix(k, col诊断) = .TextMatrix(i, col诊断) Then
                                    .Row = i: .Col = col诊断
                                    Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
                                    Exit Function
                                End If
                            End If
                        Next
                    End If
                    If Val(.TextMatrix(i, col疾病ID)) <> 0 Then str疾病IDs = str疾病IDs & "," & Val(.TextMatrix(i, col疾病ID))
                    If Val(.TextMatrix(i, col诊断ID)) <> 0 Then str诊断IDs = str诊断IDs & "," & Val(.TextMatrix(i, col诊断ID))
                    
                    blnDiagnose = True
                End If
            Next
        End With
        
        
    End If
    
    '过敏药物表格检查
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, 1)) > 60 Then
                    .Row = i: .Col = 1
                    Call ShowMessage(vsAller, "过敏药物名太长，只允许60个字符或30个汉字。")
                    Exit Function
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, 2)) > 100 Then
                    .Row = i: .Col = 2
                    Call ShowMessage(vsAller, "过敏反应内容太长，只允许100个字符或50个汉字。")
                    Exit Function
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, 1)) <> "" Then
                        If .TextMatrix(j, 1) = .TextMatrix(i, 1) Then
                            .Row = i: .Col = 1
                            Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                            Exit Function
                        ElseIf .RowData(i) <> 0 Then
                            If .RowData(j) = .RowData(i) Then
                                .Row = i: .Col = 1
                                Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    '发病时间检查
    If txt发病日期.Text <> "____-__-__" Then
        If Not IsDate(txt发病日期.Text) Then
            Call ShowMessage(txt发病日期, "请输入正确的发病日期。")
            Exit Function
        Else
            If txt发病时间.Text <> "__:__" Then
                If Not IsDate(txt发病时间.Text) Then
                    Call ShowMessage(txt发病时间, "请输入正确的发病时间。")
                    Exit Function
                End If
            End If
            
            If txt发病日期.Text & IIf(txt发病时间.Text = "__:__", "", " " & txt发病时间.Text) _
                > Format(curDate, txt发病日期.Format & IIf(txt发病时间.Text = "__:__", "", " " & txt发病时间.Format)) Then
                Call ShowMessage(txt发病日期, "发病时间应该早于当前时间。")
                Exit Function
            End If
        End If
    End If
    
    mstr疾病ID = Mid(str疾病IDs, 2)
    mstr诊断ID = Mid(str诊断IDs, 2)
    
    CheckMedRec = True
End Function

Private Function SaveMedRec() As Boolean
'功能：保存门诊首页的各种信息
    Dim arrSQL As Variant, i As Integer
    Dim curDate As Date, intIdx As Integer
    Dim str生日 As String, str发病 As String
    Dim lng单位ID As Long
    Dim blnTrans As Boolean
    Dim str生育状况 As String
    Dim str文化程度 As String
    Dim blnDiagChange As Boolean
    Dim strFilter As String, strTmp As String
    Dim str关联医嘱ID As String
    
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    
    If IsDate(txt出生日期.Text) Then
        If IsDate(txt出生时间.Text) Then
            str生日 = "To_Date('" & Format(txt出生日期.Text & " " & txt出生时间.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str生日 = "To_Date('" & Format(txt出生日期.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        End If
    Else
       str生日 = "NULL"
    End If
    
    If Trim(txtEdit(txt单位名称).Text) <> "" Then
        lng单位ID = Val(txtEdit(txt单位名称).Tag)
    End If
    
    '病人信息
    str发病 = "NULL"
    If IsDate(txt发病日期.Text) Then
        If IsDate(txt发病时间.Text) Then
            str发病 = "To_Date('" & txt发病日期.Text & " " & txt发病时间.Text & "','YYYY-MM-DD HH24:MI')"
        Else
            str发病 = "To_Date('" & txt发病日期.Text & "','YYYY-MM-DD')"
        End If
    End If
    '文化程度
    If cboEdit(cbo文化程度).ListIndex > 0 Then
        str文化程度 = Mid(cboEdit(cbo文化程度), 1, InStr(cboEdit(cbo文化程度), "-") - 1)
    End If
    '生育状况
    If cboEdit(cbo生育状况).ListIndex > 0 Then
        str生育状况 = Mid(cboEdit(cbo生育状况), 1, InStr(cboEdit(cbo生育状况), "-") - 1)
    End If
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人信息_首页整理(" & _
        mlng病人ID & ",'" & txtEdit(txt门诊号).Text & "','" & txtEdit(txt姓名).Text & "'," & _
        "'" & NeedName(cboEdit(cbo性别).Text) & "','" & txtEdit(txt年龄).Text & cboEdit(cbo年龄).Text & "'," & _
        str生日 & ",'" & txtEdit(txt出生地点).Text & "','" & txtEdit(txt身份证号).Text & "'," & _
        "'" & NeedName(cboEdit(cbo民族).Text) & "','" & NeedName(cboEdit(cbo国籍).Text) & "','" & txtEdit(txt区域).Text & "'," & _
        "'" & NeedName(cboEdit(cbo婚姻).Text) & "','" & NeedName(cboEdit(cbo职业).Text) & "'," & _
        "'" & NeedName(cboEdit(cbo付款).Text) & "','" & txtEdit(txt家庭地址).Text & "'," & _
        "'" & txtEdit(txt家庭电话).Text & "','" & txtEdit(txt家庭邮编).Text & "'," & _
        "'" & txtEdit(txt单位名称).Text & "','" & txtEdit(txt单位电话).Text & "'," & _
        "'" & txtEdit(txt单位邮编).Text & "',Null,Null,Null,Null,'" & txtEdit(txt监护人).Text & "','" & mstr挂号单 & "'," & _
        IIf(optState(opt复诊).Value, 1, 0) & ",'" & txtEdit(txt就诊摘要).Text & "'," & chkEdit(chk传染病上传).Value & "," & str发病 & ",'" & _
        Trim(txtEdit(txt其他证件).Text) & "'," & ZVal(lng单位ID) & ",'" & txtEdit(txt户口地址).Text & "','" & txtEdit(txt户口地址邮编).Text & "','" & _
        txtEdit(txt籍贯).Text & "','" & mstrEmail & "','" & mstrQQ & "','" & txtEdit(txt发病地址).Text & "')"
    
    '附加信息
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = ucPatiVitalSigns.GetSaveSQL(mlng病人ID, mlng挂号ID)
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'文化程度','" & str文化程度 & "'," & mlng挂号ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'生育状况','" & str生育状况 & "'," & mlng挂号ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'去向','" & cboEdit(cbo去向).Text & "'," & mlng挂号ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'RH','" & cboEdit(cboRh).Text & "'," & mlng挂号ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'血型','" & cboEdit(cbo血型).Text & "'," & mlng挂号ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'医学警示','" & txtEdit(txt医学警示).Text & "'," & mlng挂号ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mlng病人ID & ",'其他医学警示','" & txtEdit(txt其他医学警示).Text & "'," & mlng挂号ID & ")"
    '过敏药物
    If vsAller.Tag = "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3)"
        With vsAller
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, AC_过敏药物)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = _
                        "zl_病人过敏记录_Insert(" & mlng病人ID & "," & mlng挂号ID & "," & _
                        "3," & ZVal(.RowData(i)) & ",'" & .TextMatrix(i, AC_过敏药物) & "',1," & _
                        "To_Date('" & .Cell(flexcpData, i, AC_过敏时间) & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & .TextMatrix(i, AC_过敏反应) & "','" & .TextMatrix(i, AC_过敏源编码) & "')"
                End If
            Next
        End With
    End If
    
    '诊断记录
    If mbln中医 And vsDiagZY.Tag = "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,Null,'11')"
    End If
    If vsDiagXY.Tag = "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,Null,'1')"
        With vsDiagXY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col诊断)) <> "" Then
                    blnDiagChange = True
                    If Val(.Cell(flexcpData, i, col疑诊) & "") > 0 Then
                        strFilter = "诊断类型=1 And 记录来源=3 And 疾病id=" & ZVal(.TextMatrix(i, col疾病ID)) & " And 诊断id=" & ZVal(.TextMatrix(i, col诊断ID))

                        strTmp = IIf(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断)
                        strFilter = strFilter & " And 诊断描述= '" & strTmp & "'"
                        If IsDate(.TextMatrix(i, col发病时间)) Then
                            strFilter = strFilter & " And  发病时间= '" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "'"
                        Else
                            strFilter = strFilter & " And  发病时间= Null "
                        End If
                        
                        strFilter = strFilter & " And 是否疑诊=" & IIf(.TextMatrix(i, col疑诊) = "", 0, 1)
                        mrsXYDiag.Filter = strFilter
                        blnDiagChange = mrsXYDiag.EOF
                    End If
                     str关联医嘱ID = IIf(.TextMatrix(i, col医嘱ID) = "", "Null", "'" & .TextMatrix(i, col医嘱ID) & "'")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    If blnDiagChange Then
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                            " Null,1," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & ",Null," & _
                            "'" & IIf(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & "',Null,Null," & IIf(.TextMatrix(i, col疑诊) = "", 0, 1) & "," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str关联医嘱ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & UserInfo.姓名 & "')"
                    Else
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                            " Null,1," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & ",Null," & _
                            "'" & IIf(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & "',Null,Null," & IIf(.TextMatrix(i, col疑诊) = "", 0, 1) & "," & _
                            "To_Date('" & Format(CDate(mrsXYDiag!记录日期), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str关联医嘱ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & mrsXYDiag!记录人 & "')"
                    
                    End If
                End If
            Next
        End With
    End If
    
    If mbln中医 And vsDiagZY.Tag = "" Then
        With vsDiagZY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col诊断)) <> "" Then
                    blnDiagChange = True
                    If Val(.Cell(flexcpData, i, col疑诊) & "") > 0 Then
                        strFilter = "诊断类型=11 And 记录来源=3 And 疾病id=" & ZVal(.TextMatrix(i, col疾病ID)) & " And 诊断id=" & ZVal(.TextMatrix(i, col诊断ID))

                        strTmp = IIf(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & "(" & .TextMatrix(i, col中医证候) & ")"
                        strFilter = strFilter & " And 诊断描述= '" & strTmp & "'"

                        strFilter = strFilter & " And  证候ID= " & ZVal(.TextMatrix(i, col证候ID))
                        If IsDate(.TextMatrix(i, col发病时间)) Then
                            strFilter = strFilter & " And  发病时间= '" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "'"
                        Else
                            strFilter = strFilter & " And  发病时间= Null "
                        End If
                        
                        strFilter = strFilter & " And 是否疑诊=" & IIf(.TextMatrix(i, col疑诊) = "", 0, 1)
                        mrsXYDiag.Filter = strFilter
                        blnDiagChange = mrsZYDiag.EOF
                    End If
                    
                    str关联医嘱ID = IIf(.TextMatrix(i, col医嘱ID) = "", "Null", "'" & .TextMatrix(i, col医嘱ID) & "'")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    If blnDiagChange Then
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                            "Null,11," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & "," & _
                            ZVal(.TextMatrix(i, col证候ID)) & ",'" & IIf(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & "(" & .TextMatrix(i, col中医证候) & ")" & "',Null,Null," & _
                            IIf(.TextMatrix(i, col疑诊) = "", 0, 1) & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str关联医嘱ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & UserInfo.姓名 & "')"
                    Else
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                            "Null,11," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & "," & _
                            ZVal(.TextMatrix(i, col证候ID)) & ",'" & IIf(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & "(" & .TextMatrix(i, col中医证候) & ")" & "',Null,Null," & _
                            IIf(.TextMatrix(i, col疑诊) = "", 0, 1) & ",To_Date('" & Format(CDate(mrsZYDiag!记录日期), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str关联医嘱ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & mrsZYDiag!记录人 & "')"
                    
                    End If
                End If
            Next
        End With
    End If
    
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    
    '社区档案同步
    If Not gobjCommunity Is Nothing And mint社区 <> 0 Then
        If Not gobjCommunity.UpdateInfo(glngSys, p门诊医生站, mint社区, mstr社区号, mlng病人ID, mlng挂号ID) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnChange = False
    SaveMedRec = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cboEdit_Click(Index As Integer)
    Dim strTmp As String
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
End Sub

Private Sub cboEdit_GotFocus(Index As Integer)
    If cboEdit(Index).Style = 0 Then
        Call zlControl.TxtSelAll(cboEdit(Index))
    End If
End Sub

Private Sub cboEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngidx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        lngidx = zlControl.CboMatchIndex(cboEdit(Index).hwnd, KeyAscii)
        If lngidx = -1 And cboEdit(Index).ListCount > 0 Then lngidx = 0
        cboEdit(Index).ListIndex = lngidx
    End If
End Sub

Private Sub chkEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click(Index As Integer)
'说明：注意界面上要求CMD和对应TXT的Index相同
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, blnLevel As Boolean
    Dim strResult As String
    
    '使用Lock的方式,不采用Enabled的方式
    If Not cmdEdit(Index).Enabled Or txtEdit(Index).Locked Then
        If txtEdit(Index).Enabled Then txtEdit(Index).SetFocus
        Exit Sub
    End If
    
    Select Case Index
        Case txt出生地点, txt家庭地址, txt户口地址
            '选择地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""地区""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!名称
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt单位名称
            '选择单位信息
            strSQL = "Select ID,上级ID,末级,编码,名称,简码,地址,电话,开户银行,帐号,联系人" & _
                " From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "合约单位", , , , , True, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""合约单位""数据，请先到合约单位管理中设置。", vbInformation, gstrSysName
                End If
                txtEdit(Index).Tag = ""
                If txtEdit(Index).Enabled Then txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") > 0 Then txtEdit(Index).Tag = Val(rsTmp!ID)
                If txtEdit(txt单位电话).Text = "" Then
                    txtEdit(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
                If txtEdit(Index).Enabled Then txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt区域, txt籍贯
            '选择区域数据
            strSQL = "Select Nvl(级数,0) as 级数 From 区域 Group by Nvl(级数,0)"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.RecordCount > 1 Then blnLevel = True
            
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            If blnLevel Then
                strSQL = _
                    " Select ID,上级id,编码,名称,简码,末级" & _
                    " From (Select 编码 As ID,RPad(Substr(编码,1,Decode(Nvl(级数,0),0,0,1,2,4)),6,'0') As 上级id," & _
                    "       编码,名称,简码,Decode(Nvl(级数,0),2,1,3,1,0) as 末级" & _
                    "       From 区域 Order By 编码)" & _
                    " Start With 上级ID Is Null Connect By Prior ID=上级id"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "区域", , , , , , , vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            Else
                strSQL = "Select Rownum as ID,编码,名称,简码 From 区域 Order by 编码"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""区域""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!名称
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt医学警示
            '选择医学警示
            On Error GoTo errH
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            strSQL = "Select Rownum ID,编码,名称,简码 From 医学警示 Order by 编码"
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, True, True)

            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""医学警示""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                While Not rsTmp.EOF
                    strResult = strResult & "," & rsTmp!名称
                    rsTmp.MoveNext
                Wend
                txtEdit(Index).Text = Mid(strResult, 2)
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdMakeLog_Click()
    Dim strLog As String, i As Long
    
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col诊断) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, col诊断) & IIf(.TextMatrix(i, col疑诊) <> "", "(？)", "")
            End If
        Next
    End With
    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col诊断) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, col诊断) & IIf(.TextMatrix(i, col疑诊) <> "", "(？)", "")
            End If
        Next
    End With
    If strLog <> "" Then
        If txtEdit(txt就诊摘要).SelStart = 0 And txtEdit(txt就诊摘要).SelLength = Len(txtEdit(txt就诊摘要).Text) Then
            txtEdit(txt就诊摘要).SelStart = Len(txtEdit(txt就诊摘要).Text)
        End If
        i = txtEdit(txt就诊摘要).SelStart
        txtEdit(txt就诊摘要).SelText = Mid(strLog, 2)
        txtEdit(txt就诊摘要).SelStart = i
        txtEdit(txt就诊摘要).SelLength = Len(Mid(strLog, 2))
    End If
    
    txtEdit(txt就诊摘要).SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim blnDiagnose As Boolean
    
    If Not CheckMedRec(blnDiagnose) Then Exit Sub
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("病人的诊断信息还没有输入，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    If Not SaveMedRec Then Exit Sub
        
    mblnDiagnose = blnDiagnose
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnDiagnose Then
        On Error Resume Next
        vsDiagXY.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdMakeLog_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOk = False
        
    optInput(Val(zlDatabase.GetPara("门诊诊断输入", glngSys, p门诊医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0))).Value = True
    
    '诊断输入来源
    If gint诊断来源 > 1 Then
        optInput(0).Enabled = False
        optInput(1).Enabled = False
        If gint诊断来源 = 2 Then
            optInput(0).Value = True
        ElseIf gint诊断来源 = 3 Then
            optInput(1).Value = True
        End If
    End If
    
    If Not InitMedData Then Unload Me: Exit Sub
    If Not LoadMedRec Then Unload Me: Exit Sub
    If mblnReadOnly Then
        Call SetFaceEditable(True)
        cmdOK.Visible = False
    Else
        If InStr(mstrPrivs, "修改基本信息") = 0 Then

            If cboEdit(cbo付款).ListIndex <> -1 Then
                cboEdit(cbo付款).BackColor = Me.BackColor: cboEdit(cbo付款).Locked = True: cboEdit(cbo付款).TabStop = False
            End If
        End If
    End If
    '病人基本信息：姓名，性别，年龄，出生日期不允许修改
    txtEdit(txt姓名).BackColor = Me.BackColor: txtEdit(txt姓名).Locked = True: txtEdit(txt姓名).TabStop = False
    cboEdit(cbo性别).BackColor = Me.BackColor: cboEdit(cbo性别).Locked = True: cboEdit(cbo性别).TabStop = False
    txt出生日期.BackColor = Me.BackColor: txt出生日期.Enabled = False
    txt出生时间.BackColor = Me.BackColor: txt出生时间.Enabled = False
    txtEdit(txt年龄).BackColor = Me.BackColor: txtEdit(txt年龄).Locked = True: txtEdit(txt年龄).TabStop = False
    cboEdit(cbo年龄).BackColor = Me.BackColor: cboEdit(cbo年龄).Locked = True: cboEdit(cbo年龄).TabStop = False

    '设置过敏输入来源控件属性
    mblnUseTYT = False
    Call SetContolsByAllerPara(p门诊医生站, mint过敏输入来源, optAller(0), optAller(1))

    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("如果关闭窗体，你所作的更改将不会保存。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    Call zlDatabase.SetPara("门诊诊断输入", IIf(optInput(0).Value, 0, 1), glngSys, p门诊医生站, InStr(mstrPrivs, "参数设置") > 0)
    Call zlDatabase.SetPara("过敏输入来源", IIf(optAller(0).Value, "0", "1"), glngSys, p门诊医生站, optAller(0).Enabled And optAller(0).Visible)
End Sub

Private Sub optAller_Click(Index As Integer)
'能触发CLiCK事件说明gbytPass=3,gint过敏输入来源=0因此只需判断控件的值即可
    mblnUseTYT = Index = 1
End Sub

Private Sub optInput_LostFocus(Index As Integer)
    optInput(0).TabStop = False: optInput(1).TabStop = False '要强行代码执行一次
End Sub

Private Sub optState_Click(Index As Integer)
    Dim blnDo As Boolean
    
    If Visible Then
        '复诊：在诊断尚未录入的情况下则自动提取上次诊断
        If Index = opt复诊 Then
            If chkEdit(Index).Value = 1 Then
                blnDo = vsDiagXY.Rows = vsDiagXY.FixedRows + 1 And vsDiagZY.Rows = vsDiagZY.FixedRows + 1
                If blnDo Then blnDo = blnDo And vsDiagXY.TextMatrix(vsDiagXY.FixedRows, col诊断) = "" And vsDiagZY.TextMatrix(vsDiagZY.FixedRows, col诊断) = ""
                If blnDo Then Call LoadPatiDiag(True)
            End If
        End If
        
        mblnChange = True
    End If
End Sub

Private Sub timThis_Timer()
    Dim lngSelNum As Long
    
    If vsAller.Col = AC_过敏时间 Then
        lngSelNum = vsAller.EditSelStart
        If lngSelNum <> mlngSelNum And lngSelNum <> 16 And lngSelNum <> 0 Then
            Call Vs_EditSelChange(lngSelNum)
            mlngSelNum = lngSelNum
        End If
    End If
End Sub

Private Sub Vs_EditSelChange(ByVal lngSelNum As Long)
'当用户切换光标的时候触发
    With vsAller
        If lngSelNum <= 4 Then
            .EditSelStart = 0
            .EditSelLength = 4
            mlngNum = 0
            mlngNumBack = 4
        ElseIf lngSelNum <= 7 Then
            .EditSelStart = 5
            .EditSelLength = 2
            mlngNum = 5
            mlngNumBack = 7
        ElseIf lngSelNum <= 10 Then
            .EditSelStart = 8
            .EditSelLength = 2
            mlngNum = 8
            mlngNumBack = 10
        ElseIf lngSelNum <= 13 Then
            .EditSelStart = 11
            .EditSelLength = 2
            mlngNum = 11
            mlngNumBack = 13
        ElseIf lngSelNum < 16 Then
            .EditSelStart = 14
            .EditSelLength = 2
            mlngNum = 14
            mlngNumBack = 16
        End If
    End With
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index <> txt就诊摘要 Then
        Call zlControl.TxtSelAll(txtEdit(Index))
    ElseIf txtEdit(Index).SelLength = 0 Then
        Call zlControl.TxtSelAll(txtEdit(Index))
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Index = txt医学警示 Then
            txtEdit(txt医学警示) = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt出生地点 Or Index = txt家庭地址 Or Index = txt户口地址) And txtEdit(Index).Text <> "" Then
            '输入地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "地区", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt单位名称 And txtEdit(Index).Text <> "" Then
            '输入工作单位
            strSQL = "Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " And (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "工作单位", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") > 0 Then txtEdit(Index).Tag = Val(rsTmp!ID)
                If txtEdit(txt单位电话).Text = "" Then
                    txtEdit(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
            Else
                txtEdit(Index).Tag = ""
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf (Index = txt区域 Or Index = txt籍贯) And txtEdit(Index).Text <> "" Then
            '输入区域数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 区域 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(Index = txt区域, "区域", "籍贯"), False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Index = txt医学警示 Then
            txtEdit(txt医学警示).Text = ""
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '非控制按键
        If Index = txt医学警示 Then
            KeyAscii = 0
        End If
        '选择快捷键
        If KeyAscii = Asc("*") Then
            '注意界面上要求CMD和对应TXT的Index相同
            On Error Resume Next
            strSQL = ""
            strSQL = cmdEdit(Index).Name
            err.Clear: On Error GoTo 0
            If strSQL <> "" Then
                KeyAscii = 0
                Call cmdEdit_Click(Index)
                Exit Sub
            End If
        End If
        
        '限制输入长度
        If txtEdit(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(Index).Text) > txtEdit(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '限制输入内容
        Select Case Index
'            Case txt年龄 '允许自由录入了
'                strMask = "1234567890"
            'Case txt出生日期 'MaskEdit限制了
                'strMask = "1234567890-"
            Case txt身份证号
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Case txt家庭电话, txt单位电话
                strMask = "1234567890-()"
            Case txt家庭邮编, txt单位邮编, txt户口地址邮编
                strMask = "1234567890"
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    Else
        If Index = txt医学警示 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt发病日期_Change()
    If Visible Then mblnChange = True
    
    If IsDate(txt发病日期.Text) Then
        txt发病时间.Enabled = True
    Else
        txt发病时间.Enabled = False
        txt出生时间.Text = "__:__"
    End If
End Sub

Private Sub txt发病日期_GotFocus()
    Call zlControl.TxtSelAll(txt发病日期)
End Sub

Private Sub txt发病日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt发病日期_Validate(Cancel As Boolean)
    If txt发病日期.Text <> "____-__-__" And Not IsDate(txt发病日期.Text) Then
        txt发病日期.Text = "____-__-__": Cancel = True
    End If
End Sub

Private Sub txt发病时间_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt发病时间_GotFocus()
    Call zlControl.TxtSelAll(txt发病时间)
End Sub

Private Sub txt发病时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt发病时间_Validate(Cancel As Boolean)
    If txt发病时间.Text <> "__:__" And Not IsDate(txt发病时间.Text) Then
        txt发病时间.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = AC_过敏药物 Then mbln过敏药物Edit = True '处理回车定位到下一个单元格的问题
    Call vsAller_AfterRowColChange(-1, -1, Row, Col)
    If Col = AC_过敏药物 Then mbln过敏药物Edit = False '处理回车定位到下一个单元格的问题
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAller
        If NewCol = AC_过敏药物 Then
            If Not mbln过敏药物Edit Then
                .ComboList = "..."
                .FocusRect = flexFocusSolid
            Else '处理回车定位到下一个单元格的问题
                .ComboList = ""
                .FocusRect = flexFocusSolid
            End If
        Else
            .FocusRect = IIf(Trim(vsAller.TextMatrix(NewRow, AC_过敏药物)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     If Col = AC_过敏时间 And Trim(vsAller.Cell(flexcpData, Row, AC_过敏药物)) = "" Then Cancel = True
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int性别 As Integer
    
    With vsAller
        If mblnUseTYT Then
            strSQL = gobjPass.inputAllergy()
            If strSQL <> "" Then
                Call SetAllerInput(Row, , strSQL)
                Call AllerEnterNextCell
            End If
        Else
            If cboEdit(cbo性别).Text Like "*男*" Then
                int性别 = 1
            ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                int性别 = 2
            End If
            
            strSQL = _
                " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
                " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
                " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                " Union All" & _
                " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
                " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                " From 诊疗项目目录 A,药品特性 B" & _
                " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
                IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call SetAllerInput(Row, rsTmp)
                Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AC_过敏药物) <> "" Then
                If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    .Tag = ""
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyLeft Then
        If mlngNum <= 4 Then Exit Sub
        If mlngNum <= 7 Then Vs_EditSelChange (4): Exit Sub
        If mlngNum <= 10 Then Vs_EditSelChange (7): Exit Sub
        If mlngNum <= 13 Then Vs_EditSelChange (10): Exit Sub
        If mlngNum <= 16 Then Vs_EditSelChange (13): Exit Sub
    End If
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    With vsAller
        If KeyAscii = vbKeySpace Then  'Space
            If .Col = AC_过敏药物 And mblnUseTYT Then KeyAscii = 0: Exit Sub
        End If
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AC_过敏药物 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim blnIsNextchr As Boolean
    Dim strChr As String

    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    With vsAller
        If Col = AC_过敏时间 Then
            If KeyAscii = 13 Then .Col = .Col + 1: .ShowCell Row, Col: Exit Sub
            If KeyAscii = vbKeyBack Then
                If mlngNumBack <= 16 Then
                    If mlngNumBack = 0 Then KeyAscii = 0: Exit Sub
                    blnIsNextchr = InStr("1234567890", Mid(.TextMatrix(.Row, .Col), mlngNumBack, 1)) = 0
                    strChr = Mid(.TextMatrix(.Row, .Col), mlngNumBack - IIf(blnIsNextchr, 1, 0), 1)
                    mlngNumBack = mlngNumBack - IIf(blnIsNextchr, 2, 1)
                    .EditText = Mid(.EditText, 1, mlngNumBack) & strChr & Mid(.EditText, mlngNumBack + 2)
                    mlngNum = mlngNumBack
                    KeyAscii = 0
                    If mlngNum <= 4 Then
                        .EditSelStart = 0
                        .EditSelLength = 4
                    ElseIf mlngNum <= 8 Then
                        .EditSelStart = 5
                        .EditSelLength = 2
                    ElseIf mlngNum <= 11 Then
                        .EditSelStart = 8
                        .EditSelLength = 2
                    ElseIf mlngNum <= 14 Then
                        .EditSelStart = 11
                        .EditSelLength = 2
                    ElseIf mlngNum <= 16 Then
                        .EditSelStart = 14
                        .EditSelLength = 2
                    End If
                End If
            Else
                If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
                If Len(.EditText) <= 16 And mlngNum <> 16 Then
                    blnIsNextchr = InStr("1234567890", Mid(.TextMatrix(.Row, .Col), mlngNum + 2, 1)) = 0
                    strChr = Chr(KeyAscii)
                    .EditText = Mid(.EditText, 1, mlngNum) & strChr & Mid(.EditText, mlngNum + 2)
                    mlngNum = mlngNum + IIf(blnIsNextchr, 2, 1)
                    mlngNumBack = mlngNum
                End If
                KeyAscii = 0
                If mlngNum <= 4 Then
                    .EditSelStart = 0
                    .EditSelLength = 4
                ElseIf mlngNum <= 7 Then
                    .EditSelStart = 5
                    .EditSelLength = 2
                ElseIf mlngNum <= 10 Then
                    .EditSelStart = 8
                    .EditSelLength = 2
                ElseIf mlngNum <= 13 Then
                    .EditSelStart = 11
                    .EditSelLength = 2
                ElseIf mlngNum <= 16 Then
                    .EditSelStart = 14
                    .EditSelLength = 2
                End If
            End If
        ElseIf Col = AC_过敏药物 Then
            If KeyAscii <> 13 Then
                If mblnUseTYT Then KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = AC_过敏药物 Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
    ElseIf Col = AC_过敏时间 Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = 4
        mlngNum = 0
        mlngNumBack = 0
        timThis.Enabled = True
    End If
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AC_过敏反应 And Trim(vsAller.TextMatrix(Row, AC_过敏药物)) = "" Then Cancel = True
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int性别  As Integer
    Dim curDate As Date
    
    With vsAller
        If Col = AC_过敏药物 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call AllerEnterNextCell
            Else
                strInput = UCase(.EditText)
                If cboEdit(cbo性别).Text Like "*男*" Then
                    int性别 = 1
                ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                    int性别 = 2
                End If
                strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                    " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                    " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                    " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                    IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                    Decode(gint简码, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " Order by A.编码"
                
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%", int性别, gint简码 + 1)
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    Call SetAllerInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call AllerEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = AC_过敏时间 Then
            If Not IsDate(.EditText) And .EditText <> "" Then
                MsgBox "您输入的日期格式不正确。格式如：2010-10-10 18:30。"
                Cancel = True
                .EditText = vsAller.TextMatrix(Row, Col)
            Else
                If .EditText <> "" Then
                    curDate = zlDatabase.Currentdate
                    If CDate(.EditText) > curDate Then
                        MsgBox "您输入的日期不能大于当前时间。当前时间：" & curDate & "。"
                        Cancel = True
                        .EditText = .TextMatrix(Row, Col)
                    End If
                End If
                timThis.Enabled = False
                If .Cell(flexcpData, Row, Col) <> .EditText Then
                    .Cell(flexcpData, Row, Col) = .EditText
                    mblnChange = True
                End If
                .Tag = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col诊断 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagXY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        If .Col = Col Then Call vsDiagXY_AfterRowColChange(-1, -1, Row, Col)
    End With
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDiagXY
        If Not DiagCellEditable(vsDiagXY, NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col诊断 Then
                .ComboList = "..."
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsDiagZY.ColWidth(Col) = vsDiagXY.ColWidth(Col)
End Sub

Private Sub vsDiagXY_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col疑诊 Then Cancel = True
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str性别 As String
    
    With vsDiagXY
        If optInput(0).Value Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", mlng科室ID, , True, False)
        Else
            'D-ICD-10疾病编码
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D", mlng科室ID, cboEdit(cbo性别).Text, True)
        End If
        If rsTmp Is Nothing Then
            If optInput(0).Value Then
                MsgBox "没有疾病诊断数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call XYSetDiagInput(Row, rsTmp)
            Call DiagEnterNextCell(vsDiagXY)
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col诊断) <> "" Then
                If .TextMatrix(.Row, col医嘱ID) = "" Then
                    If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call CreatePlugInOK(p门诊医生站)
                        '删除主/次要诊断后调用外挂接口
                        If Not gobjPlugIn Is Nothing Then
                            On Error Resume Next
                            Call gobjPlugIn.DiagnosisDeleted(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(.TextMatrix(.Row, col诊断ID)), .TextMatrix(.Row, col诊断))
                            Call zlPlugInErrH(err, "DiagnosisDeleted")
                            err.Clear: On Error GoTo 0
                        End If
                        .RemoveItem .Row
                        mblnChange = True
                        .Tag = ""
                    End If
                Else
                    MsgBox "该诊断对应的处方已发送，不能删除。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagXY)
        ElseIf KeyAscii = 32 And (.Col = col疑诊) Then
            If DiagCellEditable(vsDiagXY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col疑诊 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "？", "")
                End If
            End If
        Else
            If .Col = col诊断 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(vsDiagXY, Row, Col) Then
        Cancel = True
    ElseIf Col = col疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Function GetZYSQL(ByRef strInput As String, ByRef strSQL As String, ByRef str性别 As String, Optional ByVal strType As String) As String
'功能：获得查询中医诊断的SQL
'参数：strInput-查询条件,strsql--返回的SQL，str性别--病人的性别  ,strType疾病编码种类。
'返回：strsql--查询中医诊断的SQL
    If optInput(0).Value And strType <> "Z" Then
        '按诊断输入:中医部份，一个诊断可能属于多个分类
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
        Else
            strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
        End If
        strSQL = _
            " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
            " From 疾病诊断目录 A,疾病诊断别名 B" & _
            " Where A.ID=B.诊断ID And A.类别=2" & _
            " And B.码类=[4] And (" & strSQL & ")" & _
            " Order by A.编码"
    Else
        If cboEdit(cbo性别).Text Like "*男*" Then
            str性别 = "男"
        ElseIf cboEdit(cbo性别).Text Like "*女*" Then
            str性别 = "女"
        End If
        'B-中医疾病编码
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "名称 Like [2]" '输入汉字时只匹配名称
        Else
            strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gint简码 = 0, "简码", "五笔码") & " Like [2]"
        End If
        strSQL = _
            " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
            " From 疾病编码目录" & _
            " Where 类别='" & IIf(strType = "", "B", strType) & "' And (" & strSQL & ")" & _
            IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
            " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by 编码"
    End If
    GetZYSQL = strSQL
End Function

Private Function GetXYSQL(ByRef strInput As String, ByRef strSQL As String, ByRef str性别 As String) As String
'功能：获得查询西医诊断的SQL
'参数：strInput-查询条件,strsql--返回的SQL，str性别--病人的性别
'返回：strsql--查询西医诊断的SQL
    If optInput(0).Value Then
        '按诊断输入:西医部份，一个诊断可能属于多个分类
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "B.名称 Like [2]" '输入汉字时,只匹配名称
        Else
            strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
        End If
        strSQL = _
            " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
            " From 疾病诊断目录 A,疾病诊断别名 B" & _
            " Where A.ID=B.诊断ID And A.类别=1" & _
            " And B.码类=[4] And (" & strSQL & ")" & _
            " Order by A.编码"
    Else
        If cboEdit(cbo性别).Text Like "*男*" Then
            str性别 = "男"
        ElseIf cboEdit(cbo性别).Text Like "*女*" Then
            str性别 = "女"
        End If
        'D-ICD-10疾病编码
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "名称 Like [2]" '输入汉字时,只匹配名称
        Else
            strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gint简码 = 0, "简码", "五笔码") & " Like [2]"
        End If
        strSQL = _
            " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
            " From 疾病编码目录 Where 类别='D' And (" & strSQL & ")" & _
            IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
            " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by 编码"
    End If
    GetXYSQL = strSQL
End Function

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str性别 As String, strInput As String
    Dim vPoint As POINTAPI, int诊断输入 As Integer
    
    With vsDiagXY
        If Col = col诊断 Then
            If .EditText = "" Then
                If .TextMatrix(Row, col编码) <> "" Then
                    .EditText = .Cell(flexcpData, Row, Col)
                End If
                If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
            ElseIf .TextMatrix(Row, col编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '判断加了前缀后的名称是否存在其他的诊断编码
                strInput = UCase(.EditText)
                strSQL = GetXYSQL(strInput, strSQL, str性别)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, _
                         str性别, gint简码 + 1)
                If rsTmp.RecordCount <> 1 Then
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, col诊断) = .EditText
                Else
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                End If
                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
                .Tag = ""
                mblnChange = True
            Else
                int诊断输入 = Val(Mid(gstr诊断输入, 1, 1))
                If int诊断输入 = 0 Then int诊断输入 = 1
                
                strInput = UCase(.EditText)
                strSQL = GetXYSQL(strInput, strSQL, str性别)
                If int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", str性别, gint简码 + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                    Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).Value, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", str性别, gint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                            If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col发病时间 Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                Else
                    MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。"
                    Cancel = True
                End If
            End If
            If .EditText <> .TextMatrix(Row, Col) Then mblnChange = True: vsDiagXY.Tag = ""
            If Row = 1 Then
                If .EditText <> "" Then
                    '如果填写了发病时间，则下面的发病时间则不允许填写了
                    txt发病日期.BackColor = vbButtonFace
                    txt发病日期.Enabled = False
                    txt发病时间.BackColor = vbButtonFace
                    txt发病时间.Enabled = False
                Else
                    If vsDiagZY.TextMatrix(0, col发病时间) = "" Then
                        txt发病日期.BackColor = vbWindowBackground
                        txt发病日期.Enabled = True
                        txt发病时间.BackColor = vbWindowBackground
                        txt发病时间.Enabled = True
                        txt发病日期.Text = "____-__-__"
                        txt发病时间.Text = "__:__"
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col诊断 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagZY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        If .Col = Col Then Call vsDiagZY_AfterRowColChange(-1, -1, Row, Col)
    End With
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDiagZY
        If Not DiagCellEditable(vsDiagZY, NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col诊断 Then
                .ComboList = "..."
            ElseIf NewCol = col中医证候 Then
                If .TextMatrix(NewRow, col诊断) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsDiagXY.ColWidth(Col) = vsDiagZY.ColWidth(Col)
End Sub

Private Sub vsDiagZY_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col疑诊 Then Cancel = True
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str性别 As String
    Dim blnCancle As Boolean
    
    With vsDiagZY
        If Col = col诊断 Then
            If optInput(0).Value Then
                '按诊断输入:中医部份，一个诊断可能属于多个分类
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", mlng科室ID, , True, False)
            Else
                'B-中医疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", mlng科室ID, cboEdit(cbo性别).Text, True)
            End If
            If rsTmp Is Nothing Then
                If optInput(0).Value Then
                    MsgBox "没有疾病诊断数据可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call ZYSetDiagInput(Row, rsTmp)
                Call DiagEnterNextCell(vsDiagZY)
            End If
        ElseIf Col = col中医证候 Then
            If optInput(0).Value Then
                '按诊断输入:先查是否有对应
                If Not Set中医证候(Row, Val(.TextMatrix(Row, col诊断ID))) Then
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, cboEdit(cbo性别).Text, True)
                Else
                    Exit Sub
                End If
            Else
                'Z-中医疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, cboEdit(cbo性别).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call Set中医证候(Row, 0, rsTmp)
                Call DiagEnterNextCell(vsDiagZY)
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col诊断) <> "" Then
                If .TextMatrix(.Row, col医嘱ID) = "" Then
                    If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call CreatePlugInOK(p门诊医生站)
                        '删除主/次要诊断后调用外挂接口
                        If Not gobjPlugIn Is Nothing Then
                            On Error Resume Next
                            Call gobjPlugIn.DiagnosisDeleted(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(.TextMatrix(.Row, col诊断ID)), .TextMatrix(.Row, col诊断))
                            Call zlPlugInErrH(err, "DiagnosisDeleted")
                            err.Clear: On Error GoTo 0
                        End If
                        .RemoveItem .Row
                        mblnChange = True
                        .Tag = ""
                    End If
                Else
                    MsgBox "该诊断对应的处方已发送，不能删除。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagZY)
        ElseIf KeyAscii = 32 And (.Col = col疑诊) Then
            If DiagCellEditable(vsDiagZY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col疑诊 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "？", "")
                End If
            End If
        Else
            If .Col = col诊断 Or .Col = col中医证候 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(vsDiagZY, Row, Col) Then
        Cancel = True
    ElseIf Col = col疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Function DiagCellEditable(objGrid As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With objGrid
        '隐藏列不可编辑
        If .ColHidden(lngCol) Then Exit Function
        
        If .TextMatrix(lngRow, col医嘱ID) <> "" Then
            If lngCol = col诊断 Then
                Exit Function
            End If
        End If
        '必须先输入诊断
        If .TextMatrix(lngRow, col诊断) = "" Then
            If lngCol = col疑诊 Or lngCol = col发病时间 Then
                Exit Function
            End If
        End If
        If lngCol = col编码 Then
            Exit Function
        End If
        '必须先输诊断再输证候
        If lngCol = col中医证候 Then
            If .TextMatrix(lngRow, col诊断) = "" Then Exit Function
        End If
    End With
    DiagCellEditable = True
End Function

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAller
        If .Col = AC_过敏反应 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AC_过敏药物
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub DiagEnterNextCell(objGrid As VSFlexGrid)
    Dim i As Long, j As Long
    
    With objGrid
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col诊断) To col疑诊
                If DiagCellEditable(objGrid, i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col疑诊 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'功能：处理过敏药物的输入
'参数：strTYTInput=太元通合理用药接口返回的字符串
    Dim strSQL As String, curDate As Date
    Dim arrTmp As Variant
    Dim strAllerOld As String, strAllerNew As String
    
    With vsAller
        
        strAllerOld = .Cell(flexcpData, lngRow, AC_过敏药物) & ";" & .TextMatrix(lngRow, AC_过敏源编码)
        
        If mblnUseTYT Then
            arrTmp = Split(strTYTInput, ";")
            
            If UBound(arrTmp) < 1 Then Exit Sub
            If strAllerOld <> strTYTInput Or Val(.RowData(lngRow) & "") <> 0 Then
                .TextMatrix(lngRow, AC_过敏药物) = arrTmp(1)
                .TextMatrix(lngRow, AC_过敏源编码) = arrTmp(0)
                .RowData(lngRow) = 0
            End If
        Else
            
            If Not rsInput Is Nothing Then
                .RowData(lngRow) = CLng(rsInput!ID)
                .TextMatrix(lngRow, AC_过敏药物) = Nvl(rsInput!名称)
            Else
                .RowData(lngRow) = 0
                .TextMatrix(lngRow, AC_过敏药物) = .EditText
            End If
            
            strAllerNew = .TextMatrix(lngRow, AC_过敏药物) & ";" & .TextMatrix(lngRow, AC_过敏源编码)
            
            If strAllerOld <> strAllerNew Or Val(.RowData(lngRow) & "") <> 0 Then
                .TextMatrix(lngRow, AC_过敏源编码) = ""
            End If
        End If
        
        .Cell(flexcpData, lngRow, AC_过敏药物) = .TextMatrix(lngRow, AC_过敏药物)
        If .Cell(flexcpData, lngRow, AC_过敏时间) = "" Then
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, AC_过敏时间) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, AC_过敏时间) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        
        .Tag = ""
        mblnChange = True
    End With
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理西医诊断项目的输入
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lng原诊断id As Long '0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id 的值就是修改前的 诊断ID或疾病ID
    
    With vsDiagXY
        If Not rsInput Is Nothing Then
            '检查是否允许修改
            If .TextMatrix(.Row, col医嘱ID) <> "" Then
                MsgBox "该诊断对应的处方已发送，不能修改。", vbInformation, Me.Caption
                Exit Sub
            End If
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                    .TextMatrix(lngRow, col类型) = "西医"
                    lng原诊断id = 0
                Else
                    lng原诊断id = Val(.TextMatrix(lngRow, col诊断ID))
                End If
                
                .TextMatrix(lngRow, col诊断) = Nvl(rsInput!名称)
                .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
                .TextMatrix(lngRow, col编码) = IIf(Not IsNull(rsInput!编码), rsInput!编码, "")
                '根据诊断确定疾病,或根据疾病确定诊断
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col疾病ID) = ""
                    strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col诊断ID) = ""
                    strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If optInput(0).Value Then
                        .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                Call CreatePlugInOK(p门诊医生站)
                '输入主/次要诊断后调用外挂接口
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If lngRow = .FixedRows Then
                        Call gobjPlugIn.DiagnosisEnter(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(rsInput!项目ID), .TextMatrix(lngRow, col诊断), lng原诊断id)
                        Call zlPlugInErrH(err, "DiagnosisEnter")
                    Else
                        Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(rsInput!项目ID), .TextMatrix(lngRow, col诊断), lng原诊断id)
                        Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End If
                    err.Clear: On Error GoTo errH
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col诊断) = .EditText
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
        End If
        
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col类型) = "西医"
        End If
        .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
        mblnChange = True
        .Tag = ""
    End With
    
    If optState(opt复诊).Value = False Then
        If PatiReSeeDoctor Then
            If MsgBox("病人就诊科室、医生、诊断与上次相同，要标记为复诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                optState(opt复诊).Value = True
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function PatiReSeeDoctor() As Boolean
'功能：判断病人本次是否复诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    '医生、科室与上次相同：没有转诊、续诊的
    strSQL1 = "Select 病人ID,执行人 as 医生,执行部门ID as 科室ID From 病人挂号记录 Where ID=[2] And 转诊科室ID Is Null And 续诊科室ID Is Null"
    
    strSQL2 = "Select Max(ID) as ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
            " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
    strSQL2 = "Select 病人ID,执行人 as 医生,执行部门ID as 科室ID From 病人挂号记录 Where ID=(" & strSQL2 & ") And 转诊科室ID Is Null And 续诊科室ID Is Null"
    
    strSQL = "Select 1 From (" & strSQL1 & ") A,(" & strSQL2 & ") B Where A.病人ID=B.病人ID And A.医生=B.医生 And A.科室ID=B.科室ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng病人ID, mlng挂号ID)
    If rsTmp.EOF Then Exit Function
    
    '主要诊断与上次相同
    With vsDiagXY
        If .TextMatrix(.FixedRows, col诊断) <> "" Then
            strSQL = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                    " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
            strSQL = "Select 1 From 病人诊断记录" & _
                " Where 病人ID=[1] And 主页ID=(" & strSQL & ")" & _
                " And 诊断类型=1 And 记录来源 IN(1,3) And 诊断次序=1" & _
                " And (疾病ID=[3] And 疾病ID<>0 Or 诊断ID=[4] And 诊断ID<>0 Or 诊断描述=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng病人ID, mlng挂号ID, _
                Val(.TextMatrix(.FixedRows, col疾病ID)), Val(.TextMatrix(.FixedRows, col诊断ID)), .TextMatrix(.FixedRows, col诊断))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    With vsDiagZY
        If .TextMatrix(.FixedRows, col诊断) <> "" Then
            strSQL = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                   " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
            strSQL = "Select 1 From 病人诊断记录" & _
                " Where 病人ID=[1] And 主页ID=(" & strSQL & ")" & _
                " And 诊断类型=11 And 记录来源 IN(1,3) And 诊断次序=1" & _
                " And (疾病ID=[3] And 疾病ID<>0 Or 诊断ID=[4] And 诊断ID<>0 Or 诊断描述=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng病人ID, mlng挂号ID, _
                Val(.TextMatrix(.FixedRows, col疾病ID)), Val(.TextMatrix(.FixedRows, col诊断ID)), .TextMatrix(.FixedRows, col诊断))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理中医诊断项目的输入
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, str编码 As String
    Dim i As Long
    Dim strTmp As String
    Dim lng原诊断id As Long '0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id 的值就是修改前的 诊断ID或疾病ID
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            '检查是否允许修改
            If .TextMatrix(.Row, col医嘱ID) <> "" Then
                MsgBox "该诊断对应的处方已发送，不能修改。", vbInformation, Me.Caption
                Exit Sub
            End If
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                    .TextMatrix(lngRow, col类型) = "中医"
                    lng原诊断id = 0
                Else
                    lng原诊断id = Val(.TextMatrix(lngRow, col诊断ID))
                End If
                
                If Not IsNull(rsInput!编码) Then
                    str编码 = rsInput!编码
                End If
                
                If InStr(.TextMatrix(lngRow, col诊断), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, col诊断), InStrRev(.TextMatrix(lngRow, col诊断), "("))
                End If
                .TextMatrix(lngRow, col诊断) = Nvl(rsInput!名称) & strTmp
                .TextMatrix(lngRow, col编码) = IIf(Not IsNull(rsInput!编码), rsInput!编码, "")
                
                '根据诊断确定疾病,或根据疾病确定诊断
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col疾病ID) = ""
                    strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col诊断ID) = ""
                    strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If optInput(0).Value Then
                        .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                '中医根据疾病诊断参考取证候
                Call Set中医证候(lngRow, Val(.TextMatrix(lngRow, col诊断ID)))
                                 
                Call CreatePlugInOK(p门诊医生站)
                '输入主/次要诊断后调用外挂接口
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If lngRow = .FixedRows Then
                        Call gobjPlugIn.DiagnosisEnter(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(rsInput!项目ID), .TextMatrix(lngRow, col诊断), lng原诊断id)
                        Call zlPlugInErrH(err, "DiagnosisEnter")
                    Else
                        Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(rsInput!项目ID), .TextMatrix(lngRow, col诊断), lng原诊断id)
                        Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End If
                    err.Clear: On Error GoTo errH
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col诊断) = .EditText
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
            .TextMatrix(lngRow, col证候ID) = ""
        End If
        
        '如果是出院诊断,始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col类型) = "中医"
        End If
        .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
        mblnChange = True
        .Tag = ""
    End With
    
    If optState(opt复诊).Value = False Then
        If PatiReSeeDoctor Then
            If MsgBox("病人就诊科室、医生、诊断与上次相同，要标记为复诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                optState(opt复诊).Value = True
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Set中医证候(ByVal lngRow As Long, ByVal lng诊断ID As Long, Optional ByVal rsInput As Recordset) As Boolean
'功能：中医根据疾病诊断参考取证候
'参数：rsInput-如果不为空，则输出指定的中药证候记录集
'返回：是否有对应关系
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    With vsDiagZY
        '去掉已有的证候
        If InStr(.TextMatrix(lngRow, col诊断), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, col诊断), 1, InStrRev(.TextMatrix(lngRow, col诊断), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, col诊断)
        End If
        If rsInput Is Nothing Then
            If lng诊断ID <> 0 Then
                strSQL = "Select Distinct a.证候序号 as ID,a.证候ID,a.证候名称,b.编码 as 证候编码" & _
                    " From 疾病诊断参考 A,疾病编码目录 B" & _
                    " Where a.证候ID=b.ID(+) And a.诊断ID=[1] And a.证候名称 is Not NULL" & _
                    " Order by a.证候序号"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng诊断ID)
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, col证候ID) = Nvl(rsTmp!证候id)
                    If Not IsNull(rsTmp!证候名称) Then
                        .TextMatrix(lngRow, col诊断) = strTmp
                        .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
                        .TextMatrix(lngRow, col中医证候) = Nvl(rsTmp!证候名称)
                        .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
                        If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
                        mblnChange = True
                        .Tag = ""
                    End If
                    Set中医证候 = True
                Else
                    If blnCancel Then
                        Set中医证候 = True
                        If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col中医证候)
                    Else
                        Set中医证候 = False
                    End If
                End If
            Else
                Set中医证候 = False
            End If
        Else
            .TextMatrix(lngRow, col证候ID) = Nvl(rsInput!项目ID)
            .TextMatrix(lngRow, col诊断) = strTmp
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            .TextMatrix(lngRow, col中医证候) = Nvl(rsInput!名称)
            .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
            .Tag = ""
            mblnChange = True
        End If
    End With
End Function

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim str性别 As String, int诊断输入 As Integer
    
    With vsDiagZY
        If Col = col诊断 Or Col = col中医证候 Then
            If .EditText = "" Then
                If .TextMatrix(Row, col编码) <> "" And Col = col诊断 Then
                    .EditText = .Cell(flexcpData, Row, Col)
                Else
                    '中医症候则清除备份数据
                    If Col = col中医证候 Then
                        .Cell(flexcpData, Row, Col) = ""
                    End If
                End If
                If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
            ElseIf Col = col诊断 And .TextMatrix(Row, col编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                strSQL = GetZYSQL(strInput, strSQL, str性别)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str性别, gint简码 + 1)
                If rsTmp.RecordCount = 1 Then
                    Call ZYSetDiagInput(Row, rsTmp):
                    .EditText = .Text
                Else
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, col诊断) = .EditText
                End If
                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
                .Tag = ""
                mblnChange = True
            Else
                int诊断输入 = Val(Mid(gstr诊断输入, 1, 1))
                If int诊断输入 = 0 Then int诊断输入 = 1
                
                strInput = UCase(.EditText)
                strSQL = GetZYSQL(strInput, strSQL, str性别, IIf(Col = col诊断, "B", "Z"))
                If Col = col诊断 Then
                    If int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                        On Error GoTo errH
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", str性别, gint简码 + 1)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                        End If
                        Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                        If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
                    Else
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).Value, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", str性别, gint简码 + 1)
                        If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                            Cancel = True
                        Else
                            '检查诊断输入方式
                            If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                                MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                                Cancel = True
                            Else
                                Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                                If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
                            End If
                        End If
                    End If
                ElseIf Col = col中医证候 Then
                    If optInput(0).Value Then
                        '按诊断输入:先查是否有对应
                        If Set中医证候(Row, Val(.TextMatrix(Row, col诊断ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", str性别, gint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call Set中医证候(Row, 0, rsTmp)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col发病时间 Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                Else
                    MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。"
                    Cancel = True
                End If
            End If
            If .EditText <> .TextMatrix(Row, Col) Then mblnChange = True: vsDiagZY.Tag = ""
            If Row = 0 Then
                If .EditText <> "" Then
                    '如果填写了发病时间，则下面的发病时间则不允许填写了
                    txt发病日期.BackColor = vbButtonFace
                    txt发病日期.Enabled = False
                    txt发病时间.BackColor = vbButtonFace
                    txt发病时间.Enabled = False
                Else
                    If vsDiagXY.TextMatrix(1, col发病时间) = "" Then
                        txt发病日期.BackColor = vbWindowBackground
                        txt发病日期.Enabled = True
                        txt发病时间.BackColor = vbWindowBackground
                        txt发病时间.Enabled = True
                        txt发病日期.Text = "____-__-__"
                        txt发病时间.Text = "__:__"
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetCboFromSQL(ByVal strSQL As String, ByVal arrCboIdx As Variant)
'功能：将指定数据源中的数据装入指定索引的一个或多个ComboBox
'参数：strSQL=包含"ID,简码,名称,缺省标志"字段
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, j As Long
    
    '清除原有数据
    For i = 0 To UBound(arrCboIdx)
        cboEdit(arrCboIdx(i)).Clear
    Next
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '装入数据
    For i = 1 To rsTmp.RecordCount
        For j = 0 To UBound(arrCboIdx)
            If IsNull(rsTmp!简码) Then
                cboEdit(arrCboIdx(j)).AddItem rsTmp!名称
            Else
                cboEdit(arrCboIdx(j)).AddItem rsTmp!简码 & "-" & Chr(13) & rsTmp!名称
            End If
            cboEdit(arrCboIdx(j)).ItemData(cboEdit(arrCboIdx(j)).NewIndex) = Nvl(rsTmp!ID, 0)
            If Nvl(rsTmp!缺省标志, 0) = 1 Then
                Call zlControl.CboSetIndex(cboEdit(arrCboIdx(j)).hwnd, cboEdit(arrCboIdx(j)).NewIndex)
            End If
        Next
        rsTmp.MoveNext
    Next
    '无缺省时,为未选中
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    
    If UCase(objTmp.Container.Name) <> UCase("fraInfo") Then
        If UCase(objTmp.Container.Container.Name) = UCase("fraInfo") Then sstInfo.Tab = objTmp.Container.Container.Index
    Else
        sstInfo.Tab = objTmp.Container.Index
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then objTmp.SetFocus
    Me.Refresh
End Function
