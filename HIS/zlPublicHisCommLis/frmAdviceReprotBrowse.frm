VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CO8DDC~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmAdviceReprotBrowse 
   Caption         =   "单报告查阅"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14805
   Icon            =   "frmAdviceReprotBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   14805
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPDF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1140
      ScaleWidth      =   1395
      TabIndex        =   56
      Top             =   8820
      Visible         =   0   'False
      Width           =   1395
      Begin SHDocVwCtl.WebBrowser webSub 
         Height          =   690
         Left            =   180
         TabIndex        =   57
         Top             =   150
         Width           =   810
         ExtentX         =   1429
         ExtentY         =   1217
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10485
      Left            =   180
      ScaleHeight     =   10455
      ScaleWidth      =   14595
      TabIndex        =   0
      Top             =   1380
      Width           =   14625
      Begin VB.PictureBox picComment 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3465
         Left            =   600
         ScaleHeight     =   3435
         ScaleWidth      =   4965
         TabIndex        =   11
         Top             =   3540
         Width           =   4995
         Begin VB.TextBox txtDiagnosis 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   210
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   2190
            Width           =   4665
         End
         Begin VB.TextBox txtResult 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   210
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   300
            Width           =   4665
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "诊断:"
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
            Index           =   1
            Left            =   60
            TabIndex        =   13
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "评语:"
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
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.PictureBox picCenter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5985
         Left            =   8550
         ScaleHeight     =   5985
         ScaleWidth      =   5715
         TabIndex        =   8
         Top             =   2340
         Width           =   5715
         Begin VB.PictureBox PicNegative 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   5865
            Left            =   -660
            ScaleHeight     =   5835
            ScaleWidth      =   5355
            TabIndex        =   35
            Top             =   1020
            Visible         =   0   'False
            Width           =   5385
            Begin VB.Frame frmChe 
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Caption         =   "结果选择"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Left            =   150
               TabIndex        =   50
               Top             =   2820
               Width           =   5250
               Begin VB.CheckBox chkMicroscope 
                  BackColor       =   &H80000005&
                  Caption         =   "镜检结果"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   3810
                  TabIndex        =   55
                  Top             =   300
                  Width           =   1305
               End
               Begin VB.CheckBox chkNoGerm 
                  BackColor       =   &H80000005&
                  Caption         =   "无细菌生长"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1935
                  TabIndex        =   54
                  Top             =   300
                  Width           =   1815
               End
               Begin VB.CheckBox chkPathopoiesiaGerm 
                  BackColor       =   &H80000005&
                  Caption         =   "无致病菌生长"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   60
                  TabIndex        =   53
                  Top             =   300
                  Width           =   1815
               End
               Begin VB.OptionButton optReport 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "阳性"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Index           =   0
                  Left            =   1935
                  TabIndex        =   52
                  Top             =   600
                  Width           =   885
               End
               Begin VB.OptionButton optReport 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "阴性"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Index           =   1
                  Left            =   60
                  TabIndex        =   51
                  Top             =   600
                  Value           =   -1  'True
                  Width           =   885
               End
            End
            Begin VB.Frame frmNom 
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Height          =   2655
               Left            =   180
               TabIndex        =   43
               Top             =   270
               Width           =   5250
               Begin VB.TextBox txtNormalMicrobes 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   750
                  Left            =   1050
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   46
                  Top             =   1800
                  Width           =   4065
               End
               Begin VB.TextBox txtNoFindMicrobe 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   750
                  Left            =   1050
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   45
                  Top             =   975
                  Width           =   4065
               End
               Begin VB.TextBox txtNormalMicrobe 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   690
                  Left            =   1050
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   44
                  Top             =   210
                  Width           =   4065
               End
               Begin VB.Label lblNormalMicrobes 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "补充描述"
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
                  Left            =   60
                  TabIndex        =   49
                  Top             =   1800
                  Width           =   960
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "未 检 出"
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
                  Left            =   60
                  TabIndex        =   48
                  Top             =   930
                  Width           =   960
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "常规结果"
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
                  Left            =   60
                  TabIndex        =   47
                  Top             =   210
                  Width           =   960
               End
            End
            Begin VB.Frame fraOne 
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Caption         =   "镜检结果"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   30
               TabIndex        =   36
               Top             =   3840
               Width           =   5250
               Begin VB.TextBox txtMicroscopeFinded 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   600
                  Left            =   1110
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   39
                  Top             =   690
                  Width           =   3915
               End
               Begin VB.TextBox txtMicroscopeNOFind 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   510
                  Left            =   1110
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   38
                  Top             =   1350
                  Width           =   3915
               End
               Begin VB.TextBox txtMicroscope 
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
                  Left            =   1110
                  Locked          =   -1  'True
                  TabIndex        =   37
                  Text            =   "显微镜检查"
                  Top             =   270
                  Width           =   3915
               End
               Begin VB.Label lblMicroscopeFinded 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "镜检检出"
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
                  Left            =   90
                  TabIndex        =   42
                  Top             =   660
                  Width           =   960
               End
               Begin VB.Label lblMicroscopeNOFind 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "未 检 出"
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
                  Left            =   60
                  TabIndex        =   41
                  Top             =   1290
                  Width           =   960
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "通过设备"
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
                  Left            =   90
                  TabIndex        =   40
                  Top             =   300
                  Width           =   960
               End
            End
         End
         Begin VB.PictureBox picMicrobePositive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2835
            Left            =   930
            ScaleHeight     =   2805
            ScaleWidth      =   4545
            TabIndex        =   9
            Top             =   4050
            Width           =   4575
            Begin VSFlex8Ctl.VSFlexGrid vsfMicrobePositive 
               Height          =   1785
               Left            =   660
               TabIndex        =   10
               Top             =   300
               Width           =   3285
               _cx             =   5794
               _cy             =   3149
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
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
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
      End
      Begin VB.PictureBox picPatient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   9645
         TabIndex        =   3
         Top             =   90
         Width           =   9675
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床号:"
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
            Index           =   3
            Left            =   6870
            TabIndex        =   7
            Top             =   60
            Width           =   600
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄:"
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
            Left            =   4830
            TabIndex        =   6
            Top             =   60
            Width           =   600
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别:"
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
            Index           =   1
            Left            =   2640
            TabIndex        =   5
            Top             =   60
            Width           =   600
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名:"
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
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   60
            Width           =   600
         End
      End
      Begin VB.PictureBox picTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   8790
         ScaleHeight     =   1245
         ScaleWidth      =   855
         TabIndex        =   17
         Top             =   1860
         Width           =   885
         Begin XtremeSuiteControls.TabControl tabThis 
            Height          =   1065
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   855
            _Version        =   589884
            _ExtentX        =   1508
            _ExtentY        =   1879
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4845
         Left            =   0
         ScaleHeight     =   4815
         ScaleWidth      =   5265
         TabIndex        =   1
         Top             =   450
         Width           =   5295
         Begin VSFlex8Ctl.VSFlexGrid vsfGeneral 
            Height          =   2805
            Left            =   450
            TabIndex        =   2
            Top             =   210
            Width           =   3855
            _cx             =   6800
            _cy             =   4948
            Appearance      =   1
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
      Begin VB.PictureBox picCJYM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   6540
         ScaleHeight     =   1605
         ScaleWidth      =   3375
         TabIndex        =   33
         Top             =   5580
         Width           =   3375
         Begin VSFlex8Ctl.VSFlexGrid VSFCJYM 
            Height          =   975
            Left            =   750
            TabIndex        =   34
            Top             =   330
            Width           =   1965
            _cx             =   3466
            _cy             =   1720
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
      Begin VB.PictureBox MicroorganismSmear 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   4230
         ScaleHeight     =   6375
         ScaleWidth      =   9645
         TabIndex        =   16
         Top             =   4110
         Width           =   9645
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1785
            Index           =   0
            Left            =   480
            ScaleHeight     =   1755
            ScaleWidth      =   2115
            TabIndex        =   62
            ToolTipText     =   "双击查看大图"
            Top             =   3600
            Width           =   2145
            Begin VB.Label lblLoading 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   0
               Left            =   60
               TabIndex        =   63
               Top             =   450
               Width           =   120
            End
            Begin VB.Image imgPicture 
               Height          =   1665
               Index           =   0
               Left            =   210
               Stretch         =   -1  'True
               ToolTipText     =   "点击查看大图"
               Top             =   60
               Width           =   2025
            End
         End
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1785
            Index           =   1
            Left            =   3030
            ScaleHeight     =   1755
            ScaleWidth      =   2115
            TabIndex        =   60
            ToolTipText     =   "双击查看大图"
            Top             =   3600
            Width           =   2145
            Begin VB.Label lblLoading 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   1
               Left            =   60
               TabIndex        =   61
               Top             =   450
               Width           =   120
            End
            Begin VB.Image imgPicture 
               Height          =   1635
               Index           =   1
               Left            =   60
               Stretch         =   -1  'True
               ToolTipText     =   "点击查看大图"
               Top             =   60
               Width           =   1995
            End
         End
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1785
            Index           =   2
            Left            =   5490
            ScaleHeight     =   1755
            ScaleWidth      =   2115
            TabIndex        =   58
            ToolTipText     =   "双击查看大图"
            Top             =   3600
            Width           =   2145
            Begin VB.Label lblLoading 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   2
               Left            =   60
               TabIndex        =   59
               Top             =   480
               Width           =   120
            End
            Begin VB.Image imgPicture 
               Height          =   1635
               Index           =   2
               Left            =   60
               Stretch         =   -1  'True
               ToolTipText     =   "点击查看大图"
               Top             =   60
               Width           =   1965
            End
         End
         Begin VB.Label lblAuditingTime 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   4650
            TabIndex        =   32
            Top             =   5730
            Width           =   1080
         End
         Begin VB.Label lblAuditingMan 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1230
            TabIndex        =   31
            Top             =   5730
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "报告时间:"
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
            Index           =   8
            Left            =   3540
            TabIndex        =   30
            Top             =   5730
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "报告人:"
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
            Index           =   7
            Left            =   300
            TabIndex        =   29
            Top             =   5730
            Width           =   840
         End
         Begin VB.Label lblWBC 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   4680
            TabIndex        =   28
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label lblLZ 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1860
            TabIndex        =   27
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "白 细 胞:"
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
            Index           =   5
            Left            =   3540
            TabIndex        =   26
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "鳞状上皮细胞:"
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
            Index           =   4
            Left            =   300
            TabIndex        =   25
            Top             =   630
            Width           =   1560
         End
         Begin VB.Label lblXJ 
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
            Height          =   2280
            Left            =   930
            TabIndex        =   24
            Top             =   1080
            Width           =   8400
         End
         Begin VB.Label lblXT 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   4680
            TabIndex        =   23
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblXZ 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1860
            TabIndex        =   22
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "细菌:"
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
            Index           =   6
            Left            =   300
            TabIndex        =   21
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "镜下形态:"
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
            Index           =   3
            Left            =   3540
            TabIndex        =   20
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性        状:"
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
            Left            =   300
            TabIndex        =   19
            Top             =   240
            Width           =   1560
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   300
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceReprotBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'创    建:蔡青松
'创建时间:2017/4/25
'模块功能:一个类似病例的检验报告查看模块
'---------------------------------------------------------------------------------------

Option Explicit

'动态设置是否显示窗体边框
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const const_PicRectBackColour As Long = &HE0E0E0

Private mobjFrm As Object

Private mlngAdvicID As String
Private mlngSampleID As Long        '标本号
Private mlngPaintID As Long         '病人ID
Private mintVersion As Integer      '版本号 10=老版,25=新版
Private mstrPrivs As String         '模块权限
Private mblnHaveBoder As Boolean    '是否显示窗体边框和按钮
Private mblnDoctorShow As Boolean   '是否是医生站浏览
Private mstrSupplementID As String  '补充报告指标ID

Private mlngVsfHeight As Long       'VSF的高度
Private mlngElseCrlHeight As Long   '其他控件的高度

Private mobjFTP As New clsFtp               'FTP对象
Private mblnFtp As Boolean                  'FTP是否可用
Private mstrFtpIp As String                 'FTP连接地址
Private mstrFtpUser As String               'FTP用户
Private mstrFtpPwd As String                'FTP密码
Private mstrFtpFolder As String             'FTP目录

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/25
'功    能:打开窗体
'           strAdvices      医嘱ID串，用“,”分割
'           intType         是否直接预览报告
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function ShowMe(objFrm As Object, ByVal lngAdvicID As Long, ByVal intType As Integer) As Boolean
    mblnHaveBoder = True
    mlngAdvicID = lngAdvicID
    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 2001)
    If GetSampleInformation Then
        If intType = 1 Then
            Call PrintReport(objFrm, 1)
        Else
            Me.Show vbModal, objFrm
        End If
        ShowMe = True
    Else
        Unload Me
    End If
End Function



'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-27
'功    能:  通过标本ID显示报告（报告查询模块调用）
'入    参:
'           objFrm          调用窗体
'           mblnDoctorShow  是否是医生站调用
'           lngSampleID     标本ID
'           lngPaintID      病人ID
'           intVersion      报告版本，25=新版LIS，10=老版LIS
'           intSampleType   是否是微生物报告，0=普通报告，1=微生物报告
'           intPositive     报告类型，1=药敏报告，3=PDF报告，其他=阴性报告
'           strDiagnosis    诊断
'           strResult       备注
'           intCount        老版LIS结果次数
'           strSupplementID 补充报告指标ID
'           strPrivs        人员权限
'出    参:
'           strThirdReport  三方报告
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function ShowReportByID(objFrm As Object, ByVal blnDoctorShow As Boolean, ByVal lngPaintID As Long, ByVal lngSampleID As Long, ByVal intVersion As Long, _
                               ByVal intSampleType As Integer, Optional ByVal intPositive As Integer, _
                               Optional ByVal strDiagnosis As String, Optional ByVal strResult As String, _
                               Optional ByVal intCount As Integer, Optional ByVal strSupplementID As String, _
                               Optional ByVal strPrivs As String, Optional ByRef strThirdReport As String) As Long
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo ShowReportByID_Error

2         Call YSystemMenu(Me.hWnd)

3         mstrPrivs = strPrivs
4         mblnHaveBoder = False
5         mblnDoctorShow = blnDoctorShow
6         mlngPaintID = lngPaintID
7         mlngSampleID = lngSampleID
8         mintVersion = intVersion
9         mstrSupplementID = strSupplementID
10        Set mobjFrm = objFrm

11        mlngVsfHeight = 0
12        mlngElseCrlHeight = 0

          '查询组合项目
13        If intVersion = 25 Then
14            strSQL = "Select Distinct a.申请项目 || '(' || to_char(a.申请时间, 'yyyy/mm/dd hh24:mi:Ss') || '，' || a.标本类型 || ')' 组合名称,a.核收时间,a.是否传染病,a.复核人" & vbCrLf & _
                     "   From 检验报告记录 A" & vbCrLf & _
                     "   Where a.id = [1]"
15            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "组合项目", lngSampleID)
16        ElseIf intVersion = 10 Then
17            strSQL = "Select f_List2str(Cast(Collect(b.医嘱内容) As t_Strlist)) || '(' || to_char(a.申请时间, 'yyyy/mm/dd hh24:mi:Ss') || '，' ||" & vbCrLf & _
                     "           a.标本类型 || ')' 组合名称,a.核收时间,0 是否传染病,'' 复核人" & vbCrLf & _
                     "   From 检验标本记录 A, 病人医嘱记录 B" & vbCrLf & _
                     "   Where a.医嘱id = b.Id(+) and a.id=[1]" & vbCrLf & _
                     "   Group By a.id, a.申请时间, a.标本类型,a.核收时间"
18            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "组合项目", lngSampleID)
19        End If
20        If Not rsTmp.EOF Then
21            lblPatient(0).Caption = rsTmp("组合名称") & IIf(intVersion = 25, "(新版)", "(老版)")
22            If Val(rsTmp("是否传染病") & "") = 1 Then
23                lblPatient(0).Caption = lblPatient(0).Caption & "(疑似传染病)"
24                lblPatient(0).ForeColor = vbRed
25                If rsTmp("复核人") & "" = "" Then
26                    cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = True
27                    cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = False
28                Else
29                    cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = False
30                    cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = True
31                End If
32            Else
33                lblPatient(0).ForeColor = &H80000012
34                cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = False
35                cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = False
36            End If
37        Else
38            lblPatient(0).Caption = "申请项目:"
39        End If
40        lblPatient(0).FontBold = True

41        If intVersion = 10 Then    '老版LIS
42            Call GetSampleFromOldLis(lngSampleID, intSampleType, intCount)
43        ElseIf intVersion = 25 Then    '新版LIS
44            Call GetSampleFromNewLis(lngSampleID, intSampleType, intPositive, strDiagnosis, strResult, strThirdReport)
45        End If

46        ShowReportByID = GetFrmHeight(intSampleType)


47        Exit Function
ShowReportByID_Error:
48        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(ShowReportByID)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
49        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-27
'功    能:  获取窗体的高度
'入    参:
'           intSampleType       报告类型，1=微生物报告
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Private Function GetFrmHeight(ByVal intSampleType As Integer) As Long
    Dim lngElseCrlHeight As Long
    Dim lngVsfHeight As Long
    
    lngElseCrlHeight = mlngElseCrlHeight + picComment.Height
    lngVsfHeight = mlngVsfHeight
    GetFrmHeight = lngElseCrlHeight + lngVsfHeight
    If intSampleType = 1 Then
        If GetFrmHeight < 11000 Then GetFrmHeight = 11000
    End If
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/25
'功    能:调用API动态设置窗体的border
'入    参:
'           new_Hwnd    窗体的句柄
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Exit        '退出
            Unload Me
        Case ConMenu_Browse_PrintView   '预览
            Call PrintReport(Me, 1)
        Case ConMenu_Browse_PrintSet    '打印设置
            Call PrintReport(Me, 3)
        Case ConMenu_Appfor_ClincHelp   '诊疗参考
            Call funShowClincHelp(Me, mlngSampleID, mintVersion)
        Case ConMenu_Browse_Print       '打印
            Call PrintReport(Me, 2)
        Case conFun_Sample_Auditing     '复核
            Call AuditingSample(1)
        Case conFun_Sample_unAuditing     '取消复核
            Call AuditingSample(2)
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
            Call ExePlugIn(Control.Parameter, mlngSampleID)
    End Select
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/4
'功    能:预览，打印设置，打印
'入    参:
'           objfrm          窗体对象
'           byRunMode       1=预览,2=打印，3=打印设置
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub PrintReport(objFrm As Object, ByVal byRunMode As Byte)
    Dim strErr As String

    If mintVersion = 25 Then
        Call PrintNewReport(mobjFrm, mlngSampleID, byRunMode, mblnDoctorShow, mstrPrivs, , strErr)
    Else
        Call PtintOldReport(mobjFrm, mlngSampleID, mlngPaintID, byRunMode, , strErr)
    End If
    If strErr <> "" Then MsgBox strErr, vbInformation, gSysInfo.AppName
End Sub

Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With Me.picMain
'        If mblnHaveBoder Then
            .Left = Left
            .Top = Top
            .Width = Right - Left
            .Height = Bottom - Top
'        Else
'            .Left = 0
'            .Top = 0
'            .Width = Me.Width
'            .Height = Me.Height
'        End If
    End With
    With picPDF
        If mblnHaveBoder Then
            .Left = Left
            .Top = Top
            .Width = Right - Left
            .Height = Bottom - Top
        Else
            .Left = 0
            .Top = 0
            .Width = Me.Width
            .Height = Me.Height
        End If
    End With
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfor_ClincHelp       '诊疗参考
            Control.Visible = VerCompare(gSysInfo.VersionHIS, "10.35.120") <> -1
    End Select
End Sub

Private Sub chkMicroscope_Click()
    PicNegative_Resize
End Sub

Private Sub Form_Load()
'功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------



    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True    '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "打印")
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "打印设置  ")
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_unPrint, "重置打印  ")
            cbrControl.Visible = InStr(mstrPrivs, "重置自助机报告打印次数") > 0
        End With
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "预览")
        Set cbrControl = .Add(xtpControlButton, conFun_Sample_Auditing, "复核"): cbrControl.BeginGroup = True
        cbrControl.Visible = Not mblnDoctorShow
        cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conFun_Sample_unAuditing, "取消复核")
        cbrControl.Visible = Not mblnDoctorShow
        cbrControl.Enabled = False

        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "诊疗参考")
        cbrControl.BeginGroup = True
        If mblnHaveBoder Then
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "退出")
            cbrControl.BeginGroup = True
        End If
    End With

    '创建插件按钮
    Call CreatePlugInButton(cbrToolBar)

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next


    Call intData
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/25
'功    能:初始化数据
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub intData()
    Call setVSF
    Call setTabType '设置分页
    
'    Call GetSampleInformation        '获取医嘱ID对应的标本信息
End Sub

Private Sub setTabType()
    With Me.tabThis
        .PaintManager.Appearance = xtpTabAppearanceStateButtons
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionLeft

        
        .InsertItem 0, "涂片报告", MicroorganismSmear.hWnd, 1
        .PaintManager.Layout = xtpTabLayoutAutoSize
        
        
        .InsertItem 1, "初级药敏", picCJYM.hWnd, 2
        .PaintManager.Layout = xtpTabLayoutAutoSize
        
        .InsertItem 2, "微生物报告", picCenter.hWnd, 3
        .PaintManager.Layout = xtpTabLayoutAutoSize
        
        .Item(2).Selected = True
    End With
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/25
'功    能:初始化VSF列表
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub setVSF()
      '展示普通结果的VSF

1         On Error GoTo setVSF_Error

2         With Me.vsfGeneral
3             .FixedRows = 0
4             .FixedCols = 0
5             .Rows = 1
6             .Cols = 8
7             .SelectionMode = flexSelectionFree  '自由选择
8             .AllowSelection = False    '不允许选择
9             .BorderStyle = flexBorderNone    '无边框
10            .GridLines = flexGridNone    '无网格线
11            .FontSize = 12  '小四
12            .BackColorBkg = vbWhite    '白色背景
13            .SheetBorder = vbWhite  '白色边线
14            .BorderStyle = flexBorderNone
15            .ExplorerBar = flexExSortShowAndMove    '点标题栏排序，并显示排序图标
16            .AllowUserResizing = flexResizeColumns  '可调整列宽
17            .Editable = flexEDNone                  '只读

18            .ColKey(0) = "标本ID": .TextMatrix(0, .ColIndex("标本ID")) = "标本ID": .ColWidth(.ColIndex("标本ID")) = 0: .ColHidden(.ColIndex("标本ID")) = True
19            .ColKey(1) = "检验项目": .TextMatrix(0, .ColIndex("检验项目")) = "检验项目": .ColWidth(.ColIndex("检验项目")) = 4000: .ColHidden(.ColIndex("检验项目")) = False
20            .ColKey(2) = "检验结果": .TextMatrix(0, .ColIndex("检验结果")) = "检验结果": .ColWidth(.ColIndex("检验结果")) = 1100: .ColHidden(.ColIndex("检验结果")) = False
21            .ColKey(3) = "结果单位": .TextMatrix(0, .ColIndex("结果单位")) = "结果单位": .ColWidth(.ColIndex("结果单位")) = 1100: .ColHidden(.ColIndex("结果单位")) = False
22            .ColKey(4) = "标志": .TextMatrix(0, .ColIndex("标志")) = "标志": .ColWidth(.ColIndex("标志")) = 800: .ColHidden(.ColIndex("标志")) = False
23            .ColKey(5) = "结果参考": .TextMatrix(0, .ColIndex("结果参考")) = "结果参考": .ColWidth(.ColIndex("结果参考")) = 2000: .ColHidden(.ColIndex("结果参考")) = False
24            .ColKey(6) = "临床意义": .TextMatrix(0, .ColIndex("临床意义")) = "临床意义": .ColWidth(.ColIndex("临床意义")) = 2000: .ColHidden(.ColIndex("临床意义")) = True
25            .ColKey(7) = "ID": .TextMatrix(0, .ColIndex("ID")) = "ID": .ColWidth(.ColIndex("ID")) = 2000: .ColHidden(.ColIndex("ID")) = True
26            .Cell(flexcpAlignment, 0, .ColIndex("标本ID"), 0, .ColIndex("结果参考")) = flexAlignLeftCenter  '标题靠左对齐
27        End With

          '初级药敏报告
28        With Me.VSFCJYM
29            .FixedRows = 0
30            .FixedCols = 0
31            .Rows = 1
32            .Cols = 6
33            .SelectionMode = flexSelectionFree  '自由选择
34            .AllowSelection = False    '不允许选择
35            .BorderStyle = flexBorderNone    '无边框
36            .GridLines = flexGridNone    '无网格线
37            .FontSize = 12  '小四
38            .BackColorBkg = vbWhite    '白色背景
39            .SheetBorder = vbWhite  '白色边线
40            .BorderStyle = flexBorderNone
41            .AllowUserResizing = flexResizeColumns  '可调整列宽
42            .Editable = flexEDNone                  '只读
43            .MergeCells = flexMergeRestrictRows     '允许横向合并
44            .OutlineBar = flexOutlineBarComplete    '树形结构
45            .OutlineCol = 0    '树形节点列
46            .SubtotalPosition = flexSTAbove    '树形结构样式

47            .ColKey(0) = "细菌名": .ColWidth(.ColIndex("细菌名")) = 3000: .ColAlignment(.ColIndex("细菌名")) = flexAlignLeftCenter
48            .ColKey(1) = "检验结果": .ColWidth(.ColIndex("检验结果")) = 1500: .ColAlignment(.ColIndex("检验结果")) = flexAlignLeftCenter
49            .ColKey(2) = "描述": .ColWidth(.ColIndex("描述")) = 1500: .ColAlignment(.ColIndex("描述")) = flexAlignLeftCenter
50            .ColKey(3) = "耐药机制": .ColWidth(.ColIndex("耐药机制")) = 1500: .ColAlignment(.ColIndex("耐药机制")) = flexAlignLeftCenter
51            .ColKey(4) = "参考描述": .ColWidth(.ColIndex("参考描述")) = 1500: .ColAlignment(.ColIndex("参考描述")) = flexAlignLeftCenter
52            .ColKey(5) = "Level": .ColWidth(.ColIndex("Level")) = 1500: .ColAlignment(.ColIndex("Level")) = flexAlignLeftCenter: .ColHidden(.ColIndex("Level")) = True

              '固定行
53            .TextMatrix(0, .ColIndex("细菌名")) = "细菌名"
54            .TextMatrix(0, .ColIndex("检验结果")) = "检验结果"
55            .TextMatrix(0, .ColIndex("描述")) = "描述"
56            .TextMatrix(0, .ColIndex("耐药机制")) = "耐药机制"
57            .TextMatrix(0, .ColIndex("参考描述")) = "参考描述"
58        End With

          '展示微生物结果的vsf
59        With Me.vsfMicrobePositive
60            .FixedRows = 0
61            .FixedCols = 0
62            .Rows = 1
63            .Cols = 6
64            .SelectionMode = flexSelectionFree  '自由选择
65            .AllowSelection = False    '不允许选择
66            .BorderStyle = flexBorderNone    '无边框
67            .GridLines = flexGridNone    '无网格线
68            .FontSize = 12  '小四
69            .BackColorBkg = vbWhite    '白色背景
70            .SheetBorder = vbWhite  '白色边线
71            .BorderStyle = flexBorderNone
72            .AllowUserResizing = flexResizeColumns  '可调整列宽
73            .Editable = flexEDNone                  '只读
74            .MergeCells = flexMergeRestrictRows     '允许横向合并
75            .OutlineBar = flexOutlineBarComplete    '树形结构
76            .OutlineCol = 2    '树形节点列
77            .SubtotalPosition = flexSTAbove    '树形结构样式

78            .ColKey(0) = "KEY": .TextMatrix(0, .ColIndex("KEY")) = "KEY": .ColWidth(.ColIndex("KEY")) = 0: .ColHidden(.ColIndex("KEY")) = True
79            .ColKey(1) = "抗生素ID": .TextMatrix(0, .ColIndex("抗生素ID")) = "抗生素ID": .ColWidth(.ColIndex("抗生素ID")) = 0: .ColHidden(.ColIndex("抗生素ID")) = True
80            .ColKey(2) = "抗生素名称": .TextMatrix(0, .ColIndex("抗生素名称")) = "细菌名": .ColWidth(.ColIndex("抗生素名称")) = 4000: .ColHidden(.ColIndex("抗生素名称")) = False
81            .ColKey(3) = "检验结果": .TextMatrix(0, .ColIndex("检验结果")) = "检验结果": .ColWidth(.ColIndex("检验结果")) = 2000: .ColHidden(.ColIndex("检验结果")) = False
82            .ColKey(4) = "结果类型": .TextMatrix(0, .ColIndex("结果类型")) = "描述": .ColWidth(.ColIndex("结果类型")) = 1300: .ColHidden(.ColIndex("结果类型")) = False
83            .ColKey(5) = "药敏方法": .TextMatrix(0, .ColIndex("药敏方法")) = "耐药机制": .ColWidth(.ColIndex("药敏方法")) = 1300: .ColHidden(.ColIndex("药敏方法")) = False
84            .Cell(flexcpAlignment, 0, .ColIndex("KEY"), 0, .ColIndex("药敏方法")) = flexAlignLeftCenter  '标题靠左对齐
85        End With

86        Exit Sub
setVSF_Error:
87        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(setVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
88        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/2
'功    能:获取标本信息，包括标本存在新版中还是老版中，标本号，是否是微生物标本等
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Function GetSampleInformation() As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim lngSampleID As Long           '标本ID
          Dim intSampleType As Integer      '标本类型，0=普通标本,1=微生物标本
          Dim intVersion As Integer         '版本，25=新版，10=老版
          Dim intCount As Integer           '检验次数
          Dim intPositive As Integer        '报告类型，1=药敏报告,3=PDF报告
          Dim strDiagnosis As String        '诊断
          Dim strResult As String           '备注
          Dim strSQR As String              '申请人
          Dim intIsDis As Integer           '是否是传染病
          
          '判断对应医嘱ID这老版中是否存在标本,及标本是否为微生物标本
1         On Error GoTo GetSampleInformation_Error

2         strSQL = "select a.id 标本ID,a.微生物标本,a.报告结果,a.病人ID from 检验标本记录 A where a.医嘱id=[1] and 审核人 is not null"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "标本来源", mlngAdvicID)
4         If rsTmp.RecordCount > 0 Then
5             intVersion = 10 '标本在老版LIS中
6             lngSampleID = Val(rsTmp("标本ID") & "") '获取标本ID
7             mlngPaintID = Val(rsTmp("病人ID") & "")
8             intSampleType = Val(rsTmp("微生物标本") & "")  '获取标本类型
9             intCount = Val(rsTmp("报告结果") & "")  '标本检验次数
10        Else
              '老版中没有查询到医嘱相关的就就到新版LIS中去查找
11            strSQL = "select b.id 标本ID,b.微生物,b.阳性报告,b.诊断,b.备注,a.申请人,b.是否传染病 from 检验申请组合 A,检验报告记录 B" & _
                      " where a.标本id=b.id and a.申请id=[1]  and b.审核人 is not null"
12            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "标本来源", mlngAdvicID)
13            If rsTmp.RecordCount > 0 Then
14                intVersion = 25 '标本在新版LIS中
15                lngSampleID = Val(rsTmp("标本ID") & "") '获取标本ID
16                intSampleType = Val(rsTmp("微生物") & "") '获取标本类型
17                intPositive = Val(rsTmp("阳性报告") & "") '报告类型
18                strDiagnosis = rsTmp("诊断") & ""   '诊断
19                strResult = rsTmp("备注") & ""   '备注
20                strSQR = rsTmp("申请人") & ""
21                intIsDis = Val(rsTmp("是否传染病") & "")
22            Else
23                GetSampleInformation = False
24                Exit Function
25            End If
26        End If
27        mintVersion = intVersion
28        mlngSampleID = lngSampleID
          
          '检查当前用户是否能够查看传染病报告
29        If strSQR <> gUserInfo.Name And strSQR <> "" And InStr(";" & mstrPrivs & ";", ";查看传染病报告;") <= 0 And intIsDis = 1 Then
30            If Me.Tag = "" Then
31                MsgBox "权限不足，无法查看此报告", vbInformation, Me.Caption
32                Me.Tag = "True"
33            End If
34            Exit Function
35        End If
              
          '查询病人信息
36        If mintVersion <= 0 Or lngSampleID <= 0 Then Exit Function
37        If GetPatient(intVersion, lngSampleID) = False Then Exit Function

          '查询标本记录
38        If intVersion = 10 Then '老版LIS
39            Call GetSampleFromOldLis(lngSampleID, intSampleType, intCount)
40        ElseIf intVersion = 25 Then '新版LIS
41            Call GetSampleFromNewLis(lngSampleID, intSampleType, intPositive, strDiagnosis, strResult)
42        End If

43        GetSampleInformation = True
          
44        Exit Function
GetSampleInformation_Error:
45        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(GetSampleInformation)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
46        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/2
'功    能:读取人员基本信息
'入    参:
'           intVersion          版本号
'           lngSampleType       标本号
'出    参:
'返    回:  True=查询到病人信息,False=没有查询到病人信息
'---------------------------------------------------------------------------------------
Private Function GetPatient(ByVal intVersion As Integer, ByVal lngSampleID As Long) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo GetPatient_Error

2         If intVersion = 10 Then '老版
3             strSQL = "select 姓名,性别,年龄,床号 from 检验标本记录 where ID=[1]"
4             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "人员信息", lngSampleID)
5         ElseIf intVersion = 25 Then '新版
6             strSQL = "select 姓名,decode(性别,1,'男',2,'女','未告知') 性别,年龄,床号 from 检验报告记录 where ID=[1]"
7             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "人员信息", lngSampleID)
8         End If

          '如果查询到数据则绑定数据，没有数据则返回False并退出
9         If rsTmp.RecordCount > 0 Then
10            Me.lblPatient(0).Caption = "姓名:" & rsTmp("姓名")
11            Me.lblPatient(1).Caption = "性别:" & rsTmp("性别")
12            Me.lblPatient(2).Caption = "年龄:" & rsTmp("年龄")
13            Me.lblPatient(3).Caption = "床号:" & rsTmp("床号")
14        Else
15            GetPatient = False
16            Exit Function
17        End If
          
18        GetPatient = True


19        Exit Function
GetPatient_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(GetPatient)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/2
'功    能:从老版LIS中查询标本记录
'入    参:
'           lngSampleID             标本ID
'           intSampleType           标本类型
'           intCount                标本检验次数
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub GetSampleFromOldLis(ByVal lngSampleID As Long, ByVal intSampleType As Integer, ByVal intCount As Integer)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsAntibiotic As ADODB.Recordset
          
          '检验备注
1         On Error GoTo GetSampleFromOldLis_Error

2         strSQL = "SELECT A.备注 FROM 检验标本记录 A WHERE A.ID= [1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验评语", lngSampleID)
4         If rsTmp.RecordCount > 0 Then
5             Me.txtResult.Text = rsTmp("备注") & ""
6             rsTmp.MoveNext
7         End If
          
          '诊断
8         strSQL = "Select b.医嘱id, b.项目, b.排列, b.内容 From 检验标本记录 a, 病人医嘱附件 b Where a.医嘱id = b.医嘱id and a.ID =[1] Order By 医嘱id, 排列"
9         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊断", lngSampleID)
10        Me.txtDiagnosis.Text = ""
11        If rsTmp.RecordCount > 0 Then
12            Do While Not rsTmp.EOF
13                Me.txtDiagnosis.Text = Me.txtDiagnosis.Text & NVL(rsTmp("项目")) & ":" & Replace(NVL(rsTmp("内容")), vbCrLf, vbCrLf & "    ") & vbCrLf
14                rsTmp.MoveNext
15            Loop
16        End If
          
         
17        If intSampleType = 0 Then
              '普通标本
18            strSQL = "Select b.检验项目ID ID, a.Id As 标本id, c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目, b.检验结果,d.单位 as 结果单位," & _
                       "      Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志,b.结果参考,D.临床意义" & _
                      " From 检验标本记录 A, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 诊疗项目目录 H, 检验流水线指标 E" & _
                      " Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And b.诊疗项目id = h.Id(+) And b.检验标本id = e.标本id(+) And" & _
                      "      b.检验项目id = e.项目id(+) And b.记录类型 = [1] And a.Id = [2]" & _
                      " Union All" & _
                      " Select b.检验项目ID ID, a.Id As 标本id, c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目, b.检验结果,d.单位 as 结果单位," & _
                      "       Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志,b.结果参考,D.临床意义" & _
                      " From 检验标本记录 A, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 诊疗项目目录 H, 检验流水线指标 E" & _
                      " Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And b.诊疗项目id = h.Id(+) And b.记录类型 = [1] And" & _
                      "      b.检验标本id = e.标本id(+) And b.检验项目id = e.项目id(+) And a.合并id = [2]"
19            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验普通结果", intCount, lngSampleID)
20            Call SetGeneralData(rsTmp)
21        ElseIf intSampleType = 1 Then
              '微生物标本
              
              '查询细菌
22            strSQL = "Select b.Id, b.中文名 As 细菌名, a.检验结果 As 检验结果, a.培养描述 As 描述, a.耐药机制,'' 阳性评语" & _
                       " From 检验普通结果 A, 检验细菌 B, 检验标本记录 D" & _
                       " Where a.细菌id = b.Id And a.记录类型 = [1] And d.Id = a.检验标本id And d.Id = [2]" & _
                       " Order By b.编码"
23            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验细菌", intCount, lngSampleID)

              '查询药敏
24            strSQL = "Select c.细菌id As Key, b.Id 抗生素ID, b.中文名 As 抗生素名称, a.结果 As 检验结果," & _
                       "      Decode(a.结果类型, 'R', 'R-耐药', 'I', 'I-中介', 'S', 'S-敏感', a.结果类型) As 结果类型," & _
                       "      Decode(a.药敏方法, 1, '1-MIC', 2, '2-DISK', 3, '3-K-B', '') As 药敏方法" & _
                       " From 检验药敏结果 A, 检验用抗生素 B, 检验普通结果 C" & _
                       " Where a.抗生素id = b.Id And c.Id = a.细菌结果id And c.记录类型 = a.记录类型 And c.检验标本id = [1] And c.记录类型 = [2]" & _
                       " Order By c.细菌id, b.编码"
25            Set rsAntibiotic = ComOpenSQL(Sel_His_DB, strSQL, "检验药敏", lngSampleID, intCount)
26            Call SetMicroorganismData(rsTmp, rsAntibiotic)    '绑定数据
27        End If
          
28        Call SetCrlTyep(intSampleType, 1) '设置窗体中要显示哪些控件



29        Exit Sub
GetSampleFromOldLis_Error:
30        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(GetSampleFromOldLis)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
31        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/2
'功    能:从新版LIS中查询标本记录
'入    参:
'           lngSampleID             标本ID
'           intSampleType           标本类型
'           intPositive             报告类型，1=药敏报告，3=PDF报告
'           strDiagnosis            临床诊断
'           strThirdReport          三方PDF报告
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Function GetSampleFromNewLis(ByVal lngSampleID As Long, ByVal intSampleType As Integer, _
                                     ByVal intPositive As Integer, ByVal strDiagnosis As String, _
                                     ByVal strResult As String, Optional ByRef strThirdReport As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsAntibiotic As ADODB.Recordset
          Dim lngRow As Long


1         On Error GoTo GetSampleFromNewLis_Error


          '备注
2         If strResult <> "" Then
3             Me.txtResult.Text = strResult
4         End If

          '诊断
5         If strDiagnosis <> "" Then
6             Me.txtDiagnosis.Text = strDiagnosis
7         End If

8         picMain.Visible = True
9         picPDF.Visible = False
10        If intSampleType = 0 Then
              '普通标本
11            If IsTre(lngSampleID) Then
12                strSQL = "select * from (Select Distinct c.id, a.Id 标本id,b.id 报告明细ID, c.中文名 || '(' || c.英文名 || ')' || decode(h.耐受时间,null,'', '(' || h.耐受时间 || ')')  检验项目, b.检验结果, c.单位 结果单位," & _
                         "               Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') 标志, b.结果参考,c.临床意义" & _
                         " From 检验报告记录 A, 检验报告明细 B, 检验指标 C, 检验组合项目 D, 检验申请组合 E, 流水线检验指标 F,耐受试验标本 G,检验耐受时间方案 H" & _
                         " Where a.Id = b.标本id And b.项目id = c.Id And b.组合id = d.Id(+) And b.标本id = f.标本id(+) And b.项目id = f.项目id(+) And" & _
                         "      b.标本id = e.标本id And d.Id = e.组合id and b.ID=g.报告明细id(+) and g.耐受方案id=H.id(+) And b.组合id Is Not Null And e.组合id Is Not Null And a.Id = [1]" & _
                         " Union All" & _
                         " Select Distinct c.id, a.Id 标本id,b.id 报告明细ID,  c.中文名 || '(' || c.英文名 || ')' || decode(h.耐受时间,null,'', '(' || h.耐受时间 || ')') 检验项目, b.检验结果, c.单位 结果单位," & _
                         "                Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') 标志, b.结果参考,c.临床意义" & _
                         " From 检验报告记录 A, 检验报告明细 B, 检验指标 C, 检验组合项目 D, 检验申请组合 E, 流水线检验指标 F,耐受试验标本 G,检验耐受时间方案 H" & _
                         " Where a.Id = b.标本id And b.项目id = c.Id And b.组合id = d.Id(+) And b.标本id = f.标本id(+) And b.项目id = f.项目id(+) And" & _
                         "      b.标本id = e.标本id and b.ID=g.报告明细id(+) and g.耐受方案id=H.id(+) And e.组合id Is Null And b.组合id Is Null And a.Id =[1] ) order by 报告明细ID desc"
13            Else
14                strSQL = "select * from (Select Distinct c.id, a.Id 标本id,b.id 报告明细ID, c.中文名 || '(' || c.英文名 || ')'   检验项目, b.检验结果, c.单位 结果单位," & _
                         "               Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') 标志, b.结果参考,c.临床意义" & _
                         " From 检验报告记录 A, 检验报告明细 B, 检验指标 C, 检验组合项目 D, 检验申请组合 E, 流水线检验指标 F" & _
                         " Where a.Id = b.标本id And b.项目id = c.Id And b.组合id = d.Id(+) And b.标本id = f.标本id(+) And b.项目id = f.项目id(+) And" & _
                         "      b.标本id = e.标本id And d.Id = e.组合id and b.组合id Is Not Null And e.组合id Is Not Null And a.Id = [1]" & _
                         " Union All" & _
                         " Select Distinct c.id, a.Id 标本id,b.id 报告明细ID,  c.中文名 || '(' || c.英文名 || ')'  检验项目, b.检验结果, c.单位 结果单位," & _
                         "                Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') 标志, b.结果参考,c.临床意义" & _
                         " From 检验报告记录 A, 检验报告明细 B, 检验指标 C, 检验组合项目 D, 检验申请组合 E, 流水线检验指标 F" & _
                         " Where a.Id = b.标本id And b.项目id = c.Id And b.组合id = d.Id(+) And b.标本id = f.标本id(+) And b.项目id = f.项目id(+) And" & _
                         "      b.标本id = e.标本id and e.组合id Is Null And b.组合id Is Null And a.Id =[1] ) order by 报告明细ID desc"
15            End If
16            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告明细", lngSampleID)
17            Call SetGeneralData(rsTmp)

              '在展示标本结果的VSF中添加删除线
18            With vsfGeneral
19                For lngRow = 1 To .Rows - 1
20                    If InStr("," & mstrSupplementID & ",", "," & .TextMatrix(lngRow, .ColIndex("ID")) & ",") > 0 Then
21                        vsfGeneral.Cell(flexcpFontStrikethru, lngRow, 0, lngRow, vsfGeneral.Cols - 1) = True
22                    End If
23                Next
24            End With

25        ElseIf intSampleType = 1 Then
26            If intPositive = 1 Then
                  '微生物阳性报告
27                strSQL = "Select b.Id, b.中文名 || '(' || b.英文名 || ')' 细菌名, a.检验结果, a.培养描述 描述, a.耐药机制,a.阳性评语" & _
                         " From 检验报告细菌 A, 检验细菌记录 B" & _
                         " Where a.细菌id = b.Id(+) And a.标本id = [1]"
28                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验细菌", lngSampleID)

                  '查询药敏
29                strSQL = "Select a.细菌id Key, c.Id 抗生素id, c.中文名 || '(' || c.英文名 || ')' 抗生素名称, b.结果 检验结果, b.结果类型, b.药敏方法" & _
                         " From 检验报告细菌 A, 检验报告药敏 B, 检验药敏 C, 检验药敏组用药 D" & _
                         " Where a.Id = b.结果id And b.药敏id = c.Id And b.药敏id = d.药敏id(+) And b.药敏组id = d.药敏组id(+) And a.标本id = [1]" & _
                         " Order By d.药敏组id, d.排列序号"
30                Set rsAntibiotic = ComOpenSQL(Sel_Lis_DB, strSQL, "检验药敏", lngSampleID)
31                Call SetMicroorganismData(rsTmp, rsAntibiotic)  '绑定数据
32            ElseIf intPositive = 3 Then
                  'PDF报告
33                picMain.Visible = False
34                picPDF.Visible = True
35                strThirdReport = findThirdReport(lngSampleID, webSub)
36            Else
                  '微生物阴性报告
37                strSQL = "Select  a.正常菌, a.未检出, a.补充描述, a.无致病菌, a.无细菌," & _
                           "A.镜检设备 , A.镜检检出, A.镜检未检出, A.阴性评语,a.是否镜检结果,a.结果性质" & _
                         " From 检验报告细菌 A, 检验细菌记录 B" & _
                         " Where a.细菌id = b.Id(+) And a.标本id = [1]"
38                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "阴性报告", lngSampleID)
39                Call SetPositiveData(rsTmp)
40            End If
41            Call GetfrmMicroorganismSmear(lngSampleID)  '查询涂片报告
42            Call GetMicroorganisCJYM(lngSampleID)       '初步药敏报告
43        End If

44        Call SetCrlTyep(intSampleType, intPositive)    '设置窗体中要显示那些控件

45        Exit Function
GetSampleFromNewLis_Error:
46        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(GetSampleFromNewLis)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
47        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/2
'功    能:绑定普通结果
'入    参:
'           rsTmp           数据纪录集
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SetGeneralData(ByVal rsTmp As ADODB.Recordset)

1         On Error GoTo SetGeneralData_Error

2         If Not rsTmp Is Nothing Then
3           If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            
4           With Me.vsfGeneral
                '绑定数据
5               .Rows = 1
6               Do While Not rsTmp.EOF
7                   .Rows = .Rows + 1
8                   .TextMatrix(.Rows - 1, .ColIndex("标本ID")) = rsTmp("标本ID") & ""
9                   .TextMatrix(.Rows - 1, .ColIndex("检验项目")) = rsTmp("检验项目") & ""
10                  .TextMatrix(.Rows - 1, .ColIndex("检验结果")) = rsTmp("检验结果") & ""
11                  .TextMatrix(.Rows - 1, .ColIndex("结果单位")) = rsTmp("结果单位") & ""
12                  .TextMatrix(.Rows - 1, .ColIndex("标志")) = rsTmp("标志") & ""
13                  .TextMatrix(.Rows - 1, .ColIndex("结果参考")) = rsTmp("结果参考") & ""
14                  .TextMatrix(.Rows - 1, .ColIndex("临床意义")) = rsTmp("临床意义") & ""
15                  .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
16                  .Cell(flexcpAlignment, .Rows - 1, .ColIndex("标本ID"), .Rows - 1, .ColIndex("临床意义")) = flexAlignLeftCenter  '内容靠左对齐
17                  rsTmp.MoveNext
18              Loop
          
19              lbl(0).Caption = "备注："
                
                '获取VSF高度
20              If mblnHaveBoder = False Then
21                  If mlngVsfHeight < (.Rows + 7) * .RowHeight(0) Then
22                      mlngVsfHeight = (.Rows + 7) * .RowHeight(0)
23                  End If
24              End If
25          End With
26        End If


27        Exit Sub
SetGeneralData_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(SetGeneralData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/2
'功    能:绑定微生物结果
'入    参:
'           rsBacteria          细菌纪录集
'           rsAntibiotic        药敏纪录集
'           intVersion          版本，10=老版，25=新版
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SetMicroorganismData(ByVal rsBacteria As ADODB.Recordset, ByVal rsAntibiotic As ADODB.Recordset)
          Dim lngKey As Long
          Dim lngRow As Long
          Dim lngRowCount As Long
          Dim lngRowBegin As Long

1         On Error GoTo SetMicroorganismData_Error

2         If rsBacteria.RecordCount > 0 Then rsBacteria.MoveFirst
3         If rsAntibiotic.RecordCount > 0 Then rsAntibiotic.MoveFirst
4         txtResult.Text = ""
5         With Me.vsfMicrobePositive
6             .Rows = 1
7             Do While Not rsBacteria.EOF
8                 .Rows = .Rows + 2

                  '绑定细菌数据
9                 .TextMatrix(.Rows - 2, .ColIndex("KEY")) = rsBacteria("ID") & ""
10                .TextMatrix(.Rows - 2, .ColIndex("抗生素名称")) = rsBacteria("细菌名") & ""
11                .TextMatrix(.Rows - 2, .ColIndex("检验结果")) = rsBacteria("检验结果") & ""
12                .TextMatrix(.Rows - 2, .ColIndex("结果类型")) = rsBacteria("描述") & ""
13                .TextMatrix(.Rows - 2, .ColIndex("药敏方法")) = rsBacteria("耐药机制") & ""
14                txtResult.Text = txtResult.Text & rsBacteria("阳性评语") & ""
15                .Cell(flexcpAlignment, .Rows - 2, .ColIndex("KEY"), .Rows - 2, .ColIndex("药敏方法")) = flexAlignLeftCenter  '内容靠左对齐

                  '缩进
16                .IsSubtotal(.Rows - 2) = True   '设置为树形节点
17                .RowOutlineLevel(.Rows - 2) = 3
                  '显示边框线
18                .CellBorderRange .Rows - 2, .ColIndex("抗生素名称"), .Rows - 2, .ColIndex("药敏方法"), vbBlack, 0, 0, 0, 1, 0, 0

                  '设置抗生素标题栏
19                .TextMatrix(.Rows - 1, .ColIndex("抗生素名称")) = "抗生素名称"
20                .TextMatrix(.Rows - 1, .ColIndex("检验结果")) = "检验结果"
21                .TextMatrix(.Rows - 1, .ColIndex("结果类型")) = "结果类型"
22                .TextMatrix(.Rows - 1, .ColIndex("药敏方法")) = "药敏方法"

                  '根据细菌绑定药敏
23                lngKey = Val(rsBacteria("ID") & "")
24                rsAntibiotic.Filter = "KEY=" & lngKey
25                lngRowBegin = 0
26                lngRowCount = 0
27                Do While Not rsAntibiotic.EOF
28                    .Rows = .Rows + 1
29                    lngRowCount = lngRowCount + 1
30                    If lngRowBegin = 0 Then lngRowBegin = .Rows - 2
31                    .TextMatrix(.Rows - 1, .ColIndex("KEY")) = rsAntibiotic("KEY") & ""
32                    .TextMatrix(.Rows - 1, .ColIndex("抗生素ID")) = rsAntibiotic("抗生素ID") & ""
33                    .TextMatrix(.Rows - 1, .ColIndex("抗生素名称")) = rsAntibiotic("抗生素名称") & ""
34                    .TextMatrix(.Rows - 1, .ColIndex("检验结果")) = rsAntibiotic("检验结果") & ""
35                    .TextMatrix(.Rows - 1, .ColIndex("结果类型")) = rsAntibiotic("结果类型") & ""
36                    .TextMatrix(.Rows - 1, .ColIndex("药敏方法")) = rsAntibiotic("药敏方法") & ""
37                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("KEY"), .Rows - 1, .ColIndex("药敏方法")) = flexAlignLeftCenter  '内容靠左对齐

38                    rsAntibiotic.MoveNext
39                Loop
40                rsBacteria.MoveNext
41            Loop


              '树形结构默展开
42            For lngRow = 0 To .Rows - 1
43                If .IsSubtotal(lngRow) = True Then
44                    .IsCollapsed(lngRow) = flexOutlineExpanded    '展开树形
45                End If
46            Next

              '获取VSF高度
47            If mblnHaveBoder = False Then
48                If mlngVsfHeight < (.Rows + 5) * .RowHeight(0) Then
49                    mlngVsfHeight = (.Rows + 5) * .RowHeight(0)
50                End If
51            End If
52        End With

53        Exit Sub
SetMicroorganismData_Error:
54        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(SetMicroorganismData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
55        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/3
'功    能:绑定微生物阴性报告
'入    参:
'           rsTmp           阴性报告纪录集
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SetPositiveData(ByVal rsTmp As ADODB.Recordset)

1         On Error GoTo SetPositiveData_Error

2         If rsTmp.RecordCount > 0 Then
3             rsTmp.MoveFirst
4             txtNormalMicrobe.Text = rsTmp("正常菌") & ""
5             txtNoFindMicrobe.Text = rsTmp("未检出") & ""
6             txtNormalMicrobes.Text = rsTmp("补充描述") & ""
7             chkPathopoiesiaGerm.value = IIf(Val(rsTmp("无致病菌") & "") = 1, 1, 0)
8             chkNoGerm.value = IIf(Val(rsTmp("无细菌") & "") = 1, 1, 0)
9             txtMicroscope.Text = rsTmp("镜检设备") & ""
10            txtMicroscopeNOFind.Text = rsTmp("镜检未检出") & ""
11            txtMicroscopeFinded.Text = rsTmp("镜检检出") & ""
12          If Val(rsTmp("是否镜检结果") & "") = 0 Then
13              chkMicroscope.value = 0
14          Else
15              chkMicroscope.value = 1
16          End If
17          If Val(rsTmp("结果性质") & "") = 0 Then
18              optReport(1).value = True
19          Else
20              optReport(0).value = True
21          End If
22          optReportShow

23            txtResult.Text = rsTmp("阴性评语") & ""
24        End If
          
          '阴性报告高度
25        mlngElseCrlHeight = 6000

26        Exit Sub
SetPositiveData_Error:
27        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(SetPositiveData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
28        Err.Clear
End Sub

Private Sub optReportShow()
    If optReport(0).value = True Then
        txtNormalMicrobe.ForeColor = vbRed
        txtNormalMicrobe.FontBold = True
        txtNoFindMicrobe.ForeColor = vbRed
        txtNoFindMicrobe.FontBold = True
        txtNormalMicrobes.ForeColor = vbRed
        txtNormalMicrobes.FontBold = True
        txtMicroscope.ForeColor = vbRed
        txtMicroscope.FontBold = True
        txtMicroscopeFinded.ForeColor = vbRed
        txtMicroscopeFinded.FontBold = True
        txtMicroscopeNOFind.ForeColor = vbRed
        txtMicroscopeNOFind.FontBold = True
    Else
        txtNormalMicrobe.ForeColor = vbBlack
        txtNormalMicrobe.FontBold = False
        txtNoFindMicrobe.ForeColor = vbBlack
        txtNoFindMicrobe.FontBold = False
        txtNormalMicrobes.ForeColor = vbBlack
        txtNormalMicrobes.FontBold = False
        txtMicroscope.ForeColor = vbBlack
        txtMicroscope.FontBold = False
        txtMicroscopeFinded.ForeColor = vbBlack
        txtMicroscopeFinded.FontBold = False
        txtMicroscopeNOFind.ForeColor = vbBlack
        txtMicroscopeNOFind.FontBold = False
    End If

End Sub


'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/3
'功    能:设置控件的状态
'入    参:
'           intSampleType       标本状态，0=普通标本,1=微生物标本
'           intPositive         报告类型,1=药敏报告，3=PDF报告
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SetCrlTyep(ByVal intSampleType As Integer, ByVal intPositive As Integer)
    If intSampleType = 0 Then   '普通报告
        Me.picTab.Visible = False
        Me.picGeneral.Visible = True
    ElseIf intSampleType = 1 Then   '微生物报告
        Me.picTab.Visible = True
        Me.picGeneral.Visible = False
        If intPositive <> 1 Then '阴性报告
            picMicrobePositive.Visible = False
            PicNegative.Visible = True
        ElseIf intPositive = 1 Then     '阳性报告
            picMicrobePositive.Visible = True
            PicNegative.Visible = False
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/8/1
'功    能:获取微生物镜检结果
'入    参:
'           lngSmapleID     标本ID
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub GetfrmMicroorganismSmear(ByVal lngSmapleID As Long)
          Dim objFSO As New FileSystemObject
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim intloop As Integer
          Dim strFloder As String
          Dim strImgPath As String
          Dim strImgData As String
          Dim intNO As Integer
          Dim intReturn As Integer

1         On Error GoTo GetfrmMicroorganismSmear_Error

2         Call ConnFtp        '检查FTP是否可用

3         strSQL = "Select a.性状, a.镜下形态, a.鳞状上皮细胞, a.白细胞, a.检出细菌, a.审核人, a.审核时间" & vbCrLf & _
                 "       From 微生物涂片报告 A, 微生物涂片细菌 B, 检验细菌记录 C" & vbCrLf & _
                 "       Where a.标本id = b.标本id(+) And b.细菌id = c.Id(+) And a.标本id = [1]" & IIf(mblnDoctorShow, " And a.审核人 Is Not Null", "")
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "微生物涂片报告", lngSmapleID)
5         If rsTmp.RecordCount <= 0 Then Exit Sub
6         lblXZ.Caption = rsTmp("性状") & ""
7         lblXT.Caption = rsTmp("镜下形态") & ""
8         lblLZ.Caption = rsTmp("鳞状上皮细胞") & ""
9         lblWBC.Caption = rsTmp("白细胞") & ""
10        lblXJ.Caption = Replace(Replace(rsTmp("检出细菌") & "", ",", vbCrLf), "()", "")
11        lblAuditingMan.Caption = rsTmp("审核人") & ""
12        lblAuditingTime.Caption = rsTmp("审核时间") & ""

          '查询操作过程
13        imgPicture(0).Tag = ""
14        imgPicture(1).Tag = ""
15        imgPicture(2).Tag = ""
16        strFloder = App.Path & "\MicroorganismPicture"
17        strSQL = "select b.id,序号, b.图像位置 from 微生物涂片报告 A,微生物镜检过程 B where a.id=b.报告id and a.标本ID=[1] order by b.序号"
18        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "镜检过程", lngSmapleID)
19        Do While Not rsTmp.EOF
20            intNO = Val(rsTmp("序号") & "")
              '检查本地是否有缓存文件，有缓则，则优先读取缓存文件
21            DoEvents        '加载一张显示一张
22            strImgPath = strFloder & "\" & mlngSampleID & "_" & intloop & ".bmp"
23            lblLoading(intloop).Caption = "loading..."
24            If Not objFSO.FileExists(strImgPath) Then
25                If mblnFtp And rsTmp("图像位置") & "" <> "" Then
                      '从FTP读取
26                    intReturn = mobjFTP.FuncDownloadFile(rsTmp("图像位置") & "", strImgPath, mlngSampleID & "_" & intloop & ".bmp")
27                    If intReturn = 1 Then
28                        MsgBox "FTP连接失败", vbInformation, gSysInfo.AppName
29                        Exit Sub
30                    ElseIf intReturn = 2 Then
31                        MsgBox "图像下载失败", vbInformation, gSysInfo.AppName
32                        Exit Sub
33                    End If
34                Else
                      '从数据库读取
35                    strImgData = gobjHisComLib.ReadLob(2500, 0, Val(rsTmp("id") & ""), strImgPath, 1, 0)
                      '解码图像
36                    If Replace(strImgData, " ", "") <> "" Then
37                        strImgPath = getBase64Img(strFloder, strImgPath, strImgData)
38                    End If
39                End If
40            End If
41            If objFSO.FileExists(strImgPath) Then
42                Me.imgPicture(intNO - 1).Picture = LoadPicture(strImgPath)
43                Me.imgPicture(intNO - 1).Tag = strImgPath
44            End If
45            lblLoading(intloop).Caption = ""
46            intloop = intloop + 1
47            rsTmp.MoveNext
48        Loop

49        Call DeleteImg  '删除过期图片
50        mlngVsfHeight = mlngVsfHeight + MicroorganismSmear.Height

51        Exit Sub
GetfrmMicroorganismSmear_Error:
52        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(GetfrmMicroorganismSmear)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
53        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/8/14
'功    能:获取初级药敏报告
'入    参:
'           lngSmapleID     标本ID
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub GetMicroorganisCJYM(ByVal lngSmapleID As Long)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
          '查询数据
1         On Error GoTo GetMicroorganisCJYM_Error

2         strSQL = " Select 细菌名, 检验结果, 描述, 耐药机制, 参考描述, Level" & vbCrLf & _
                   " From (Select 0 ID, b.Id 上级id, '抗生素名' 细菌名, '检验结果' 检验结果, '药敏方法' 耐药机制, '结果类型' 描述, '参考描述' 参考描述" & vbCrLf & _
                   "        From 微生物涂片报告 A, 微生物涂片细菌 B, 检验细菌记录 C" & vbCrLf & _
                   "        Where A.标本id = B.标本id And B.细菌id = C.ID And A.标本id = [1] And A.初步药敏审核人 Is Not Null" & vbCrLf & _
                   "        Union all" & vbCrLf & _
                   "        Select b.Id, Null 上级id, '鉴定结果：' || c.中文名 || '(' || c.英文名 || ')' 细菌名, b.检验结果, b.耐药机制, b.描述, '' 参考描述" & vbCrLf & _
                   "        From 微生物涂片报告 A, 微生物涂片细菌 B, 检验细菌记录 C" & vbCrLf & _
                   "        Where A.标本id = B.标本id And B.细菌id = C.ID And A.标本id = [1] And A.初步药敏审核人 Is Not Null" & vbCrLf & _
                   "        Union all" & vbCrLf & _
                   "        Select 0 ID, c.结果id 上级id, d.中文名 || '(' || d.英文名 || ')' 细菌名, nvl(c.结果,' ') 检验结果, c.药敏方法 耐药机制, c.结果类型 描述, c.参考描述" & vbCrLf & _
                   "        From 微生物涂片报告 A, 微生物涂片细菌 B, 微生物涂片药敏 C, 检验药敏 D" & vbCrLf & _
                   "        Where a.标本id = b.标本id And b.Id = c.结果id And c.药敏id = d.Id And a.标本id = [1] And a.初步药敏审核人 Is Not Null)" & vbCrLf & _
                   " Connect By Prior ID = 上级id" & vbCrLf & _
                   " Start With 上级id Is Null"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "初步药敏结果", lngSmapleID)
          
          '绑定数据
4         With Me.VSFCJYM
              '绑定数据
5             .Rows = 1
6             Do While Not rsTmp.EOF
7                 .Rows = .Rows + 1
8                 .TextMatrix(.Rows - 1, .ColIndex("细菌名")) = rsTmp("细菌名") & ""
9                 .TextMatrix(.Rows - 1, .ColIndex("检验结果")) = rsTmp("检验结果") & ""
10                .TextMatrix(.Rows - 1, .ColIndex("描述")) = rsTmp("描述") & ""
11                .TextMatrix(.Rows - 1, .ColIndex("耐药机制")) = rsTmp("耐药机制") & ""
12                .TextMatrix(.Rows - 1, .ColIndex("参考描述")) = rsTmp("参考描述") & ""
13                .TextMatrix(.Rows - 1, .ColIndex("Level")) = rsTmp("Level") & ""
                  
                  '如果level=1,则这行为父级行（细菌），level=2表示子级行（抗生素）
14                If Val(rsTmp("Level") & "") = 1 Then
                      '缩进
15                    .IsSubtotal(.Rows - 1) = True   '设置为树形节点
16                    .RowOutlineLevel(.Rows - 1) = 3
                      '显示边框线
17                    .CellBorderRange .Rows - 1, .ColIndex("细菌名"), .Rows - 1, .ColIndex("参考描述"), vbBlack, 0, 0, 0, 1, 0, 0
18                End If
                  
19                rsTmp.MoveNext
20            Loop
              
              '获取VSF高度
21            If mblnHaveBoder = False Then
22                If mlngVsfHeight < (.Rows + 5) * .RowHeight(0) Then
23                    mlngVsfHeight = (.Rows + 5) * .RowHeight(0)
24                End If
25            End If
26        End With


27        Exit Sub
GetMicroorganisCJYM_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(GetMicroorganisCJYM)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
          
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mintVersion = 0
    mlngSampleID = 0
    mlngPaintID = 0
    mlngAdvicID = 0
    mstrFtpIp = ""
    mstrFtpUser = ""
    mstrFtpPwd = ""
    mstrFtpFolder = ""
    mblnFtp = False
    Call mobjFTP.FuncFtpDisConnect   '断开FTP连接
    Set mobjFTP = Nothing
End Sub

Private Sub imgPicture_Click(Index As Integer)
    If imgPicture(Index).Tag <> "" Then
        Call frmAdviceReprotBrowseShowPic.ShowMe(Me, imgPicture(Index).Tag)
    End If
End Sub

Private Sub picCenter_Resize()
    On Error Resume Next
    With Me.PicNegative
        .Left = 0
        .Top = 0
        .Height = Me.picCenter.Height
        .Width = Me.picCenter.Width
        .BorderStyle = 0
    End With
    
    With Me.picMicrobePositive
        .Left = 0
        .Top = 0
        .Width = Me.picCenter.Width
        .Height = Me.picCenter.Height
        .BorderStyle = 0
    End With
End Sub

Private Sub picCJYM_Resize()
     With Me.VSFCJYM
        .Left = 0
        .Top = 0
        .Width = Me.picCJYM.Width
        .Height = Me.picCJYM.Height
     End With
End Sub

Private Sub picComment_Resize()
    On Error Resume Next
    With Me.txtResult
        .Width = Me.picComment.Width - .Left
        .BackColor = Me.picComment.BackColor
        .BorderStyle = 0
    End With
    
    With Me.txtDiagnosis
        .Width = Me.txtResult.Width
        .BackColor = Me.picComment.BackColor
        .BorderStyle = 0
    End With
End Sub

Private Sub picGeneral_Resize()
    On Error Resume Next
    With Me.vsfGeneral
        .Left = 0
        .Top = 0
        .Width = Me.picGeneral.Width
        .Height = Me.picGeneral.Height
    End With
End Sub

Private Sub picImg_Resize(Index As Integer)
    On Error Resume Next
    With imgPicture(Index)
        .Left = 0
        .Top = 0
        .Width = picImg(Index).Width
        .Height = picImg(Index).Height
    End With
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next
    With Me.picPatient
        .Left = 300
        .Top = 0
        .Width = Me.picMain.Width - 600
        .BorderStyle = 0
    End With
    
    With Me.picComment
        .Left = picPatient.Left
        .Top = Me.picMain.Height - .Height
        .Width = picPatient.Width
        .BorderStyle = 0
    End With
    
    With Me.picGeneral
        .Left = Me.picPatient.Left
        .Top = Me.picPatient.Top + Me.picPatient.Height
        .Width = Me.picPatient.Width
        .Height = Me.picMain.Height - .Top - Me.picComment.Height - 500
        .BorderStyle = 0
    End With
    
'    With Me.picCenter
'        .Left = Me.picPatient.Left
'        .Top = Me.picPatient.Top + Me.picPatient.Height
'        .Width = Me.picPatient.Width
'        .Height = Me.picMain.Height - .Top - Me.picComment.Height - 500
'        .BorderStyle = 0
'    End With
    
    With Me.picTab
        .Left = 10
        .Top = Me.picPatient.Top + Me.picPatient.Height - 30
        .Width = Me.picPatient.Width + 250
        .Height = Me.picMain.Height - .Top - Me.picComment.Height - 500
        .BorderStyle = 0
    End With
    
    
End Sub

Private Sub picMicrobePositive_Resize()
    On Error Resume Next
    With Me.vsfMicrobePositive
        .Left = 0
        .Top = 0
        .Width = Me.picMicrobePositive.Width
        .Height = Me.picMicrobePositive.Height
    End With
End Sub

Private Sub PicNegative_Resize()
    On Error Resume Next
    With frmNom
        .Top = 20
        .Left = 60
        .Width = PicNegative.ScaleWidth - 60
    End With
    txtNormalMicrobe.Width = frmNom.Width - Label21.Width - 300
    txtNoFindMicrobe.Width = txtNormalMicrobe.Width
    txtNormalMicrobes.Width = txtNormalMicrobe.Width
    With frmChe
        .Top = frmNom.Top + frmNom.Height + 20
        .Left = 60
        .Width = PicNegative.ScaleWidth - 60
    End With

    If chkMicroscope.value = 1 Then
        fraOne.Visible = True
        With fraOne
            .Top = frmChe.Top + frmChe.Height + 20
            .Left = 60
            .Width = PicNegative.ScaleWidth - 60
'            .Height = PicNegative.ScaleHeight - frmNom.Height - frmChe.Height - 300
        End With
        txtMicroscope.Width = fraOne.Width - Label21.Width - 500
        txtMicroscopeFinded.Width = txtMicroscope.Width
        txtMicroscopeNOFind.Width = txtMicroscope.Width

    Else
        fraOne.Visible = False
'        frmChe.Height = PicNegative.ScaleHeight - frmNom.Height - 300
    End If
End Sub


Private Sub picPatient_Resize()
    If Not mblnHaveBoder Then
        If lblPatient(0).Caption = "姓名:" Then lblPatient(0).Caption = "申请项目:"
        lblPatient(1).Visible = False
        lblPatient(2).Visible = False
        lblPatient(3).Visible = False
    End If
End Sub

Private Sub picPDF_Resize()
    On Error Resume Next
    With webSub
        .Left = 0
        .Top = 0
        .Width = picPDF.Width
        .Height = picPDF.Height
    End With
End Sub

Private Sub picTab_Resize()
    With Me.tabThis
        .Left = 0
        .Top = 0
        .Width = Me.picTab.Width + 50
        .Height = Me.picTab.Height + 50
    End With
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-19
'功    能:  显示诊疗参考
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim objAdvice As Object
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemCode As String
          Dim strItemIDs As String

1         On Error GoTo ShowClincHelp_Error

2         If mlngSampleID <> 0 Then
3             If mintVersion = 25 Then    '新版中去查询
4                 strSQL = "Select f_List2str(Cast(Collect(b.诊疗编码 || '') As t_Strlist)) 编码" & vbCrLf & _
                         "   From 检验申请组合 A, 检验组合项目 B" & vbCrLf & _
                         "   Where A.组合ID = b.id And a.标本id = [1]"
5                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", mlngSampleID)
6                 If Not rsTmp.EOF Then
7                     strItemCode = rsTmp("编码") & ""
8                 End If

9                 If strItemCode <> "" Then
                      '通过诊疗编码查询诊疗项目ID
10                    strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                               "f_List2str(Cast(Collect(a.ID || '') As t_Strlist)) ID" & vbCrLf & _
                             " From 诊疗项目目录 A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                             " Where A.编码 = B.Column_Value"
11                    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", strItemCode)
12                    If Not rsTmp.EOF Then strItemIDs = rsTmp("ID") & ""
13                End If
14            ElseIf mintVersion = 10 Then    '老版中去查询
15                strSQL = " select f_List2str(Cast(Collect(b.诊疗项目ID || '') As t_Strlist)) 诊疗项目ID from 检验标本记录 A, 病人医嘱记录 B where a.医嘱id=b.id and a.id=[1]"
16                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目ID", mlngSampleID)
17                If Not rsTmp.EOF Then
18                    strItemIDs = rsTmp("诊疗项目ID") & ""
19                End If
20            End If
21        End If



          '调用接口
22        If Not rsTmp.EOF Then
23            If objAdvice Is Nothing Then
24                Set objAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
25                If Not objAdvice Is Nothing Then
26                    On Error Resume Next
27                    Call objAdvice.ShowClincHelp(1, Me, 0, False, strItemIDs)
28                    If Err.Number = 438 Then
29                        MsgBox "HIS版本过低", vbInformation, gSysInfo.AppName
30                        Exit Sub
31                    End If
32                End If
33            End If
34        End If


35        Exit Sub
ShowClincHelp_Error:
36        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(ShowClincHelp)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
37        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-12
'功    能:  检查FTP服务器是否可用
'入    参:
'出    参:
'返    回:  True=可用，False=不可用
'调整影响:
'---------------------------------------------------------------------------------------
Private Function ConnFtp() As Boolean
          Dim strFTP As String
          Dim intReturn As Integer

1         On Error GoTo ConnFtp_Error

2         strFTP = ComGetPara(Sel_Lis_DB, "FTP设置", 2500, 2500, "")
3         If UBound(Split(strFTP, ";")) >= 3 Then
4             mstrFtpUser = Split(strFTP, ";")(0)
5             mstrFtpPwd = Split(strFTP, ";")(1)
6             If mstrFtpPwd Like "ZLSV*:*" Then
7                 mstrFtpPwd = gobjHisComLib.zlStr.Sm4DecryptEcb(Split(strFTP, ";")(1))
8             Else
9                 mstrFtpPwd = Split(strFTP, ";")(1)
10            End If
11            mstrFtpIp = Split(strFTP, ";")(2)
12            mstrFtpFolder = Split(strFTP, ";")(3)
13            If mobjFTP.FuncFtpConnect(mstrFtpIp, mstrFtpUser, mstrFtpPwd) > 0 Then
14                mblnFtp = True
                  
                  '创建图片保存目录
15                intReturn = mobjFTP.FuncFtpMkDir(mstrFtpFolder, "MicroorganismPicture")
16                If intReturn = 1 Then
17                    MsgBox "FTP连接失败", vbInformation, gSysInfo.AppName
18                    Exit Function
19                End If
20            End If
21        End If



22        Exit Function
ConnFtp_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(ConnFtp)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
24        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/7/24
'功    能:删除创建时间超过30天的图片文件
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub DeleteImg()
          Dim objFSO As New FileSystemObject
          Dim strFolder As String
          Dim dateNow As Date
          Dim objFolder As Folder
          Dim objFiles As Files
          Dim objFile As File

1         On Error GoTo DeleteImg_Error

2         strFolder = App.Path & "\MicroorganismPicture"
3         dateNow = gobjHisDatabase.Currentdate
          '判断文件夹是否存在
4         If Not objFSO.FolderExists(strFolder) Then Exit Sub
          '便利文件夹下的所有文件，如果创建时间大于等于30天，则删除该文件
5         Set objFolder = objFSO.GetFolder(strFolder)
6         Set objFiles = objFolder.Files
7         For Each objFile In objFiles
8             If DateDiff("d", objFile.DateCreated, dateNow) >= 30 Then
9                 objFSO.DeleteFile (objFile.Path)
10            End If
11        Next


12        Exit Sub
DeleteImg_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(DeleteImg)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear

End Sub

Private Sub vsfGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfGeneral
        lngRow = .MouseRow
        lngCol = .MouseCol
        If .ColIndex("临床意义") < 0 Then Exit Sub
        If lngRow < 1 Or lngCol < 0 Then
            Call gobjHisComLib.zlCommFun.ShowTipInfo(0, "")
            Exit Sub
        End If
        Call gobjHisComLib.zlCommFun.ShowTipInfo(.hWnd, .TextMatrix(lngRow, .ColIndex("临床意义")), True)
    End With
End Sub

Private Function AuditingSample(ByVal intType As Integer) As Boolean
      '复核/取消复核
      'intType    1=复核,2=取消复核

          Dim strSQL As String

1         On Error GoTo AuditingSample_Error

2         strSQL = "Zl_检验传染病复核_Edit(" & intType & "," & mlngSampleID & ",'" & gUserInfo.Name & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "传染病报告复核")

4         SaveDBLog 18, 6, Val(mlngPaintID), IIf(intType = 1, "复核", "取消复核"), IIf(intType = 1, "复核", "取消复核"), 2500, "临床实验室管理"

5         AuditingSample = True
          
6         If intType = 1 Then
7             cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = False
8             cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = True
9         Else
10            cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = True
11            cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = False
12        End If
          
13        Exit Function
AuditingSample_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "执行(AuditingSample)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear

End Function




